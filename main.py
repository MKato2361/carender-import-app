import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
# 変更: excel_parser から必要な関数を個別にインポート
from excel_parser import (
    process_excel_data_for_calendar, # 新しいメイン処理関数
    _load_and_merge_dataframes,      # ファイルロード＆マージのヘルパー関数
    get_available_columns_for_event_name, # イベント名選択用列取得
    check_event_name_columns         # イベント名列の有無チェック
)
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service, # tasks_serviceを返す関数を直接インポート
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from googleapiclient.discovery import build
from googleapiciapient.errors import HttpError

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")


# Firebaseの初期化
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

# Firebase認証フォームの表示とユーザーIDの取得
user_id = get_firebase_user_id()

if not user_id:
    # ユーザーがログインしていない場合、認証フォームを表示して停止
    firebase_auth_form()
    st.stop()

# --- ここから下の処理は、Firebase認証が完了した場合にのみ実行されます ---

google_auth_placeholder = st.empty()

with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    # Firebase認証後に、Google認証に進む
    creds = authenticate_google()

    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()
        st.sidebar.success("✅ Googleカレンダーに認証済みです！")


# カレンダーサービスの初期化
def initialize_calendar_service():
    """カレンダーサービスを初期化する"""
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        
        editable_calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') != 'reader'
        }
        
        return service, editable_calendar_options
    except HttpError as e:
        st.error(f"カレンダーサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.error(f"カレンダーサービスの初期化に失敗しました: {e}")
        return None, None

# タスクサービスの初期化
def initialize_tasks_service_wrapper(): # 名前を衝突しないように変更
    """タスクサービスを初期化する"""
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
            
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        
        # デフォルトのタスクリストを探す
        for task_list in task_lists.get('items', []):
            # 'My Tasks' は英語環境でのデフォルト名。日本語環境では 'My Tasks' とは限らない場合があるため注意
            # 'My Tasks' に相当するものを厳密に判断するには、Google Tasks APIのドキュメントを確認するか、
            # ユーザーに選択させるUIを用意するのがより堅牢
            if task_list.get('title') == 'My Tasks': # これはGoogle Tasksのデフォルトリスト名
                default_task_list_id = task_list['id']
                break
                
        # 'My Tasks'が見つからない場合、最初のタスクリストを使用
        if not default_task_list_id and task_lists.get('items'):
            default_task_list_id = task_lists['items'][0]['id']
        
        return tasks_service, default_task_list_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
        return None, None

# カレンダーサービスの初期化または取得
if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    service, editable_calendar_options = initialize_calendar_service()
    
    if not service:
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop()
    
    st.session_state['calendar_service'] = service
    st.session_state['editable_calendar_options'] = editable_calendar_options
else:
    service = st.session_state['calendar_service']
    # 既にサービスが存在する場合でも、カレンダーオプションを再度読み込むことで最新の状態を反映
    # initialize_calendar_serviceが毎回呼び出されるべきか、セッション間でキャッシュすべきかはユースケースによる
    # 現状は毎回認証するが、もしパフォーマンスが問題ならキャッシュ検討
    # ここでは、initialize_calendar_service()はすでにcredsに依存しているため、再度実行しても問題ない
    _, st.session_state['editable_calendar_options'] = initialize_calendar_service()


# タスクサービスの初期化または取得
if 'tasks_service' not in st.session_state or not st.session_state.get('tasks_service'): # Noneチェックを追加
    tasks_service, default_task_list_id = initialize_tasks_service_wrapper() # ラッパー関数を呼び出す
    
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
    
    if not tasks_service:
        st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
else:
    tasks_service = st.session_state['tasks_service']

# メインタブの作成
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

# セッション状態の初期化
if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame() # 新しくマージ済みDFを保持

with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        
        try:
            # 選択肢表示のために、アップロードされたファイルを統合
            # _load_and_merge_dataframes はエラーをスローする可能性があるのでtry-except
            st.session_state['merged_df_for_selector'] = _load_and_merge_dataframes(uploaded_files)
            
            # 説明文の列プールの更新
            st.session_state['description_columns_pool'] = st.session_state['merged_df_for_selector'].columns.tolist()

            if st.session_state['merged_df_for_selector'].empty:
                st.warning("アップロードされたファイルに有効なデータがありませんでした。")

        except (ValueError, IOError) as e:
            st.error(f"ファイルの読み込みまたは結合に失敗しました: {e}")
            st.session_state['uploaded_files'] = []
            st.session_state['description_columns_pool'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            
    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みのファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")
        if not st.session_state['merged_df_for_selector'].empty:
             st.info(f"読み込まれたデータには {len(st.session_state['merged_df_for_selector'].columns)} 列 {len(st.session_state['merged_df_for_selector'])} 行のデータがあります。")

        # ファイル削除機能の追加
        if st.button("🗑️ アップロード済みファイルをクリア", help="アップロードされたExcelファイルの情報をアプリケーションから削除します。"):
            st.session_state['uploaded_files'] = []
            st.session_state['description_columns_pool'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.success("アップロードされたExcelファイルがクリアされました。")
            st.rerun() # 変更を反映するために再実行

with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False) # 名称変更を反映
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        # 説明文に含める列の選択
        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', []),
            default=[col for col in ["内容", "詳細"] if col in st.session_state.get('description_columns_pool', [])]
        )
        
        # イベント名の代替列選択UIをここに配置
        fallback_event_name_column = None
        has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
        
        if not (has_mng_data and has_name_data):
            st.subheader("イベント名の設定")
            if not has_mng_data and not has_name_data:
                st.info("ファイルに「管理番号」と「物件名」のデータが見つかりませんでした。イベント名に使用する列を選択してください。")
            elif not has_mng_data:
                st.info("ファイルに「管理番号」のデータが見つかりませんでした。物件名と合わせてイベント名に使用する列を選択できます。")
            elif not has_name_data:
                st.info("ファイルに「物件名」のデータが見つかりませんでした。管理番号と合わせてイベント名に使用する列を選択できます。")

            available_event_name_cols = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options = ["選択しない"] + available_event_name_cols
            
            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=0, # デフォルトは「選択しない」
                key="event_name_selector_register" # Keyをユニークにする
            )
            if selected_event_name_col != "選択しない":
                fallback_event_name_column = selected_event_name_col
        else:
            st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")


        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("✅ ToDoリスト連携設定 (オプション)")
            create_todo = st.checkbox("このイベントに対応するToDoリストを作成する", value=False, key="create_todo_checkbox")

            # ToDoの選択肢を「点検通知」のみに固定
            fixed_todo_types = ["点検通知"] # 今後増える可能性を考慮しリスト形式で維持
            
            if create_todo:
                st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
            else:
                st.markdown(f"ToDoリストの作成は無効です。")


            deadline_offset_options = {
                "2週間前": 14,
                "10日前": 10,
                "1週間前": 7,
                "カスタム日数前": None
            }
            selected_offset_key = st.selectbox(
                "ToDoリストの期限をイベント開始日の何日前に設定しますか？",
                list(deadline_offset_options.keys()),
                disabled=not create_todo,
                key="deadline_offset_select"
            )

            custom_offset_days = None
            if selected_offset_key == "カスタム日数前":
                custom_offset_days = st.number_input(
                    "何日前に設定しますか？ (日数)",
                    min_value=0,
                    value=3,
                    disabled=not create_todo,
                    key="custom_offset_input"
                )


            st.subheader("➡️ イベント登録")
            if st.button("Googleカレンダーに登録する"):
                with st.spinner("イベントデータを処理中..."):
                    # 変更: process_excel_data_for_calendar を呼び出す
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns, 
                            all_day_event_override, # 名称変更
                            private_event, 
                            fallback_event_name_column # 新しい引数
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame() # エラー時は空のDFにする

                    if df.empty:
                        st.warning("有効なイベントデータがありません。処理を中断しました。")
                    else:
                        st.info(f"{len(df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        successful_todo_creations = 0

                        for i, row in df.iterrows():
                            event_start_date_obj = None
                            event_end_date_obj = None
                            event_time_str = "" # ToDo詳細用の時間文字列
                            event_id = None # ToDoに紐付けるイベントID

                            try:
                                if row['All Day Event'] == "True":
                                    event_start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                                    event_end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                                    
                                    start_date_str = event_start_date_obj.strftime("%Y-%m-%d")
                                    # Outlook CSV形式に合わせたend_dateなので、Google Calendar APIは+1日が必要
                                    end_date_for_api = (event_end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d") 

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_for_api},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                    # ToDo詳細表示用: イベント日時
                                    event_time_str = f"{event_start_date_obj.strftime('%Y/%-m/%-d')}"
                                    if event_start_date_obj != event_end_date_obj:
                                        event_time_str += f"～{event_end_date_obj.strftime('%Y/%-m/%-d')}"

                                else:
                                    event_start_datetime_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                                    event_end_datetime_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                                    
                                    event_start_date_obj = event_start_datetime_obj.date()
                                    event_end_date_obj = event_end_datetime_obj.date()

                                    start_iso = event_start_datetime_obj.isoformat()
                                    end_iso = event_end_datetime_obj.isoformat()

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'dateTime': start_iso, 'timeZone': 'Asia/Tokyo'},
                                        'end': {'dateTime': end_iso, 'timeZone': 'Asia/Tokyo'},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                    # ToDo詳細表示用: イベント日時
                                    event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%H:%M')}"
                                    if event_start_datetime_obj.date() != event_end_datetime_obj.date():
                                        event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}"

                                # イベント登録し、イベントIDを取得
                                created_event = add_event_to_calendar(service, calendar_id, event_data)
                                if created_event:
                                    successful_registrations += 1
                                    event_id = created_event.get('id')

                                # ToDoリストの作成ロジック
                                if create_todo and tasks_service and st.session_state.get('default_task_list_id') and event_id: # event_idがある場合のみToDo作成
                                    if event_start_date_obj: # ToDo期限計算の基準となる日付
                                        offset_days = deadline_offset_options.get(selected_offset_key)
                                        if selected_offset_key == "カスタム日数前" and custom_offset_days is not None:
                                            offset_days = custom_offset_days

                                        if offset_days is not None:
                                            todo_due_date = event_start_date_obj - timedelta(days=offset_days)
                                            
                                            # 全ての固定ToDoタイプを追加
                                            for todo_item in fixed_todo_types:
                                                todo_summary = f"{todo_item} - {row['Subject']}"
                                                # ToDo詳細にイベントIDを含める
                                                todo_notes = (
                                                    f"関連イベントID: {event_id}\n" # イベントIDを保存
                                                    f"イベント日時: {event_time_str}\n"
                                                    f"場所: {row['Location']}"
                                                )

                                                add_task_to_todo_list(
                                                    tasks_service,
                                                    st.session_state['default_task_list_id'],
                                                    todo_summary,
                                                    todo_due_date,
                                                    notes=todo_notes
                                                )
                                                successful_todo_creations += 1
                                        else:
                                            st.warning(f"ToDoの期限が設定されませんでした。カスタム日数が無効です。")
                                    else:
                                        st.warning(f"ToDoの期限を設定できませんでした。イベント開始日が不明です。")

                            except Exception as e:
                                st.error(f"{row['Subject']} の登録またはToDoリスト作成に失敗しました: {e}")
                            progress.progress((i + 1) / len(df))

                        st.success(f"✅ {successful_registrations} 件のイベント登録が完了しました！")
                        if create_todo:
                            st.success(f"✅ {successful_todo_creations} 件のToDoリストが作成されました！")

with tabs[2]:
    st.header("イベントを削除")

    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today_date = date.today() # datetime.today()ではなくdate.today()を使用
        delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
        delete_end_date = st.date_input("削除終了日", value=today_date)
        
        # ToDoリストも削除するかどうかのチェックボックス
        delete_related_todos = st.checkbox("関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)", value=False)


        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")
            if st.button("選択期間のイベントを削除する"):
                calendar_service = st.session_state['calendar_service']
                tasks_service = st.session_state['tasks_service']
                default_task_list_id = st.session_state.get('default_task_list_id')

                # まず期間内のイベントを取得
                # JST_OFFSET = timedelta(hours=9) # calendar_utilsで処理しているので不要
                # time_minとtime_maxはUTCでなければならない
                start_dt_utc = datetime.combine(delete_start_date, datetime.min.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(datetime.timezone.utc)
                end_dt_utc = datetime.combine(delete_end_date, datetime.max.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(datetime.timezone.utc)
                
                time_min_utc = start_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                time_max_utc = end_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')


                events_to_delete = fetch_all_events(calendar_service, calendar_id_del, time_min_utc, time_max_utc)
                
                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")
                    # st.stop() # ここでstopすると、その後のログアウトボタンなどが動作しない
                    # continue # forループではないのでこれも不要。単純にリターンまたはメッセージ表示でOK

                deleted_events_count = 0
                deleted_todos_count = 0 # ToDoの削除数をカウント
                total_events = len(events_to_delete)
                
                if total_events > 0:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for i, event in enumerate(events_to_delete):
                        event_summary = event.get('summary', '不明なイベント')
                        event_id = event['id']
                        
                        status_text.text(f"イベント '{event_summary}' を削除中... ({i+1}/{total_events})")

                        try:
                            # 関連ToDoの削除
                            if delete_related_todos and tasks_service and default_task_list_id:
                                deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                    tasks_service,
                                    default_task_list_id,
                                    event_id
                                )
                                deleted_todos_count += deleted_task_count_for_event
                            
                            # イベント自体の削除
                            calendar_service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                            deleted_events_count += 1
                        except Exception as e:
                            st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")
                        
                        progress_bar.progress((i + 1) / total_events)
                    
                    status_text.empty() # 処理完了後にステータステキストをクリア

                    if deleted_events_count > 0:
                        st.success(f"✅ {deleted_events_count} 件のイベントが削除されました。")
                        if delete_related_todos:
                            if deleted_todos_count > 0:
                                st.success(f"✅ {deleted_todos_count} 件の関連ToDoタスクが削除されました。")
                            else:
                                st.info("関連するToDoタスクは見つからなかったか、すでに削除されていました。")
                    else:
                        st.info("指定期間内に削除するイベントはありませんでした。")
                else:
                    st.info("指定期間内に削除するイベントはありませんでした。")


with tabs[3]:
    st.header("イベントを更新")

    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        # 更新タブでの設定も、登録タブと同様に名称と挙動を統一
        all_day_event_override_update = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event_update = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")
        
        description_columns_update = st.multiselect(
            "説明欄に含める列", 
            st.session_state['description_columns_pool'], 
            default=[col for col in ["内容", "詳細"] if col in st.session_state.get('description_columns_pool', [])],
            key="update_desc_cols"
        )

        # イベント名の代替列選択UIをここに配置 (更新タブ用)
        fallback_event_name_column_update = None
        has_mng_data_update, has_name_data_update = check_event_name_columns(st.session_state['merged_df_for_selector'])
        
        if not (has_mng_data_update and has_name_data_update):
            st.subheader("更新時のイベント名の設定")
            st.info("Excelデータからのイベント名生成に、以下の列を代替として使用できます。")

            available_event_name_cols_update = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options_update = ["選択しない"] + available_event_name_cols_update
            
            selected_event_name_col_update = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options_update,
                index=0,
                key="event_name_selector_update"
            )
            if selected_event_name_col_update != "選択しない":
                fallback_event_name_column_update = selected_event_name_col_update
        else:
            st.info("「管理番号」と「物件名」のデータが存在するため、それらがイベント名として使用されます。")


        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name_upd = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
            calendar_id_upd = st.session_state['editable_calendar_options'][selected_calendar_name_upd]

            if st.button("イベントを照合・更新"):
                with st.spinner("イベントを処理中..."):
                    try:
                        # 変更: process_excel_data_for_calendar を呼び出す
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns_update, # 更新タブ用の列
                            all_day_event_override_update, # 更新タブ用の設定
                            private_event_update,         # 更新タブ用の設定
                            fallback_event_name_column_update # 新しい引数
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame() # エラー時は空のDFにする

                    if df.empty:
                        st.warning("有効なイベントデータがありません。更新を中断しました。")
                        st.stop()

                    # 検索期間を広げる。作業指示書での紐付けなので、ある程度の期間をカバーする必要がある
                    today_for_update = datetime.now()
                    # 現在から過去2年、未来2年の範囲で検索
                    time_min = (today_for_update - timedelta(days=365*2)).isoformat() + 'Z'
                    time_max = (today_for_update + timedelta(days=365*2)).isoformat() + 'Z'
                    events = fetch_all_events(service, calendar_id_upd, time_min, time_max)

                    worksheet_to_event = {}
                    for event in events:
                        desc = event.get('description', '')
                        # 作業指示書は数値型で抽出される場合があるので、\d+ に変更し、厳密に数値部分を捉える
                        match = re.search(r"作業指示書[：:]\s*(\d+)", desc) # 半角・全角コロン、スペースに対応
                        if match:
                            worksheet_id = match.group(1)
                            # 同じ作業指示書IDのイベントが複数ある場合、古いものを上書きしないようにリスト化するか、
                            # 最新のものだけを保持するかなどのロジックを検討する必要があるが、
                            # 今回は単純に最新（fetch_all_eventsで取得順序に依存）を保持。
                            worksheet_to_event[worksheet_id] = event

                    update_count = 0
                    progress_bar = st.progress(0)
                    for i, row in df.iterrows():
                        # process_excel_data_for_calendar で生成された 'Description' 列から作業指示書IDを抽出
                        # ここもformat_worksheet_valueで整形された文字列を想定
                        match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
                        if not match:
                            progress_bar.progress((i + 1) / len(df)) # 進捗バーを更新
                            continue # 作業指示書IDが見つからない行はスキップ
                        
                        worksheet_id = match.group(1)
                        matched_event = worksheet_to_event.get(worksheet_id)
                        if not matched_event:
                            progress_bar.progress((i + 1) / len(df)) # 進捗バーを更新
                            continue # マッチする既存イベントがない場合はスキップ

                        # カレンダーイベントのデータ構造を構築
                        event_data = {
                            'summary': row['Subject'],
                            'location': row['Location'],
                            'description': row['Description'],
                            'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                        }
                        
                        # 日時の設定
                        if row['All Day Event'] == "True":
                            start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                            end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                            
                            start_date_str = start_date_obj.strftime("%Y-%m-%d")
                            # Google Calendar APIの終日イベントの終了日は排他的なため、Outlook CSV形式の終了日+1が必要
                            end_date_for_api = (end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d")
                            
                            event_data['start'] = {'date': start_date_str}
                            event_data['end'] = {'date': end_date_for_api}
                        else:
                            start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                            end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                            
                            event_data['start'] = {'dateTime': start_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
                            event_data['end'] = {'dateTime': end_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

                        try:
                            # update_event_if_neededは既存のeventオブジェクトと更新データを受け取る
                            if update_event_if_needed(service, calendar_id_upd, matched_event, event_data):
                                update_count += 1
                        except Exception as e:
                            st.error(f"イベント '{row['Subject']}' (作業指示書: {worksheet_id}) の更新に失敗しました: {e}")
                        
                        progress_bar.progress((i + 1) / len(df))

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")


# サイドバーに認証情報表示
with st.sidebar:
    st.header("🔐 認証状態")
    st.success("✅ Firebase認証済み")
    
    if st.session_state.get('calendar_service'):
        st.success("✅ Googleカレンダー認証済み")
    else:
        st.warning("⚠️ Googleカレンダー認証が未完了です")
    
    if st.session_state.get('tasks_service'):
        st.success("✅ ToDoリスト利用可能")
    else:
        st.warning("⚠️ ToDoリスト利用不可")
    
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    st.metric("アップロード済みファイル", uploaded_count)
    
    # ログアウトボタン
    if st.button("🚪 ログアウト", type="secondary"):
        # セッション状態をクリア
        for key in list(st.session_state.keys()):
            # 全てのセッション情報をクリアして完全にログアウトする
            # 'user_id' や 'creds' など、ログインセッションを維持したいものは残すことも検討
            del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
