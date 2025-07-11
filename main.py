import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
from excel_parser import process_excel_files
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from googleapiclient.discovery import build # build関数をインポート
from googleapiclient.errors import HttpError

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
def initialize_tasks_service():
    """タスクサービスを初期化する"""
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
            
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        
        # デフォルトのタスクリストを探す
        for task_list in task_lists.get('items', []):
            if task_list.get('title') == 'My Tasks':
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

# タスクサービスの初期化または取得
if 'tasks_service' not in st.session_state:
    tasks_service, default_task_list_id = initialize_tasks_service()
    
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

with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        description_columns_pool = set()
        for file in uploaded_files:
            try:
                df_temp = pd.read_excel(file, engine="openpyxl")
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                description_columns_pool.update(df_temp.columns)
            except Exception as e:
                st.warning(f"{file.name} の読み込みに失敗しました: {e}")
        st.session_state['description_columns_pool'] = list(description_columns_pool)
    elif 'uploaded_files' not in st.session_state:
        st.session_state['uploaded_files'] = []
        st.session_state['description_columns_pool'] = []

    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みのファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")

with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', [])
        )

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("✅ ToDoリスト連携設定 (オプション)")
            create_todo = st.checkbox("このイベントに対応するToDoリストを作成する", value=False, key="create_todo_checkbox")

            # ToDoの選択肢を「点検通知」のみに固定
            fixed_todo_types = ["点検通知"]
            
            st.markdown(f"以下のToDoが**常にすべて**作成されます: {', '.join(fixed_todo_types)}")


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
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                    else:
                        st.info(f"{len(df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        successful_todo_creations = 0

                        for i, row in df.iterrows():
                            event_start_date_obj = None
                            event_end_date_obj = None
                            event_time_str = "" # ToDo詳細用の時間文字列
                            event_id = None # 追加: ToDoに紐付けるイベントID

                            try:
                                if row['All Day Event'] == "True":
                                    event_start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                                    event_end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                                    
                                    start_date_str = event_start_date_obj.strftime("%Y-%m-%d")
                                    end_date_str = (event_end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d") # 終日イベントは終了日+1

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_str},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                    event_time_str = f"{event_start_date_obj.strftime('%-m/%-d')}" # 例: 6/30
                                    if event_start_date_obj != event_end_date_obj:
                                        event_time_str += f"～{event_end_date_obj.strftime('%-m/%-d')}"

                                else:
                                    event_start_datetime_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                                    event_end_datetime_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                                    
                                    event_start_date_obj = event_start_datetime_obj.date()
                                    event_end_date_obj = event_end_datetime_obj.date()

                                    start = event_start_datetime_obj.isoformat()
                                    end = event_end_datetime_obj.isoformat()

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'dateTime': start, 'timeZone': 'Asia/Tokyo'},
                                        'end': {'dateTime': end, 'timeZone': 'Asia/Tokyo'},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                    # 例: 6/30 9:00～10:00
                                    event_time_str = f"{event_start_datetime_obj.strftime('%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%H:%M')}"
                                    if event_start_datetime_obj.date() != event_end_datetime_obj.date():
                                        event_time_str = f"{event_start_datetime_obj.strftime('%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%-m/%-d %H:%M')}"

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
        st.error("削除可能なカレンダーが見つかりませんでした。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today = date.today()
        delete_start_date = st.date_input("削除開始日", value=today - timedelta(days=30))
        delete_end_date = st.date_input("削除終了日", value=today)
        
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
                JST_OFFSET = timedelta(hours=9)
                start_dt_jst = datetime.combine(delete_start_date, datetime.min.time())
                end_dt_jst = datetime.combine(delete_end_date, datetime.max.time())
                time_min_utc = (start_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'
                time_max_utc = (end_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'

                events_to_delete = fetch_all_events(calendar_service, calendar_id_del, time_min_utc, time_max_utc)
                
                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")
                    st.stop()

                deleted_events_count = 0
                deleted_todos_count = 0 # ToDoの削除数をカウント
                total_events = len(events_to_delete)
                
                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, event in enumerate(events_to_delete):
                    event_summary = event.get('summary', '不明なイベント')
                    event_id = event['id']
                    
                    status_text.text(f"イベント '{event_summary}' を削除中... ({i+1}/{total_events})")

                    try:
                        # 関連ToDoの削除
                        if delete_related_todos and tasks_service and default_task_list_id:
                            # ToDo詳細に保存されたイベントIDに基づいてタスクを検索し削除
                            deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                tasks_service,
                                default_task_list_id,
                                event_id
                            )
                            deleted_todos_count += deleted_task_count_for_event
                            # 各イベントごとのToDo削除メッセージは表示せず、最終的な合計のみ表示するように変更
                            # if deleted_task_count_for_event > 0:
                            #     st.info(f"イベント '{event_summary}' に関連するToDoタスクを {deleted_task_count_for_event} 件削除しました。")
                        
                        # イベント自体の削除
                        calendar_service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                        deleted_events_count += 1
                    except Exception as e:
                        st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")
                    
                    progress_bar.progress((i + 1) / total_events)
                
                status_text.empty() # 処理完了後にステータステキストをクリア

                if deleted_events_count > 0:
                    st.success(f"✅ {deleted_events_count} 件のイベントが削除されました。")
                    if delete_related_todos: # チェックボックスがオンの場合のみToDo削除結果を表示
                        if deleted_todos_count > 0:
                            st.success(f"✅ {deleted_todos_count} 件の関連ToDoタスクが削除されました。")
                        else:
                            st.info("関連するToDoタスクは見つからなかったか、すでに削除されていました。")
                else:
                    st.info("指定期間内に削除するイベントはありませんでした。")


with tabs[3]:
    st.header("イベントを更新")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        all_day_event = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")
        description_columns = st.multiselect("説明欄に含める列", st.session_state['description_columns_pool'], key="update_desc_cols")

        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name_upd = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
            calendar_id_upd = st.session_state['editable_calendar_options'][selected_calendar_name_upd]

            if st.button("イベントを照合・更新"):
                with st.spinner("イベントを処理中..."):
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                        st.stop()

                    today = datetime.now()
                    time_min = (today - timedelta(days=180)).isoformat() + 'Z'
                    time_max = (today + timedelta(days=180)).isoformat() + 'Z'
                    events = fetch_all_events(service, calendar_id_upd, time_min, time_max)

                    worksheet_to_event = {}
                    for event in events:
                        desc = event.get('description', '')
                        match = re.search(r"作業指示書：(\d+)", desc)
                        if match:
                            worksheet_to_event[match.group(1)] = event

                    update_count = 0
                    for i, row in df.iterrows():
                        match = re.search(r"作業指示書：(\d+)", row['Description'])
                        if not match:
                            continue
                        worksheet_id = match.group(1)
                        matched_event = worksheet_to_event.get(worksheet_id)
                        if not matched_event:
                            continue

                        if row['All Day Event'] == "True":
                            start_date = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                            end_date = (datetime.strptime(row['End Date'], "%Y/%m/%d") + timedelta(days=1)).strftime("%Y-%m-%d")
                            event_data = {
                                'start': {'date': start_date},
                                'end': {'date': end_date}
                            }
                        else:
                            start_dt = f"{row['Start Date']} {row['Start Time']}"
                            end_dt = f"{row['End Date']} {row['End Time']}"
                            event_data = {
                                'start': {'dateTime': datetime.strptime(start_dt, "%Y/%m/%d %H:%M").isoformat(), 'timeZone': 'Asia/Tokyo'},
                                'end': {'dateTime': datetime.strptime(end_dt, "%Y/%m/%d %H:%M").isoformat(), 'timeZone': 'Asia/Tokyo'}
                            }

                        try:
                            if update_event_if_needed(service, calendar_id_upd, matched_event, event_data):
                                update_count += 1
                        except Exception as e:
                            st.error(f"{row['Subject']} の更新に失敗: {e}")

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")


# サイドバーに認証情報表示
with st.sidebar:
    st.header("🔐 認証状態")
    st.success("✅ Firebase認証済み")
    st.success("✅ Google認証済み")
    
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
            if key in st.session_state:
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
