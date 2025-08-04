import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
import re
from excel_parser import (
    process_excel_data_for_calendar,
    _load_and_merge_dataframes,
    get_available_columns_for_event_name,
    check_event_name_columns,
    format_worksheet_value # この関数が必要になります
)
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id # ToDo関連は今回の変更で複雑になる可能性があるため、一旦イベントのみに焦点を当てる
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from firebase_admin import firestore

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")


# Firebaseの初期化
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

# Firestoreクライアントの取得
db = firestore.client()

# Firebase認証フォームの表示とユーザーIDの取得
user_id = get_firebase_user_id()

if not user_id:
    # ユーザーがログインしていない場合、認証フォームを表示して停止
    firebase_auth_form()
    st.stop()

# ユーザー固有の設定をFirestoreから読み込む
def load_user_settings(user_id):
    """Firestoreからユーザー設定を読み込み、st.session_stateに設定する"""
    if not user_id:
        return

    doc_ref = db.collection('user_settings').document(user_id)
    doc = doc_ref.get()

    if doc.exists:
        settings = doc.to_dict()
        # 各選択項目のキーがユーザーIDに紐付くように修正
        if 'description_columns_selected' in settings:
            st.session_state[f'description_columns_selected_{user_id}'] = settings['description_columns_selected']
        if 'event_name_col_selected' in settings:
            st.session_state[f'event_name_col_selected_{user_id}'] = settings['event_name_col_selected']
        if 'event_name_col_selected_update' in settings: # 更新タブ用の設定も考慮
            st.session_state[f'event_name_col_selected_update_{user_id}'] = settings['event_name_col_selected_update']
    else:
        # ドキュメントがない場合はデフォルト値を設定
        st.session_state[f'description_columns_selected_{user_id}'] = ["内容", "詳細"]
        st.session_state[f'event_name_col_selected_{user_id}'] = "選択しない"
        st.session_state[f'event_name_col_selected_update_{user_id}'] = "選択しない" # 更新タブ用デフォルト

def save_user_setting(user_id, setting_key, setting_value):
    """ユーザー設定をFirestoreに保存する"""
    if not user_id:
        return

    doc_ref = db.collection('user_settings').document(user_id)
    try:
        doc_ref.set({setting_key: setting_value}, merge=True) # merge=True で既存のフィールドを上書きせず更新
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

load_user_settings(user_id)


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
def initialize_tasks_service_wrapper():
    """タスクサービスを初期化する"""
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
            
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        
        # デフォルトのタスクリストを探す
        for task_list in task_lists.get('items', []):
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
    _, st.session_state['editable_calendar_options'] = initialize_calendar_service()


# タスクサービスの初期化または取得
if 'tasks_service' not in st.session_state or not st.session_state.get('tasks_service'): # Noneチェックを追加
    tasks_service, default_task_list_id = initialize_tasks_service_wrapper()
    
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
    
    if not tasks_service:
        st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
else:
    tasks_service = st.session_state['tasks_service']

# メインタブの作成
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録", # このタブでUpsertロジックを実装
    "3. イベントの削除",
    "4. イベントの更新" # このタブはそのまま残すか、2と統合するか要検討
])

# セッション状態の初期化
if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame() # 新しくマージ済みDFを保持



with tabs[0]:
    st.header("ファイルをアップロード")
    st.info("""
    ☐作業指示書一覧をアップロードすると管理番号+物件名をイベント名として任意のカレンダーに登録します。
    
    ☐イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。（複数選択可能,次回から同じ項目が選択されます）
    
    ☐イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。
    
    ☐作業外予定の一覧をアップロードすると、イベント名を選択することができます。

    ☐ToDoリストを作成すると、点検通知のリマインドが可能です（ToDoとしてイベント登録されます）
    """)

    import os
    from pathlib import Path
    from io import BytesIO

    def get_local_excel_files():
        current_dir = Path(__file__).parent
        return [f for f in current_dir.glob("*.xlsx") if f.is_file()]

    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    local_excel_files = get_local_excel_files()
    selected_local_files = []
    if local_excel_files:
        st.subheader("📁 サーバーにあるExcelファイル")
        local_file_names = [f.name for f in local_excel_files]
        selected_names = st.multiselect(
            "以下のファイルを処理対象に含める（アップロードと同様に扱われます）",
            local_file_names
        )
        for name in selected_names:
            full_path = next((f for f in local_excel_files if f.name == name), None)
            if full_path:
                with open(full_path, "rb") as f:
                    file_bytes = f.read()
                    file_obj = BytesIO(file_bytes)
                    file_obj.name = name
                    selected_local_files.append(file_obj)

    all_files = []
    if uploaded_files:
        all_files.extend(uploaded_files)
    if selected_local_files:
        all_files.extend(selected_local_files)

    if all_files:
        st.session_state['uploaded_files'] = all_files
        try:
            st.session_state['merged_df_for_selector'] = _load_and_merge_dataframes(all_files)
            st.session_state['description_columns_pool'] = st.session_state['merged_df_for_selector'].columns.tolist()

            if st.session_state['merged_df_for_selector'].empty:
                st.warning("読み込まれたファイルに有効なデータがありませんでした。")
        except (ValueError, IOError) as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []

    if st.session_state.get('uploaded_files'):
        st.subheader("📄 処理対象ファイル一覧")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")
        if not st.session_state['merged_df_for_selector'].empty:
            st.info(f"📊 データ列数: {len(st.session_state['merged_df_for_selector'].columns)}、行数: {len(st.session_state['merged_df_for_selector'])}")

        if st.button("🗑️ アップロード済みファイルをクリア", help="選択中のファイルとデータを削除します。"):
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()
with tabs[1]:
    st.header("イベントを登録・更新") # タブ名を変更
    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)
        # 作業タイプ列をイベント名の先頭に追加するかのチェックボックスを追加
        prepend_event_type = st.checkbox("イベント名の先頭に作業タイプを追加する", value=False)

        # 説明文に含める列の選択 (ユーザーごとに記憶)
        current_description_cols_selection = st.session_state.get(f'description_columns_selected_{user_id}', [])
        
        # description_columns を初期化
        description_columns = []
        if st.session_state.get('description_columns_pool'):
            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                st.session_state.get('description_columns_pool', []),
                default=[col for col in current_description_cols_selection if col in st.session_state.get('description_columns_pool', [])],
                key=f"description_selector_register_{user_id}", # ユーザー固有のキー
            )
        else:
            st.info("説明欄に含める列の候補がありません。ファイルをアップロードしてください。")
            description_columns = current_description_cols_selection # 候補がない場合でも既存の設定は保持

        # イベント名の代替列選択UIをここに配置 (ユーザーごとに記憶)
        fallback_event_name_column = None
        has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
        
        # selected_event_name_col を初期化
        selected_event_name_col = st.session_state.get(f'event_name_col_selected_{user_id}', "選択しない")

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
            
            # current_event_name_selection の代わりに selected_event_name_col を使用
            default_index = event_name_options.index(selected_event_name_col) if selected_event_name_col in event_name_options else 0
            
            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=default_index,
                key=f"event_name_selector_register_{user_id}", # ユーザー固有のキー
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
            fixed_todo_types = ["点検通知"]
            
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


            st.subheader("➡️ イベント登録・更新実行") # ボタンの表記も変更
            if st.button("Googleカレンダーに登録・更新する"):
                # ここでFirestoreに選択項目を保存
                save_user_setting(user_id, 'description_columns_selected', description_columns)
                save_user_setting(user_id, 'event_name_col_selected', selected_event_name_col)


                with st.spinner("イベントデータを処理中..."):
                    # process_excel_data_for_calendar を呼び出す
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns, 
                            all_day_event_override,
                            private_event, 
                            fallback_event_name_column,
                            prepend_event_type
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame() # エラー時は空のDFにする

                    if df.empty:
                        st.warning("有効なイベントデータがありません。処理を中断しました。")
                    else:
                        st.info(f"{len(df)} 件のイベントを処理します。")
                        progress = st.progress(0)
                        successful_operations = 0
                        successful_todo_creations = 0

                        # Googleカレンダーから既存イベントを作業指示書IDで検索するための準備
                        # 広めの期間で検索（例: 過去1年～未来5年など、実運用に合わせて調整）
                        now_for_search = datetime.now()
                        search_time_min = (now_for_search - timedelta(days=365)).isoformat() + 'Z' # 過去1年
                        search_time_max = (now_for_search + timedelta(days=365*5)).isoformat() + 'Z' # 未来5年
                        
                        existing_events = fetch_all_events(service, calendar_id, search_time_min, search_time_max)
                        
                        # 作業指示書IDをキーとした既存イベントの辞書を作成
                        worksheet_id_to_existing_event = {}
                        for event in existing_events:
                            desc = event.get('description', '')
                            # '作業指示書:' の後に続く数値を抽出
                            match = re.search(r"作業指示書[：:]\s*(\d+)", desc) 
                            if match:
                                worksheet_id = match.group(1)
                                worksheet_id_to_existing_event[worksheet_id] = event


                        for i, row in df.iterrows():
                            event_summary = row['Subject']
                            event_start_date_obj = None
                            event_end_date_obj = None
                            event_time_str = "" # ToDo詳細用の時間文字列
                            event_id_for_todo = None # ToDoに紐付けるイベントID

                            # Excelデータの 'Description' から作業指示書IDを抽出 (format_worksheet_valueで付与されていることを前提)
                            excel_description = row['Description']
                            excel_worksheet_match = re.search(r"作業指示書[：:]\s*(\d+)", excel_description)
                            excel_worksheet_id = excel_worksheet_match.group(1) if excel_worksheet_match else None

                            event_data_to_process = None
                            operation_type = "新規登録"

                            if excel_worksheet_id and excel_worksheet_id in worksheet_id_to_existing_event:
                                existing_event = worksheet_id_to_existing_event[excel_worksheet_id]
                                existing_event_id = existing_event['id']
                                
                                # 更新対象イベントのデータ構造を構築
                                updated_event_data = {
                                    'summary': event_summary,
                                    'location': row['Location'],
                                    'description': row['Description'],
                                    'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                }

                                # 日時の設定
                                if row['All Day Event'] == "True":
                                    event_start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                                    event_end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                                    
                                    start_date_str = event_start_date_obj.strftime("%Y-%m-%d")
                                    end_date_for_api = (event_end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d") 
                                    
                                    updated_event_data['start'] = {'date': start_date_str}
                                    updated_event_data['end'] = {'date': end_date_for_api}
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

                                    updated_event_data['start'] = {'dateTime': start_iso, 'timeZone': 'Asia/Tokyo'}
                                    updated_event_data['end'] = {'dateTime': end_iso, 'timeZone': 'Asia/Tokyo'}
                                    event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%H:%M')}"
                                    if event_start_datetime_obj.date() != event_end_datetime_obj.date():
                                        event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}"
                                
                                # update_event_if_neededを使って、変更がある場合のみ更新
                                updated_or_existing_event = update_event_if_needed(service, calendar_id, existing_event_id, updated_event_data)
                                if updated_or_existing_event and updated_or_existing_event != existing_event:
                                    successful_operations += 1
                                    operation_type = "更新"
                                    event_id_for_todo = updated_or_existing_event.get('id')
                                else:
                                    # 変更がない場合はスキップ
                                    progress.progress((i + 1) / len(df))
                                    continue

                            else:
                                # 新規イベントデータ構築
                                if row['All Day Event'] == "True":
                                    event_start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                                    event_end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                                    
                                    start_date_str = event_start_date_obj.strftime("%Y-%m-%d")
                                    end_date_for_api = (event_end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d") 

                                    event_data_to_process = {
                                        'summary': event_summary,
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_for_api},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
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

                                    event_data_to_process = {
                                        'summary': event_summary,
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'dateTime': start_iso, 'timeZone': 'Asia/Tokyo'},
                                        'end': {'dateTime': end_iso, 'timeZone': 'Asia/Tokyo'},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                    event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%H:%M')}"
                                    if event_start_datetime_obj.date() != event_end_datetime_obj.date():
                                        event_time_str = f"{event_start_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}～{event_end_datetime_obj.strftime('%Y/%-m/%-d %H:%M')}"
                                
                                created_event = add_event_to_calendar(service, calendar_id, event_data_to_process)
                                if created_event:
                                    successful_operations += 1
                                    event_id_for_todo = created_event.get('id')
                                else:
                                    progress.progress((i + 1) / len(df))
                                    continue # イベント登録失敗時はToDoも作成しない

                            # ToDoリストの作成ロジック (更新の場合も新規作成される)
                            if create_todo and tasks_service and st.session_state.get('default_task_list_id') and event_id_for_todo: 
                                if event_start_date_obj: 
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
                                                f"関連イベントID: {event_id_for_todo}\n" 
                                                f"イベント日時: {event_time_str}\n"
                                                f"場所: {row['Location']}"
                                            )
                                            # TODO: 既存のToDoタスクがある場合に更新するロジックをここに追加する必要がある
                                            # 現状のコードでは常に新規作成になる
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
                            
                            progress.progress((i + 1) / len(df))

                        st.success(f"✅ {successful_operations} 件のイベントが処理されました (新規登録/更新)。")
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
        today_date = date.today()
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
                start_dt_utc = datetime.combine(delete_start_date, datetime.min.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                end_dt_utc = datetime.combine(delete_end_date, datetime.max.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                
                time_min_utc = start_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                time_max_utc = end_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')


                events_to_delete = fetch_all_events(calendar_service, calendar_id_del, time_min_utc, time_max_utc)
                
                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")

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
    st.info("このタブは、主に既存イベントの情報をExcelデータに基づいて**上書き**したい場合に使用します。新規イベントの作成は行いません。")

    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        # 更新タブでの設定も、登録タブと同様に名称と挙動を統一
        all_day_event_override_update = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event_update = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")
        
        # 説明欄に含める列 (更新タブ用、ユーザーごとに記憶)
        current_description_cols_selection_update = st.session_state.get(f'description_columns_selected_{user_id}', [])

        # description_columns_update を初期化
        description_columns_update = []
        if st.session_state.get('description_columns_pool'):
            description_columns_update = st.multiselect(
                "説明欄に含める列", 
                st.session_state['description_columns_pool'], 
                default=[col for col in current_description_cols_selection_update if col in st.session_state.get('description_columns_pool', [])],
                key=f"update_desc_cols_{user_id}", # ユーザー固有のキー
            )
        else:
            st.info("説明欄に含める列の候補がありません。ファイルをアップロードしてください。")
            description_columns_update = current_description_cols_selection_update # 候補がない場合でも既存の設定は保持

        # イベント名の代替列選択UIをここに配置 (更新タブ用、ユーザーごとに記憶)
        fallback_event_name_column_update = None
        has_mng_data_update, has_name_data_update = check_event_name_columns(st.session_state['merged_df_for_selector'])
        
        # selected_event_name_col_update を初期化
        selected_event_name_col_update = st.session_state.get(f'event_name_col_selected_update_{user_id}', "選択しない")

        if not (has_mng_data_update and has_name_data_update):
            st.subheader("更新時のイベント名の設定")
            st.info("Excelデータからのイベント名生成に、以下の列を代替として使用できます。")

            available_event_name_cols_update = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options_update = ["選択しない"] + available_event_name_cols_update
            
            # st.session_stateに保存された値を使用
            current_event_name_selection_update = st.session_state.get(f'event_name_col_selected_update_{user_id}', "選択しない")
            
            # 現在の選択がオプションリストにあるか確認し、なければデフォルトにフォールバック
            default_index_update = event_name_options_update.index(current_event_name_selection_update) if current_event_name_selection_update in event_name_options_update else 0

            selected_event_name_col_update = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options_update,
                index=default_index_update,
                key=f"event_name_selector_update_{user_id}", # ユーザー固有のキー
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
                # ここでFirestoreに選択項目を保存
                save_user_setting(user_id, 'description_columns_selected_update', description_columns_update)
                save_user_setting(user_id, 'event_name_col_selected_update', selected_event_name_col_update)

                with st.spinner("イベントを処理中..."):
                    try:
                        # process_excel_data_for_calendar を呼び出す
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns_update, # 更新タブ用の列
                            all_day_event_override_update, # 更新タブ用の設定
                            private_event_update,         # 更新タブ用の設定
                            fallback_event_name_column_update, # 新しい引数
                            prepend_event_type # 登録タブと同じ変数を使用
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
        # ユーザー固有の設定もクリア
        if user_id:
            # Firestoreからユーザー設定を削除する（オプション、通常は残します）
            # try:
            #     db.collection('user_settings').document(user_id).delete()
            #     st.info("ユーザー設定をFirestoreから削除しました。")
            # except Exception as e:
            #     st.error(f"ユーザー設定の削除に失敗しました。")

            if f'description_columns_selected_{user_id}' in st.session_state:
                del st.session_state[f'description_columns_selected_{user_id}']
            if f'event_name_col_selected_{user_id}' in st.session_state:
                del st.session_state[f'event_name_col_selected_{user_id}']
            if f'event_name_col_selected_update_{user_id}' in st.session_state:
                del st.session_state[f'event_name_col_selected_update_{user_id}']

        # その他のセッション状態をクリア
        for key in list(st.session_state.keys()):
            # 認証関連のキーは残すか、Firebase認証ロジックと連携して適切に処理
            if not key.startswith("google_auth") and not key.startswith("firebase_"):
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
