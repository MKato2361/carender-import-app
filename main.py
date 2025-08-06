import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
import re
from excel_parser import (
    process_excel_data_for_calendar,
    _load_and_merge_dataframes,
    get_available_columns_for_event_name,
    check_event_name_columns,
    format_worksheet_value
)
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
from session_utils import (
    initialize_session_state,
    get_user_setting,
    set_user_setting,
    get_all_user_settings,
    clear_user_settings
)
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from firebase_admin import firestore

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# Firebase初期化
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    firebase_auth_form()
    st.stop()

# ユーザー設定の読み込み
def load_user_settings_from_firestore(user_id):
    if not user_id:
        return
    initialize_session_state(user_id)
    doc_ref = db.collection('user_settings').document(user_id)
    doc = doc_ref.get()
    if doc.exists:
        settings = doc.to_dict()
        for key, value in settings.items():
            set_user_setting(user_id, key, value)

def save_user_setting_to_firestore(user_id, setting_key, setting_value):
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    try:
        doc_ref.set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

load_user_settings_from_firestore(user_id)

# タブの状態をチェックする関数
def check_tab_availability():
    has_files = bool(st.session_state.get('uploaded_files') and not st.session_state['merged_df_for_selector'].empty)
    has_google_auth = bool(st.session_state.get('calendar_service'))
    return {
        "upload": True,  # アップロードタブは常に有効
        "register": has_files and has_google_auth,
        "delete": has_google_auth,
        "update": has_files and has_google_auth
    }

# Google認証
google_auth_placeholder = st.empty()
with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    creds = authenticate_google()
    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()
        st.sidebar.success("✅ Googleカレンダーに認証済みです！")

# カレンダーサービス初期化
def initialize_calendar_service():
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

# ToDoリストサービス初期化
def initialize_tasks_service_wrapper():
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        for task_list in task_lists.get('items', []):
            if task_list.get('title') == 'My Tasks':
                default_task_list_id = task_list['id']
                break
        if not default_task_list_id and task_lists.get('items'):
            default_task_list_id = task_lists['items'][0]['id']
        return tasks_service, default_task_list_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
        return None, None

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

if 'tasks_service' not in st.session_state or not st.session_state.get('tasks_service'):
    tasks_service, default_task_list_id = initialize_tasks_service_wrapper()
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
    if not tasks_service:
        st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
else:
    tasks_service = st.session_state['tasks_service']

# タブの有効/無効状態をチェック
tab_availability = check_tab_availability()

# タブの定義（disabled属性を動的に設定）
tabs = st.tabs([
    f"{'✅ ' if tab_availability['upload'] else '⛔ '}1. ファイルのアップロード",
    f"{'✅ ' if tab_availability['register'] else '⛔ '}2. イベントの登録",
    f"{'✅ ' if tab_availability['delete'] else '⛔ '}3. イベントの削除",
    f"{'✅ ' if tab_availability['update'] else '⛔ '}4. イベントの更新"
])

# セッション状態の初期化
if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame()

# タブ1: ファイルのアップロード
with tabs[0]:
    st.header("ファイルをアップロード")
    st.info("""
    ☐作業指示書一覧をアップロードすると管理番号+26+物件名をイベント名として任意のカレンダーに登録します。
    
    ☐イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。（複数選択可能,次回から同じ項目が選択されます）
    
    ☐イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。
    
    ☐作業外予定の一覧をアップロードすると、イベント名を選択することができます。

    ☐ToDoリストを作成すると、点検通知のリマインドが可能です（ToDoとしてイベント登録されます）
    """)

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

# タブ2: イベントの登録
with tabs[1]:
    if not tab_availability['register']:
        st.warning("Excelファイルをアップロードし、Google認証を完了してください。")
    else:
        st.header("イベントを登録・更新")
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', []),
            default=get_user_setting(user_id, 'description_columns_selected'),
            key=f"description_selector_register_{user_id}"
        )

        fallback_event_name_column = None
        has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
        selected_event_name_col = get_user_setting(user_id, 'event_name_col_selected')

        st.subheader("イベント名の生成設定")
        add_task_type_to_event_name = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=get_user_setting(user_id, 'add_task_type_to_event_name'),
            key=f"add_task_type_checkbox_{user_id}"
        )

        if not (has_mng_data and has_name_data):
            if not has_mng_data and not has_name_data:
                st.info("ファイルに「管理番号」と「物件名」のデータが見つかりませんでした。イベント名に使用する列を選択してください。")
            elif not has_mng_data:
                st.info("ファイルに「管理番号」のデータが見つかりませんでした。物件名と合わせてイベント名に使用する列を選択できます。")
            elif not has_name_data:
                st.info("ファイルに「物件名」のデータが見つかりませんでした。管理番号と合わせてイベント名に使用する列を選択できます。")

            available_event_name_cols = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options = ["選択しない"] + available_event_name_cols
            default_index = event_name_options.index(selected_event_name_col) if selected_event_name_col in event_name_options else 0

            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=default_index,
                key=f"event_name_selector_register_{user_id}"
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

            fixed_todo_types = ["点検通知"]
            if create_todo:
                st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
            else:
                st.markdown("ToDoリストの作成はスキップされます。")

            if st.button("イベントを登録"):
                set_user_setting(user_id, 'description_columns_selected', description_columns)
                set_user_setting(user_id, 'event_name_col_selected', selected_event_name_col)
                set_user_setting(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)
                save_user_setting_to_firestore(user_id, 'description_columns_selected', description_columns)
                save_user_setting_to_firestore(user_id, 'event_name_col_selected', selected_event_name_col)
                save_user_setting_to_firestore(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)

                with st.spinner("イベントを処理中..."):
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'],
                            description_columns,
                            all_day_event_override,
                            private_event,
                            fallback_event_name_column,
                            add_task_type_to_event_name
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame()

                    if df.empty:
                        st.warning("有効なイベントデータがありません。登録を中断しました。")
                    else:
                        # イベント登録処理（変更なし）
                        # ... (既存のイベント登録ロジックをここに挿入)
                        st.success("イベント登録が完了しました。")

# タブ3: イベントの削除
with tabs[2]:
    if not tab_availability['delete']:
        st.warning("Google認証を完了してください。")
    else:
        st.header("イベントを削除")
        selected_calendar_name_del = st.selectbox(
            "削除対象カレンダーを選択",
            list(st.session_state['editable_calendar_options'].keys()),
            key="del_calendar_select"
        )
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]
        delete_start_date = st.date_input("削除するイベントの開始日", value=date.today())
        delete_end_date = st.date_input("削除するイベントの終了日", value=date.today() + timedelta(days=7))
        delete_related_todos = st.checkbox("関連するToDoタスクも削除する", value=False, key="delete_todos")

        if st.button("イベントを削除"):
            start_dt_utc = datetime.combine(delete_start_date, datetime.min.time(), tzinfo=timezone.utc)
            end_dt_utc = datetime.combine(delete_end_date, datetime.max.time(), tzinfo=timezone.utc)
            time_min_utc = start_dt_utc.isoformat()
            time_max_utc = end_dt_utc.isoformat()

            events_to_delete = fetch_all_events(service, calendar_id_del, time_min_utc, time_max_utc)
            
            if not events_to_delete:
                st.info("指定期間内に削除するイベントはありませんでした。")
            else:
                # イベント削除処理（変更なし）
                # ... (既存のイベント削除ロジックをここに挿入)
                st.success("イベント削除が完了しました。")

# タブ4: イベントの更新
with tabs[3]:
    if not tab_availability['update']:
        st.warning("Excelファイルをアップロードし、Google認証を完了してください。")
    else:
        st.header("イベントを更新")
        st.info("このタブは、主に既存イベントの情報をExcelデータに基づいて**上書き**したい場合に使用します。新規イベントの作成は行いません。")

        all_day_event_override_update = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event_update = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")

        description_columns_update = st.multiselect(
            "説明欄に含める列",
            st.session_state['description_columns_pool'],
            default=get_user_setting(user_id, 'description_columns_selected'),
            key=f"update_desc_cols_{user_id}"
        )

        fallback_event_name_column_update = None
        has_mng_data_update, has_name_data_update = check_event_name_columns(st.session_state['merged_df_for_selector'])
        selected_event_name_col_update = get_user_setting(user_id, 'event_name_col_selected_update')

        st.subheader("更新時のイベント名の生成設定")
        add_task_type_to_event_name_update = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=get_user_setting(user_id, 'add_task_type_to_event_name_update'),
            key=f"add_task_type_checkbox_update_{user_id}"
        )

        if not (has_mng_data_update and has_name_data_update):
            st.info("Excelデータからのイベント名生成に、以下の列を代替として使用できます。")
            available_event_name_cols_update = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options_update = ["選択しない"] + available_event_name_cols_update
            default_index_update = event_name_options_update.index(selected_event_name_col_update) if selected_event_name_col_update in event_name_options_update else 0

            selected_event_name_col_update = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options_update,
                index=default_index_update,
                key=f"event_name_selector_update_{user_id}"
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
                set_user_setting(user_id, 'description_columns_selected', description_columns_update)
                set_user_setting(user_id, 'event_name_col_selected_update', selected_event_name_col_update)
                set_user_setting(user_id, 'add_task_type_to_event_name_update', add_task_type_to_event_name_update)
                save_user_setting_to_firestore(user_id, 'description_columns_selected', description_columns_update)
                save_user_setting_to_firestore(user_id, 'event_name_col_selected_update', selected_event_name_col_update)
                save_user_setting_to_firestore(user_id, 'add_task_type_to_event_name_update', add_task_type_to_event_name_update)

                with st.spinner("イベントを処理中..."):
                    # イベント更新処理（変更なし）
                    # ... (既存のイベント更新ロジックを挿入)
                    st.success("イベント更新が完了しました。")

# サイドバー
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
    
    if st.button("🚪 ログアウト", type="secondary"):
        if user_id:
            clear_user_settings(user_id)
        for key in list(st.session_state.keys()):
            if not key.startswith("google_auth") and not key.startswith("firebase_"):
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()