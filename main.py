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
    firebase_auth_form()
    st.stop()

# ユーザー固有の設定をFirestoreから読み込む
def load_user_settings(user_id):
    if not user_id:
        return

    doc_ref = db.collection('user_settings').document(user_id)
    doc = doc_ref.get()

    if doc.exists:
        settings = doc.to_dict()
        if 'description_columns_selected' in settings:
            st.session_state[f'description_columns_selected_{user_id}'] = settings['description_columns_selected']
        if 'event_name_col_selected' in settings:
            st.session_state[f'event_name_col_selected_{user_id}'] = settings['event_name_col_selected']
        if 'event_name_col_selected_update' in settings:
            st.session_state[f'event_name_col_selected_update_{user_id}'] = settings['event_name_col_selected_update']
    else:
        st.session_state[f'description_columns_selected_{user_id}'] = ["内容", "詳細"]
        st.session_state[f'event_name_col_selected_{user_id}'] = "選択しない"
        st.session_state[f'event_name_col_selected_update_{user_id}'] = "選択しない"

def save_user_setting(user_id, setting_key, setting_value):
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    try:
        doc_ref.set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

load_user_settings(user_id)

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

tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame()

with tabs[0]:
    st.header("ファイルをアップロード")
    st.info("""
    ☐作業指示書一覧をアップロードすると管理番号+物件名をイベント名として任意のカレンダーに登録します。
    ☐イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。
    ☐イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。
    ☐作業外予定の一覧をアップロードすると、イベント名を選択することができます。
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

        if st.button("🗑️ アップロード済みファイルをクリア"):
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()

with tabs[1]:
    st.header("イベントを登録・更新")
    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)
        include_work_type = st.checkbox("イベント名の先頭に作業タイプを追加", value=False)

        current_description_cols_selection = st.session_state.get(f'description_columns_selected_{user_id}', [])
        description_columns = []
        if st.session_state.get('description_columns_pool'):
            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                st.session_state.get('description_columns_pool', []),
                default=[col for col in current_description_cols_selection if col in st.session_state.get('description_columns_pool', [])],
                key=f"description_selector_register_{user_id}",
            )
        else:
            st.info("説明欄に含める列の候補がありません。ファイルをアップロードしてください。")
            description_columns = current_description_cols_selection

        fallback_event_name_column = None
        has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
        selected_event_name_col = st.session_state.get(f'event_name_col_selected_{user_id}', "選択しない")

        if not (has_mng_data and has_name_data):
            st.subheader("イベント名の設定")
            available_event_name_cols = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options = ["選択しない"] + available_event_name_cols
            default_index = event_name_options.index(selected_event_name_col) if selected_event_name_col in event_name_options else 0
            
            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=default_index,
                key=f"event_name_selector_register_{user_id}",
            )
            if selected_event_name_col != "選択しない":
                fallback_event_name_column = selected_event_name_col
        else:
            st.info("「管理番号」と「物件名」が利用できます。")

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("➡️ イベント登録・更新実行")
            if st.button("Googleカレンダーに登録・更新する"):
                save_user_setting(user_id, 'description_columns_selected', description_columns)
                save_user_setting(user_id, 'event_name_col_selected', selected_event_name_col)

                with st.spinner("イベントデータを処理中..."):
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'],
                            description_columns,
                            all_day_event_override,
                            private_event,
                            fallback_event_name_column,
                            include_work_type=include_work_type
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラー: {e}")
                        df = pd.DataFrame()

                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                    else:
                        st.info(f"{len(df)} 件のイベントを処理します。")
                        # --- ここから先のGoogleカレンダー登録ロジックは元コードを維持 ---
                        # （既存イベントの照合・新規登録・更新処理）
                        # ...
