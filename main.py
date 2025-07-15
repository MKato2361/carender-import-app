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
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("\U0001F4C5 Googleカレンダー一括イベント登録・削除")

# Firebase初期化
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

# Firebase認証フォーム
user_id = get_firebase_user_id()
if not user_id:
    firebase_auth_form()
    st.stop()

# Google認証
google_auth_placeholder = st.empty()
with google_auth_placeholder.container():
    st.subheader("\U0001F510 Googleカレンダー認証")
    creds = authenticate_google()
    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()
        st.sidebar.success("✅ Googleカレンダーに認証済みです！")

# Google Calendar サービス初期化
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
    except:
        return None, None

# Tasks サービス初期化
def initialize_tasks_service():
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
    except:
        return None, None

# セッションステートでサービス初期化
if 'calendar_service' not in st.session_state:
    service, editable_calendar_options = initialize_calendar_service()
    st.session_state['calendar_service'] = service
    st.session_state['editable_calendar_options'] = editable_calendar_options
else:
    service = st.session_state['calendar_service']

if 'tasks_service' not in st.session_state:
    tasks_service, default_task_list_id = initialize_tasks_service()
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
else:
    tasks_service = st.session_state['tasks_service']

# タブ定義（ここが重要！）
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

with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        all_day_event = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', [])
        )

        # ボタンの前に preview_df を生成（これによりドロップダウンが表示される）
        with st.spinner("イベントデータを読み込み中..."):
            preview_df = process_excel_files(
                st.session_state['uploaded_files'],
                description_columns,
                all_day_event,
                private_event
            )

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()))
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            if st.button("Googleカレンダーに登録する"):
                if preview_df.empty:
                    st.warning("有効なイベントデータがありません。")
                else:
                    st.success(f"{len(preview_df)} 件のイベントを登録できます（ここに登録処理を追加）")
                    # 実際の登録処理は add_event_to_calendar をループで呼び出す形で実装してください
