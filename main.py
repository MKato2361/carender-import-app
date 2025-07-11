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

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# Firebaseの初期化
if not initialize_firebase():
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

# ToDoリストサービスの初期化を試みる
if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    try:
        service = build("calendar", "v3", credentials=creds)
        st.session_state['calendar_service'] = service
        calendar_list = service.calendarList().list().execute()

        editable_calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') != 'reader'
        }
        st.session_state['editable_calendar_options'] = editable_calendar_options

    except Exception as e:
        st.error(f"カレンダーサービスの取得またはカレンダーリストの取得に失敗しました: {e}")
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop()
else:
    service = st.session_state['calendar_service']

# ToDoリストサービスをここでビルド
if 'tasks_service' not in st.session_state or not st.session_state['tasks_service']:
    try:
        tasks_service = build_tasks_service(creds)
        st.session_state['tasks_service'] = tasks_service
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        for task_list in task_lists.get('items', []):
            if task_list.get('title') == 'My Tasks':
                default_task_list_id = task_list['id']
                break
        if not default_task_list_id and task_lists.get('items'):
            default_task_list_id = task_lists['items'][0]['id']

        st.session_state['default_task_list_id'] = default_task_list_id

    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました。ToDoリスト機能は利用できません: {e}")
        st.session_state['tasks_service'] = None
        st.session_state['default_task_list_id'] = None
else:
    tasks_service = st.session_state['tasks_service']


tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)
    # ... (既存のファイルアップロードロジックは変更なし)
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
    # ... (既存のイベント登録ロジックは変更なし)

with tabs[2]:
    st.header("イベントを削除")
    # ... (既存のイベント削除ロジックは変更なし)

with tabs[3]:
    st.header("イベントを更新")
    # ... (既存のイベント更新ロジックは変更なし)
