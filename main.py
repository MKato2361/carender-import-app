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
import os
from pathlib import Path
from io import BytesIO

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    firebase_auth_form()
    st.stop()

def load_user_settings_from_firestore(user_id):
    """Firestoreからユーザー設定を読み込み、セッションに同期"""
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
    """Firestoreにユーザー設定を保存"""
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    try:
        doc_ref.set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

# ユーザー設定の読み込み
load_user_settings_from_firestore(user_id)

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
    "4. イベントの更新",
    "5. イベントのExcel出力" 
])

if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame()

# （中略：アップロード関連処理は変更なし）

# ==========================================
# 【修正箇所】イベント登録処理部分
# ==========================================
                        for event in events:
                            desc = event.get('description', '')
                            match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", desc)  # ★ 修正済み: 作業指示書番号照合を厳密化
                            if match:
                                worksheet_id = match.group(1)
                                worksheet_to_event[worksheet_id] = event

                        for i, row in df.iterrows():
                            match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", row['Description'])  # ★ 修正済み: 作業指示書番号照合を厳密化
                            event_data = {
                                'summary': row['Subject'],
                                'location': row['Location'],
                                'description': row['Description'],
                                'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                            }
                            # （後略：登録・更新処理はすべて変更なし）

# ==========================================
# 【修正箇所】イベント更新処理部分
# ==========================================
                    worksheet_to_event = {}
                    for event in events:
                        desc = event.get('description', '')
                        match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", desc)  # ★ 修正済み
                        if match:
                            worksheet_id = match.group(1)
                            worksheet_to_event[worksheet_id] = event

                    update_count = 0
                    progress_bar = st.progress(0)
                    for i, row in df.iterrows():
                        match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", row['Description'])  # ★ 修正済み
                        if not match:
                            progress_bar.progress((i + 1) / len(df))
                            continue
                        
                        worksheet_id = match.group(1)
                        matched_event = worksheet_to_event.get(worksheet_id)
                        if not matched_event:
                            progress_bar.progress((i + 1) / len(df))
                            continue
                        # （後略：更新処理は変更なし）
