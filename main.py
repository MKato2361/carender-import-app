import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
import re
import json
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from firebase_admin import firestore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from excel_parser import (
    process_excel_data_for_calendar,
    _load_and_merge_dataframes,
    get_available_columns_for_event_name,
    check_event_name_columns
)
from calendar_utils import (
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id,
    delete_event_from_calendar
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id

# ===== 設定 =====
SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# ===== Firebase初期化 =====
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    firebase_auth_form()
    st.stop()

# ===== Google OAuth フロー =====
def google_oauth_flow():
    doc_ref = db.collection('google_tokens').document(user_id)

    # セッションから
    if 'credentials' in st.session_state:
        creds = st.session_state['credentials']
        if creds and creds.valid:
            return creds
        elif creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state['credentials'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            return creds

    # Firestoreから
    try:
        doc = doc_ref.get()
        if doc.exists:
            creds = Credentials.from_authorized_user_info(doc.to_dict(), SCOPES)
            if creds and creds.valid:
                st.session_state['credentials'] = creds
                return creds
            elif creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                doc_ref.set(json.loads(creds.to_json()))
                return creds
    except Exception as e:
        st.error(f"Firestoreトークン読み込み失敗: {e}")

    # OAuth開始
    client_config = {
        "web": {
            "client_id": st.secrets["google"]["client_id"],
            "project_id": st.secrets["google"]["project_id"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": st.secrets["google"]["client_secret"],
            "redirect_uris": [st.secrets["google"]["redirect_uri"]]
        }
    }
    flow = Flow.from_client_config(client_config, SCOPES)
    flow.redirect_uri = st.secrets["google"]["redirect_uri"]

    params = st.experimental_get_query_params()
    if "code" not in params:
        auth_url, _ = flow.authorization_url(
            prompt='consent',
            access_type='offline',
            include_granted_scopes='true'
        )
        st.markdown(f"[Googleでログインする]({auth_url})")
        st.stop()
    else:
        code = params["code"][0]
        flow.fetch_token(code=code)
        creds = flow.credentials
        st.session_state['credentials'] = creds
        doc_ref.set(json.loads(creds.to_json()))
        st.success("Google認証完了！")
        st.experimental_set_query_params()
        return creds

# 実行
creds = google_oauth_flow()
if not creds:
    st.warning("Googleカレンダー認証を完了してください。")
    st.stop()
else:
    st.sidebar.success("✅ Googleカレンダー認証済み！")

# ===== Googleサービス初期化 =====
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
    except Exception as e:
        st.error(f"カレンダーサービス初期化失敗: {e}")
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
    except Exception as e:
        st.warning(f"ToDoリストサービス初期化失敗: {e}")
        return None, None

service, editable_calendar_options = initialize_calendar_service()
tasks_service, default_task_list_id = initialize_tasks_service_wrapper()

if not service:
    st.stop()

# ===== ユーザー設定管理 =====
def load_user_settings(user_id):
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    doc = doc_ref.get()
    if doc.exists:
        settings = doc.to_dict()
        for k, v in settings.items():
            st.session_state[f'{k}_{user_id}'] = v

def save_user_setting(user_id, setting_key, setting_value):
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    doc_ref.set({setting_key: setting_value}, merge=True)

load_user_settings(user_id)

# ===== タブ構成 =====
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

# === タブ1: アップロード ===
with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)
    if uploaded_files:
        try:
            merged_df = _load_and_merge_dataframes(uploaded_files)
            st.session_state['uploaded_files'] = uploaded_files
            st.session_state['merged_df_for_selector'] = merged_df
            st.session_state['description_columns_pool'] = merged_df.columns.tolist()
            st.success(f"ファイル読み込み成功: {len(merged_df)}件")
        except Exception as e:
            st.error(f"読み込み失敗: {e}")

# === タブ2: イベント登録 ===
with tabs[1]:
    st.header("イベントを登録・更新")
    if not st.session_state.get('uploaded_files'):
        st.info("先にタブ1でファイルをアップロードしてください。")
    else:
        # 設定読み込み
        selected_calendar_name = st.selectbox("登録先カレンダー", list(editable_calendar_options.keys()))
        calendar_id = editable_calendar_options[selected_calendar_name]

        create_todo = st.checkbox("ToDoリストを作成する", value=False)
        deadline_offset_options = {"2週間前": 14, "10日前": 10, "1週間前": 7, "カスタム日数前": None}
        selected_offset_key = st.selectbox("ToDo期限", list(deadline_offset_options.keys()), disabled=not create_todo)
        custom_offset_days = None
        if selected_offset_key == "カスタム日数前":
            custom_offset_days = st.number_input("日数", min_value=0, value=3, disabled=not create_todo)

        if st.button("Googleカレンダーに登録・更新する"):
            df = process_excel_data_for_calendar(
                st.session_state['uploaded_files'],
                [],
                False,
                True
            )
            existing_events = fetch_all_events(service, calendar_id)
            worksheet_id_to_event = {}
            for e in existing_events:
                desc = e.get('description', '')
                match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                if match:
                    worksheet_id_to_event[match.group(1)] = e

            for _, row in df.iterrows():
                excel_desc = row['Description']
                match = re.search(r"作業指示書[：:]\s*(\d+)", excel_desc)
                worksheet_id = match.group(1) if match else None
                event_data = {
                    'summary': row['Subject'],
                    'location': row['Location'],
                    'description': row['Description'],
                    'start': {'dateTime': row['Start Date'], 'timeZone': 'Asia/Tokyo'},
                    'end': {'dateTime': row['End Date'], 'timeZone': 'Asia/Tokyo'}
                }
                if worksheet_id and worksheet_id in worksheet_id_to_event:
                    update_event_if_needed(service, calendar_id, worksheet_id_to_event[worksheet_id]['id'], event_data)
                else:
                    add_event_to_calendar(service, calendar_id, event_data)

                if create_todo and tasks_service and default_task_list_id:
                    offset_days = deadline_offset_options.get(selected_offset_key)
                    if selected_offset_key == "カスタム日数前" and custom_offset_days is not None:
                        offset_days = custom_offset_days
                    if offset_days is not None:
                        due_date = datetime.fromisoformat(row['Start Date']).date() - timedelta(days=offset_days)
                        task_data = {
                            'title': f"点検通知 - {row['Subject']}",
                            'notes': f"関連イベントID: {worksheet_id}",
                            'due': due_date.isoformat() + 'Z'
                        }
                        add_task_to_todo_list(tasks_service, default_task_list_id, task_data)

            st.success("処理完了")

# === タブ3: 削除 ===
with tabs[2]:
    st.header("イベントを削除")
    selected_calendar_name = st.selectbox("削除対象カレンダー", list(editable_calendar_options.keys()))
    calendar_id = editable_calendar_options[selected_calendar_name]
    delete_related_todos = st.checkbox("関連ToDoも削除する", value=False)
    if st.button("削除実行"):
        events = fetch_all_events(service, calendar_id)
        for e in events:
            if delete_related_todos and tasks_service and default_task_list_id:
                find_and_delete_tasks_by_event_id(tasks_service, default_task_list_id, e['id'])
            delete_event_from_calendar(service, calendar_id, e['id'])
        st.success("削除完了")

# === タブ4: 更新 ===
with tabs[3]:
    st.header("イベントを更新")
    selected_calendar_name = st.selectbox("更新対象カレンダー", list(editable_calendar_options.keys()))
    calendar_id = editable_calendar_options[selected_calendar_name]
    if st.button("更新実行"):
        df = process_excel_data_for_calendar(st.session_state['uploaded_files'], [], False, True)
        events = fetch_all_events(service, calendar_id)
        worksheet_to_event = {}
        for e in events:
            desc = e.get('description', '')
            match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
            if match:
                worksheet_to_event[match.group(1)] = e
        for _, row in df.iterrows():
            match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
            if match and match.group(1) in worksheet_to_event:
                update_event_if_needed(service, calendar_id, worksheet_to_event[match.group(1)]['id'], {
                    'summary': row['Subject'],
                    'location': row['Location'],
                    'description': row['Description']
                })
        st.success("更新完了")
