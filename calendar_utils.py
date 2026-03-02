import os
import json
import pickle
from pathlib import Path
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from firebase_admin import firestore
from firebase_auth import get_firebase_user_id
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta, timezone

# Google API スコープ
SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks",
    "https://www.googleapis.com/auth/spreadsheets",
]

CLIENT_SECRET_FILE = "client_secret.json"
REDIRECT_URI = "http://localhost:8501"


# ==============================
# Google 認証（PKCE対応・Session保持）
# ==============================
def authenticate_google():
    if "google_creds" in st.session_state:
        return st.session_state["google_creds"]

    # Flow を session_state に保存（PKCE対策）
    if "google_flow" not in st.session_state:
        flow = Flow.from_client_secrets_file(
            CLIENT_SECRET_FILE,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
        )

        auth_url, _ = flow.authorization_url(
            prompt="consent",
            access_type="offline",
            include_granted_scopes="true",
        )

        st.session_state["google_flow"] = flow

        st.markdown("### 🔐 Google認証が必要です")
        st.markdown(f"[👉 Googleでログイン]({auth_url})")
        return None

    flow = st.session_state["google_flow"]

    query_params = st.query_params

    if "code" in query_params:
        try:
            flow.fetch_token(code=query_params["code"])

            creds = flow.credentials
            st.session_state["google_creds"] = creds

            # URLから code を削除
            st.query_params.clear()

            return creds

        except Exception as e:
            st.error(f"Google認証に失敗しました: {e}")
            return None

    return None


# ==============================
# イベント操作関数群
# ==============================

def add_event_to_calendar(service, calendar_id, event_data):
    try:
        return service.events().insert(calendarId=calendar_id, body=event_data).execute()
    except HttpError as e:
        st.error(f"イベント追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント追加失敗: {e}")
    return None


def fetch_all_events(service, calendar_id, time_min=None, time_max=None):
    events = []
    page_token = None
    try:
        while True:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                orderBy='startTime',
                pageToken=page_token
            ).execute()
            events.extend(events_result.get('items', []))
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        return events
    except HttpError as e:
        st.error(f"イベント取得失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント取得失敗: {e}")
    return []


def update_event_if_needed(service, calendar_id, event_id, new_event_data):
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

        def normalize(val):
            return val or ""

        needs_update = False

        if normalize(existing_event.get("summary")) != normalize(new_event_data.get("summary")):
            needs_update = True
        elif normalize(existing_event.get("description")) != normalize(new_event_data.get("description")):
            needs_update = True

        if not needs_update:
            if normalize(existing_event.get("transparency")) != normalize(new_event_data.get("transparency")):
                needs_update = True

        if not needs_update:
            if (existing_event.get("recurrence") or []) != (new_event_data.get("recurrence") or []):
                needs_update = True

        if not needs_update:
            if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
                needs_update = True

        if not needs_update:
            if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
                needs_update = True

        if needs_update:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=new_event_data
            ).execute()
            return updated_event

        return existing_event

    except HttpError as e:
        st.error(f"イベント更新失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント更新失敗: {e}")
    return None


def delete_event_from_calendar(service, calendar_id, event_id):
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        return True
    except HttpError as e:
        st.error(f"イベント削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント削除失敗: {e}")
    return False


# ==============================
# ToDoリスト操作関数群
# ==============================

def build_tasks_service(creds):
    try:
        if not creds:
            return None
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        st.warning(f"Google Tasks サービスのビルドに失敗しました: {e}")
        return None


def add_task_to_todo_list(tasks_service, task_list_id, task_data):
    try:
        return tasks_service.tasks().insert(tasklist=task_list_id, body=task_data).execute()
    except HttpError as e:
        st.error(f"タスク追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク追加失敗: {e}")
    return None


def find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id):
    try:
        tasks_result = tasks_service.tasks().list(tasklist=task_list_id).execute()
        tasks = tasks_result.get('items', [])
        deleted_count = 0
        for task in tasks:
            if (event_id in task.get('notes', '') or event_id in task.get('title', '')):
                tasks_service.tasks().delete(tasklist=task_list_id, task=task['id']).execute()
                deleted_count += 1
        return deleted_count
    except HttpError as e:
        st.error(f"タスク検索・削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク検索・削除失敗: {e}")
    return 0
