import streamlit as st
import firebase_admin
from firebase_admin import auth, credentials, firestore
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
import os

# --- Firebase Initialization ---
def init_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate("firebase_credentials.json")
        firebase_admin.initialize_app(cred)

# --- Google OAuth 2.0 Flow ---
def get_google_auth_url():
    """Generates the Google OAuth 2.0 authorization URL."""
    init_firebase()
    
    # We need to get the redirect URI, which is the current app URL.
    if st.secrets.get("google_oauth", {}).get("redirect_uri"):
        redirect_uri = st.secrets.google_oauth.redirect_uri
    else:
        st.error("Please configure `redirect_uri` in your Streamlit secrets.")
        return "#"

    flow = Flow.from_client_config(
        client_config={
            "web": {
                "client_id": st.secrets.google_oauth.client_id,
                "project_id": st.secrets.firebase.project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets.google_oauth.client_secret,
                "redirect_uris": [redirect_uri],
            }
        },
        scopes=st.secrets.google_oauth.scopes,
        redirect_uri=redirect_uri,
    )

    authorization_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
    )
    return authorization_url

def handle_google_auth_callback(auth_code):
    """Exchanges the authorization code for tokens and handles user creation/login."""
    init_firebase()
    
    # We need to get the redirect URI, which is the current app URL.
    redirect_uri = st.secrets.google_oauth.redirect_uri

    flow = Flow.from_client_config(
        client_config={
            "web": {
                "client_id": st.secrets.google_oauth.client_id,
                "project_id": st.secrets.firebase.project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets.google_oauth.client_secret,
                "redirect_uris": [redirect_uri],
            }
        },
        scopes=st.secrets.google_oauth.scopes,
        redirect_uri=redirect_uri,
    )
    
    flow.fetch_token(code=auth_code)
    
    creds = flow.credentials
    user_info = auth.get_user(creds.id_token) # Get user info from id_token

    db = firestore.client()
    doc_ref = db.collection("google_tokens").document(user_info.uid)
    doc_ref.set({
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
        "expires_at": creds.expiry.isoformat(),
    })
    
    google_auth_info = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
        "expiry": creds.expiry,
    }
    
    return user_info, google_auth_info
    

# ==============================
# 以下、元の関数群（変更なし）
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
    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        return events_result.get('items', [])
    except HttpError as e:
        st.error(f"イベント取得失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント取得失敗: {e}")
    return []

def update_event_if_needed(service, calendar_id, event_id, updated_event_data):
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        needs_update = any(existing_event.get(k) != v for k, v in updated_event_data.items())
        if needs_update:
            return service.events().update(calendarId=calendar_id, eventId=event_id, body=updated_event_data).execute()
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
