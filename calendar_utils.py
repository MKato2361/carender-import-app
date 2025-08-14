import os
import json
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from firebase_admin import firestore
from firebase_auth import get_firebase_user_id
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta, timezone

SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/tasks"]

def _get_flow():
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
    return flow

def _refresh_and_save_creds(creds, doc_ref):
    try:
        creds.refresh(Request())
        st.session_state['creds'] = creds
        doc_ref.set(json.loads(creds.to_json()))
        return creds
    except Exception as e:
        st.warning(f"トークンの更新に失敗しました: {e}")
        doc_ref.delete()
        st.session_state.pop('creds', None)
        return None

def authenticate_google():
    user_id = get_firebase_user_id()
    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # セッションからクレデンシャルをロード
    if 'creds' in st.session_state and st.session_state['creds']:
        creds = st.session_state['creds']
        if creds.valid:
            return creds
        if creds.expired and creds.refresh_token:
            return _refresh_and_save_creds(creds, doc_ref)

    # Firestoreからクレデンシャルをロード
    try:
        doc = doc_ref.get()
        if doc.exists:
            creds_data = doc.to_dict()
            if creds_data:
                creds = Credentials.from_authorized_user_info(creds_data, SCOPES)
                st.session_state['creds'] = creds
                if creds.expired and creds.refresh_token:
                    return _refresh_and_save_creds(creds, doc_ref)
                return creds
    except Exception as e:
        if "invalid_grant" in str(e):
            st.warning("保存されたGoogleトークンが無効化されました。再認証します。")
            doc_ref.delete()
            st.session_state.pop('creds', None)
        else:
            st.error(f"Firestoreからトークン取得に失敗: {e}")

    # OAuthフロー開始
    flow = _get_flow()
    params = st.query_params

    if "code" not in params:
        auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline', include_granted_scopes='true')
        st.markdown(f"[Googleでログインする]({auth_url})", unsafe_allow_html=True)
        st.stop()
    else:
        try:
            flow.fetch_token(code=params["code"])
            creds = flow.credentials
            st.session_state['creds'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            st.success("Google認証が完了しました！")
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"認証中にエラーが発生しました: {e}")
            st.stop()
    return None

def build_tasks_service(creds):
    if not creds:
        return None
    try:
        return build("tasks", "v1", credentials=creds)
    except Exception as e:
        st.warning(f"Google Tasks サービスのビルドに失敗しました: {e}")
        return None

def add_event_to_calendar(service, calendar_id, event_data, timeout=30):
    try:
        event = {
            'summary': event_data['Subject'],
            'start': {
                'dateTime': datetime.strptime(f"{event_data['Start Date']} {event_data['Start Time']}", "%Y/%m/%d %H:%M:%S").isoformat(),
                'timeZone': 'Asia/Tokyo',
            },
            'end': {
                'dateTime': datetime.strptime(f"{event_data['End Date']} {event_data['End Time']}", "%Y/%m/%d %H:%M:%S").isoformat(),
                'timeZone': 'Asia/Tokyo',
            },
            'description': event_data.get('Description', ''),
            'location': event_data.get('Location', ''),
            'visibility': 'private' if event_data.get('Private', 'False') == "True" else 'default'
        }
        if event_data.get('All Day Event') == "True":
            event['start'] = {'date': event_data['Start Date'].replace("/", "-")}
            event['end'] = {'date': event_data['End Date'].replace("/", "-")}

        created_event = service.events().insert(calendarId=calendar_id, body=event, timeout=timeout).execute()
        return created_event.get('id')
    except HttpError as e:
        st.error(f"イベント追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント追加失敗: {e}")
    return None

def update_event_if_needed(service, calendar_id, event_id, new_event_data, timeout=30):
    try:
        old_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        
        # タイムゾーンを考慮した日時比較
        old_start = old_event['start'].get('dateTime', old_event['start'].get('date'))
        old_end = old_event['end'].get('dateTime', old_event['end'].get('date'))
        
        # 新しいイベントデータの整形
        new_start_dict = new_event_data['start']
        new_end_dict = new_event_data['end']
        new_start = new_start_dict.get('dateTime', new_start_dict.get('date'))
        new_end = new_end_dict.get('dateTime', new_end_dict.get('date'))
        
        is_updated = False
        if old_event.get('summary') != new_event_data.get('summary'):
            old_event['summary'] = new_event_data['summary']
            is_updated = True
        if old_event.get('description') != new_event_data.get('description'):
            old_event['description'] = new_event_data['description']
            is_updated = True
        if old_event.get('location') != new_event_data.get('location'):
            old_event['location'] = new_event_data['location']
            is_updated = True
        if str(old_start) != str(new_start) or str(old_end) != str(new_end):
            old_event['start'] = new_event_data['start']
            old_event['end'] = new_event_data['end']
            is_updated = True
        
        if is_updated:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=old_event,
                timeout=timeout
            ).execute()
            return True
        return False
    except HttpError as e:
        st.error(f"イベント更新失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント更新失敗: {e}")
    return False

def delete_event_from_calendar(service, calendar_id, event_id, timeout=30):
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id, timeout=timeout).execute()
        return True
    except HttpError as e:
        if e.resp.status == 404:
            st.warning(f"イベントID '{event_id}' はすでに存在しません。")
            return True
        st.error(f"イベント削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント削除失敗: {e}")
    return False

def fetch_all_events(service, calendar_id, time_min, max_results=2500):
    all_events = []
    page_token = None
    try:
        while True:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min.isoformat(),
                maxResults=max_results,
                singleEvents=True,
                orderBy='startTime',
                pageToken=page_token
            ).execute()
            all_events.extend(events_result.get('items', []))
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
    except HttpError as e:
        st.error(f"イベントの取得に失敗しました (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベントの取得に失敗しました: {e}")
    return all_events

def add_task_to_todo_list(tasks_service, task_list_id, title, deadline, notes):
    try:
        task_data = {
            'title': title,
            'notes': notes,
            'due': deadline
        }
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
            if event_id in task.get('notes', ''):
                tasks_service.tasks().delete(tasklist=task_list_id, task=task['id']).execute()
                deleted_count += 1
        return deleted_count
    except HttpError as e:
        st.error(f"タスク削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク削除失敗: {e}")
    return 0
