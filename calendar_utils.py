import logging
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from backoff import on_exception, expo
from datetime import datetime, timedelta
import re

logging.basicConfig(level=logging.INFO, filename="app.log")

def authenticate_google():
    try:
        import streamlit as st
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        import pickle
        import os

        SCOPES = [
            'https://www.googleapis.com/auth/calendar',
            'https://www.googleapis.com/auth/tasks'
        ]
        
        creds = None
        if 'google_auth' in st.session_state:
            creds = Credentials(**st.session_state['google_auth'])
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                st.session_state['google_auth'] = {
                    'token': creds.token,
                    'refresh_token': creds.refresh_token,
                    'token_uri': creds.token_uri,
                    'client_id': creds.client_id,
                    'client_secret': creds.client_secret,
                    'scopes': creds.scopes
                }
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            st.session_state['google_auth'] = {
                'token': creds.token,
                'refresh_token': creds.refresh_token,
                'token_uri': creds.token_uri,
                'client_id': creds.client_id,
                'client_secret': creds.client_secret,
                'scopes': creds.scopes
            }
        return creds
    except Exception as e:
        logging.error(f"Google認証中にエラーが発生: {e}")
        return None

def build_tasks_service(creds):
    try:
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        logging.error(f"Tasksサービス構築中にエラーが発生: {e}")
        return None

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def add_event_to_calendar(service, calendar_id, event_data):
    import streamlit as st
    try:
        return service.events().insert(calendarId=calendar_id, body=event_data).execute()
    except HttpError as e:
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in add_event_to_calendar: {e}")
        st.error(f"イベント追加に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"イベント追加中に予期しないエラーが発生: {e}")
        st.error(f"イベント追加に失敗しました: {str(e)}")
        raise

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def fetch_all_events(service, calendar_id, time_min, time_max):
    import streamlit as st
    try:
        events = []
        page_token = None
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
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in fetch_all_events: {e}")
        st.error(f"イベント取得に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"イベント取得中に予期しないエラーが発生: {e}")
        st.error(f"イベント取得に失敗しました: {str(e)}")
        raise

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def update_event_if_needed(service, calendar_id, event_id, event_data):
    import streamlit as st
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        needs_update = (
            existing_event.get('summary') != event_data.get('summary') or
            existing_event.get('description') != event_data.get('description') or
            existing_event.get('location') != event_data.get('location') or
            existing_event.get('start') != event_data.get('start') or
            existing_event.get('end') != event_data.get('end') or
            existing_event.get('transparency') != event_data.get('transparency')
        )
        if needs_update:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=event_data
            ).execute()
            return updated_event
        return None
    except HttpError as e:
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in update_event_if_needed: {e}")
        st.error(f"イベント更新に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"イベント更新中に予期しないエラーが発生: {e}")
        st.error(f"イベント更新に失敗しました: {str(e)}")
        raise

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def delete_event_from_calendar(service, calendar_id, event_id):
    import streamlit as st
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        return True
    except HttpError as e:
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in delete_event_from_calendar: {e}")
        st.error(f"イベント削除に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"イベント削除中に予期しないエラーが発生: {e}")
        st.error(f"イベント削除に失敗しました: {str(e)}")
        raise

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def add_task_to_todo_list(service, task_list_id, task_data):
    import streamlit as st
    try:
        return service.tasks().insert(tasklist=task_list_id, body=task_data).execute()
    except HttpError as e:
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in add_task_to_todo_list: {e}")
        st.error(f"ToDo追加に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"ToDo追加中に予期しないエラーが発生: {e}")
        st.error(f"ToDo追加に失敗しました: {str(e)}")
        raise

@on_exception(expo, HttpError, max_tries=5, giveup=lambda e: e.resp.status not in [429, 503])
def find_and_delete_tasks_by_event_id(service, task_list_id, event_id):
    import streamlit as st
    try:
        tasks = service.tasks().list(tasklist=task_list_id).execute()
        deleted_count = 0
        for task in tasks.get('items', []):
            if task.get('notes') and event_id in task['notes']:
                service.tasks().delete(tasklist=task_list_id, task=task['id']).execute()
                deleted_count += 1
        return deleted_count
    except HttpError as e:
        error_code = e.resp.status
        logging.error(f"HTTPエラー ({error_code}) in find_and_delete_tasks_by_event_id: {e}")
        st.error(f"ToDo削除に失敗しました: HTTPエラー {error_code}。詳細はログを確認してください。")
        raise
    except Exception as e:
        logging.exception(f"ToDo削除中に予期しないエラーが発生: {e}")
        st.error(f"ToDo削除に失敗しました: {str(e)}")
        raise