
import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import re

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]

def authenticate_google():
    creds = None

    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                st.rerun()
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました: {e}")
                st.session_state['credentials'] = None
                creds = None
        else:
            try:
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"]
                    }
                }
                flow = Flow.from_client_config(client_config, SCOPES)
                flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
                auth_url, _ = flow.authorization_url(prompt='consent')

                st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
                st.write(auth_url)
                code = st.text_input("認証コードを貼り付けてください:")

                if code:
                    flow.fetch_token(code=code)
                    creds = flow.credentials
                    st.session_state['credentials'] = creds
                    st.success("Google認証が完了しました！")
                    st.rerun()
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['credentials'] = None
                return None

    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event.get("htmlLink")

def delete_events_from_calendar(service, calendar_id, start_date: datetime, end_date: datetime):
    JST_OFFSET = timedelta(hours=9)

    start_dt_jst = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt_jst = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    time_min_utc = (start_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'
    time_max_utc = (end_dt_jst - JST_OFFSET).isoformat(timespec='microseconds') + 'Z'

    deleted_count = 0
    all_events_to_delete = []
    page_token = None

    with st.spinner(f"{start_date.strftime('%Y/%m/%d')}から{end_date.strftime('%Y/%m/%d')}までの削除対象イベントを検索中..."):
        while True:
            try:
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min_utc,
                    timeMax=time_max_utc,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                ).execute()
                events = events_result.get('items', [])
                all_events_to_delete.extend(events)

                page_token = events_result.get('nextPageToken')
                if not page_token:
                    break
            except Exception as e:
                st.error(f"イベントの検索中にエラーが発生しました: {e}")
                return 0

    total_events = len(all_events_to_delete)

    if total_events == 0:
        return 0

    progress_bar = st.progress(0)

    for i, event in enumerate(all_events_to_delete):
        event_summary = event.get('summary', '不明なイベント')
        try:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            deleted_count += 1
        except Exception as e:
            st.warning(f"イベント '{event_summary}' の削除に失敗しました: {e}")

        progress_bar.progress((i + 1) / total_events)

    return deleted_count

def fetch_all_events(service, calendar_id, time_min, time_max):
    events = []
    page_token = None
    while True:
        result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
            pageToken=page_token
        ).execute()
        events.extend(result.get('items', []))
        page_token = result.get('nextPageToken')
        if not page_token:
            break
    return events

def update_event_if_needed(service, calendar_id, event, new_event_data):
    updated = False

    if 'date' in event['start'] and 'date' in new_event_data['start']:
        if event['start']['date'] != new_event_data['start']['date'] or event['end']['date'] != new_event_data['end']['date']:
            event['start']['date'] = new_event_data['start']['date']
            event['end']['date'] = new_event_data['end']['date']
            updated = True
    elif 'dateTime' in event['start'] and 'dateTime' in new_event_data['start']:
        if event['start']['dateTime'] != new_event_data['start']['dateTime'] or event['end']['dateTime'] != new_event_data['end']['dateTime']:
            event['start']['dateTime'] = new_event_data['start']['dateTime']
            event['end']['dateTime'] = new_event_data['end']['dateTime']
            updated = True

    if updated:
        service.events().update(calendarId=calendar_id, eventId=event['id'], body=event).execute()
    return updated

def create_tasks_for_event(service, task_service, title, due_datetime):
    task_titles = [
        f"{title} - 点検通知（FAX）",
        f"{title} - 点検通知（電話）",
        f"{title} - 貼紙"
    ]
    task_list_id = task_service.tasklists().list().execute()["items"][0]["id"]
    for task_title in task_titles:
        task = {
            'title': task_title,
            'due': due_datetime.isoformat() + 'Z'
        }
        task_service.tasks().insert(tasklist=task_list_id, body=task).execute()
