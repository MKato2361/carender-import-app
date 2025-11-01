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
    "https://www.googleapis.com/auth/tasks"
]

# ==============================
# Google 認証（Webリダイレクト型 + トークン自動削除）
# ==============================
def authenticate_google():
    creds = None
    user_id = get_firebase_user_id()

    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # --- セッションから ---
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds
        elif creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                doc_ref.set(json.loads(creds.to_json()))
                return creds
            except Exception as e:
                st.warning(f"リフレッシュトークンの更新に失敗: {e}")
                doc_ref.delete()
                st.session_state.pop('credentials', None)
                return authenticate_google()

    # --- Firestoreから ---
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            st.session_state['credentials'] = creds

            if creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    st.session_state['credentials'] = creds
                    doc_ref.set(json.loads(creds.to_json()))
                    st.info("Google認証トークンを更新しました。")
                    st.rerun()
                except Exception as e:
                    st.warning(f"Firestoreトークンの更新に失敗: {e}")
                    doc_ref.delete()
                    st.session_state.pop('credentials', None)
                    return authenticate_google()

            return creds
    except Exception as e:
        if "invalid_grant" in str(e):
            st.warning("保存されたGoogleトークンが無効化されました。再認証します。")
            doc_ref.delete()
            st.session_state.pop('credentials', None)
            return authenticate_google()
        else:
            st.error(f"Firestoreからトークン取得に失敗: {e}")
            creds = None

    # --- 新しいOAuthフロー（Webリダイレクト型） ---
    try:
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

        params = st.query_params
        if "code" not in params:
            auth_url, _ = flow.authorization_url(
                prompt='consent',
                access_type='offline',
                include_granted_scopes='true'
            )
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()
        else:
            code = params["code"]
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.session_state['credentials'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            st.success("Google認証が完了しました！")
            st.query_params.clear()
            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state['credentials'] = None
        return None

    return creds

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

# 💡 修正点: イベント全件取得のためのページネーションを追加
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
                pageToken=page_token  # ページトークンを指定
            ).execute()
            events.extend(events_result.get('items', []))
            page_token = events_result.get('nextPageToken') # 次のページトークンを取得
            if not page_token:
                break # トークンがなければ終了
        return events
    except HttpError as e:
        st.error(f"イベント取得失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント取得失敗: {e}")
    return []

def update_event_if_needed(service, calendar_id, event_id, new_event_data):
    """
    既存イベントと new_event_data を比較し、差分がある場合のみ Google Calendar を更新する。

    比較対象:
      - summary（タイトル）
      - description（説明）
      - start（終日/時間指定/タイムゾーン含め厳密比較）
      - end（終日/時間指定/タイムゾーン含め厳密比較）
      - transparency（公開/非公開設定）
      - recurrence（繰り返し設定があれば）

    ※ Location（場所）は比較対象外（あなたの要望）
    ※ 差分がある場合のみ update API を実行し、更新ログを st.info で表示
    """
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

        # ===== 差分判定 =====
        def normalize(val):
            return val or ""

        # 1) summary
        if normalize(existing_event.get("summary")) != normalize(new_event_data.get("summary")):
            needs_update = True
        # 2) description
        elif normalize(existing_event.get("description")) != normalize(new_event_data.get("description")):
            needs_update = True
        else:
            needs_update = False

        # 3) transparency（非公開/公開）
        if not needs_update:
            if normalize(existing_event.get("transparency")) != normalize(new_event_data.get("transparency")):
                needs_update = True

        # 4) recurrence（繰り返し設定）
        if not needs_update:
            existing_recur = existing_event.get("recurrence") or []
            new_recur = new_event_data.get("recurrence") or []
            if existing_recur != new_recur:
                needs_update = True

        # 5) start
        if not needs_update:
            if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
                needs_update = True

        # 6) end
        if not needs_update:
            if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
                needs_update = True

        # ===== 更新実行 or スキップ =====
        if needs_update:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=new_event_data
            ).execute()

            # --- 更新ログ（簡潔型） ---
            summary = new_event_data.get("summary") or "(無題)"

            return updated_event

        # 差分なし → 更新不要
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
