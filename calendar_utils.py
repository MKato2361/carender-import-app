import os
import json
import time
import random
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from requests_oauthlib import OAuth2Session
from firebase_admin import firestore
from firebase_auth import get_firebase_user_id
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta, timezone

# requests_oauthlib の PKCE を強制無効化
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = "1"

# Google API スコープ
SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks",
    "https://www.googleapis.com/auth/spreadsheets"
]

# ==============================
# リトライ設定
# ==============================
_RETRYABLE_STATUS = {403, 429, 500, 502, 503, 504}
_MAX_RETRIES = 5
_BACKOFF_BASE = 2.0   # 秒
_BACKOFF_MAX  = 64.0  # 秒（上限）


def _is_rate_limit_error(e: HttpError) -> bool:
    """403 rateLimitExceeded / 429 Too Many Requests を判定する"""
    if e.resp.status not in _RETRYABLE_STATUS:
        return False
    # 403 の場合は reason を確認（forbidden と区別）
    if e.resp.status == 403:
        try:
            details = json.loads(e.content).get("error", {}).get("errors", [])
            return any(
                d.get("reason") in ("rateLimitExceeded", "userRateLimitExceeded")
                for d in details
            )
        except Exception:
            return False
    return True  # 429, 5xx は無条件リトライ


def _call_with_retry(api_call, *args, **kwargs):
    """
    指数バックオフ付きリトライでGoogle API呼び出しを実行する。

    - リトライ対象: 403 rateLimitExceeded / 429 / 5xx
    - 非リトライ対象: 404, 401 など → そのまま raise
    - 最大リトライ回数: _MAX_RETRIES
    """
    last_error = None
    for attempt in range(_MAX_RETRIES + 1):
        try:
            return api_call(*args, **kwargs)
        except HttpError as e:
            if not _is_rate_limit_error(e):
                raise  # リトライ対象外はそのまま上に投げる
            last_error = e
            if attempt == _MAX_RETRIES:
                break
            wait = min(_BACKOFF_BASE ** attempt + random.uniform(0, 1), _BACKOFF_MAX)
            st.warning(f"レートリミット到達。{wait:.1f} 秒後にリトライします… (試行 {attempt + 1}/{_MAX_RETRIES})")
            time.sleep(wait)

    raise last_error  # 全リトライ失敗


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

    # --- 新しいOAuthフロー（OAuth2Session直接使用・PKCE完全無効） ---
    try:
        client_id     = st.secrets["google"]["client_id"]
        client_secret = st.secrets["google"]["client_secret"]
        redirect_uri  = st.secrets["google"]["redirect_uri"]

        AUTH_URI  = "https://accounts.google.com/o/oauth2/auth"
        TOKEN_URI = "https://oauth2.googleapis.com/token"

        params = st.query_params

        if "code" not in params:
            oauth = OAuth2Session(
                client_id=client_id,
                redirect_uri=redirect_uri,
                scope=SCOPES,
            )
            auth_url, state = oauth.authorization_url(
                AUTH_URI,
                access_type="offline",
                prompt="consent",
                include_granted_scopes="true",
            )
            st.session_state["oauth_state"] = state
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()

        else:
            state = st.session_state.get("oauth_state", "")
            oauth = OAuth2Session(
                client_id=client_id,
                redirect_uri=redirect_uri,
                scope=SCOPES,
                state=state,
            )
            current_url = redirect_uri + "?" + "&".join(
                f"{k}={v}" for k, v in params.items()
            )
            token = oauth.fetch_token(
                TOKEN_URI,
                authorization_response=current_url,
                client_secret=client_secret,
            )
            creds = Credentials(
                token=token["access_token"],
                refresh_token=token.get("refresh_token"),
                token_uri=TOKEN_URI,
                client_id=client_id,
                client_secret=client_secret,
                scopes=SCOPES,
            )
            st.session_state["credentials"] = creds
            doc_ref.set(json.loads(creds.to_json()))
            st.success("Google認証が完了しました！")
            st.query_params.clear()
            st.session_state.pop("oauth_state", None)
            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state["credentials"] = None
        return None

    return creds


# ==============================
# イベント操作関数群
# ==============================

def add_event_to_calendar(service, calendar_id, event_data):
    try:
        return _call_with_retry(
            lambda: service.events().insert(calendarId=calendar_id, body=event_data).execute()
        )
    except HttpError as e:
        st.error(f"イベント追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント追加失敗: {e}")
    return None


def fetch_all_events(service, calendar_id, time_min=None, time_max=None):
    """イベント全件取得（ページネーション＋リトライ対応）"""
    events = []
    page_token = None
    try:
        while True:
            result = _call_with_retry(
                lambda pt=page_token: service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=True,
                    orderBy="startTime",
                    pageToken=pt,
                ).execute()
            )
            events.extend(result.get("items", []))
            page_token = result.get("nextPageToken")
            if not page_token:
                break
        return events
    except HttpError as e:
        st.error(f"イベント取得失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント取得失敗: {e}")
    return []


def update_event_if_needed(service, calendar_id, event_id, new_event_data):
    """
    既存イベントと new_event_data を比較し、差分がある場合のみ更新する。
    API呼び出しはレートリミット対応のリトライ付き。
    """
    try:
        existing_event = _call_with_retry(
            lambda: service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        )

        nz = lambda v: v or ""
        needs_update = False

        for field in ("summary", "description", "transparency"):
            if nz(existing_event.get(field)) != nz(new_event_data.get(field)):
                needs_update = True
                break

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
            return _call_with_retry(
                lambda: service.events().update(
                    calendarId=calendar_id,
                    eventId=event_id,
                    body=new_event_data,
                ).execute()
            )

        return existing_event

    except HttpError as e:
        st.error(f"イベント更新失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント更新失敗: {e}")
    return None


def delete_event_from_calendar(service, calendar_id, event_id):
    try:
        _call_with_retry(
            lambda: service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        )
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
        return _call_with_retry(
            lambda: tasks_service.tasks().insert(tasklist=task_list_id, body=task_data).execute()
        )
    except HttpError as e:
        st.error(f"タスク追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク追加失敗: {e}")
    return None


def find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id):
    try:
        tasks_result = _call_with_retry(
            lambda: tasks_service.tasks().list(tasklist=task_list_id).execute()
        )
        tasks = tasks_result.get('items', [])
        deleted_count = 0
        for task in tasks:
            if event_id in task.get('notes', '') or event_id in task.get('title', ''):
                _call_with_retry(
                    lambda tid=task['id']: tasks_service.tasks().delete(
                        tasklist=task_list_id, task=tid
                    ).execute()
                )
                deleted_count += 1
        return deleted_count
    except HttpError as e:
        st.error(f"タスク検索・削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク検索・削除失敗: {e}")
    return 0
