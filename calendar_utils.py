import os
import json
import pickle
from pathlib import Path
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
# Google 認証（Webリダイレクト型 + トークン自動削除）
# ==============================
def authenticate_google():

    if "credentials" not in st.session_state:
        st.session_state["credentials"] = None

    if "credentials_user_id" not in st.session_state:
        st.session_state["credentials_user_id"] = None

    # -------------------------------
    # 強制リセット
    # -------------------------------
    if st.query_params.get("clear_auth") == "1":
        user_id = get_firebase_user_id()
        if user_id:
            try:
                firestore.client().collection('google_tokens').document(user_id).delete()
            except Exception:
                pass

        st.session_state.pop('credentials', None)
        st.session_state.pop('credentials_user_id', None)
        st.session_state.pop('oauth_state', None)

        st.query_params.clear()
        st.rerun()

    user_id = get_firebase_user_id()
    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    creds = None

    # ============================================================
    # ① セッションから取得（user_id一致チェック）
    # ============================================================
    if (
        hasattr(st, "session_state") and
        "credentials" in st.session_state and
        st.session_state["credentials"] and
        st.session_state.get("credentials_user_id") == user_id
    ):
        creds = st.session_state.get('credentials')

        if creds:
            if not creds.refresh_token:
                st.session_state.pop('credentials', None)
                st.session_state.pop('credentials_user_id', None)
                try:
                    doc_ref.delete()
                except Exception:
                    pass
                creds = None

            elif creds.valid:
                return creds

            elif creds.expired:
                try:
                    creds.refresh(Request())
                    st.session_state['credentials'] = creds
                    st.session_state['credentials_user_id'] = user_id
                    doc_ref.set(json.loads(creds.to_json()))
                    return creds
                except Exception:
                    st.session_state.pop('credentials', None)
                    st.session_state.pop('credentials_user_id', None)
                    try:
                        doc_ref.delete()
                    except Exception:
                        pass
                    creds = None

    # ============================================================
    # ② Firestoreから取得
    # ============================================================
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()

            try:
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            except Exception:
                creds = None

            if creds:
                if not creds.refresh_token:
                    doc_ref.delete()
                    creds = None

                elif creds.expired:
                    try:
                        creds.refresh(Request())
                        st.session_state['credentials'] = creds
                        st.session_state['credentials_user_id'] = user_id
                        doc_ref.set(json.loads(creds.to_json()))
                        return creds
                    except Exception:
                        doc_ref.delete()
                        creds = None

                elif creds.valid:
                    st.session_state['credentials'] = creds
                    st.session_state['credentials_user_id'] = user_id
                    return creds

    except Exception:
        pass

    # ============================================================
    # ③ OAuthフロー
    # ============================================================
    try:
        client_id     = st.secrets["google"]["client_id"]
        client_secret = st.secrets["google"]["client_secret"]
        redirect_uri  = st.secrets["google"]["redirect_uri"]

        AUTH_URI  = "https://accounts.google.com/o/oauth2/auth"
        TOKEN_URI = "https://oauth2.googleapis.com/token"

        params = st.query_params

        # -------------------------------
        # 認証URL生成
        # -------------------------------
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

            st.query_params["oauth_state"] = state
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()

        # -------------------------------
        # コールバック処理
        # -------------------------------
        else:
            state = st.query_params.get("oauth_state")
        if not state:
            st.warning("セッションが切れました。再度ログインしてください。")
            st.session_state.pop("credentials", None)
            st.session_state.pop("credentials_user_id", None)

            st.query_params.clear()
            st.stop()
            
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
                token=token.get("access_token"),
                refresh_token=token.get("refresh_token"),
                token_uri=TOKEN_URI,
                client_id=client_id,
                client_secret=client_secret,
                scopes=SCOPES,
            )

            if not creds or not creds.refresh_token:
                st.error("認証に失敗しました（トークン不正）")
                return None

            st.session_state["credentials"] = creds
            st.session_state["credentials_user_id"] = user_id

            doc_ref.set(json.loads(creds.to_json()))

            st.success("Google認証が完了しました！")

            st.query_params.clear()
            st.session_state.pop("oauth_state", None)

            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state.pop('credentials', None)
        st.session_state.pop('credentials_user_id', None)
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
    """イベント全件取得（ページネーション対応）"""
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
    """
    既存イベントと new_event_data を比較し、差分がある場合のみ Google Calendar を更新する。

    比較対象:
      - summary（タイトル）
      - description（説明）
      - start（終日/時間指定/タイムゾーン含め厳密比較）
      - end（終日/時間指定/タイムゾーン含め厳密比較）
      - transparency（公開/非公開設定）
      - recurrence（繰り返し設定があれば）

    ※ Location（場所）は比較対象外
    ※ 差分がある場合のみ update API を実行
    """
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

        def normalize(val):
            return val or ""

        needs_update = False

        # 1) summary
        if normalize(existing_event.get("summary")) != normalize(new_event_data.get("summary")):
            needs_update = True

        # 2) description
        if not needs_update:
            if normalize(existing_event.get("description")) != normalize(new_event_data.get("description")):
                needs_update = True

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

        if needs_update:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=new_event_data
            ).execute()
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
