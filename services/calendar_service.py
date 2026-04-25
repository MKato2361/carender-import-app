from __future__ import annotations
"""
services/calendar_service.py
カレンダー・タスク操作のサービス層

tabs から直接呼ぶ唯一の入口。
- エラーを st.* で表示し、呼び出し元はリターン値だけ見ればよい
- ロジックは core/calendar/ に委譲
- calendar_utils.py の後継として全 tabs を差し替える
"""
import logging
from typing import Optional
import streamlit as st
from googleapiclient.errors import HttpError

from core.calendar.crud import (
    fetch_all_events,
    add_event,
    update_event_if_changed,
    delete_event,
)
from core.calendar.tasks import (
    build_tasks_service,
    get_default_task_list_id,
    add_task,
    find_and_delete_tasks_by_event_id,
)

logger = logging.getLogger(__name__)


# ── エラーハンドラ ──────────────────────────────────────────

def _http_error_msg(e: HttpError, action: str) -> str:
    try:
        status = e.resp.status
    except Exception:
        status = None
    msgs = {
        401: f"{action}に失敗しました。Googleセッションが切れています。ページを再読み込みして再連携してください。",
        403: f"{action}に失敗しました。このカレンダーへの書き込み権限がありません。",
        404: f"{action}に失敗しました。対象のイベントが見つかりません（すでに削除済みの可能性があります）。",
        429: f"{action}に失敗しました。APIのリクエスト上限に達しました。しばらく待ってから再試行してください。",
    }
    return msgs.get(status, f"{action}に失敗しました（エラーコード: {status}）。しばらく待ってから再試行してください。")


def _generic_error_msg(e: Exception, action: str) -> str:
    msg = str(e).lower()
    if "invalid_grant" in msg or "token has been expired" in msg:
        # セッションのトークンをクリアして再認証を促す
        try:
            user_id = st.session_state.get("user_info")
            if user_id:
                from core.auth.google_oauth import _clear_creds
                _clear_creds(user_id)
        except Exception:
            pass
        return f"{action}に失敗しました。Googleアカウントの連携が切れています。ページを再読み込みして再連携してください。"
    return f"{action}に失敗しました。しばらく待ってから再試行してください。"


# ── イベント CRUD ───────────────────────────────────────────

def get_events(
    service,
    calendar_id: str,
    time_min: Optional[str] = None,
    time_max: Optional[str] = None,
) -> list[dict]:
    """イベントを全件取得する。失敗時は空リストを返す。"""
    try:
        return fetch_all_events(service, calendar_id, time_min, time_max)
    except HttpError as e:
        st.error(_http_error_msg(e, "イベントの取得"))
    except Exception as e:
        st.error(_generic_error_msg(e, "イベントの取得"))
    return []


def add_event_to_calendar(
    service, calendar_id: str, event_data: dict
) -> Optional[dict]:
    """イベントを追加する。失敗時は None を返す。"""
    try:
        return add_event(service, calendar_id, event_data)
    except HttpError as e:
        st.error(_http_error_msg(e, "イベントの追加"))
    except Exception as e:
        st.error(_generic_error_msg(e, "イベントの追加"))
    return None


def update_event_if_needed(
    service, calendar_id: str, event_id: str, new_data: dict
) -> Optional[dict]:
    """差分があればイベントを更新する。失敗時は None を返す。"""
    try:
        return update_event_if_changed(service, calendar_id, event_id, new_data)
    except HttpError as e:
        st.error(_http_error_msg(e, "イベントの更新"))
    except Exception as e:
        st.error(_generic_error_msg(e, "イベントの更新"))
    return None


def delete_event_from_calendar(
    service, calendar_id: str, event_id: str
) -> bool:
    """イベントを削除する。成功時 True、失敗時 False を返す。"""
    try:
        delete_event(service, calendar_id, event_id)
        return True
    except HttpError as e:
        st.error(_http_error_msg(e, "イベントの削除"))
    except Exception as e:
        st.error(_generic_error_msg(e, "イベントの削除"))
    return False


# ── タスク操作 ──────────────────────────────────────────────

def add_task_to_todo_list(
    tasks_service, task_list_id: str, task_data: dict
) -> Optional[dict]:
    """タスクを追加する。失敗時は None を返す。"""
    try:
        return add_task(tasks_service, task_list_id, task_data)
    except HttpError as e:
        st.error(_http_error_msg(e, "タスクの追加"))
    except Exception as e:
        st.error(_generic_error_msg(e, "タスクの追加"))
    return None


def delete_tasks_by_event_id(
    tasks_service, task_list_id: str, event_id: str
) -> int:
    """event_id に紐づくタスクを削除する。削除件数を返す。失敗時は 0。"""
    try:
        return find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id)
    except HttpError as e:
        st.error(_http_error_msg(e, "タスクの削除"))
    except Exception as e:
        st.error(_generic_error_msg(e, "タスクの削除"))
    return 0


# ── Tasks サービス構築 ──────────────────────────────────────

def init_tasks_service(creds):
    """Tasks API サービスを構築して返す。失敗時は None。"""
    try:
        svc = build_tasks_service(creds)
        return svc, get_default_task_list_id(svc)
    except Exception as e:
        logger.warning("Tasks サービス構築失敗: %s", e)
        return None, None
