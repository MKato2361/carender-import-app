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

# ── キャッシュ対応イベント取得（セッションキャッシュ / TTL 5分）──
def _get_cache_key(calendar_id: str, time_min, time_max) -> str:
    return f"{calendar_id}::{time_min}::{time_max}"


def get_events(
    service,
    calendar_id: str,
    time_min: Optional[str] = None,
    time_max: Optional[str] = None,
    use_cache: bool = True,
) -> list[dict]:
    """
    イベントを全件取得する。失敗時は空リストを返す。
    use_cache=True の場合、5分間 session_state にキャッシュして API 呼び出しを削減する。
    """
    import streamlit as st

    cache_key = f"_ev_cache_{_get_cache_key(calendar_id, time_min, time_max)}"
    ts_key    = f"_ev_cache_ts_{_get_cache_key(calendar_id, time_min, time_max)}"

    if use_cache:
        import time as _time
        cached    = st.session_state.get(cache_key)
        cached_ts = st.session_state.get(ts_key, 0)
        if cached is not None and (_time.time() - cached_ts) < 300:  # 5分TTL
            return cached

    try:
        result = fetch_all_events(service, calendar_id, time_min, time_max)
        if use_cache:
            import time as _time
            st.session_state[cache_key] = result
            st.session_state[ts_key]    = _time.time()
        return result
    except HttpError as e:
        st.error(_http_error_msg(e, "イベントの取得"))
    except Exception as e:
        st.error(_generic_error_msg(e, "イベントの取得"))
    return []


def invalidate_events_cache(calendar_id: str = None) -> None:
    """
    イベントキャッシュを手動で無効化する。
    calendar_id を指定するとそのカレンダーのみ、None なら全て削除。
    登録・削除・更新完了後に呼ぶ。
    """
    import streamlit as st
    keys_to_del = [
        k for k in list(st.session_state.keys())
        if k.startswith("_ev_cache_")
        and (calendar_id is None or calendar_id in k)
    ]
    for k in keys_to_del:
        del st.session_state[k]


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
