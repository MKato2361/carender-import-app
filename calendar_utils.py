from ui.components import handle_http_error as _handle_http_error
"""
calendar_utils.py — 後方互換ラッパー

実体は以下に移行済み:
  core/auth/google_oauth.py     … トークン管理ロジック
  services/auth_service.py      … authenticate_google (UI 付き)
  core/calendar/crud.py         … fetch_all_events / add_event / update_event / delete_event
  core/calendar/tasks.py        … build_tasks_service / add_task / find_and_delete_tasks_by_event_id

既存の import が壊れないようエイリアスを公開する。
新規コードでは直接 core / services を import すること。
"""
from __future__ import annotations
from typing import Optional

import streamlit as st
from googleapiclient.errors import HttpError

# ── ロジック層 ──
from services.auth_service import authenticate_google          # noqa: F401
from core.calendar.crud  import (
    fetch_all_events,                                          # noqa: F401
    add_event         as _add_event,
    update_event_if_changed as _update_event_if_changed,
    delete_event      as _delete_event,
)
from core.calendar.tasks import (
    build_tasks_service,                                       # noqa: F401
    add_task          as _add_task,
    find_and_delete_tasks_by_event_id,                        # noqa: F401
)



# ── UI 向けラッパー（エラーを st.error で表示して None を返す） ──

def _handle_generic_error(e: Exception, action: str) -> None:
    """HttpError 以外の例外を種別判定してユーザー向けメッセージで表示する。"""
    msg = str(e).lower()
    if "invalid_grant" in msg or "token has been expired or revoked" in msg:
        # セッションとFirestoreのトークンを削除して再認証を促す
        try:
            user_id = st.session_state.get("user_info")
            if user_id:
                from core.auth.google_oauth import _clear_creds
                _clear_creds(user_id)
        except Exception:
            pass
        st.error(
            f"{action}に失敗しました。Googleアカウントの連携が切れています。"
            "ページを再読み込みして再連携してください。"
        )
    else:
        st.error(f"{action}に失敗しました。しばらく待ってから再試行してください。")


def add_event_to_calendar(service, calendar_id: str,
                           event_data: dict) -> Optional[dict]:
    try:
        return _add_event(service, calendar_id, event_data)
    except HttpError as e:
        _handle_http_error(e, "イベントの追加")
    except Exception as e:
        _handle_generic_error(e, "イベントの追加")
    return None


def update_event_if_needed(service, calendar_id: str,
                            event_id: str, new_event_data: dict) -> Optional[dict]:
    try:
        return _update_event_if_changed(service, calendar_id, event_id, new_event_data)
    except HttpError as e:
        _handle_http_error(e, "イベントの更新")
    except Exception as e:
        _handle_generic_error(e, "イベントの更新")
    return None


def delete_event_from_calendar(service, calendar_id: str,
                                event_id: str) -> bool:
    try:
        _delete_event(service, calendar_id, event_id)
        return True
    except HttpError as e:
        _handle_http_error(e, "イベントの削除")
    except Exception as e:
        _handle_generic_error(e, "イベントの削除")
    return False


def add_task_to_todo_list(tasks_service, task_list_id: str,
                          task_data: dict) -> Optional[dict]:
    try:
        return _add_task(tasks_service, task_list_id, task_data)
    except HttpError as e:
        _handle_http_error(e, "タスクの追加")
    except Exception as e:
        _handle_generic_error(e, "タスクの追加")
    return None
