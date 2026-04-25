from __future__ import annotations
"""
services/auth_service.py
Google 認証フロー制御（st.* 表示 + セッション管理）

auth_manager.py の ensure_google_services / authenticate_google の UI 部分を担う。
ロジックは core/auth/google_oauth.py に委譲。
"""
from typing import Optional

import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from core.auth.google_oauth import (
    get_valid_credentials,
    build_auth_url,
    handle_oauth_callback,
)
from core.calendar.tasks import build_tasks_service, get_default_task_list_id
from core.auth.firebase_client import get_user_id as get_firebase_user_id


def authenticate_google() -> Optional[Credentials]:
    """
    Google OAuth 認証を実行し、有効な Credentials を返す。
    認証が必要な場合は st.link_button を表示して st.stop() する。
    calendar_utils.authenticate_google の後継。
    """
    # ?clear_auth=1 による強制リセット
    if st.query_params.get("clear_auth") == "1":
        user_id = get_firebase_user_id()
        from core.auth.google_oauth import _clear_creds
        _clear_creds(user_id)
        st.query_params.clear()
        st.rerun()

    user_id = get_firebase_user_id()
    if not user_id:
        return None

    # コールバック処理（?code= がある場合）
    if "code" in st.query_params:
        creds = handle_oauth_callback(user_id)
        if creds:
            st.query_params.clear()
            st.rerun()
        else:
            st.warning("セッションが切れました。再度ログインしてください。")
            st.query_params.clear()
            st.stop()

    # 有効なトークンを取得
    creds = get_valid_credentials(user_id)
    if creds:
        return creds

    # 新規認証が必要 → UI 表示
    auth_url = build_auth_url()
    # セッションに invalid_grant フラグがあれば理由を明示
    if st.session_state.pop("_invalid_grant", False):
        st.warning("Googleアカウントの連携が切れました。再度連携してください。")
    else:
        st.info("Googleカレンダーへのアクセス許可が必要です。下のボタンからGoogleアカウントで連携してください。")
    st.link_button("Googleアカウントで連携する", auth_url, use_container_width=True, type="primary")
    st.stop()
    return None


def build_google_services(creds: Credentials) -> dict:
    """
    Calendar / Tasks / Sheets の各サービスを構築して返す。
    失敗したサービスは None。Calendar は必須（失敗時は空 dict を返す）。

    返り値: {
        "calendar_service": ...,
        "editable_calendar_options": {...},
        "tasks_service": ...,
        "default_task_list_id": ...,
        "sheets_service": ...,
    }
    """
    from googleapiclient.errors import HttpError

    result = {
        "calendar_service": None,
        "editable_calendar_options": {},
        "tasks_service": None,
        "default_task_list_id": None,
        "sheets_service": None,
    }

    # Calendar（必須）
    try:
        svc = build("calendar", "v3", credentials=creds)
        cal_list = svc.calendarList().list().execute()
        result["calendar_service"] = svc
        result["editable_calendar_options"] = {
            c["summary"]: c["id"]
            for c in cal_list.get("items", [])
            if c.get("accessRole") != "reader"
        }
        st.session_state["calendar_service"]          = svc
        st.session_state["editable_calendar_options"] = result["editable_calendar_options"]
    except HttpError as e:
        status = e.resp.status if hasattr(e, "resp") else None
        if status in (401, 403):
            st.error("Googleカレンダーへのアクセス権限がありません。ページを再読み込みして再連携してください。")
        else:
            st.error(f"Googleカレンダーへの接続に失敗しました（エラーコード: {status}）。しばらく待ってから再試行してください。")
        return result
    except Exception:
        st.error("Googleカレンダーへの接続に失敗しました。ネットワーク接続を確認してください。")
        return result

    # Tasks（任意）
    try:
        tasks_svc = build_tasks_service(creds)
        if tasks_svc:
            result["tasks_service"]         = tasks_svc
            result["default_task_list_id"]  = get_default_task_list_id(tasks_svc)
            st.session_state["tasks_service"]        = tasks_svc
            st.session_state["default_task_list_id"] = result["default_task_list_id"]
    except Exception:
        pass

    # Sheets（任意）
    try:
        sheets_svc = build("sheets", "v4", credentials=creds)
        result["sheets_service"] = sheets_svc
        st.session_state["sheets_service"] = sheets_svc
    except Exception:
        pass

    return result
