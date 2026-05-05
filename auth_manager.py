from __future__ import annotations
"""
auth_manager.py（完全修正版）
- ユーザーごとの状態混在を防止
- cache_resource を完全撤廃
- session_stateベースに統一
"""

import streamlit as st

from core.auth.firebase_client import (
    initialize_firebase,
    get_user_id as get_firebase_user_id,
)
from services.auth_service import authenticate_google, build_google_services
from services.settings_service import (
    _ensure_initialized as initialize_session_state,
    set_setting as set_user_setting,
)


class AuthManager:
    """
    認証・サービス管理（ステートレス寄り設計）
    - インスタンスは軽量（毎回生成OK）
    - 実データは session_state に保持
    """

    # ── Firebase ──

    def sync_with_session(self) -> str | None:
        """Firebase 認証状態を確認し、セッション初期化"""
        if not initialize_firebase():
            return None

        user_id = get_firebase_user_id()

        if user_id:
            initialize_session_state(user_id)

        return user_id

    # ── Google ──

    def ensure_google_services(self) -> bool:
        """Google認証 + APIサービス初期化"""

        user_id = get_firebase_user_id()
        if not user_id:
            return False

        # 認証（内部で session_state 使用）
        creds = authenticate_google()
        if not creds:
            return False

        # 既に初期化済みならスキップ（user_id で照合してユーザー混在を防ぐ）
        if st.session_state.get("_google_services_initialized") == user_id:
            return True

        with st.spinner("Googleサービスに接続中..."):
            result = build_google_services(creds)

        if not result["calendar_service"]:
            return False

        # session_stateに保存（ユーザー単位）
        st.session_state["calendar_service"] = result["calendar_service"]
        st.session_state["tasks_service"] = result["tasks_service"]
        st.session_state["sheets_service"] = result["sheets_service"]
        st.session_state["editable_calendar_options"] = result["editable_calendar_options"]
        st.session_state["default_task_list_id"] = result["default_task_list_id"]

        st.session_state["_google_services_initialized"] = user_id

        return True

    # ── 設定保存 ──

    def save_user_setting(self, user_id: str, key: str, value) -> None:
        set_user_setting(user_id, key, value)

    # ── 状態判定 ──

    @property
    def is_authenticated(self) -> bool:
        return bool(get_firebase_user_id() and st.session_state.get("credentials"))
        

def get_auth_manager() -> AuthManager:
    return AuthManager()
