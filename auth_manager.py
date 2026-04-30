from __future__ import annotations
"""
auth_manager.py（完全修正版）
- ユーザーごとの状態混在を防止
- cache_resource を完全撤廃
- session_stateベースに統一
"""

import streamlit as st
from firebase_admin import firestore

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

    def __init__(self):
        self.db = None

    # ── Firebase ──

    def sync_with_session(self) -> str | None:
        """Firebase 認証状態を確認し、セッション初期化"""
        if not initialize_firebase():
            return None

        self.db = firestore.client()
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

        # 既に初期化済みならスキップ
        if st.session_state.get("_google_services_initialized"):
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

        st.session_state["_google_services_initialized"] = True

        return True

    # ── 設定保存 ──

    def save_user_setting(self, *args, **kwargs) -> None:
        """柔軟な引数対応"""
        user_id = get_firebase_user_id()

        key = value = None
        if len(args) == 3:
            user_id, key, value = args
        elif len(args) == 2:
            key, value = args
        else:
            key = kwargs.get("key")
            value = kwargs.get("value")
            user_id = kwargs.get("user_id") or user_id

        if user_id and key is not None:
            set_user_setting(user_id, key, value)

    # ── 状態判定 ──

    @property
    def is_authenticated(self) -> bool:
        return bool(get_firebase_user_id() and st.session_state.get("credentials"))
        

# ❌ キャッシュ削除（これが一番重要）
def get_auth_manager() -> AuthManager:
    """毎回新しいインスタンスを返す（安全）"""
    return AuthManager()
