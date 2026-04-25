"""
auth_manager.py
Firebase + Google 認証と API サービスを一括管理するクラス。

認証ロジックは services/auth_service.py に委譲。
このクラスはサービスインスタンスを保持するシェルとして機能する。
"""
from __future__ import annotations
import streamlit as st
from firebase_admin import firestore

from firebase_auth import initialize_firebase, get_firebase_user_id
from services.auth_service import authenticate_google, build_google_services
from services.settings_service import (
    _ensure_initialized as initialize_session_state,
    set_setting as set_user_setting,
)


class AuthManager:
    """
    Firebase 認証・Google 認証・各 API サービスを一括管理する。
    Streamlit の session_state を跨いでサービスを保持し、再認証フローを簡略化する。
    """

    def __init__(self):
        self.firebase_user_id        = None
        self.google_creds            = None
        self.calendar_service        = None
        self.tasks_service           = None
        self.sheets_service          = None
        self.editable_calendar_options: dict = {}
        self.default_task_list_id    = None
        self.db                      = None
        self._initialized_services   = False

    # ── Firebase ──

    def sync_with_session(self) -> str | None:
        """Firebase 認証状態を確認し、ユーザー設定をセッションにロードする。"""
        if not initialize_firebase():
            return None
        self.db               = firestore.client()
        self.firebase_user_id = get_firebase_user_id()
        if self.firebase_user_id:
            initialize_session_state(self.firebase_user_id)
        return self.firebase_user_id

    # ── Google ──

    def ensure_google_services(self) -> bool:
        """Google 認証を確認し、未初期化なら各 API サービスを構築する。"""
        if not self.firebase_user_id:
            return False

        self.google_creds = authenticate_google()
        if not self.google_creds:
            return False

        if not self._initialized_services:
            with st.spinner("Googleサービスに接続中..."):
                result = build_google_services(self.google_creds)

            self.calendar_service         = result["calendar_service"]
            self.editable_calendar_options = result["editable_calendar_options"]
            self.tasks_service            = result["tasks_service"]
            self.default_task_list_id     = result["default_task_list_id"]
            self.sheets_service           = result["sheets_service"]

            if not self.calendar_service:
                return False
            self._initialized_services = True

        return True

    # ── 設定保存（呼び出し元の多様な引数形式に対応） ──

    def save_user_setting(self, *args, **kwargs) -> None:
        """(user_id, key, value) or (key, value) 形式どちらでも受け付ける。"""
        user_id = self.firebase_user_id
        key = value = None
        if len(args) == 3:
            user_id, key, value = args
        elif len(args) == 2:
            key, value = args
        else:
            key     = kwargs.get("key")
            value   = kwargs.get("value")
            user_id = kwargs.get("user_id") or self.firebase_user_id
        if user_id and key is not None:
            set_user_setting(user_id, key, value)

    @property
    def is_authenticated(self) -> bool:
        return bool(self.firebase_user_id and self.google_creds)


@st.cache_resource
def get_auth_manager() -> AuthManager:
    """AuthManager のシングルトンインスタンスを取得する。"""
    return AuthManager()
