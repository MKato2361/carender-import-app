from __future__ import annotations
import streamlit as st
import json
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from firebase_admin import firestore

# 既存の認証関連ユーティリティをインポート
from calendar_utils import authenticate_google, build_tasks_service
from firebase_auth import initialize_firebase, get_firebase_user_id
from session_utils import initialize_session_state, set_user_setting, load_user_settings_from_firestore

class AuthManager:
    """
    Firebase認証、Google認証、およびAPIサービス（Calendar, Tasks, Sheets）を一括管理するクラス。
    Streamlitのセッションを跨いでサービスを保持し、再認証フローを簡略化します。
    """
    def __init__(self):
        self.firebase_user_id = None
        self.google_creds = None
        self.calendar_service = None
        self.tasks_service = None
        self.sheets_service = None
        self.editable_calendar_options = {}
        self.default_task_list_id = None
        self.db = None
        self._initialized_services = False

    def sync_with_session(self) -> str | None:
        """
        Firebase認証状態を確認し、必要に応じてセッションを初期化する。
        """
        if not initialize_firebase():
            return None
        
        self.db = firestore.client()
        self.firebase_user_id = get_firebase_user_id()
        
        if self.firebase_user_id:
            # ユーザー設定のロード
            initialize_session_state(self.firebase_user_id)
        
        return self.firebase_user_id

    def ensure_google_services(self) -> bool:
        """
        Google認証を確認し、各種APIサービスを初期化する。
        """
        if not self.firebase_user_id:
            return False

        # Google認証 (既存の authenticate_google を利用)
        self.google_creds = authenticate_google()
        if not self.google_creds:
            return False

        # サービスが未初期化の場合に構築（Calendarが必須。失敗した場合はリトライできるよう _initialized_services をセットしない）
        if not self._initialized_services:
            with st.spinner("Googleサービスに接続中..."):
                self._build_all_services()
            if self.calendar_service:
                self._initialized_services = True
            else:
                return False

        return True

    def _build_all_services(self):
        """各種Google APIサービスを構築"""
        from googleapiclient.errors import HttpError

        # 1. Calendar（必須サービス）
        try:
            self.calendar_service = build("calendar", "v3", credentials=self.google_creds)
            cal_list = self.calendar_service.calendarList().list().execute()
            self.editable_calendar_options = {
                cal["summary"]: cal["id"]
                for cal in cal_list.get("items", [])
                if cal.get("accessRole") != "reader"
            }
            st.session_state["calendar_service"] = self.calendar_service
            st.session_state["editable_calendar_options"] = self.editable_calendar_options
        except HttpError as e:
            self.calendar_service = None
            if e.resp.status in (401, 403):
                st.error("Googleカレンダーへのアクセス権限がありません。ページを再読み込みしてGoogleアカウントを再連携してください。")
            else:
                st.error(f"Googleカレンダーへの接続に失敗しました（エラーコード: {e.resp.status}）。しばらく待ってから再試行してください。")
        except Exception:
            self.calendar_service = None
            st.error("Googleカレンダーへの接続に失敗しました。ネットワーク接続を確認してください。")

        # 2. Tasks（任意サービス）
        try:
            self.tasks_service = build_tasks_service(self.google_creds)
            if self.tasks_service:
                task_lists = self.tasks_service.tasklists().list().execute()
                for item in task_lists.get("items", []):
                    if item.get("title") == "My Tasks":
                        self.default_task_list_id = item["id"]
                        break
                if not self.default_task_list_id and task_lists.get("items"):
                    self.default_task_list_id = task_lists["items"][0]["id"]
                st.session_state["tasks_service"] = self.tasks_service
                st.session_state["default_task_list_id"] = self.default_task_list_id
        except Exception:
            self.tasks_service = None

        # 3. Sheets（任意サービス）
        try:
            self.sheets_service = build("sheets", "v4", credentials=self.google_creds)
            st.session_state["sheets_service"] = self.sheets_service
        except Exception:
            self.sheets_service = None

    def save_user_setting(self, *args, **kwargs):
        """
        設定を保存。引数の数が2つ(key, value)でも3つ(user_id, key, value)でも柔軟に対応。
        これにより呼び出し側の古いコードとの互換性を保ち、TypeErrorを防止します。
        """
        user_id = self.firebase_user_id
        key = None
        value = None

        if len(args) == 3:
            # (user_id, key, value) 形式
            user_id, key, value = args
        elif len(args) == 2:
            # (key, value) 形式
            key, value = args
        else:
            # キーワード引数からの取得
            key = kwargs.get("key")
            value = kwargs.get("value")
            user_id = kwargs.get("user_id") or self.firebase_user_id

        if user_id and key is not None:
            set_user_setting(user_id, key, value)

    @property
    def is_authenticated(self) -> bool:
        return bool(self.firebase_user_id and self.google_creds)

@st.cache_resource
def get_auth_manager() -> AuthManager:
    """AuthManagerのシングルトンインスタンスを取得"""
    return AuthManager()
