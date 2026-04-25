from __future__ import annotations
"""
core/auth/firebase_client.py
Firebase 認証ロジック（st.* 禁止）

firebase_auth.py から認証ロジックのみを抽出した実装。
セッション状態の getter も含む（st.session_state は Data として許容）。
"""
import logging
from typing import Optional

import requests
import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

logger = logging.getLogger(__name__)


# ── Firebase 初期化 ──────────────────────────────────────────

def initialize_firebase() -> bool:
    """Firebase Admin SDK を初期化する。すでに初期化済みなら True を返すだけ。"""
    if firebase_admin._apps:
        return True
    try:
        s = st.secrets["firebase"]
        cred_dict = {k: s[k] for k in (
            "type", "project_id", "private_key_id", "private_key",
            "client_email", "client_id", "auth_uri", "token_uri",
            "auth_provider_x509_cert_url", "client_x509_cert_url", "universe_domain",
        )}
        firebase_admin.initialize_app(credentials.Certificate(cred_dict))
        return True
    except Exception as e:
        logger.error("Firebase 初期化失敗: %s", e)
        return False


# ── REST API 認証 ────────────────────────────────────────────

def sign_in(email: str, password: str) -> dict:
    """
    メール・パスワードで Firebase Auth にサインインする。
    成功時: {"success": True, "user_id": ..., "email": ..., "id_token": ...}
    失敗時: {"success": False, "error": "<エラーコード>"}
    """
    api_key = st.secrets["web_api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}"
    try:
        resp = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True})
        if resp.status_code == 200:
            d = resp.json()
            return {"success": True, "user_id": d["localId"], "email": d["email"], "id_token": d["idToken"]}
        return {"success": False, "error": resp.json().get("error", {}).get("message", "UNKNOWN_ERROR")}
    except Exception as e:
        return {"success": False, "error": str(e)}


def sign_up(email: str, password: str) -> dict:
    """
    メール・パスワードで Firebase Auth にアカウントを作成する。
    成功時: {"success": True, "user_id": ..., "email": ...}
    失敗時: {"success": False, "error": "<エラーコード>"}
    """
    api_key = st.secrets["web_api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={api_key}"
    try:
        resp = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True})
        if resp.status_code == 200:
            d = resp.json()
            return {"success": True, "user_id": d["localId"], "email": d["email"]}
        return {"success": False, "error": resp.json().get("error", {}).get("message", "UNKNOWN_ERROR")}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── セッション getter ────────────────────────────────────────

def get_user_id() -> Optional[str]:
    """セッションからログイン中のユーザー ID を返す。未ログインなら None。"""
    return st.session_state.get("user_info")


def get_user_email() -> Optional[str]:
    """セッションからログイン中のメールアドレスを返す。"""
    return st.session_state.get("user_email")


def get_id_token() -> Optional[str]:
    """セッションから Firebase ID トークンを返す。"""
    return st.session_state.get("id_token")


def is_authenticated() -> bool:
    """ログイン済みかどうかを返す。"""
    return bool(st.session_state.get("user_info"))


# ── Firestore クライアント ────────────────────────────────────

def get_firestore_client():
    """Firestore クライアントを返す。未初期化なら初期化してから返す。"""
    initialize_firebase()
    return firestore.client()
