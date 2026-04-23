"""
core/auth/firebase_client.py
Firebase REST API クライアント（st.* 禁止）

firebase_auth.py の UI 部分を除いた認証ロジックのみを担う。
"""
from __future__ import annotations
import json
from typing import Optional
import requests
import streamlit as st  # secrets の取得のみに使用（UI呼び出しは不可）

# ── Firebase Admin SDK 初期化 ──
import firebase_admin
from firebase_admin import credentials, firestore


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
        raise RuntimeError(f"Firebase の初期化に失敗しました: {e}") from e


def sign_in(email: str, password: str) -> dict:
    """
    メール・パスワードで Firebase Auth にサインインする。
    成功時: {"success": True, "user_id": ..., "email": ..., "id_token": ...}
    失敗時: {"success": False, "error": "<Firebase エラーコード>"}
    """
    api_key = st.secrets["web_api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}"
    resp = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True})
    if resp.status_code == 200:
        d = resp.json()
        return {"success": True, "user_id": d["localId"], "email": d["email"], "id_token": d["idToken"]}
    return {"success": False, "error": resp.json().get("error", {}).get("message", "UNKNOWN_ERROR")}


def sign_up(email: str, password: str) -> dict:
    """
    メール・パスワードで Firebase Auth にアカウントを作成する。
    成功時: {"success": True, "user_id": ..., "email": ...}
    失敗時: {"success": False, "error": "<Firebase エラーコード>"}
    """
    api_key = st.secrets["web_api_key"]
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={api_key}"
    resp = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True})
    if resp.status_code == 200:
        d = resp.json()
        return {"success": True, "user_id": d["localId"], "email": d["email"]}
    return {"success": False, "error": resp.json().get("error", {}).get("message", "UNKNOWN_ERROR")}


def get_firestore_client():
    """Firestore クライアントを返す。未初期化なら初期化してから返す。"""
    initialize_firebase()
    return firestore.client()
