from __future__ import annotations
"""
core/auth/google_oauth.py
Google OAuth トークンの検証・取得・更新ロジック（st.* 禁止）

authenticate_google() から純粋ロジック部分を抽出。
セッション状態の読み書きは許可（st.session_state は Data only）。
st.info / st.error 等の UI 表示は呼び出し元（services/auth_service.py）が担う。
"""
import json
import urllib.parse
from typing import Optional

import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from requests_oauthlib import OAuth2Session

from core.storage.firestore_client import load_token, save_token, delete_token

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks",
    "https://www.googleapis.com/auth/spreadsheets",
]

AUTH_URI  = "https://accounts.google.com/o/oauth2/auth"
TOKEN_URI = "https://oauth2.googleapis.com/token"

import os
try:
    _redirect = st.secrets.get("google", {}).get("redirect_uri", "")
    if _redirect.startswith("http://"):
        os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
except Exception:
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"


def _save_creds_to_session(creds: Credentials, user_id: str) -> None:
    st.session_state["credentials"]         = creds
    st.session_state["credentials_user_id"] = user_id


def _clear_creds(user_id: Optional[str] = None) -> None:
    st.session_state.pop("credentials",         None)
    st.session_state.pop("credentials_user_id", None)
    if user_id:
        delete_token(user_id)


class InvalidGrantError(Exception):
    """リフレッシュトークンが失効・失効している場合に raise する"""
    pass


def get_valid_credentials(user_id: str) -> Optional[Credentials]:
    """
    有効な Google OAuth 認証情報を返す。
    ① セッション → ② Firestore → None（OAuthフロー必要）の順に探す。

    返り値:
      Credentials  : 有効なトークンあり
      None         : 新規認証が必要（呼び出し元が OAuth フローを開始する）
    """
    st.session_state.setdefault("credentials",         None)
    st.session_state.setdefault("credentials_user_id", None)

    # ① セッションから
    if (st.session_state["credentials"]
            and st.session_state.get("credentials_user_id") == user_id):
        creds: Credentials = st.session_state["credentials"]

        if not creds.refresh_token:
            _clear_creds(user_id)
        elif creds.valid:
            return creds
        elif creds.expired:
            try:
                creds.refresh(Request())
                _save_creds_to_session(creds, user_id)
                save_token(user_id, json.loads(creds.to_json()))
                return creds
            except Exception as e:
                # invalid_grant = トークン失効 → 削除して再認証
                _clear_creds(user_id)

    # ② Firestore から
    token_data = load_token(user_id)
    if token_data:
        try:
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
        except Exception:
            creds = None

        if creds:
            if not creds.refresh_token:
                delete_token(user_id)
            elif creds.expired:
                try:
                    creds.refresh(Request())
                    _save_creds_to_session(creds, user_id)
                    save_token(user_id, json.loads(creds.to_json()))
                    return creds
                except Exception as e:
                    # invalid_grant = トークン失効 → 削除して再認証
                    delete_token(user_id)
            elif creds.valid:
                _save_creds_to_session(creds, user_id)
                return creds

    return None


def build_auth_url() -> str:
    """Google OAuth 認証 URL を生成して返す。"""
    client_id    = st.secrets["google"]["client_id"]
    redirect_uri = st.secrets["google"]["redirect_uri"]
    oauth = OAuth2Session(client_id=client_id, redirect_uri=redirect_uri, scope=SCOPES)
    auth_url, _ = oauth.authorization_url(
        AUTH_URI, access_type="offline", prompt="consent",
        include_granted_scopes="true",
    )
    return auth_url


def handle_oauth_callback(user_id: str) -> Optional[Credentials]:
    """
    OAuth コールバック（?code=... が URL にある状態）を処理して Credentials を返す。
    失敗時は None を返す（エラー表示は呼び出し元が行う）。
    """
    params = st.query_params
    state  = params.get("state")
    if not state:
        _clear_creds()
        return None

    client_id     = st.secrets["google"]["client_id"]
    client_secret = st.secrets["google"]["client_secret"]
    redirect_uri  = st.secrets["google"]["redirect_uri"]

    oauth = OAuth2Session(
        client_id=client_id, redirect_uri=redirect_uri,
        scope=SCOPES, state=state,
    )
    current_url = redirect_uri + "?" + "&".join(
        f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items()
    )

    try:
        token = oauth.fetch_token(
            TOKEN_URI, authorization_response=current_url,
            client_secret=client_secret,
        )
        creds = Credentials(
            token=token.get("access_token"),
            refresh_token=token.get("refresh_token"),
            token_uri=TOKEN_URI,
            client_id=client_id,
            client_secret=client_secret,
            scopes=SCOPES,
        )
        if not creds or not creds.refresh_token:
            return None

        _save_creds_to_session(creds, user_id)
        save_token(user_id, json.loads(creds.to_json()))
        return creds
    except Exception:
        return None
