import os
import json
import pickle
from pathlib import Path
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from firebase_admin import firestore
from firebase_auth import get_firebase_user_id
from googleapiclient.discovery import build

# Google API スコープ
SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]

def authenticate_google():
    creds = None
    user_id = get_firebase_user_id()

    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # --- セッションから ---
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds
        elif creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                doc_ref.set(json.loads(creds.to_json()))
                return creds
            except Exception as e:
                st.warning(f"リフレッシュトークンの更新に失敗しました: {e}")
                doc_ref.delete()
                st.session_state.pop('credentials', None)
                return authenticate_google()  # 再試行

    # --- Firestoreから ---
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            st.session_state['credentials'] = creds

            if creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    st.session_state['credentials'] = creds
                    doc_ref.set(json.loads(creds.to_json()))
                    st.info("Google認証トークンを更新しました。")
                    st.rerun()
                except Exception as e:
                    st.warning(f"Firestoreトークンの更新に失敗しました: {e}")
                    doc_ref.delete()
                    st.session_state.pop('credentials', None)
                    return authenticate_google()  # 再試行

            return creds
    except Exception as e:
        if "invalid_grant" in str(e):
            st.warning("保存されたGoogleトークンが無効化されました。再認証します。")
            doc_ref.delete()
            st.session_state.pop('credentials', None)
            return authenticate_google()  # 再試行
        else:
            st.error(f"Firestoreからトークン取得に失敗しました: {e}")
            creds = None

    # --- 新しいOAuthフロー（Webリダイレクト型） ---
    try:
        client_config = {
            "web": {
                "client_id": st.secrets["google"]["client_id"],
                "project_id": st.secrets["google"]["project_id"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets["google"]["client_secret"],
                "redirect_uris": [st.secrets["google"]["redirect_uri"]]
            }
        }

        flow = Flow.from_client_config(client_config, SCOPES)
        flow.redirect_uri = st.secrets["google"]["redirect_uri"]

        params = st.experimental_get_query_params()
        if "code" not in params:
            auth_url, _ = flow.authorization_url(
                prompt='consent',
                access_type='offline',
                include_granted_scopes='true'
            )
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()
        else:
            code = params["code"][0]
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.session_state['credentials'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            st.success("Google認証が完了しました！")
            st.experimental_set_query_params()  # URLからcodeを削除
            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state['credentials'] = None
        return None

    return creds
