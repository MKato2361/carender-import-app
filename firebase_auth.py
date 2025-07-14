import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth
from google.oauth2.credentials import Credentials

# streamlit-firebase-auth をインポート
from streamlit_firebase_auth import with_firebase_auth
import json

# StreamlitのシークレットからFirebaseのサービスアカウント情報を取得
FIREBASE_SECRETS = st.secrets["firebase"]

# Firebase Admin SDKの初期化
def initialize_firebase():
    """Firebase Admin SDKの初期化"""
    if not firebase_admin._apps:
        try:
            cred_dict = {
                "type": FIREBASE_SECRETS["type"],
                "project_id": FIREBASE_SECRETS["project_id"],
                "private_key_id": FIREBASE_SECRETS["private_key_id"],
                "private_key": FIREBASE_SECRETS["private_key"],
                "client_email": FIREBASE_SECRETS["client_email"],
                "client_id": FIREBASE_SECRETS["client_id"],
                "auth_uri": FIREBASE_SECRETS["auth_uri"],
                "token_uri": FIREBASE_SECRETS["token_uri"],
                "auth_provider_x509_cert_url": FIREBASE_SECRETS["auth_provider_x509_cert_url"],
                "client_x509_cert_url": FIREBASE_SECRETS["client_x509_cert_url"],
                "universe_domain": FIREBASE_SECRETS["universe_domain"]
            }
            
            cred = credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(cred)
            return True
        except Exception as e:
            st.error(f"Firebaseの初期化に失敗しました: {e}")
            return False
    return True

# FirebaseのAdmin SDKを初期化
if initialize_firebase():
    
    # with_firebase_auth を使用して認証フォームとセッション管理を統合
    firebase_config = {
        "apiKey": st.secrets["firebase"]["api_key"],
        "authDomain": st.secrets["firebase"]["auth_domain"],
        "projectId": st.secrets["firebase"]["project_id"],
        "storageBucket": st.secrets["firebase"]["storage_bucket"],
        "messagingSenderId": st.secrets["firebase"]["messaging_sender_id"],
        "appId": st.secrets["firebase"]["app_id"]
    }
    
    # 認証状態とユーザー情報を取得
    st.session_state.auth_state, st.session_state.user = with_firebase_auth(firebase_config)

    # 認証状態に応じたUIを表示
    if st.session_state.auth_state["status"] == "signed_in":
        st.success(f"ログイン済みユーザー: {st.session_state.user['email']}")
        if st.button("ログアウト"):
            st.session_state.auth_state["status"] = "signed_out"
            st.rerun()

    elif st.session_state.auth_state["status"] == "signed_out":
        st.info("ログインまたは新規登録してください。")

    elif st.session_state.auth_state["status"] == "not_signed_in":
        st.info("ログインしてください。")
        
    elif st.session_state.auth_state["status"] == "loading":
        st.info("認証情報を確認中です...")
        
    elif st.session_state.auth_state["status"] == "not_authenticated":
        st.error("認証に失敗しました。")

# calendar_utils.pyとの互換性のために関数を再追加
def get_firebase_user_id():
    """認証済みユーザーのIDを返す"""
    if "user" in st.session_state and st.session_state.user:
        return st.session_state.user["uid"]
    return None
