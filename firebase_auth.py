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
            # 辞書を再構成
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
            
            # 辞書を使って認証情報を初期化
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
    # FirebaseのAPIキーはStreamlitのsecrets.tomlに追加してください
    # firebase_config = {"apiKey": "YOUR_API_KEY", ...}
    firebase_config = {
        "apiKey": st.secrets["firebase"]["api_key"],
        "authDomain": st.secrets["firebase"]["auth_domain"],
        "projectId": st.secrets["firebase"]["project_id"],
        "storageBucket": st.secrets["firebase"]["storage_bucket"],
        "messagingSenderId": st.secrets["firebase"]["messaging_sender_id"],
        "appId": st.secrets["firebase"]["app_id"]
    }
    
    auth_state, user = with_firebase_auth(firebase_config)

    if auth_state["status"] == "signed_in":
        # ログイン済みの場合の処理
        st.success(f"ログイン済みユーザー: {user['email']}")
        st.write("認証が永続化されました。ページをリロードしてもログイン状態は維持されます。")
        if st.button("ログアウト"):
            # ライブラリのログアウト機能を使用
            auth_state["status"] = "signed_out"
            st.info("ログアウトしました。")
            st.rerun()

    elif auth_state["status"] == "signed_out":
        # ログアウト状態の場合、ライブラリが自動的にログイン/新規登録UIを表示します
        st.info("ログインまたは新規登録してください。")

    elif auth_state["status"] == "not_signed_in":
        # 認証されていない状態
        st.info("ログインしてください。")
        
    elif auth_state["status"] == "loading":
        st.info("認証情報を確認中です...")
        
    elif auth_state["status"] == "not_authenticated":
        st.error("認証に失敗しました。")
        
    # ライブラリはセッションステートを内部で管理し、永続化します
    # ユーザー情報は auth_state["user"] で取得できます
    if user:
        st.write("---")
        st.subheader("ユーザー情報（デバッグ用）")
        st.json(user)
