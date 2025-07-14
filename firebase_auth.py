import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth
from google.oauth2.credentials import Credentials

# StreamlitのシークレットからFirebaseのサービスアカウント情報を取得
FIREBASE_SECRETS = st.secrets["firebase"]

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
initialize_firebase()

# 認証済みユーザーのIDを返す関数（calendar_utils.pyから呼び出されることを想定）
def get_firebase_user_id():
    """認証済みユーザーのIDを返す"""
    if "user_info" in st.session_state and st.session_state.user_info:
        return st.session_state.user_info
    return None

def firebase_auth_form():
    """Firebase Authenticationのフォームを表示"""
    st.subheader("ログイン / 新規登録")
    st.write("---")

    col1, col2 = st.columns(2)
    with col1:
        st.header("既存アカウントでログイン")
        email = st.text_input("メールアドレス")
        password = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            if email and password:
                try:
                    user = auth.get_user_by_email(email)
                    st.session_state.user_info = user.uid
                    st.session_state.user_email = email
                    st.success("ログインしました！")
                    st.rerun()
                except auth.UserNotFoundError:
                    st.error("ユーザーが見つかりません。")
                except Exception as e:
                    st.error(f"ログインに失敗しました: {e}")
            else:
                st.warning("メールアドレスとパスワードを入力してください。")

    with col2:
        st.header("新規アカウントを作成")
        new_email = st.text_input("新しいメールアドレス")
        new_password = st.text_input("新しいパスワード", type="password")
        if st.button("新規登録"):
            if new_email and new_password:
                try:
                    user = auth.create_user(email=new_email, password=new_password)
                    st.success(f"ユーザー {user.uid} の新規登録が完了しました。ログインしてください。")
                except Exception as e:
                    st.error(f"新規登録に失敗しました: {e}")
            else:
                st.warning("メールアドレスとパスワードを入力してください。")

if "user_info" in st.session_state and st.session_state.user_info:
    st.success(f"ログイン済みユーザー: {st.session_state.user_email}")
    if st.button("ログアウト"):
        st.session_state.user_info = None
        st.session_state.user_email = None
        st.info("ログアウトしました。")
        st.rerun()
else:
    firebase_auth_form()
