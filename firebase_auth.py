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

def firebase_auth_form():
    """ログイン/サインアップのUIを表示し、認証状態を管理する"""
    st.title("Firebase認証")
    # 認証情報をセッションステートで管理
    if "user_info" not in st.session_state:
        st.session_state.user_info = None

    if st.session_state.user_info is None:
        choice = st.selectbox("選択してください", ["ログイン", "新規登録"])
        
        if choice == "新規登録":
            new_email = st.text_input("新しいメールアドレス", key="signup_email")
            new_password = st.text_input("新しいパスワード", type="password", key="signup_password")
            if st.button("新規登録"):
                if new_email and new_password:
                    try:
                        user = auth.create_user(email=new_email, password=new_password)
                        st.success(f"ユーザー {user.uid} の新規登録が完了しました。ログインしてください。")
                    except Exception as e:
                        st.error(f"新規登録に失敗しました: {e}")
                else:
                    st.warning("メールアドレスとパスワードを入力してください。")
        else: # ログイン
            email = st.text_input("メールアドレス", key="login_email")
            password = st.text_input("パスワード", type="password", key="login_password")
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
    else:
        st.success(f"ログイン済みユーザー: {st.session_state.user_email}")
        if st.button("ログアウト"):
            st.session_state.user_info = None
            st.session_state.user_email = None
            if 'credentials' in st.session_state:
                del st.session_state.credentials
            st.info("ログアウトしました。")
            st.rerun()

def get_firebase_user_id():
    """現在の認証済みユーザーIDを返す"""
    return st.session_state.get("user_info")

def get_firebase_user_id():
    """現在の認証済みユーザーIDを返す"""
    return st.session_state.get("user_info")
