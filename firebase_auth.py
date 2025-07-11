import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth
from google.oauth2.credentials import Credentials

# StreamlitのシークレットからFirebaseのサービスアカウント情報を取得
FIREBASE_SERVICE_ACCOUNT_KEY = st.secrets["firebase"]["service_account_key"]
# Google認証用のシークレットも同時に取得
GOOGLE_CLIENT_ID = st.secrets["google"]["client_id"]
GOOGLE_CLIENT_SECRET = st.secrets["google"]["client_secret"]

def initialize_firebase():
    """Firebase Admin SDKの初期化"""
    if not firebase_admin._apps:
        try:
            # シークレットから取得したJSON文字列を使って認証情報を初期化
            cred = credentials.Certificate(FIREBASE_SERVICE_ACCOUNT_KEY)
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
        email = st.text_input("メールアドレス")
        password = st.text_input("パスワード", type="password")
        
        if choice == "新規登録":
            if st.button("新規登録"):
                if email and password:
                    try:
                        user = auth.create_user(email=email, password=password)
                        st.success(f"ユーザー {user.uid} の新規登録が完了しました。ログインしてください。")
                    except Exception as e:
                        st.error(f"新規登録に失敗しました: {e}")
                else:
                    st.warning("メールアドレスとパスワードを入力してください。")
        else: # ログイン
            if st.button("ログイン"):
                if email and password:
                    try:
                        # Streamlitで直接ユーザーのパスワードを認証する安全な方法がないため、
                        # 仮のロジックとして、Firebaseのカスタム認証トークンを発行し、ユーザーの存在を確認する
                        # 実際のWebアプリケーションではクライアント側（JavaScript）で認証を行うのが一般的
                        user = auth.get_user_by_email(email)
                        # この時点でパスワードの正当性は検証できないため、
                        # 厳密な認証には、クライアントサイドでの実装が必須
                        st.session_state.user_info = user.uid
                        st.session_state.user_email = email
                        st.success("ログインしました！")
                        st.experimental_rerun()
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
            st.experimental_rerun()

def get_firebase_user_id():
    """現在の認証済みユーザーIDを返す"""
    return st.session_state.get("user_info")