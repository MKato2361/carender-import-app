import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import requests
import json

def initialize_firebase():
    """Firebase Admin SDKの初期化"""
    if not firebase_admin._apps:
        try:
            cred_dict = st.secrets["firebase"]
            cred = credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(cred)
            return True
        except Exception as e:
            st.error(f"Firebaseの初期化に失敗しました: {e}")
            return False
    return True

def _call_firebase_api(endpoint, payload):
    """Firebase Authentication REST APIを呼び出すヘルパー関数"""
    try:
        api_key = st.secrets["web_api_key"]
        url = f"https://identitytoolkit.googleapis.com/v1/accounts:{endpoint}?key={api_key}"
        payload["returnSecureToken"] = True
        response = requests.post(url, json=payload)
        response.raise_for_status()
        return {"success": True, "data": response.json()}
    except requests.exceptions.RequestException as e:
        try:
            error_message = e.response.json()["error"]["message"]
        except (requests.exceptions.JSONDecodeError, KeyError):
            error_message = str(e)
        return {"success": False, "error": error_message}

def authenticate_user(email, password):
    return _call_firebase_api("signInWithPassword", {"email": email, "password": password})

def create_user_account(email, password):
    return _call_firebase_api("signUp", {"email": email, "password": password})

def firebase_auth_form():
    """ログイン/サインアップのUIを表示し、認証状態を管理する"""
    st.title("Firebase認証")
    if "user_info" not in st.session_state:
        st.session_state.user_info = None

    if st.session_state.user_info is None:
        _display_auth_form()
    else:
        _display_logout_button()

def _display_auth_form():
    choice = st.selectbox("選択してください", ["ログイン", "新規登録"])
    if choice == "新規登録":
        new_email = st.text_input("新しいメールアドレス", key="signup_email")
        new_password = st.text_input("新しいパスワード", type="password", key="signup_password")
        if st.button("新規登録"):
            if new_email and new_password:
                result = create_user_account(new_email, new_password)
                if result["success"]:
                    st.success(f"新規登録が完了しました。ログインしてください。")
                else:
                    st.error(f"新規登録に失敗しました: {result['error']}")
            else:
                st.warning("メールアドレスとパスワードを入力してください。")
    else:
        email = st.text_input("メールアドレス", key="login_email")
        password = st.text_input("パスワード", type="password", key="login_password")
        if st.button("ログイン"):
            if email and password:
                result = authenticate_user(email, password)
                if result["success"]:
                    st.session_state.user_info = result["data"]["localId"]
                    st.session_state.user_email = result["data"]["email"]
                    st.session_state.id_token = result["data"]["idToken"]
                    st.success("ログインしました！")
                    st.rerun()
                else:
                    st.error(f"ログインに失敗しました: {result['error']}")
            else:
                st.warning("メールアドレスとパスワードを入力してください。")

def _display_logout_button():
    st.success(f"ログイン済みユーザー: {st.session_state.user_email}")
    if st.button("ログアウト"):
        st.session_state.user_info = None
        st.session_state.user_email = None
        if 'id_token' in st.session_state:
            del st.session_state.id_token
        if 'creds' in st.session_state:
            del st.session_state.creds
        st.info("ログアウトしました。")
        st.rerun()

def get_firebase_user_id():
    return st.session_state.get("user_info")

def get_firebase_user_email():
    return st.session_state.get("user_email")

def get_firebase_id_token():
    return st.session_state.get("id_token")

def is_user_authenticated():
    return st.session_state.get("user_info") is not None
