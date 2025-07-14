import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth, firestore
import requests
import json

# StreamlitのシークレットからFirebaseのサービスアカウント情報を取得
FIREBASE_SECRETS = st.secrets["firebase"]

def initialize_firebase():
    """Firebase Admin SDKの初期化"""
    if not firebase_admin._apps:
        try:
            cred = credentials.Certificate({
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
            })
            firebase_admin.initialize_app(cred)
            return True
        except Exception as e:
            st.error(f"Firebaseの初期化に失敗しました: {e}")
            return False
    return True

def authenticate_user(email, password):
    """Firebase REST APIを使用してユーザー認証"""
    try:
        API_KEY = st.secrets["web_api_key"]
        url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={API_KEY}"
        payload = {
            "email": email,
            "password": password,
            "returnSecureToken": True
        }
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            user_data = response.json()
            return {
                "success": True,
                "user_id": user_data["localId"],
                "email": user_data["email"],
                "id_token": user_data["idToken"]
            }
        else:
            error_data = response.json()
            return {
                "success": False,
                "error": error_data.get("error", {}).get("message", "認証に失敗しました")
            }
    except Exception as e:
        return {"success": False, "error": str(e)}

def create_user_account(email, password):
    """Firebase REST APIを使用してユーザー作成"""
    try:
        API_KEY = st.secrets["web_api_key"]
        url = f"https://identitytoolkit.googleapis.com/v1/accounts:signUp?key={API_KEY}"
        payload = {
            "email": email,
            "password": password,
            "returnSecureToken": True
        }
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            user_data = response.json()
            return {
                "success": True,
                "user_id": user_data["localId"],
                "email": user_data["email"]
            }
        else:
            error_data = response.json()
            return {
                "success": False,
                "error": error_data.get("error", {}).get("message", "アカウント作成に失敗しました")
            }
    except Exception as e:
        return {"success": False, "error": str(e)}

def firebase_auth_form():
    """ログイン/サインアップのUIを表示し、認証状態を管理する"""
    st.title("Firebase認証")

    if "user_info" not in st.session_state:
        st.session_state.user_info = None
    if "user_email" not in st.session_state:
        st.session_state.user_email = None

    if st.session_state.user_info is None:
        choice = st.selectbox("選択してください", ["ログイン", "新規登録"])

        if choice == "新規登録":
            new_email = st.text_input("新しいメールアドレス", key="signup_email")
            new_password = st.text_input("新しいパスワード", type="password", key="signup_password")
            if st.button("新規登録"):
                if new_email and new_password:
                    result = create_user_account(new_email, new_password)
                    if result["success"]:
                        st.success("新規登録が完了しました。ログインしてください。")
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
                        st.session_state.user_info = result["user_id"]
                        st.session_state.user_email = result["email"]
                        st.session_state.id_token = result["id_token"]
                        st.success("ログインしました！")
                        st.rerun()
                    else:
                        st.error(f"ログインに失敗しました: {result['error']}")
                else:
                    st.warning("メールアドレスとパスワードを入力してください。")
    else:
        st.success(f"ログイン済みユーザー: {st.session_state.user_email}")
        if st.button("ログアウト"):
            st.session_state.user_info = None
            st.session_state.user_email = None
            st.session_state.pop('id_token', None)
            st.session_state.pop('credentials', None)
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

def safe_load_tokens_from_firestore(user_id):
    """Firestoreからトークンを安全に読み込む"""
    try:
        db = firestore.client()
        doc_ref = db.collection('users').document(user_id)
        doc = doc_ref.get()

        if not doc.exists:
            st.warning("ユーザードキュメントが見つかりません。")
            return {}

        token_data = doc.get('tokens') or doc.get('token') or doc.get('credentials')

        if token_data is None:
            st.info("トークンデータがありません。")
            return {}

        st.write(f"トークンデータの型: {type(token_data)}")

        if isinstance(token_data, dict):
            return token_data

        elif isinstance(token_data, str):
            try:
                parsed = json.loads(token_data)
                if isinstance(parsed, dict):
                    return parsed
                else:
                    st.warning("パースされたデータが辞書ではありません。")
                    return {"raw_token": token_data}
            except json.JSONDecodeError as e:
                st.error(f"JSONパースエラー: {e}")
                return {"raw_token": token_data}

        elif isinstance(token_data, list):
            return {"tokens": token_data}

        else:
            st.error(f"未対応のデータ型: {type(token_data)}")
            return {}

    except Exception as e:
        st.error(f"Firestoreからのトークン読み込みに失敗しました: {e}")
        return {}

def save_tokens_to_firestore(user_id, tokens):
    """トークンをFirestoreに保存する"""
    try:
        db = firestore.client()
        doc_ref = db.collection('users').document(user_id)
        doc_ref.set({
            'tokens': tokens,
            'updated_at': firestore.SERVER_TIMESTAMP
        }, merge=True)
        st.success("トークンが正常に保存されました。")
        return True
    except Exception as e:
        st.error(f"トークンの保存に失敗しました: {e}")
        return False

def process_tokens_safely(user_id):
    """トークンを安全に処理する関数"""
    tokens = safe_load_tokens_from_firestore(user_id)
    if not tokens:
        st.info("利用可能なトークンがありません。")
        return

    st.write("取得したトークン:")
    try:
        if isinstance(tokens, dict):
            for key, value in tokens.items():
                st.write(f"- {key}: {value}")
        else:
            st.write(f"トークンデータ: {tokens}")
    except AttributeError as e:
        st.error(f"トークンデータの処理でエラーが発生しました: {e}")
        st.write(f"データの型: {type(tokens)}")
        st.write(f"データの内容: {tokens}")

def main():
    st.title("Firebase認証 & Firestoreトークン管理")

    if not initialize_firebase():
        st.stop()

    firebase_auth_form()

    if is_user_authenticated():
        st.divider()
        st.subheader("Firestoreトークン管理")

        user_id = get_firebase_user_id()

        col1, col2 = st.columns(2)

        with col1:
            if st.button("トークンを読み込む"):
                process_tokens_safely(user_id)

        with col2:
            if st.button("テストトークンを保存"):
                test_tokens = {
                    "access_token": "test_access_token_123",
                    "refresh_token": "test_refresh_token_456",
                    "expires_at": "2025-07-15T12:00:00Z"
                }
                save_tokens_to_firestore(user_id, test_tokens)

if __name__ == "__main__":
    main()
