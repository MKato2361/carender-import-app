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

def authenticate_user(email, password):
    """Firebase REST APIを使用してユーザー認証"""
    try:
        # Firebase Web API Key（st.secretsに追加が必要）
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
        return {
            "success": False,
            "error": str(e)
        }

def create_user_account(email, password):
    """Firebase REST APIを使用してユーザー作成"""
    try:
        # Firebase Web API Key
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
        return {
            "success": False,
            "error": str(e)
        }

_FIREBASE_ERROR_MAP = {
    "EMAIL_NOT_FOUND":             "メールアドレスが登録されていません。",
    "INVALID_PASSWORD":            "パスワードが正しくありません。",
    "INVALID_LOGIN_CREDENTIALS":   "メールアドレスまたはパスワードが正しくありません。",
    "USER_DISABLED":               "このアカウントは無効化されています。管理者にお問い合わせください。",
    "TOO_MANY_ATTEMPTS_TRY_LATER": "ログイン試行が多すぎます。しばらく待ってから再試行してください。",
    "EMAIL_EXISTS":                "このメールアドレスはすでに登録されています。",
    "WEAK_PASSWORD":               "パスワードは6文字以上で設定してください。",
    "INVALID_EMAIL":               "メールアドレスの形式が正しくありません。",
}

def _localize_firebase_error(raw: str) -> str:
    for code, msg in _FIREBASE_ERROR_MAP.items():
        if code in raw:
            return msg
    return "エラーが発生しました。しばらく待ってから再試行してください。"


def firebase_auth_form():
    """ログイン/サインアップのUIを表示し、認証状態を管理する"""
    st.subheader("ログイン")

    if "user_info" not in st.session_state:
        st.session_state.user_info = None
    if "user_email" not in st.session_state:
        st.session_state.user_email = None

    if st.session_state.user_info is None:
        tab_login, tab_signup = st.tabs(["ログイン", "新規登録"])

        with tab_login:
            email = st.text_input("メールアドレス", key="login_email", placeholder="example@mail.com")
            password = st.text_input("パスワード", type="password", key="login_password")

            if st.button("ログイン", use_container_width=True, type="primary"):
                if email and password:
                    with st.spinner("ログイン中..."):
                        result = authenticate_user(email, password)
                    if result["success"]:
                        st.session_state.user_info = result["user_id"]
                        st.session_state.user_email = result["email"]
                        st.session_state.id_token = result["id_token"]
                        st.rerun()
                    else:
                        st.error(_localize_firebase_error(result["error"]))
                else:
                    st.warning("メールアドレスとパスワードを入力してください。")

        with tab_signup:
            new_email = st.text_input("メールアドレス", key="signup_email", placeholder="example@mail.com")
            new_password = st.text_input("パスワード（6文字以上）", type="password", key="signup_password")

            if st.button("アカウントを作成", use_container_width=True, type="primary"):
                if new_email and new_password:
                    with st.spinner("アカウントを作成中..."):
                        result = create_user_account(new_email, new_password)
                    if result["success"]:
                        # 登録後にそのままログイン
                        login_result = authenticate_user(new_email, new_password)
                        if login_result["success"]:
                            st.session_state.user_info = login_result["user_id"]
                            st.session_state.user_email = login_result["email"]
                            st.session_state.id_token = login_result["id_token"]
                            st.rerun()
                        else:
                            st.success("登録が完了しました。ログインタブからサインインしてください。")
                    else:
                        st.error(_localize_firebase_error(result["error"]))
                else:
                    st.warning("メールアドレスとパスワードを入力してください。")

    else:
        st.success(f"ログイン済み: {st.session_state.user_email}")
        if st.button("ログアウト"):
            st.session_state.user_info = None
            st.session_state.user_email = None
            for key in ("id_token", "credentials"):
                st.session_state.pop(key, None)
            st.rerun()
def get_firebase_user_id():
    """現在の認証済みユーザーIDを返す"""
    return st.session_state.get("user_info")

def get_firebase_user_email():
    """現在の認証済みユーザーのメールアドレスを返す"""
    return st.session_state.get("user_email")

def get_firebase_id_token():
    """現在の認証済みユーザーのIDトークンを返す"""
    return st.session_state.get("id_token")

def is_user_authenticated():
    """ユーザーが認証済みかどうかを確認"""
    return st.session_state.get("user_info") is not None

# Firestore関連の関数
def safe_load_tokens_from_firestore(user_id):
    """Firestoreからトークンを安全に読み込む"""
    try:
        db = firestore.client()
        doc_ref = db.collection('users').document(user_id)
        doc = doc_ref.get()

        if not doc.exists:
            return {}

        token_data = doc.get('tokens') or doc.get('token') or doc.get('credentials')
        if token_data is None:
            return {}

        if isinstance(token_data, dict):
            return token_data
        elif isinstance(token_data, str):
            try:
                parsed = json.loads(token_data)
                return parsed if isinstance(parsed, dict) else {}
            except json.JSONDecodeError:
                return {}
        elif isinstance(token_data, list):
            return {"tokens": token_data}
        return {}

    except Exception as e:
        return {}
def save_tokens_to_firestore(user_id, tokens):
    """トークンをFirestoreに保存する"""
    try:
        db = firestore.client()
        doc_ref = db.collection('users').document(user_id)
        
        # トークンデータを辞書として保存
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
    
    # トークンデータのキー一覧のみ表示（値は非表示）
    if isinstance(tokens, dict):
        st.caption(f"トークンキー: {list(tokens.keys())}")
    else:
        st.caption("トークン取得済み")

def main():
    """メイン関数"""
    st.title("Firebase認証 & Firestoreトークン管理")
    
    # Firebase初期化
    if not initialize_firebase():
        st.stop()
    
    # 認証フォーム
    firebase_auth_form()
    
    # 認証済みユーザーの場合のみFirestore操作を表示
    if is_user_authenticated():
        st.divider()
        st.subheader("Firestoreトークン管理")
        
        user_id = get_firebase_user_id()
        
        # 現在のユーザー情報を表示
        st.caption(f"ユーザーID: {user_id}")
        
        # トークン操作ボタン
        if st.button("トークンを読み込む"):
            process_tokens_safely(user_id)
        
        if st.button("テストトークンを保存"):
            test_tokens = {
                "access_token": "test_access_token_123",
                "refresh_token": "test_refresh_token_456",
                "expires_at": "2025-07-15T12:00:00Z"
            }
            save_tokens_to_firestore(user_id, test_tokens)
        
        # Firestoreの接続状態を確認
        if st.button("Firestore接続テスト"):
            try:
                db = firestore.client()
                # テストドキュメントを作成
                test_ref = db.collection('test').document('connection_test')
                test_ref.set({
                    'timestamp': firestore.SERVER_TIMESTAMP,
                    'user_id': user_id
                })
                st.success("Firestoreへの接続が成功しました。")
            except Exception as e:
                st.error(f"Firestore接続エラー: {e}")

if __name__ == "__main__":
    main()
