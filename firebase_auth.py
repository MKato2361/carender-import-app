import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth
import json
from typing import Optional, Dict, Any

def initialize_firebase() -> bool:
    """
    Firebaseアプリを初期化する
    
    Returns:
        初期化が成功したかどうか
    """
    try:
        # 既に初期化されているかチェック
        if firebase_admin._apps:
            return True
        
        # Firebase認証情報を取得
        if "firebase" not in st.secrets:
            st.error("Firebase認証情報が設定されていません。secrets.tomlを確認してください。")
            return False
        
        # 認証情報を辞書形式で取得
        firebase_config = dict(st.secrets["firebase"])
        
        # 認証情報オブジェクトを作成
        cred = credentials.Certificate(firebase_config)
        
        # Firebaseアプリを初期化
        firebase_admin.initialize_app(cred)
        
        return True
        
    except Exception as e:
        st.error(f"Firebase初期化エラー: {e}")
        return False

def firebase_auth_form():
    """
    Firebase認証フォームを表示する
    """
    st.subheader("🔐 ユーザー認証")
    
    # タブでログインとサインアップを分ける
    login_tab, signup_tab = st.tabs(["ログイン", "サインアップ"])
    
    with login_tab:
        st.markdown("### ログイン")
        login_email = st.text_input("メールアドレス", key="login_email")
        login_password = st.text_input("パスワード", type="password", key="login_password")
        
        if st.button("ログイン", key="login_button"):
            if login_email and login_password:
                user = authenticate_user(login_email, login_password)
                if user:
                    st.session_state['firebase_user'] = user
                    st.success("ログインに成功しました！")
                    st.rerun()
                else:
                    st.error("ログインに失敗しました。メールアドレスとパスワードを確認してください。")
            else:
                st.error("メールアドレスとパスワードを入力してください。")
    
    with signup_tab:
        st.markdown("### サインアップ")
        signup_email = st.text_input("メールアドレス", key="signup_email")
        signup_password = st.text_input("パスワード", type="password", key="signup_password")
        signup_password_confirm = st.text_input("パスワード（確認）", type="password", key="signup_password_confirm")
        
        if st.button("サインアップ", key="signup_button"):
            if signup_email and signup_password and signup_password_confirm:
                if signup_password == signup_password_confirm:
                    user = create_user(signup_email, signup_password)
                    if user:
                        st.session_state['firebase_user'] = user
                        st.success("サインアップに成功しました！")
                        st.rerun()
                    else:
                        st.error("サインアップに失敗しました。")
                else:
                    st.error("パスワードが一致しません。")
            else:
                st.error("すべてのフィールドを入力してください。")

def authenticate_user(email: str, password: str) -> Optional[Dict[str, Any]]:
    """
    ユーザーを認証する（簡易版）
    
    Note: この実装は簡易版です。実際の本番環境では、
    Firebase Client SDKを使用したフロントエンド認証を推奨します。
    
    Args:
        email: メールアドレス
        password: パスワード
    
    Returns:
        ユーザー情報、または認証失敗時はNone
    """
    try:
        # Firebase Admin SDKでは直接パスワード認証はできないため、
        # 実際のアプリケーションでは Firebase Auth REST API を使用するか、
        # フロントエンドでFirebase Client SDKを使用する必要があります。
        
        # ここでは簡易的な実装として、メールアドレスでユーザーを取得
        user = auth.get_user_by_email(email)
        
        # 実際の認証は別途実装が必要
        # この例では、ユーザーが存在すれば認証成功とみなす
        return {
            'uid': user.uid,
            'email': user.email,
            'display_name': user.display_name,
            'email_verified': user.email_verified
        }
        
    except auth.UserNotFoundError:
        st.error("ユーザーが見つかりません。")
        return None
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return None

def create_user(email: str, password: str) -> Optional[Dict[str, Any]]:
    """
    新しいユーザーを作成する
    
    Args:
        email:
