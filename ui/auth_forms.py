"""
ui/auth_forms.py
認証 UI フォーム（firebase_auth.py の UI パートを分離）

st.* の使用は許可。ロジックは core/auth/firebase_client.py に委譲。
"""
from __future__ import annotations
import streamlit as st
from core.auth.firebase_client import sign_in, sign_up

_ERROR_MAP = {
    "EMAIL_NOT_FOUND":              "メールアドレスが登録されていません。",
    "INVALID_PASSWORD":             "パスワードが正しくありません。",
    "INVALID_LOGIN_CREDENTIALS":    "メールアドレスまたはパスワードが正しくありません。",
    "USER_DISABLED":                "このアカウントは無効化されています。管理者にお問い合わせください。",
    "TOO_MANY_ATTEMPTS_TRY_LATER":  "ログイン試行が多すぎます。しばらく待ってから再試行してください。",
    "EMAIL_EXISTS":                 "このメールアドレスはすでに登録されています。",
    "WEAK_PASSWORD":                "パスワードは 6 文字以上で設定してください。",
    "INVALID_EMAIL":                "メールアドレスの形式が正しくありません。",
}

def _localize(raw: str) -> str:
    for code, msg in _ERROR_MAP.items():
        if code in raw:
            return msg
    return "エラーが発生しました。しばらく待ってから再試行してください。"


def login_form() -> None:
    """ログイン / 新規登録フォームを描画し、認証状態を session_state に反映する。"""
    st.subheader("ログイン")

    st.session_state.setdefault("user_info", None)
    st.session_state.setdefault("user_email", None)

    if st.session_state["user_info"] is not None:
        st.success(f"ログイン済み: {st.session_state['user_email']}")
        if st.button("ログアウト"):
            st.session_state["user_info"] = None
            st.session_state["user_email"] = None
            for key in ("id_token", "credentials"):
                st.session_state.pop(key, None)
            st.rerun()
        return

    tab_login, tab_signup = st.tabs(["ログイン", "新規登録"])

    with tab_login:
        email    = st.text_input("メールアドレス", key="login_email", placeholder="example@mail.com")
        password = st.text_input("パスワード", type="password", key="login_password")
        if st.button("ログイン", use_container_width=True, type="primary"):
            if email and password:
                with st.spinner("ログイン中..."):
                    result = sign_in(email, password)
                if result["success"]:
                    st.session_state["user_info"]  = result["user_id"]
                    st.session_state["user_email"] = result["email"]
                    st.session_state["id_token"]   = result["id_token"]
                    st.rerun()
                else:
                    st.error(_localize(result["error"]))
            else:
                st.warning("メールアドレスとパスワードを入力してください。")

    with tab_signup:
        new_email    = st.text_input("メールアドレス", key="signup_email", placeholder="example@mail.com")
        new_password = st.text_input("パスワード（6 文字以上）", type="password", key="signup_password")
        if st.button("アカウントを作成", use_container_width=True, type="primary"):
            if new_email and new_password:
                with st.spinner("アカウントを作成中..."):
                    result = sign_up(new_email, new_password)
                if result["success"]:
                    login_result = sign_in(new_email, new_password)
                    if login_result["success"]:
                        st.session_state["user_info"]  = login_result["user_id"]
                        st.session_state["user_email"] = login_result["email"]
                        st.session_state["id_token"]   = login_result["id_token"]
                        st.rerun()
                    else:
                        st.success("登録が完了しました。ログインタブからサインインしてください。")
                else:
                    st.error(_localize(result["error"]))
            else:
                st.warning("メールアドレスとパスワードを入力してください。")
