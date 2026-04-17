import json
import urllib.parse
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from requests_oauthlib import OAuth2Session
from google.cloud import firestore
import streamlit as st


def authenticate_google():
    # ============================================================
    # セッション初期化（必須）
    # ============================================================
    if "credentials" not in st.session_state:
        st.session_state["credentials"] = None

    if "credentials_user_id" not in st.session_state:
        st.session_state["credentials_user_id"] = None

    # ============================================================
    # 強制リセット
    # ============================================================
    if st.query_params.get("clear_auth") == "1":
        user_id = get_firebase_user_id()
        if user_id:
            try:
                firestore.client().collection('google_tokens').document(user_id).delete()
            except Exception:
                pass

        st.session_state.pop('credentials', None)
        st.session_state.pop('credentials_user_id', None)

        st.query_params.clear()
        st.rerun()

    # ============================================================
    # user_id取得
    # ============================================================
    user_id = get_firebase_user_id()
    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    creds = None

    # ============================================================
    # ① セッションから取得（user一致チェック）
    # ============================================================
    if (
        "credentials" in st.session_state and
        st.session_state["credentials"] and
        st.session_state.get("credentials_user_id") == user_id
    ):
        creds = st.session_state["credentials"]

        if creds:
            if not creds.refresh_token:
                st.session_state.pop('credentials', None)
                st.session_state.pop('credentials_user_id', None)
                try:
                    doc_ref.delete()
                except Exception:
                    pass
                creds = None

            elif creds.valid:
                return creds

            elif creds.expired:
                try:
                    creds.refresh(Request())
                    st.session_state["credentials"] = creds
                    st.session_state["credentials_user_id"] = user_id
                    doc_ref.set(json.loads(creds.to_json()))
                    return creds
                except Exception:
                    st.session_state.pop('credentials', None)
                    st.session_state.pop('credentials_user_id', None)
                    try:
                        doc_ref.delete()
                    except Exception:
                        pass
                    creds = None

    # ============================================================
    # ② Firestoreから取得
    # ============================================================
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()

            try:
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            except Exception:
                creds = None

            if creds:
                if not creds.refresh_token:
                    doc_ref.delete()
                    creds = None

                elif creds.expired:
                    try:
                        creds.refresh(Request())
                        st.session_state["credentials"] = creds
                        st.session_state["credentials_user_id"] = user_id
                        doc_ref.set(json.loads(creds.to_json()))
                        return creds
                    except Exception:
                        doc_ref.delete()
                        creds = None

                elif creds.valid:
                    st.session_state["credentials"] = creds
                    st.session_state["credentials_user_id"] = user_id
                    return creds
    except Exception:
        pass

    # ============================================================
    # ③ OAuthフロー
    # ============================================================
    try:
        client_id     = st.secrets["google"]["client_id"]
        client_secret = st.secrets["google"]["client_secret"]
        redirect_uri  = st.secrets["google"]["redirect_uri"]

        AUTH_URI  = "https://accounts.google.com/o/oauth2/auth"
        TOKEN_URI = "https://oauth2.googleapis.com/token"

        params = st.query_params

        # -------------------------------
        # 認証開始
        # -------------------------------
        if "code" not in params:
            oauth = OAuth2Session(
                client_id=client_id,
                redirect_uri=redirect_uri,
                scope=SCOPES,
            )

            auth_url, state = oauth.authorization_url(
                AUTH_URI,
                access_type="offline",
                prompt="consent",
                include_granted_scopes="true",
            )

            # stateはGoogleがコールバック時に返してくれるので保存不要
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()

        # -------------------------------
        # コールバック処理
        # -------------------------------
        else:
            # GoogleがコールバックURLに付けて返してくれるstateをそのまま使う
            state = params.get("state")

            if not state:
                st.warning("セッションが切れました。再度ログインしてください。")
                st.session_state.pop("credentials", None)
                st.session_state.pop("credentials_user_id", None)
                st.query_params.clear()
                st.stop()

            oauth = OAuth2Session(
                client_id=client_id,
                redirect_uri=redirect_uri,
                scope=SCOPES,
                state=state,
            )

            # URLエンコードして正確にcurrent_urlを組み立てる
            current_url = redirect_uri + "?" + "&".join(
                f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items()
            )

            token = oauth.fetch_token(
                TOKEN_URI,
                authorization_response=current_url,
                client_secret=client_secret,
            )

            creds = Credentials(
                token=token.get("access_token"),
                refresh_token=token.get("refresh_token"),
                token_uri=TOKEN_URI,
                client_id=client_id,
                client_secret=client_secret,
                scopes=SCOPES,
            )

            if not creds or not creds.refresh_token:
                st.error("認証に失敗しました（トークン不正）")
                return None

            st.session_state["credentials"] = creds
            st.session_state["credentials_user_id"] = user_id

            doc_ref.set(json.loads(creds.to_json()))

            st.success("Google認証が完了しました！")

            st.query_params.clear()
            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state.pop('credentials', None)
        st.session_state.pop('credentials_user_id', None)
        return None

    return None
    
    from googleapiclient.discovery import build

def build_tasks_service(creds):
    """Google Tasks API サービスオブジェクトを生成して返す"""
    return build("tasks", "v1", credentials=creds)
