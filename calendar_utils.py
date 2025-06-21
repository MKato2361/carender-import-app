import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import json # client_configを直接扱うために追加

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def authenticate_google():
    creds = None

    # 1. まず現在のStreamlitセッションの認証情報がst.session_stateにあるか確認します
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']

        # 認証情報が有効かどうかをチェック
        # Google Credential オブジェクトは自身で有効期限とリフレッシュトークンの有無を持つ
        if creds.valid:
            return creds
        elif creds.expired and creds.refresh_token:
            # トークンが期限切れでリフレッシュトークンがある場合、トークンをリフレッシュします
            try:
                st.info("期限切れの認証トークンを更新しようとしています...")
                creds.refresh(Request())
                st.session_state['credentials'] = creds # リフレッシュされた認証情報を保存
                st.success("認証トークンを更新しました。")
                return creds
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None # 無効な認証情報をクリア
                creds = None # credsをNoneにして再認証フローへ進める
        else:
            # リフレッシュトークンがない、または無効な認証情報の場合
            st.warning("認証トークンが期限切れまたは無効です。再認証が必要です。")
            st.session_state['credentials'] = None # 無効な認証情報をクリア
            creds = None

    # 2. 有効な認証情報がない場合、新しい認証フローを開始します
    if creds is None:
        try:
            # Streamlit Secretsからクライアント情報を取得
            # secrets.tomlの構造と合わせる
            client_config = {
                "installed": {
                    "client_id": st.secrets["google"]["client_id"],
                    "client_secret": st.secrets["google"]["client_secret"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"] # コンソール認証用
                }
            }
            # Flowオブジェクトを作成 (セッションキャッシュを有効にする)
            flow = Flow.from_client_config(client_config, SCOPES,
                                            redirect_uri="urn:ietf:wg:oauth:2.0:oob")

            auth_url, _ = flow.authorization_url(prompt='consent')

            st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
            st.markdown(f"[{auth_url}]({auth_url})") # リンクとして表示
            code = st.text_input("認証コードを貼り付けてください:", key="auth_code_input")

            if code:
                try:
                    flow.fetch_token(code=code)
                    creds = flow.credentials
                    # 認証情報をsession_stateに保存
                    st.session_state['credentials'] = creds
                    st.success("Google認証が完了しました！")
                    st.rerun() # 認証成功後、アプリを再読み込みして認証済み状態にする
                    return creds # 認証が完了したのでcredsを返す
                except Exception as e:
                    st.error(f"認証コードの検証に失敗しました: {e}")
                    st.session_state['credentials'] = None
                    return None # 認証失敗

        except Exception as e:
            st.error(f"Google認証フローの開始に失敗しました。secrets.tomlの設定を確認してください: {e}")
            st.session_state['credentials'] = None
            return None

    return creds # この行はcredsが既に有効な場合にのみ到達する
