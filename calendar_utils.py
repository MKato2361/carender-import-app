import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import re
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/tasks"]

def authenticate_google():
    creds = None
    
    # 認証情報がセッションステートに存在し、有効な場合はそれを使用
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds
        # トークンが期限切れの場合はリフレッシュを試みる
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                return creds
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証が必要です: {e}")
                st.session_state['credentials'] = None
                creds = None
                
    # 認証情報がない場合、OAuthフローを開始
    if not creds:
        try:
            client_config = {
                "installed": {
                    "client_id": st.secrets["google"]["client_id"],
                    "client_secret": st.secrets["google"]["client_secret"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"]
                }
            }
            flow = Flow.from_client_config(client_config, SCOPES)
            flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
            
            # 修正：authorization_urlメソッドの呼び出し
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
            st.write(auth_url)
            
            code = st.text_input("認証コードを貼り付けてください:")
            
            if code:
                try:
                    flow.fetch_token(code=code)
                    creds = flow.credentials
                    st.session_state['credentials'] = creds
                    st.success("Google認証が完了しました！")
                    st.rerun()
                except Exception as e:
                    st.error(f"認証コードの処理に失敗しました: {e}")
                    return None
                    
        except Exception as e:
            st.error(f"Google認証に失敗しました: {e}")
            st.session_state['credentials'] = None
            return None
            
    return creds

# 使用例とヘルパー関数
def get_calendar_service(creds):
    """カレンダーサービスを取得"""
    if creds:
        return build('calendar', 'v3', credentials=creds)
    return None

def get_tasks_service(creds):
    """タスクサービスを取得"""
    if creds:
        return build('tasks', 'v1', credentials=creds)
    return None

def logout_google():
    """Google認証をログアウト"""
    if 'credentials' in st.session_state:
        del st.session_state['credentials']
        st.success("ログアウトしました。")
        st.rerun()

# メイン関数での使用例
def main():
    st.title("Google Calendar & Tasks Integration")
    
    # 認証状態の確認
    if 'credentials' in st.session_state and st.session_state['credentials']:
        st.success("Google認証済み")
        if st.button("ログアウト"):
            logout_google()
    else:
        st.warning("Google認証が必要です")
        creds = authenticate_google()
        if creds:
            st.success("認証が完了しました！")

if __name__ == "__main__":
    main()
