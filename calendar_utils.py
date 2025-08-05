import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth, firestore
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import json
import re
from datetime import datetime, timedelta, timezone
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests

from firebase_auth import get_firebase_user_id, initialize_firebase

# Firebaseの初期化
initialize_firebase()

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]

def authenticate_google():
    """
    Google OAuth認証を処理し、認証情報を返す関数。
    Webアプリケーション向けの認証フローに修正
    """
    creds = None
    user_id = get_firebase_user_id()

    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # 1. セッションステートから認証情報を確認
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds

    # 2. Firestoreからトークンを読み込む
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            st.session_state['credentials'] = creds
            
            # トークンの有効期限を確認し、必要ならリフレッシュ
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                token_data = json.loads(creds.to_json())
                doc_ref.set(token_data) # Firestoreに更新後のトークンを保存
                st.session_state['credentials'] = creds
            if creds.valid:
                return creds
            
    except Exception as e:
        st.error(f"Firestoreからの認証情報読み込みに失敗しました: {e}")
        st.session_state['credentials'] = None
        return None

    # 3. 認証情報がない場合、OAuthフローを開始
    client_config = st.secrets["google_oauth"]
    
    # Webアプリケーション向けにリダイレクトURIを設定
    redirect_uri = client_config["redirect_uris"][0]
    
    flow = Flow.from_client_config(
        client_config, 
        scopes=SCOPES, 
        redirect_uri=redirect_uri
    )

    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    # 認証URLをユーザーに表示
    st.write("Googleカレンダーと連携するには、以下のリンクをクリックしてください。")
    st.markdown(f"[Google認証ページへ移動する]({auth_url})", unsafe_allow_html=True)
    
    # ユーザーが認証を完了し、リダイレクトされた後、URLに認証コードが付与される
    if "code" in st.query_params:
        try:
            flow.fetch_token(code=st.query_params["code"])
            creds = flow.credentials
            st.session_state['credentials'] = creds

            # 認証情報をFirestoreに保存
            token_data = json.loads(creds.to_json())
            doc_ref.set(token_data)
            
            st.experimental_rerun() # 認証完了後、画面を再描画して認証済み状態にする
        except Exception as e:
            st.error(f"トークンの取得に失敗しました: {e}")
            st.session_state['credentials'] = None
    
    return None

def get_google_service(creds, service_name='calendar', version='v3'):
    """
    Google APIサービスを認証情報を使って構築する
    """
    if creds is None or not creds.valid:
        return None
    try:
        service = build(service_name, version, credentials=creds)
        return service
    except HttpError as error:
        st.error(f"Google APIサービス構築中にエラーが発生しました: {error}")
        return None

def get_all_calendars(service):
    """
    ユーザーがアクセス可能なすべてのカレンダーのリストを取得する
    """
    if not service:
        st.error("Google認証がされていません。")
        return []
    
    try:
        calendar_list_result = service.calendarList().list().execute()
        calendars = calendar_list_result.get('items', [])
        return calendars
    except HttpError as e:
        st.error(f"カレンダーリストの取得に失敗しました: {e}")
        return []

def get_calendar_id_by_summary(calendars, summary):
    """
    指定されたsummary（カレンダー名）を持つカレンダーのIDを検索する
    """
    for calendar in calendars:
        if calendar.get('summary') == summary:
            return calendar.get('id')
    return None

def create_event(service, calendar_id, event_summary, start_time_str, end_time_str, timezone_str):
    """
    指定されたカレンダーに新しいイベントを作成する
    """
    if not service:
        st.error("Google認証がされていません。")
        return None
    
    try:
        event = {
            'summary': event_summary,
            'start': {
                'dateTime': start_time_str,
                'timeZone': timezone_str,
            },
            'end': {
                'dateTime': end_time_str,
                'timeZone': timezone_str,
            },
        }

        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        return created_event
    except HttpError as e:
        st.error(f"イベント作成中にエラーが発生しました: {e}")
        return None

def delete_events_by_summary(service, calendar_id, event_summary):
    """
    指定されたカレンダーから、指定されたサマリーを持つすべてのイベントを削除する
    """
    if not service:
        st.error("Google認証がされていません。")
        return 0
    
    try:
        # イベントを検索
        events_result = service.events().list(
            calendarId=calendar_id,
            q=event_summary,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        
        events_to_delete = events_result.get('items', [])
        
        if not events_to_delete:
            st.info(f"'{event_summary}' というイベントは見つかりませんでした。")
            return 0

        deleted_count = 0
        for event in events_to_delete:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            deleted_count += 1
            
        return deleted_count
    except HttpError as e:
        st.error(f"イベント削除中にエラーが発生しました: {e}")
        return 0

def convert_excel_date_to_datetime_utc(excel_date):
    """
    Excelのシリアル値をUTCのdatetimeオブジェクトに変換する
    """
    try:
        base_date = datetime(1899, 12, 30, tzinfo=timezone.utc)
        if isinstance(excel_date, (int, float)):
            delta = timedelta(days=excel_date)
            # Excelの閏年バグ（1900年2月29日）を考慮して1日引く
            if excel_date > 60:
                delta -= timedelta(days=1)
            utc_datetime = base_date + delta
            return utc_datetime
    except Exception as e:
        st.error(f"日付の変換に失敗しました: {e}")
    return None

def normalize_date_string(date_str):
    """
    さまざまな日付文字列を 'YYYY-MM-DD' 形式に正規化する
    """
    formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日",
        "%m/%d/%Y", "%m-%d-%Y",
        "%B %d, %Y", "%d %B, %Y"
    ]
    
    # 全角数字を半角に変換
    date_str = date_str.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    
    st.error(f"日付形式 '{date_str}' を解析できませんでした。'YYYY-MM-DD' または 'YYYY/MM/DD' 形式を使用してください。")
    return None
