import streamlit as st
import firebase_admin
from firebase_admin import credentials, auth, firestore # firestoreを追加
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import json # 必要に応じて使用
import re
from datetime import datetime, timedelta, timezone

# 外部モジュールからFirebaseユーザーIDを取得
from firebase_auth import get_firebase_user_id

# 認証スコープ
SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/tasks"]

def authenticate_google():
    creds = None
    user_id = get_firebase_user_id()
    
    if not user_id:
        # Firebaseユーザーが認証されていない場合はGoogle認証も行わない
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # 1. セッションステートから認証情報を確認 (高速化のため)
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds

    # 2. Firestoreから永続化された認証情報を読み込む
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            st.session_state['credentials'] = creds # セッションステートにも保存
            
            # トークンが期限切れの場合はリフレッシュを試みる
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                # 成功したらFirestoreのトークンも更新
                doc_ref.set(creds.to_json())
                st.info("認証トークンを更新しました。")
                st.rerun()
                
            return creds
    except Exception as e:
        st.error(f"Firestoreからのトークン読み込みに失敗しました: {e}")
        creds = None

    # 3. 認証情報がまだない場合、OAuthフローを開始
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
            auth_url, _ = flow.authorization_url(prompt='consent')

            st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
            st.write(auth_url)
            code = st.text_input("認証コードを貼り付けてください:")

            if code:
                flow.fetch_token(code=code)
                creds = flow.credentials
                st.session_state['credentials'] = creds
                
                # 認証が完了したらFirestoreに保存
                doc_ref.set(json.loads(creds.to_json()))
                
                st.success("Google認証が完了しました！")
                st.rerun()
        except Exception as e:
            st.error(f"Google認証に失敗しました: {e}")
            st.session_state['credentials'] = None
            return None

    return creds

def build_tasks_service(creds):
    """Google Tasks サービスをビルドする"""
    try:
        if not creds:
            return None
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        st.warning(f"Google Tasks サービスのビルドに失敗しました: {e}")
        return None

def add_event_to_calendar(service, calendar_id, event_data):
    """カレンダーにイベントを追加する"""
    try:
        event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
        return event
    except HttpError as e:
        st.error(f"イベントの追加に失敗しました (HTTPエラー): {e}")
        return None
    except Exception as e:
        st.error(f"イベントの追加に失敗しました: {e}")
        return None

def fetch_all_events(service, calendar_id, time_min=None, time_max=None):
    """カレンダーからすべてのイベントを取得する"""
    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        
        events = events_result.get('items', [])
        return events
    except HttpError as e:
        st.error(f"イベントの取得に失敗しました (HTTPエラー): {e}")
        return []
    except Exception as e:
        st.error(f"イベントの取得に失敗しました: {e}")
        return []

def update_event_if_needed(service, calendar_id, event_id, updated_event_data):
    """必要に応じてイベントを更新する"""
    try:
        # 既存のイベントを取得
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        
        # 更新が必要かチェック（簡単な比較）
        needs_update = False
        for key, value in updated_event_data.items():
            if existing_event.get(key) != value:
                needs_update = True
                break
        
        if needs_update:
            updated_event = service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=updated_event_data
            ).execute()
            return updated_event
        else:
            return existing_event
            
    except HttpError as e:
        st.error(f"イベントの更新に失敗しました (HTTPエラー): {e}")
        return None
    except Exception as e:
        st.error(f"イベントの更新に失敗しました: {e}")
        return None

def add_task_to_todo_list(tasks_service, task_list_id, task_data):
    """ToDoリストにタスクを追加する"""
    try:
        if not tasks_service:
            return None
            
        task = tasks_service.tasks().insert(
            tasklist=task_list_id,
            body=task_data
        ).execute()
        
        return task
    except HttpError as e:
        st.error(f"タスクの追加に失敗しました (HTTPエラー): {e}")
        return None
    except Exception as e:
        st.error(f"タスクの追加に失敗しました: {e}")
        return None

def find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id):
    """イベントIDに基づいてタスクを検索・削除する"""
    try:
        if not tasks_service:
            return False
            
        # タスクを取得
        tasks_result = tasks_service.tasks().list(tasklist=task_list_id).execute()
        tasks = tasks_result.get('items', [])
        
        deleted_count = 0
        for task in tasks:
            # タスクのノートやタイトルにevent_idが含まれているかチェック
            if (event_id in task.get('notes', '') or 
                event_id in task.get('title', '')):
                try:
                    tasks_service.tasks().delete(
                        tasklist=task_list_id,
                        task=task['id']
                    ).execute()
                    deleted_count += 1
                except Exception as e:
                    st.warning(f"タスクの削除に失敗しました: {e}")
        
        return deleted_count > 0
        
    except HttpError as e:
        st.error(f"タスクの検索・削除に失敗しました (HTTPエラー): {e}")
        return False
    except Exception as e:
        st.error(f"タスクの検索・削除に失敗しました: {e}")
        return False

def delete_event_from_calendar(service, calendar_id, event_id):
    """カレンダーからイベントを削除する"""
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        return True
    except HttpError as e:
        st.error(f"イベントの削除に失敗しました (HTTPエラー): {e}")
        return False
    except Exception as e:
        st.error(f"イベントの削除に失敗しました: {e}")
        return False

def format_event_for_calendar(title, start_datetime, end_datetime, description="", location=""):
    """カレンダーイベント用のデータ形式を作成する"""
    event_data = {
        'summary': title,
        'start': {
            'dateTime': start_datetime.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'end': {
            'dateTime': end_datetime.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'description': description,
    }
    
    if location:
        event_data['location'] = location
    
    return event_data

def format_task_for_todo_list(title, notes="", due_date=None):
    """ToDoリスト用のタスクデータ形式を作成する"""
    task_data = {
        'title': title,
        'notes': notes,
    }
    
    if due_date:
        task_data['due'] = due_date.isoformat() + 'Z'
    
    return task_data

def get_calendar_colors():
    """カレンダーで使用可能な色を取得する"""
    return {
        'デフォルト': '1',
        'ラベンダー': '2',
        'セージ': '3',
        'ぶどう': '4',
        'フラミンゴ': '5',
        'バナナ': '6',
        'マンダリン': '7',
        'ピーコック': '8',
        'グラファイト': '9',
        'バジル': '10',
        'トマト': '11'
    }

def validate_datetime(date_str, time_str):
    """日付と時刻の文字列を検証し、datetimeオブジェクトを返す"""
    try:
        # 日付の解析
        if isinstance(date_str, str):
            date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
        else:
            date_obj = date_str
        
        # 時刻の解析
        if isinstance(time_str, str):
            time_obj = datetime.strptime(time_str, '%H:%M').time()
        else:
            time_obj = time_str
        
        # datetimeオブジェクトの作成
        dt = datetime.combine(date_obj, time_obj)
        
        # タイムゾーンを設定（日本時間）
        dt = dt.replace(tzinfo=timezone(timedelta(hours=9)))
        
        return dt
        
    except ValueError as e:
        st.error(f"日付または時刻の形式が正しくありません: {e}")
        return None
