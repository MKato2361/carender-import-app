import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from datetime import datetime, timedelta, timezone
import json

SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/tasks"]

def authenticate_google():
    """Google認証を処理する"""
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
            # secrets.tomlからクライアント情報を取得
            if "google" not in st.secrets:
                st.error("Google認証情報が設定されていません。secrets.tomlを確認してください。")
                return None
                
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
            
            # 認証URLを生成
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
            st.code(auth_url)
            
            code = st.text_input("認証コードを貼り付けてください:", type="password")
            
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
