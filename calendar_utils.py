import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ===== Google Tasks サービス構築 =====
def build_tasks_service(creds):
    """Google Tasks API サービスを構築"""
    try:
        if not creds:
            return None
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        st.warning(f"Google Tasks サービスのビルドに失敗しました: {e}")
        return None

# ===== カレンダー操作 =====
def add_event_to_calendar(service, calendar_id, event_data):
    """Googleカレンダーにイベントを追加"""
    try:
        return service.events().insert(calendarId=calendar_id, body=event_data).execute()
    except HttpError as e:
        st.error(f"イベント追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント追加失敗: {e}")
    return None

def fetch_all_events(service, calendar_id, time_min=None, time_max=None):
    """指定期間内のイベントを取得"""
    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        return events_result.get('items', [])
    except HttpError as e:
        st.error(f"イベント取得失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント取得失敗: {e}")
    return []

def update_event_if_needed(service, calendar_id, event_id, updated_event_data):
    """イベントの内容が変わっていれば更新"""
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        needs_update = False
        for key, value in updated_event_data.items():
            if existing_event.get(key) != value:
                needs_update = True
                break
        if needs_update:
            return service.events().update(calendarId=calendar_id, eventId=event_id, body=updated_event_data).execute()
        return existing_event
    except HttpError as e:
        st.error(f"イベント更新失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント更新失敗: {e}")
    return None

def delete_event_from_calendar(service, calendar_id, event_id):
    """イベントを削除"""
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        return True
    except HttpError as e:
        st.error(f"イベント削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"イベント削除失敗: {e}")
    return False

# ===== タスク操作 =====
def add_task_to_todo_list(tasks_service, task_list_id, task_data):
    """ToDoリストにタスクを追加"""
    try:
        return tasks_service.tasks().insert(tasklist=task_list_id, body=task_data).execute()
    except HttpError as e:
        st.error(f"タスク追加失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク追加失敗: {e}")
    return None

def find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id):
    """イベントIDに関連するタスクを検索して削除"""
    try:
        tasks_result = tasks_service.tasks().list(tasklist=task_list_id).execute()
        tasks = tasks_result.get('items', [])
        deleted_count = 0
        for task in tasks:
            if (event_id in task.get('notes', '') or event_id in task.get('title', '')):
                tasks_service.tasks().delete(tasklist=task_list_id, task=task['id']).execute()
                deleted_count += 1
        return deleted_count
    except HttpError as e:
        st.error(f"タスク検索・削除失敗 (HTTPエラー): {e}")
    except Exception as e:
        st.error(f"タスク検索・削除失敗: {e}")
    return 0
