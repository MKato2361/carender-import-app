"""
core/calendar/tasks.py
Google Tasks API 操作（st.* 禁止）

calendar_utils.py の Tasks 関連を抽出。
"""
from __future__ import annotations
from typing import Optional
from googleapiclient.discovery import build


def build_tasks_service(creds):
    """Tasks API サービスを構築して返す。失敗時は Exception を raise。"""
    return build("tasks", "v1", credentials=creds)


def get_default_task_list_id(tasks_service) -> Optional[str]:
    """'My Tasks' のタスクリスト ID を返す。見つからなければ最初のリストの ID。"""
    resp = tasks_service.tasklists().list().execute()
    items = resp.get("items", [])
    for item in items:
        if item.get("title") == "My Tasks":
            return item["id"]
    return items[0]["id"] if items else None


def add_task(tasks_service, task_list_id: str, task_data: dict) -> dict:
    """タスクを追加する。失敗時は Exception を raise。"""
    return tasks_service.tasks().insert(tasklist=task_list_id, body=task_data).execute()


def find_and_delete_tasks_by_event_id(
    tasks_service, task_list_id: str, event_id: str
) -> int:
    """
    notes に [EVENT_ID:{event_id}] を含むタスクを検索して削除する。
    削除した件数を返す。
    """
    deleted = 0
    page_token = None
    while True:
        resp = tasks_service.tasks().list(
            tasklist=task_list_id, maxResults=100,
            showCompleted=True, showDeleted=False,
            showHidden=False, pageToken=page_token,
        ).execute()
        for task in resp.get("items", []):
            notes = task.get("notes") or ""
            if f"[EVENT_ID:{event_id}]" in notes:
                tasks_service.tasks().delete(
                    tasklist=task_list_id, task=task["id"]
                ).execute()
                deleted += 1
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return deleted
