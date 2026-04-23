"""
core/calendar/crud.py
Google Calendar API CRUD 操作（st.* 禁止）

calendar_utils.py から UI 呼び出しを除去して抽出。
エラーは例外として呼び出し元に伝え、表示は ui 層が担う。
"""
from __future__ import annotations
from typing import Optional
from googleapiclient.errors import HttpError


def fetch_all_events(service, calendar_id: str, time_min: str, time_max: str) -> list[dict]:
    """
    指定期間内のイベントをページネーションで全件取得する。
    失敗時は HttpError を raise する。
    """
    events: list[dict] = []
    page_token: Optional[str] = None
    while True:
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
            maxResults=250,
            pageToken=page_token,
        ).execute()
        events.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return events


def add_event(service, calendar_id: str, event_data: dict) -> dict:
    """イベントを追加する。失敗時は HttpError / Exception を raise。"""
    return service.events().insert(calendarId=calendar_id, body=event_data).execute()


def update_event(service, calendar_id: str, event_id: str, event_data: dict) -> dict:
    """イベントを上書き更新する。失敗時は HttpError / Exception を raise。"""
    return service.events().update(
        calendarId=calendar_id, eventId=event_id, body=event_data
    ).execute()


def delete_event(service, calendar_id: str, event_id: str) -> None:
    """イベントを削除する。失敗時は HttpError / Exception を raise。"""
    service.events().delete(calendarId=calendar_id, eventId=event_id).execute()


def get_calendar_list(service) -> list[dict]:
    """書き込み可能なカレンダー一覧を返す。"""
    resp = service.calendarList().list().execute()
    return [c for c in resp.get("items", []) if c.get("accessRole") != "reader"]
