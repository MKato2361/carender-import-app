"""
core/calendar/crud.py
Google Calendar API CRUD 操作（st.* 禁止）

全関数がエラーを raise する。st.error 等の表示は呼び出し元が担う。
"""
from __future__ import annotations
from typing import Optional
from googleapiclient.errors import HttpError


def fetch_all_events(service, calendar_id: str,
                     time_min: Optional[str] = None,
                     time_max: Optional[str] = None) -> list[dict]:
    """指定期間のイベントをページネーションで全件取得する。"""
    events, page_token = [], None
    while True:
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min, timeMax=time_max,
            singleEvents=True, orderBy="startTime",
            maxResults=250, pageToken=page_token,
        ).execute()
        events.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return events


def add_event(service, calendar_id: str, event_data: dict) -> dict:
    """イベントを追加する。"""
    return service.events().insert(calendarId=calendar_id, body=event_data).execute()


def update_event(service, calendar_id: str, event_id: str, event_data: dict) -> dict:
    """イベントを上書き更新する。"""
    return service.events().update(
        calendarId=calendar_id, eventId=event_id, body=event_data
    ).execute()


def update_event_if_changed(service, calendar_id: str,
                             event_id: str, new_data: dict) -> Optional[dict]:
    """
    既存イベントと new_data を比較し、差分がある場合のみ更新する。
    差分なし → 既存イベントをそのまま返す。
    """
    existing = service.events().get(calendarId=calendar_id, eventId=event_id).execute()

    nz = lambda v: v or ""
    changed = (
        nz(existing.get("summary"))      != nz(new_data.get("summary"))
        or nz(existing.get("description")) != nz(new_data.get("description"))
        or nz(existing.get("transparency")) != nz(new_data.get("transparency"))
        or (existing.get("recurrence") or []) != (new_data.get("recurrence") or [])
        or (existing.get("start") or {})  != (new_data.get("start") or {})
        or (existing.get("end")   or {})  != (new_data.get("end")   or {})
    )
    if changed:
        return update_event(service, calendar_id, event_id, new_data)
    return existing


def delete_event(service, calendar_id: str, event_id: str) -> None:
    """イベントを削除する。"""
    service.events().delete(calendarId=calendar_id, eventId=event_id).execute()


def get_calendar_list(service) -> list[dict]:
    """書き込み可能なカレンダー一覧を返す。"""
    resp = service.calendarList().list().execute()
    return [c for c in resp.get("items", []) if c.get("accessRole") != "reader"]
