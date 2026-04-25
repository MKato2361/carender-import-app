from __future__ import annotations
"""
calendar_utils.py — 後方互換ラッパー（最終形）

実体は以下に移行済み:
  services/auth_service.py       … authenticate_google
  services/calendar_service.py  … CRUD・Tasks UI ラッパー
  core/calendar/crud.py          … Calendar API 操作
  core/calendar/tasks.py         … Tasks API 操作

このファイルは既存 import を壊さないために残す。
新規コードでは services/calendar_service を直接 import すること。
"""
from services.auth_service import authenticate_google          # noqa: F401
from services.calendar_service import (                        # noqa: F401
    get_events           as fetch_all_events,
    add_event_to_calendar,
    update_event_if_needed,
    delete_event_from_calendar,
    add_task_to_todo_list,
    delete_tasks_by_event_id as find_and_delete_tasks_by_event_id,
)
from core.calendar.tasks import build_tasks_service            # noqa: F401
