# 省略：既存の import や認証関数などはそのまま

SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks"
]

# 省略：add_event_to_calendar, delete_events_from_calendar, fetch_all_events などもそのまま

def create_tasks_for_event(service, task_service, title, due_datetime):
    task_titles = [
        f"{title} - 点検通知（FAX）",
        f"{title} - 点検通知（電話）",
        f"{title} - 貼紙"
    ]
    task_list_id = task_service.tasklists().list().execute()["items"][0]["id"]
    for task_title in task_titles:
        task = {
            'title': task_title,
            'due': due_datetime.isoformat() + 'Z'
        }
        task_service.tasks().insert(tasklist=task_list_id, body=task).execute()
