import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import re
from icalendar import Calendar, Event, vUri # Import Calendar, Event, and vUri from icalendar

# SCOPESにGoogle Tasksのスコープを追加
SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/tasks"]

def authenticate_google():
    creds = None

    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                st.rerun()
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None
                creds = None
        else:
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
                    st.success("Google認証が完了しました！")
                    st.rerun()
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['credentials'] = None
                return None

    return creds

def build_tasks_service(creds):
    """Google ToDoリストサービスを構築する"""
    return build("tasks", "v1", credentials=creds)

def add_task_to_todo_list(service, task_list_id, title, due_date: datetime.date = None, notes: str = None):
    """
    指定されたToDoリストにタスクを追加する。
    :param service: Google Tasks APIサービスオブジェクト
    :param task_list_id: タスクを追加するToDoリストのID
    :param title: タスクのタイトル
    :param due_date: タスクの期限 (datetime.dateオブジェクト)
    :param notes: タスクの詳細（メモ）
    """
    task_body = {
        'title': title
    }
    if due_date:
        # RFC 3339 format (YYYY-MM-DDTHH:MM:SS.sssZ) に変換
        # ToDoリストのdueはUTCで時刻まで必要なので、JSTの0時0分0秒に設定し、UTCに変換
        # 日本時間 (JST) のタイムゾーンオフセット
        JST = timezone(timedelta(hours=9))
        # 期限日の開始時刻をJSTで指定
        due_datetime_jst = datetime(due_date.year, due_date.month, due_date.day, 0, 0, 0, tzinfo=JST)
        # UTCに変換
        due_datetime_utc = due_datetime_jst.astimezone(timezone.utc)
        task_body['due'] = due_datetime_utc.isoformat(timespec='milliseconds').replace('+00:00', 'Z')

    if notes:
        task_body['notes'] = notes

    try:
        task = service.tasks().insert(tasklist=task_list_id, body=task_body).execute()
        return task
    except Exception as e:
        st.error(f"ToDoリストへのタスク追加に失敗しました ('{title}'): {e}")
        return None

def find_and_delete_tasks_by_event_id(tasks_service, task_list_id: str, event_id: str) -> int:
    """
    指定されたイベントIDを含むToDoタスクを検索し、削除する。
    :param tasks_service: Google Tasks APIサービスオブジェクト
    :param task_list_id: 検索対象のToDoリストID
    :param event_id: 検索するイベントID
    :return: 削除されたタスクの数
    """
    deleted_count = 0
    try:
        # ToDoリスト内の全てのタスクを取得
        tasks_result = tasks_service.tasks().list(tasklist=task_list_id, showCompleted=False, showHidden=False).execute()
        tasks = tasks_result.get('items', [])

        for task in tasks:
            notes = task.get('notes', '')
            # ToDoの詳細（notes）にイベントIDが含まれているか確認
            if f"関連イベントID: {event_id}" in notes:
                task_id = task['id']
                try:
                    tasks_service.tasks().delete(tasklist=task_list_id, task=task_id).execute()
                    deleted_count += 1
                except Exception as e:
                    st.warning(f"ToDoタスク '{task.get('title', '不明')}' (ID: {task_id}) の削除に失敗しました: {e}")
    except Exception as e:
        st.error(f"ToDoリストからタスクを検索または削除中にエラーが発生しました: {e}")
    return deleted_count

def add_event_to_calendar(service, calendar_id, event_data):
    # イベントオブジェクト全体を返すように変更
    event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
    return event # event.get("htmlLink") ではなくイベントオブジェクト全体を返す

def delete_events_from_calendar(service, calendar_id, start_date: datetime, end_date: datetime):
    # この関数は削除対象のイベントを「取得」するだけで、実際の削除はmain.pyで行うため、削除ロジックを削除
    # ただし、fetch_all_eventsは既に存在するため、そちらを利用する
    # この関数自体は不要になるか、引数を変える必要があるが、今回はmain.pyで直接fetch_all_eventsを呼び出す形で対応
    pass # この関数はメインの削除ロジックから呼び出されないので、実質的には不要

def fetch_all_events(service, calendar_id, time_min, time_max):
    events = []
    page_token = None
    while True:
        result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
            pageToken=page_token
        ).execute()
        events.extend(result.get('items', []))
        page_token = result.get('nextPageToken')
        if not page_token:
            break
    return events

def update_event_if_needed(service, calendar_id, event, new_event_data):
    updated = False

    if 'date' in event['start'] and 'date' in new_event_data['start']:
        if event['start']['date'] != new_event_data['start']['date'] or event['end']['date'] != new_event_data['end']['date']:
            event['start']['date'] = new_event_data['start']['date']
            event['end']['date'] = new_event_data['end']['date']
            updated = True
    elif 'dateTime' in event['start'] and 'dateTime' in new_event_data['start']:
        if event['start']['dateTime'] != new_event_data['start']['dateTime'] or event['end']['dateTime'] != new_event_data['end']['dateTime']:
            event['start']['dateTime'] = new_event_data['start']['dateTime']
            event['end']['dateTime'] = new_event_data['end']['dateTime']
            updated = True

    if updated:
        service.events().update(calendarId=calendar_id, eventId=event['id'], body=event).execute()
    return updated

def generate_ics_content(events: list) -> str:
    """
    Google Calendar APIから取得したイベントリストからICSファイルの内容を生成する。
    """
    cal = Calendar()
    cal.add('prodid', '-//Google Calendar Events Export//jp')
    cal.add('version', '2.0')

    # 日本のタイムゾーンを定義 (Asia/Tokyo)
    tokyo_tz = timezone(timedelta(hours=9))

    for event_data in events:
        event = Event()
        event.add('summary', event_data.get('summary', ''))
        event.add('description', event_data.get('description', ''))
        event.add('location', event_data.get('location', ''))
        event.add('uid', event_data.get('id') + '@google.com') # GoogleイベントIDをUIDとして使用

        # 日付/時刻情報の処理
        start = event_data['start']
        end = event_data['end']

        if 'dateTime' in start: # 通常イベント (日付と時刻)
            try:
                # Google Calendar APIから取得したdateTimeはISO 8601形式 (例: 2023-10-27T09:00:00+09:00)
                # icalendarはタイムゾーン情報を持つdatetimeオブジェクトを直接受け入れる
                start_dt = datetime.fromisoformat(start['dateTime'])
                end_dt = datetime.fromisoformat(end['dateTime'])
                
                event.add('dtstart', start_dt)
                event.add('dtend', end_dt)

            except ValueError as e:
                st.warning(f"イベント '{event_data.get('summary')}' の日時解析に失敗しました: {e}. このイベントはスキップされます。")
                continue # 解析失敗したイベントはスキップ

        elif 'date' in start: # 終日イベント (日付のみ)
            try:
                # 終日イベントの場合、icalendarはPythonのdateオブジェクトまたはnaive datetimeオブジェクトを期待する
                # Google Calendar APIの終日イベントのend.dateはイベントの最終日の翌日を指すため、-1日する
                start_date = datetime.strptime(start['date'], '%Y-%m-%d').date()
                end_date = datetime.strptime(end['date'], '%Y-%m-%d').date() - timedelta(days=1)

                event.add('dtstart', start_date)
                event.add('dtend', end_date)
                # 終日イベントのプロパティを追加 (X-GUESTYLE-ALLDAY: TRUE などは非標準)
                # icalendarはdtstart/dtendが日付オブジェクトの場合、自動的に終日イベントとして扱います。
            except ValueError as e:
                st.warning(f"イベント '{event_data.get('summary')}' の終日イベント日時解析に失敗しました: {e}. このイベントはスキップされます。")
                continue # 解析失敗したイベントはスキップ
        else:
            st.warning(f"イベント '{event_data.get('summary')}' の開始/終了日時形式が不明です。このイベントはスキップされます。")
            continue

        cal.add_component(event)

    return str(cal.to_ical().decode('utf-8'))
