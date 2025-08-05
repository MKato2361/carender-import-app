def authenticate_google():
    creds = None
    user_id = get_firebase_user_id()

    if not user_id:
        return None

    db = firestore.client()
    doc_ref = db.collection('google_tokens').document(user_id)

    # セッションから認証情報を取得
    if 'credentials' in st.session_state and st.session_state['credentials']:
        creds = st.session_state['credentials']
        if creds.valid:
            return creds
        elif creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state['credentials'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            return creds

    # Firestoreから認証情報を取得
    try:
        doc = doc_ref.get()
        if doc.exists:
            token_data = doc.to_dict()
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
            st.session_state['credentials'] = creds

            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                st.session_state['credentials'] = creds
                doc_ref.set(json.loads(creds.to_json()))
                st.info("Google認証トークンを更新しました。")
                st.rerun()

            return creds
    except Exception as e:
        st.error(f"Firestoreからトークン取得に失敗しました: {e}")
        creds = None

    # 新しいOAuthフロー（Webリダイレクト型）
    try:
        client_config = {
            "web": {
                "client_id": st.secrets["google"]["client_id"],
                "project_id": st.secrets["google"]["project_id"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets["google"]["client_secret"],
                "redirect_uris": [st.secrets["google"]["redirect_uri"]]
            }
        }

        flow = Flow.from_client_config(client_config, SCOPES)
        flow.redirect_uri = st.secrets["google"]["redirect_uri"]

        # URLパラメータから認証コード取得
        params = st.experimental_get_query_params()
        if "code" not in params:
            auth_url, _ = flow.authorization_url(
                prompt='consent',
                access_type='offline',
                include_granted_scopes='true'
            )
            st.markdown(f"[Googleでログインする]({auth_url})")
            st.stop()
        else:
            code = params["code"][0]
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.session_state['credentials'] = creds
            doc_ref.set(json.loads(creds.to_json()))
            st.success("Google認証が完了しました！")
            st.experimental_set_query_params()  # 認証コードをURLから消す
            st.rerun()

    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        st.session_state['credentials'] = None
        return None

    return creds


def build_tasks_service(creds):
    try:
        if not creds:
            return None
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        st.warning(f"Google Tasks サービスのビルドに失敗しました: {e}")
        return None

def add_event_to_calendar(service, calendar_id, event_data):
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
        st.error(f"イベントの取得に失敗しました (HTTPエラー): {e}")
        return []
    except Exception as e:
        st.error(f"イベントの取得に失敗しました: {e}")
        return []

def update_event_if_needed(service, calendar_id, event_id, updated_event_data):
    try:
        existing_event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        needs_update = any(
            existing_event.get(key) != value
            for key, value in updated_event_data.items()
        )
        if needs_update:
            return service.events().update(
                calendarId=calendar_id,
                eventId=event_id,
                body=updated_event_data
            ).execute()
        else:
            return existing_event
    except HttpError as e:
        st.error(f"イベントの更新に失敗しました (HTTPエラー): {e}")
        return None
    except Exception as e:
        st.error(f"イベントの更新に失敗しました: {e}")
        return None

def add_task_to_todo_list(tasks_service, task_list_id, task_data):
    try:
        if not tasks_service:
            return None
        return tasks_service.tasks().insert(
            tasklist=task_list_id,
            body=task_data
        ).execute()
    except HttpError as e:
        st.error(f"タスクの追加に失敗しました (HTTPエラー): {e}")
        return None
    except Exception as e:
        st.error(f"タスクの追加に失敗しました: {e}")
        return None

def find_and_delete_tasks_by_event_id(tasks_service, task_list_id, event_id):
    try:
        if not tasks_service:
            return False
        tasks_result = tasks_service.tasks().list(tasklist=task_list_id).execute()
        tasks = tasks_result.get('items', [])
        deleted_count = 0
        for task in tasks:
            if (event_id in task.get('notes', '') or event_id in task.get('title', '')):
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
    task_data = {'title': title, 'notes': notes}
    if due_date:
        task_data['due'] = due_date.isoformat() + 'Z'
    return task_data

def get_calendar_colors():
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
    try:
        if isinstance(date_str, str):
            date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
        else:
            date_obj = date_str

        if isinstance(time_str, str):
            time_obj = datetime.strptime(time_str, '%H:%M').time()
        else:
            time_obj = time_str

        dt = datetime.combine(date_obj, time_obj)
        return dt.replace(tzinfo=timezone(timedelta(hours=9)))
    except ValueError as e:
        st.error(f"日付または時刻の形式が正しくありません: {e}")
        return None
