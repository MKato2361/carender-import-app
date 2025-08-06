import streamlit as st
import calendar_utils
import session_utils
import firebase_auth # This is now the Google Auth module

# --- Main App ---
st.title("Streamlit Google Calendar & Task App")

# Check for authentication redirect
auth_code = st.query_params.get("code")
if auth_code:
    try:
        # Exchange auth code for tokens and save to session
        user_info, google_auth_info = firebase_auth.handle_google_auth_callback(auth_code)

        if user_info and google_auth_info:
            st.session_state.user_info = user_info
            st.session_state.google_auth_info = google_auth_info

            # Initialize user-specific settings after authentication
            session_utils.initialize_session_state(st.session_state.user_info.uid)

            st.success("Successfully logged in!")
            st.query_params.clear()
            st.rerun()

    except Exception as e:
        st.error(f"Authentication failed: {e}")
        st.query_params.clear()

# Check if user is authenticated
if st.session_state.get("user_info") and st.session_state.get("google_auth_info"):
    # User is logged in, show main content
    st.sidebar.success(f"Logged in as: {st.session_state.user_info.email}")
    st.sidebar.button("Logout", on_click=session_utils.logout)

    st.subheader("Google Calendar Events")
    calendar_utils.display_calendar_events(
        st.session_state.user_info.uid, st.session_state.google_auth_info
    )

    st.subheader("Google Tasks")
    calendar_utils.display_task_lists(
        st.session_state.user_info.uid, st.session_state.google_auth_info
    )

else:
    # User is not logged in, show login page
    st.info("Please sign in with your Google account to use the app.")
    auth_url = firebase_auth.get_google_auth_url()
    st.link_button("Sign in with Google", auth_url)

tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame()

with tabs[0]:
    st.header("ファイルをアップロード")
    st.info("""
    ☐作業指示書一覧をアップロードすると管理番号+物件名をイベント名として任意のカレンダーに登録します。
    
    ☐イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。（複数選択可能,次回から同じ項目が選択されます）
    
    ☐イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。
    
    ☐作業外予定の一覧をアップロードすると、イベント名を選択することができます。

    ☐ToDoリストを作成すると、点検通知のリマインドが可能です（ToDoとしてイベント登録されます）
    """)

    import os
    from pathlib import Path
    from io import BytesIO

    def get_local_excel_files():
        current_dir = Path(__file__).parent
        return [f for f in current_dir.glob("*.xlsx") if f.is_file()]

    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    local_excel_files = get_local_excel_files()
    selected_local_files = []
    if local_excel_files:
        st.subheader("📁 サーバーにあるExcelファイル")
        local_file_names = [f.name for f in local_excel_files]
        selected_names = st.multiselect(
            "以下のファイルを処理対象に含める（アップロードと同様に扱われます）",
            local_file_names
        )
        for name in selected_names:
            full_path = next((f for f in local_excel_files if f.name == name), None)
            if full_path:
                with open(full_path, "rb") as f:
                    file_bytes = f.read()
                    file_obj = BytesIO(file_bytes)
                    file_obj.name = name
                    selected_local_files.append(file_obj)

    all_files = []
    if uploaded_files:
        all_files.extend(uploaded_files)
    if selected_local_files:
        all_files.extend(selected_local_files)

    if all_files:
        st.session_state['uploaded_files'] = all_files
        try:
            st.session_state['merged_df_for_selector'] = _load_and_merge_dataframes(all_files)
            st.session_state['description_columns_pool'] = st.session_state['merged_df_for_selector'].columns.tolist()
            if st.session_state['merged_df_for_selector'].empty:
                st.warning("読み込まれたファイルに有効なデータがありませんでした。")
        except (ValueError, IOError) as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []

    if st.session_state.get('uploaded_files'):
        st.subheader("📄 処理対象ファイル一覧")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")
        if not st.session_state['merged_df_for_selector'].empty:
            st.info(f"📊 データ列数: {len(st.session_state['merged_df_for_selector'].columns)}、行数: {len(st.session_state['merged_df_for_selector'])}")

        if st.button("🗑️ アップロード済みファイルをクリア", help="選択中のファイルとデータを削除します。"):
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()

with tabs[1]:
    st.header("イベントを登録・更新")
    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', []),
            default=get_user_setting(user_id, 'description_columns_selected'),
            key=f"description_selector_register_{user_id}"
        )

        fallback_event_name_column = None
        has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
        selected_event_name_col = get_user_setting(user_id, 'event_name_col_selected')

        st.subheader("イベント名の生成設定")
        add_task_type_to_event_name = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=get_user_setting(user_id, 'add_task_type_to_event_name'),
            key=f"add_task_type_checkbox_{user_id}"
        )

        if not (has_mng_data and has_name_data):
            if not has_mng_data and not has_name_data:
                st.info("ファイルに「管理番号」と「物件名」のデータが見つかりませんでした。イベント名に使用する列を選択してください。")
            elif not has_mng_data:
                st.info("ファイルに「管理番号」のデータが見つかりませんでした。物件名と合わせてイベント名に使用する列を選択できます。")
            elif not has_name_data:
                st.info("ファイルに「物件名」のデータが見つかりませんでした。管理番号と合わせてイベント名に使用する列を選択できます。")

            available_event_name_cols = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options = ["選択しない"] + available_event_name_cols
            default_index = event_name_options.index(selected_event_name_col) if selected_event_name_col in event_name_options else 0

            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=default_index,
                key=f"event_name_selector_register_{user_id}"
            )

            if selected_event_name_col != "選択しない":
                fallback_event_name_column = selected_event_name_col
        else:
            st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("✅ ToDoリスト連携設定 (オプション)")
            create_todo = st.checkbox("このイベントに対応するToDoリストを作成する", value=False, key="create_todo_checkbox")

            fixed_todo_types = ["点検通知"]
            if create_todo:
                st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
            else:
                st.markdown(f"ToDoリストの作成は無効です。")

            deadline_offset_options = {
                "2週間前": 14,
                "10日前": 10,
                "1週間前": 7,
                "カスタム日数前": None
            }
            selected_offset_key = st.selectbox(
                "ToDoリストの期限をイベント開始日の何日前に設定しますか？",
                list(deadline_offset_options.keys()),
                disabled=not create_todo,
                key="deadline_offset_select"
            )

            custom_offset_days = None
            if selected_offset_key == "カスタム日数前":
                custom_offset_days = st.number_input(
                    "何日前に設定しますか？ (日数)",
                    min_value=0,
                    value=3,
                    disabled=not create_todo,
                    key="custom_offset_input"
                )

            st.subheader("➡️ イベント登録・更新実行")
            if st.button("Googleカレンダーに登録・更新する"):
                set_user_setting(user_id, 'description_columns_selected', description_columns)
                set_user_setting(user_id, 'event_name_col_selected', selected_event_name_col)
                set_user_setting(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)
                save_user_setting_to_firestore(user_id, 'description_columns_selected', description_columns)
                save_user_setting_to_firestore(user_id, 'event_name_col_selected', selected_event_name_col)
                save_user_setting_to_firestore(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)

                with st.spinner("イベントデータを処理中..."):
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns, 
                            all_day_event_override,
                            private_event, 
                            fallback_event_name_column,
                            add_task_type_to_event_name
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame()

                    if df.empty:
                        st.warning("有効なイベントデータがありません。処理を中断しました。")
                    else:
                        st.info(f"{len(df)} 件のイベントを処理します。")
                        progress = st.progress(0)
                        successful_operations = 0
                        successful_todo_creations = 0

                        worksheet_to_event = {}
                        time_min = (datetime.now(timezone.utc) - timedelta(days=365*2)).isoformat()
                        time_max = (datetime.now(timezone.utc) + timedelta(days=365*2)).isoformat()
                        events = fetch_all_events(service, calendar_id, time_min, time_max)

                        for event in events:
                            desc = event.get('description', '')
                            match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                            if match:
                                worksheet_id = match.group(1)
                                worksheet_to_event[worksheet_id] = event

                        for i, row in df.iterrows():
                            match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
                            event_data = {
                                'summary': row['Subject'],
                                'location': row['Location'],
                                'description': row['Description'],
                                'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                            }

                            if row['All Day Event'] == "True":
                                start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                                end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                                start_date_str = start_date_obj.strftime("%Y-%m-%d")
                                end_date_for_api = (end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d")
                                event_data['start'] = {'date': start_date_str}
                                event_data['end'] = {'date': end_date_for_api}
                            else:
                                start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                                end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                                event_data['start'] = {'dateTime': start_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
                                event_data['end'] = {'dateTime': end_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

                            worksheet_id = match.group(1) if match else None
                            existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

                            if existing_event:
                                try:
                                    updated_event = update_event_if_needed(service, calendar_id, existing_event['id'], event_data)
                                    if updated_event:
                                        successful_operations += 1
                                except Exception as e:
                                    st.error(f"イベント '{row['Subject']}' (作業指示書: {worksheet_id}) の更新に失敗しました: {e}")
                            else:
                                try:
                                    added_event = add_event_to_calendar(service, calendar_id, event_data)
                                    if added_event:
                                        successful_operations += 1
                                        worksheet_id = match.group(1) if match else None
                                        if worksheet_id:
                                            worksheet_to_event[worksheet_id] = added_event
                                except Exception as e:
                                    st.error(f"イベント '{row['Subject']}' の追加に失敗しました: {e}")

                            if create_todo and tasks_service and default_task_list_id:
                                start_date_str = row['Start Date']
                                try:
                                    start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
                                    offset_days = custom_offset_days if selected_offset_key == "カスタム日数前" else deadline_offset_options.get(selected_offset_key)
                                    if offset_days is not None:
                                        todo_due_date = (start_date - timedelta(days=offset_days)).strftime("%Y-%m-%d")
                                        for todo_type in fixed_todo_types:
                                            todo_summary = f"{todo_type}: {row['Subject']}"
                                            todo_notes = f"イベントID: {worksheet_to_event.get(worksheet_id, {}).get('id', '不明')}\n詳細: {row['Description']}"
                                            task_data = {
                                                'title': todo_summary,
                                                'due': todo_due_date,
                                                'notes': todo_notes
                                            }
                                            try:
                                                if add_task_to_todo_list(tasks_service, default_task_list_id, task_data):
                                                    successful_todo_creations += 1
                                            except Exception as e:
                                                st.error(f"ToDo '{todo_summary}' の追加に失敗しました: {e}")
                                    else:
                                        st.warning(f"ToDoの期限が設定されませんでした。カスタム日数が無効です。")
                                except Exception as e:
                                    st.warning(f"ToDoの期限を設定できませんでした。イベント開始日が不明です: {e}")

                            progress.progress((i + 1) / len(df))

                        st.success(f"✅ {successful_operations} 件のイベントが処理されました (新規登録/更新)。")
                        if create_todo:
                            st.success(f"✅ {successful_todo_creations} 件のToDoリストが作成されました！")

with tabs[2]:
    st.header("イベントを削除")
    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]
        st.subheader("🗓️ 削除期間の選択")
        today_date = date.today()
        delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
        delete_end_date = st.date_input("削除終了日", value=today_date)
        delete_related_todos = st.checkbox("関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)", value=False)

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")
            if st.button("選択期間のイベントを削除する"):
                calendar_service = st.session_state['calendar_service']
                tasks_service = st.session_state['tasks_service']
                default_task_list_id = st.session_state.get('default_task_list_id')

                start_dt_utc = datetime.combine(delete_start_date, datetime.min.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                end_dt_utc = datetime.combine(delete_end_date, datetime.max.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                
                time_min_utc = start_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                time_max_utc = end_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')

                events_to_delete = fetch_all_events(calendar_service, calendar_id_del, time_min_utc, time_max_utc)
                
                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")

                deleted_events_count = 0
                deleted_todos_count = 0
                total_events = len(events_to_delete)
                
                if total_events > 0:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for i, event in enumerate(events_to_delete):
                        event_summary = event.get('summary', '不明なイベント')
                        event_id = event['id']
                        
                        status_text.text(f"イベント '{event_summary}' を削除中... ({i+1}/{total_events})")

                        try:
                            if delete_related_todos and tasks_service and default_task_list_id:
                                deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                    tasks_service,
                                    default_task_list_id,
                                    event_id
                                )
                                deleted_todos_count += deleted_task_count_for_event
                            
                            calendar_service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                            deleted_events_count += 1
                        except Exception as e:
                            st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")
                        
                        progress_bar.progress((i + 1) / total_events)
                    
                    status_text.empty()

                    if deleted_events_count > 0:
                        st.success(f"✅ {deleted_events_count} 件のイベントが削除されました。")
                        if delete_related_todos:
                            if deleted_todos_count > 0:
                                st.success(f"✅ {deleted_todos_count} 件の関連ToDoタスクが削除されました。")
                            else:
                                st.info("関連するToDoタスクは見つからなかったか、すでに削除されていました。")
                    else:
                        st.info("指定期間内に削除するイベントはありませんでした。")
                else:
                    st.info("指定期間内に削除するイベントはありませんでした。")

with tabs[3]:
    st.header("イベントを更新")
    st.info("このタブは、主に既存イベントの情報をExcelデータに基づいて**上書き**したい場合に使用します。新規イベントの作成は行いません。")

    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        all_day_event_override_update = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event_update = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")

        description_columns_update = st.multiselect(
            "説明欄に含める列", 
            st.session_state['description_columns_pool'], 
            default=get_user_setting(user_id, 'description_columns_selected'),
            key=f"update_desc_cols_{user_id}"
        )

        fallback_event_name_column_update = None
        has_mng_data_update, has_name_data_update = check_event_name_columns(st.session_state['merged_df_for_selector'])
        selected_event_name_col_update = get_user_setting(user_id, 'event_name_col_selected_update')

        st.subheader("更新時のイベント名の生成設定")
        add_task_type_to_event_name_update = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=get_user_setting(user_id, 'add_task_type_to_event_name_update'),
            key=f"add_task_type_checkbox_update_{user_id}"
        )

        if not (has_mng_data_update and has_name_data_update):
            st.info("Excelデータからのイベント名生成に、以下の列を代替として使用できます。")
            available_event_name_cols_update = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
            event_name_options_update = ["選択しない"] + available_event_name_cols_update
            default_index_update = event_name_options_update.index(selected_event_name_col_update) if selected_event_name_col_update in event_name_options_update else 0

            selected_event_name_col_update = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options_update,
                index=default_index_update,
                key=f"event_name_selector_update_{user_id}"
            )

            if selected_event_name_col_update != "選択しない":
                fallback_event_name_column_update = selected_event_name_col_update
        else:
            st.info("「管理番号」と「物件名」のデータが存在するため、それらがイベント名として使用されます。")

        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name_upd = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
            calendar_id_upd = st.session_state['editable_calendar_options'][selected_calendar_name_upd]

            if st.button("イベントを照合・更新"):
                set_user_setting(user_id, 'description_columns_selected', description_columns_update)
                set_user_setting(user_id, 'event_name_col_selected_update', selected_event_name_col_update)
                set_user_setting(user_id, 'add_task_type_to_event_name_update', add_task_type_to_event_name_update)
                save_user_setting_to_firestore(user_id, 'description_columns_selected', description_columns_update)
                save_user_setting_to_firestore(user_id, 'event_name_col_selected_update', selected_event_name_col_update)
                save_user_setting_to_firestore(user_id, 'add_task_type_to_event_name_update', add_task_type_to_event_name_update)

                with st.spinner("イベントを処理中..."):
                    try:
                        df = process_excel_data_for_calendar(
                            st.session_state['uploaded_files'], 
                            description_columns_update,
                            all_day_event_override_update,
                            private_event_update,
                            fallback_event_name_column_update,
                            add_task_type_to_event_name_update
                        )
                    except (ValueError, IOError) as e:
                        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                        df = pd.DataFrame()

                    if df.empty:
                        st.warning("有効なイベントデータがありません。更新を中断しました。")
                        st.stop()

                    today_for_update = datetime.now()
                    time_min = (today_for_update - timedelta(days=365*2)).isoformat() + 'Z'
                    time_max = (today_for_update + timedelta(days=365*2)).isoformat() + 'Z'
                    events = fetch_all_events(service, calendar_id_upd, time_min, time_max)

                    worksheet_to_event = {}
                    for event in events:
                        desc = event.get('description', '')
                        match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                        if match:
                            worksheet_id = match.group(1)
                            worksheet_to_event[worksheet_id] = event

                    update_count = 0
                    progress_bar = st.progress(0)
                    for i, row in df.iterrows():
                        match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
                        if not match:
                            progress_bar.progress((i + 1) / len(df))
                            continue
                        
                        worksheet_id = match.group(1)
                        matched_event = worksheet_to_event.get(worksheet_id)
                        if not matched_event:
                            progress_bar.progress((i + 1) / len(df))
                            continue

                        event_data = {
                            'summary': row['Subject'],
                            'location': row['Location'],
                            'description': row['Description'],
                            'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                        }
                        
                        if row['All Day Event'] == "True":
                            start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                            end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                            start_date_str = start_date_obj.strftime("%Y-%m-%d")
                            end_date_for_api = (end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d")
                            event_data['start'] = {'date': start_date_str}
                            event_data['end'] = {'date': end_date_for_api}
                        else:
                            start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                            end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                            event_data['start'] = {'dateTime': start_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
                            event_data['end'] = {'dateTime': end_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

                        try:
                            if update_event_if_needed(service, calendar_id_upd, matched_event['id'], event_data):
                                update_count += 1
                        except Exception as e:
                            st.error(f"イベント '{row['Subject']}' (作業指示書: {worksheet_id}) の更新に失敗しました: {e}")
                        
                        progress_bar.progress((i + 1) / len(df))

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")

with st.sidebar:
    st.header("🔐 認証状態")
    st.success("✅ Firebase認証済み")
    
    if st.session_state.get('calendar_service'):
        st.success("✅ Googleカレンダー認証済み")
    else:
        st.warning("⚠️ Googleカレンダー認証が未完了です")
    
    if st.session_state.get('tasks_service'):
        st.success("✅ ToDoリスト利用可能")
    else:
        st.warning("⚠️ ToDoリスト利用不可")
    
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    st.metric("アップロード済みファイル", uploaded_count)
    
    if st.button("🚪 ログアウト", type="secondary"):
        if user_id:
            clear_user_settings(user_id)
        for key in list(st.session_state.keys()):
            if not key.startswith("google_auth") and not key.startswith("firebase_"):
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
