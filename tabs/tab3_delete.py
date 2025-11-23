import streamlit as st
from datetime import datetime, date, timedelta, timezone

JST = timezone(timedelta(hours=9))


def render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id):
    st.subheader("イベントを削除")

    if not editable_calendar_options:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    # -------------------------------
    # カレンダー選択
    # -------------------------------
    calendar_names = list(editable_calendar_options.keys())
    default_index = 0
    saved_name = st.session_state.get("selected_calendar_name")
    if saved_name and saved_name in calendar_names:
        default_index = calendar_names.index(saved_name)

    selected_calendar_name_del = st.selectbox(
        "削除対象カレンダーを選択",
        calendar_names,
        index=default_index,
        key="del_calendar_select",
    )
    st.session_state["selected_calendar_name"] = selected_calendar_name_del
    calendar_id_del = editable_calendar_options[selected_calendar_name_del]

    # -------------------------------
    # イベント削除の期間指定
    # -------------------------------
    st.subheader("🗓️ 削除期間の選択（イベント）")
    today_date = date.today()
    delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
    delete_end_date = st.date_input("削除終了日", value=today_date)
    delete_related_todos = st.checkbox(
        "関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)",
        value=False,
    )

    if delete_start_date > delete_end_date:
        st.error("削除開始日は終了日より前に設定してください。")
        return

    # UTC 変換ヘルパー（イベント・ToDo共通で使用）
    def to_utc_range(d1: date, d2: date):
        sdt = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
        edt = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
        return (
            sdt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
            edt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
        )

    # -------------------------------
    # イベント削除 実行セクション
    # -------------------------------
    st.subheader("🗑️ イベント削除の実行")

    if "confirm_delete" not in st.session_state:
        st.session_state["confirm_delete"] = False

    if not st.session_state["confirm_delete"]:
        if st.button("選択期間のイベントを削除する", type="primary", key="events_delete_request"):
            st.session_state["confirm_delete"] = True
            st.rerun()
    else:
        st.warning(
            f"""
⚠️ **イベント削除の確認**

以下のイベントを削除します:
- **カレンダー名**: {selected_calendar_name_del}
- **期間**: {delete_start_date.strftime('%Y年%m月%d日')} ～ {delete_end_date.strftime('%Y年%m月%d日')}
- **ToDoリストも削除**: {'はい' if delete_related_todos else 'いいえ'}

この操作は取り消せません。本当に削除しますか？
"""
        )

        col1, col2 = st.columns([1, 1])

        with col1:
            if st.button("✅ 実行", type="primary", use_container_width=True, key="events_delete_execute"):
                st.session_state["confirm_delete"] = False

                time_min_utc, time_max_utc = to_utc_range(delete_start_date, delete_end_date)
                events_to_delete = service.events().list(
                    calendarId=calendar_id_del,
                    timeMin=time_min_utc,
                    timeMax=time_max_utc,
                    singleEvents=True,
                ).execute().get("items", [])

                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")
                else:
                    deleted_events_count = 0
                    deleted_todos_count = 0
                    total_events = len(events_to_delete or [])

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for i, event in enumerate(events_to_delete, start=1):
                        event_summary = event.get("summary", "不明なイベント")
                        event_id = event["id"]
                        status_text.text(f"イベント '{event_summary}' を削除中... ({i}/{total_events})")
                        try:
                            if delete_related_todos and tasks_service and default_task_list_id:
                                from calendar_utils import find_and_delete_tasks_by_event_id
                                deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                    tasks_service,
                                    default_task_list_id,
                                    event_id,
                                )
                                deleted_todos_count += deleted_task_count_for_event

                            service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                            deleted_events_count += 1
                        except Exception as e:
                            st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")

                        progress_bar.progress(i / total_events)

                    status_text.empty()

                    if deleted_events_count > 0:
                        st.success(f"✅ {deleted_events_count} 件のイベントが削除されました。")
                        if delete_related_todos:
                            if deleted_todos_count > 0:
                                st.success(f"✅ {deleted_todos_count} 件の関連ToDoタスクが削除されました。")
                            else:
                                st.info("関連するToDoタスクは見つかりませんでした。")
                    else:
                        st.info("指定期間内に削除するイベントはありませんでした。")

        with col2:
            if st.button("❌ キャンセル", use_container_width=True, key="events_delete_cancel"):
                st.session_state["confirm_delete"] = False
                st.rerun()

    # -------------------------------
    # ToDo一括削除セクション
    # -------------------------------
    st.subheader("✅ ToDo の一括削除")

    if not tasks_service or not default_task_list_id:
        st.info("Google ToDo リストサービスが利用できないため、ToDo一括削除機能は使用できません。")
        return

    # ToDo 削除の条件入力
    todo_delete_start = st.date_input(
        "ToDo削除開始日（期限）",
        value=today_date - timedelta(days=30),
        key="todo_delete_start",
    )
    todo_delete_end = st.date_input(
        "ToDo削除終了日（期限）",
        value=today_date + timedelta(days=30),
        key="todo_delete_end",
    )
    delete_completed_todos = st.checkbox(
        "完了済みToDoも削除対象に含める",
        value=True,
        key="todo_delete_completed",
    )

    # ★ 追加: 削除対象の範囲オプション
    delete_scope = st.radio(
        "削除対象",
        (
            "このアプリが作成したToDoのみ（[EVENT_ID:...]付き）",
            "指定期間のToDoをすべて削除する（注意）",
        ),
        index=0,
        key="todo_delete_scope",
    )

    if todo_delete_start > todo_delete_end:
        st.error("ToDo削除開始日は終了日より前に設定してください。")
        return

    if "confirm_delete_todo" not in st.session_state:
        st.session_state["confirm_delete_todo"] = False

    if not st.session_state["confirm_delete_todo"]:
        if st.button("選択期間のToDoを一括削除する", type="secondary", key="todo_delete_request"):
            st.session_state["confirm_delete_todo"] = True
            st.rerun()
    else:
        target_desc = (
            "本アプリが作成した ToDo（notes 内に `[EVENT_ID:xxx]` を含むもの）"
            if "このアプリが作成した" in delete_scope
            else "指定期間に該当する **すべての ToDo**"
        )

        st.warning(
            f"""
⚠️ **ToDo一括削除の確認**

以下の条件で ToDo を削除します:
- **期限日**: {todo_delete_start.strftime('%Y年%m月%d日')} ～ {todo_delete_end.strftime('%Y年%m月%d日')}
- **完了済みも削除**: {'はい' if delete_completed_todos else 'いいえ'}
- **対象**: {target_desc}

この操作は取り消せません。本当に削除しますか？
"""
        )

        colt1, colt2 = st.columns([1, 1])

        with colt1:
            if st.button("✅ ToDo削除を実行", type="primary", use_container_width=True, key="todo_delete_execute"):
                st.session_state["confirm_delete_todo"] = False

                due_min_utc, due_max_utc = to_utc_range(todo_delete_start, todo_delete_end)

                tasks_to_delete = []
                page_token = None

                # ToDo を due 範囲＋（必要に応じて [EVENT_ID:...]）で絞り込み
                while True:
                    params = dict(
                        tasklist=default_task_list_id,
                        maxResults=100,
                        showCompleted=True,
                        showDeleted=False,
                        showHidden=False,
                        pageToken=page_token,
                        dueMin=due_min_utc,
                        dueMax=due_max_utc,
                    )
                    resp = tasks_service.tasks().list(**params).execute()
                    items = resp.get("items", [])

                    for t in items:
                        notes = (t.get("notes") or "")

                        # このアプリが作成したものだけ削除するモード
                        if "このアプリが作成した" in delete_scope:
                            if "[EVENT_ID:" not in notes:
                                continue

                        # 完了済みを除外したい場合
                        if (not delete_completed_todos) and t.get("status") == "completed":
                            continue

                        tasks_to_delete.append(t)

                    page_token = resp.get("nextPageToken")
                    if not page_token:
                        break

                total_tasks = len(tasks_to_delete)
                if total_tasks == 0:
                    st.info("指定期間内に削除対象のToDoは見つかりませんでした。")
                else:
                    deleted_tasks_count = 0
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for i, task in enumerate(tasks_to_delete, start=1):
                        title = task.get("title", "無題のToDo")
                        status_text.text(f"ToDo '{title}' を削除中... ({i}/{total_tasks})")
                        try:
                            tasks_service.tasks().delete(
                                tasklist=default_task_list_id,
                                task=task["id"],
                            ).execute()
                            deleted_tasks_count += 1
                        except Exception as e:
                            st.error(f"ToDo '{title}' の削除に失敗しました: {e}")
                        progress_bar.progress(i / total_tasks)

                    status_text.empty()
                    st.success(f"✅ {deleted_tasks_count} 件のToDoを削除しました。")

        with colt2:
            if st.button("❌ ToDo削除をキャンセル", use_container_width=True, key="todo_delete_cancel"):
                st.session_state["confirm_delete_todo"] = False
                st.rerun()
