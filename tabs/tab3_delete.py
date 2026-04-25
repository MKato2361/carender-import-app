from ui.components import calendar_card
from core.utils.datetime_utils import to_utc_range
from services.settings_service import get_setting as get_user_setting, set_setting as set_user_setting
import streamlit as st
from services.calendar_service import get_events as fetch_all_events
from datetime import datetime, date, timedelta, timezone

def _get_current_user_key(fallback: str = "") -> str:
    """設定保存用のユーザーキーを取得（優先: uid -> email）。"""
    return (
        st.session_state.get("user_id")
        or st.session_state.get("firebase_uid")
        or st.session_state.get("localId")
        or st.session_state.get("uid")
        or st.session_state.get("user_email")
        or fallback
        or ""
    )

JST = timezone(timedelta(hours=9))



def render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id):
    if not editable_calendar_options:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    # -------------------------------
    # カレンダー選択（サイドバー設定と連動）
    # -------------------------------
    calendar_names = list(editable_calendar_options.keys())

    # サイドバーの「タブ間で選択を共有」と連動
    share_on = st.session_state.get("share_calendar_selection_across_tabs", True)

    # サイドバーで設定した「基準カレンダー」を初期値として使う（タブ側の選択は永続化しない）
    base_calendar = (
        st.session_state.get("base_calendar_name")
        or st.session_state.get("selected_calendar_name")
        or calendar_names[0]
    )
    if base_calendar not in calendar_names:
        base_calendar = calendar_names[0]

    select_key = "del_calendar_select"
    if share_on:
        st.session_state[select_key] = base_calendar
    elif (select_key not in st.session_state) or (st.session_state.get(select_key) not in calendar_names):
        st.session_state[select_key] = base_calendar

    selected_calendar_name_del = calendar_card(
        calendar_names=calendar_names,
        session_key=select_key,
        base_calendar=base_calendar,
        label="削除対象カレンダー",
        share_on=share_on,
    )

    calendar_id_del = editable_calendar_options[selected_calendar_name_del]

    # -------------------------------
    # イベント削除の期間指定
    # -------------------------------
    st.divider()
    st.markdown('<div class="section-heading"><span class="mi">date_range</span>削除期間</div>', unsafe_allow_html=True)
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


    # -------------------------------
    # イベント削除 実行セクション
    # -------------------------------
    st.markdown('<div class="section-heading"><span class="mi">delete</span>削除の実行</div>', unsafe_allow_html=True)

    if "confirm_delete" not in st.session_state:
        st.session_state["confirm_delete"] = False

    if not st.session_state["confirm_delete"]:
        if st.button("選択期間のイベントを削除する", type="primary", key="events_delete_request"):
            st.session_state["confirm_delete"] = True
            st.rerun()
    else:
        _d1 = delete_start_date.strftime('%Y/%m/%d')
        _d2 = delete_end_date.strftime('%Y/%m/%d')
        _todo_txt = "＋関連ToDo" if delete_related_todos else ""
        st.warning(f"「{selected_calendar_name_del}」 {_d1}〜{_d2} のイベントを削除します{_todo_txt}。この操作は取り消せません。")

        col1, col2 = st.columns([3, 1])

        with col1:
            if st.button("削除を実行する", type="primary", use_container_width=True, key="events_delete_execute"):
                st.session_state["confirm_delete"] = False

                time_min_utc, time_max_utc = to_utc_range(delete_start_date, delete_end_date)
                events_to_delete = fetch_all_events(service, calendar_id_del, time_min_utc, time_max_utc)

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
                                from services.calendar_service import delete_tasks_by_event_id as find_and_delete_tasks_by_event_id
                                deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                    tasks_service,
                                    default_task_list_id,
                                    event_id,
                                )
                                deleted_todos_count += deleted_task_count_for_event

                            service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                            deleted_events_count += 1
                        except Exception as e:
                            st.warning(f"イベント '{event_summary}' の削除に失敗しました（スキップして続行します）。")

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
            if st.button("キャンセル", use_container_width=True, key="events_delete_cancel"):
                st.session_state["confirm_delete"] = False
                st.rerun()

    # -------------------------------
    # ToDo一括削除セクション
    # -------------------------------
    st.divider()
    st.markdown('<div class="section-heading"><span class="mi">checklist</span>ToDo の一括削除</div>', unsafe_allow_html=True)

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
        _td1 = todo_delete_start.strftime('%Y/%m/%d')
        _td2 = todo_delete_end.strftime('%Y/%m/%d')
        _scope_txt = "本アプリ作成分のみ" if "このアプリが作成した" in delete_scope else "全Todo"
        _comp_txt = "（完了済み含む）" if delete_completed_todos else ""
        st.warning(f"期限 {_td1}〜{_td2} の ToDo を削除します。対象：{_scope_txt}{_comp_txt}。この操作は取り消せません。")

        colt1, colt2 = st.columns([3, 1])

        with colt1:
            if st.button("ToDo削除を実行", type="primary", use_container_width=True, key="todo_delete_execute"):
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
                            st.warning(f"ToDo '{title}' の削除に失敗しました（スキップして続行します）。")
                        progress_bar.progress(i / total_tasks)

                    status_text.empty()
                    st.success(f"✅ {deleted_tasks_count} 件のToDoを削除しました。")

        with colt2:
            if st.button("キャンセル", use_container_width=True, key="todo_delete_cancel"):
                st.session_state["confirm_delete_todo"] = False
                st.rerun()
