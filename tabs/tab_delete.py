import streamlit as st
from datetime import date, datetime, timedelta, timezone
from typing import Dict
from state.calendar_state import get_calendar, set_calendar

from utils.event_utils import fetch_all_events
from utils.todo_utils import find_and_delete_tasks_by_event_id
from utils.timezone import JST


def render_tab_delete(service, editable_calendar_options, user_id, current_calendar_name: str):
    st.subheader("イベントを削除")

    # ---- カレンダー選択（タブ上部 × サイドバー同期）----
    if not editable_calendar_options:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return
    else:
        calendar_names = list(editable_calendar_options.keys())
        try:
            idx = calendar_names.index(current_calendar_name)
        except Exception:
            idx = 0

        selected_tab_calendar = st.selectbox(
            "削除対象カレンダーを選択",
            calendar_names,
            index=idx,
            key="del_calendar_select_tab"
        )

        # サイドバー & 全タブ同期
        if selected_tab_calendar != current_calendar_name:
            set_calendar(user_id, selected_tab_calendar)
            st.session_state["selected_calendar_name"] = selected_tab_calendar
            st.rerun()

        selected_calendar_name_del = selected_tab_calendar
        calendar_id_del = editable_calendar_options[selected_calendar_name_del]

    # ---- 以下、あなたが貼ったコード：ロジック改変なし ----

    st.subheader("🗓️ 削除期間の選択")
    today_date = date.today()
    delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
    delete_end_date = st.date_input("削除終了日", value=today_date)
    delete_related_todos = st.checkbox(
        "関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)",
        value=False
    )

    if delete_start_date > delete_end_date:
        st.error("削除開始日は終了日より前に設定してください。")
        return

    st.subheader("🗑️ 削除実行")
    if "confirm_delete" not in st.session_state:
        st.session_state["confirm_delete"] = False

    if not st.session_state["confirm_delete"]:
        if st.button("選択期間のイベントを削除する", type="primary"):
            st.session_state["confirm_delete"] = True
            st.rerun()

    if st.session_state["confirm_delete"]:
        st.warning(
            f"""
⚠️ **削除確認**

以下のイベントを削除します:
- **カレンダー名**: {selected_calendar_name_del}
- **期間**: {delete_start_date.strftime('%Y年%m月%d日')} ～ {delete_end_date.strftime('%Y年%m月%d日')}
- **ToDoリストも削除**: {'はい' if delete_related_todos else 'いいえ'}

この操作は取り消せません。本当に削除しますか？
"""
        )

        col1, col2 = st.columns([1, 1])

        def to_utc_range_btn(d1: date, d2: date):
            sdt = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
            edt = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
            return (
                sdt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
                edt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
            )

        with col1:
            if st.button("✅ 実行", type="primary", use_container_width=True):
                st.session_state["confirm_delete"] = False

                time_min_utc, time_max_utc = to_utc_range_btn(delete_start_date, delete_end_date)
                events_to_delete = fetch_all_events(service, calendar_id_del, time_min_utc, time_max_utc)

                if not events_to_delete:
                    st.info("指定期間内に削除するイベントはありませんでした。")
                    return

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
                        if delete_related_todos and st.session_state.get("tasks_service") and st.session_state.get("default_task_list_id"):
                            deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                st.session_state["tasks_service"],
                                st.session_state["default_task_list_id"],
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
                            st.info("関連するToDoタスクは見つからなかったか、すでに削除されていました。")
                else:
                    st.info("指定期間内に削除するイベントはありませんでした。")

        with col2:
            if st.button("❌ キャンセル", use_container_width=True):
                st.session_state["confirm_delete"] = False
                st.rerun()
