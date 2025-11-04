import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

# ===== 他モジュール依存 =====
from utils.helpers import safe_get
from utils.parsers import extract_worksheet_id_from_text
from excel_parser import (
    process_excel_data_for_calendar,
    get_available_columns_for_event_name,
    check_event_name_columns,
)
from calendar_utils import (
    fetch_all_events,
    add_event_to_calendar,
    update_event_if_needed,
)
from session_utils import (
    get_user_setting,
    set_user_setting,
)
from firebase_admin import firestore


JST = timezone(timedelta(hours=9))


def is_event_changed(existing_event: dict, new_event_data: dict) -> bool:
    nz = lambda v: (v or "")

    if nz(existing_event.get("summary")) != nz(new_event_data.get("summary")):
        return True

    if nz(existing_event.get("description")) != nz(new_event_data.get("description")):
        return True

    if nz(existing_event.get("transparency")) != nz(new_event_data.get("transparency")):
        return True

    if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
        return True

    if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
        return True

    return False


def default_fetch_window_years(years: int = 2):
    from datetime import datetime, timezone, timedelta

    now_utc = datetime.now(timezone.utc)
    return (
        (now_utc - timedelta(days=365 * years)).isoformat(),
        (now_utc + timedelta(days=365 * years)).isoformat(),
    )


def extract_worksheet_id_from_description(desc: str) -> str | None:
    import re
    import unicodedata

    RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")

    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return unicodedata.normalize("NFKC", m.group(1)).strip()


def render_tab2_register(user_id: str, editable_calendar_options: dict, service, tasks_service=None, default_task_list_id=None):
    st.subheader("イベントを登録・更新")

    if not st.session_state.get("uploaded_files") or st.session_state["merged_df_for_selector"].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    # --- カレンダー選択 ---
    calendar_options = list(editable_calendar_options.keys())
    saved_calendar_name = get_user_setting(user_id, "selected_calendar_name")
    try:
        default_index = calendar_options.index(saved_calendar_name)
    except Exception:
        default_index = 0

    selected_calendar_name = st.selectbox(
        "登録先カレンダーを選択",
        calendar_options,
        index=default_index,
        key="reg_calendar_select"
    )
    calendar_id = editable_calendar_options[selected_calendar_name]

    set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)

    # --- 設定ロード ---
    description_columns_pool = st.session_state.get("description_columns_pool", [])
    saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
    saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
    saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")
    saved_create_todo_flag = get_user_setting(user_id, "create_todo_checkbox_state")

    expand_event_setting = not bool(saved_description_cols)
    expand_name_setting = not (saved_event_name_col or saved_task_type_flag)
    expand_todo_setting = bool(saved_create_todo_flag)

    # --- イベント設定 ---
    with st.expander("📝 イベント設定", expanded=expand_event_setting):
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        default_selection = [col for col in saved_description_cols if col in description_columns_pool]
        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            description_columns_pool,
            default=default_selection,
            key=f"description_selector_register_{user_id}",
        )

    # --- イベント名生成設定 ---
    with st.expander("🧱 イベント名の生成設定", expanded=expand_name_setting):
        has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
        selected_event_name_col = saved_event_name_col
        add_task_type_to_event_name = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=bool(saved_task_type_flag),
            key=f"add_task_type_checkbox_{user_id}",
        )
        fallback_event_name_column = None

        if not (has_mng_data and has_name_data):
            available_event_name_cols = get_available_columns_for_event_name(
                st.session_state["merged_df_for_selector"]
            )
            event_name_options = ["選択しない"] + available_event_name_cols
            try:
                name_index = event_name_options.index(selected_event_name_col) if selected_event_name_col else 0
            except Exception:
                name_index = 0
            selected_event_name_col = st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=event_name_options,
                index=name_index,
                key=f"event_name_selector_register_{user_id}",
            )
            if selected_event_name_col != "選択しない":
                fallback_event_name_column = selected_event_name_col
        else:
            st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

    # --- ToDo設定 (UIのみ移植 / 処理ロジック維持) ---
    st.subheader("✅ ToDoリスト連携設定 (オプション)")
    with st.expander("ToDoリスト作成オプション", expanded=expand_todo_setting):
        create_todo = st.checkbox(
            "このイベントに対応するToDoリストを作成する",
            value=bool(saved_create_todo_flag),
            key="create_todo_checkbox",
        )
        set_user_setting(user_id, "create_todo_checkbox_state", create_todo)

        fixed_todo_types = ["点検通知"]
        if create_todo:
            st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
        else:
            st.markdown("ToDoリストの作成は無効です。")

        deadline_offset_options = {"2週間前": 14, "10日前": 10, "1週間前": 7, "カスタム日数前": None}
        selected_offset_key = st.selectbox(
            "ToDoリストの期限をイベント開始日の何日前に設定しますか？",
            list(deadline_offset_options.keys()),
            disabled=not create_todo,
            key="deadline_offset_select",
        )
        custom_offset_days = None
        if selected_offset_key == "カスタム日数前":
            custom_offset_days = st.number_input(
                "何日前に設定しますか？ (日数)",
                min_value=0,
                value=3,
                disabled=not create_todo,
                key="custom_offset_input",
            )

    # --- 実行ボタン ---
    st.subheader("➡️ イベント登録・更新実行")
    if st.button("Googleカレンダーに登録・更新する"):

        # 保存
        set_user_setting(user_id, "description_columns_selected", description_columns)
        set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
        set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

        # ===== イベント登録処理 =====
        with st.spinner("イベントデータを処理中..."):
            try:
                df = process_excel_data_for_calendar(
                    st.session_state["uploaded_files"],
                    description_columns,
                    all_day_event_override,
                    private_event,
                    fallback_event_name_column,
                    add_task_type_to_event_name,
                )
            except Exception as e:
                st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                return

            if df.empty:
                st.warning("有効なイベントデータがありません。処理を中断しました。")
                return

            st.info(f"{len(df)} 件のイベントを処理します。")
            progress = st.progress(0)

            added_count = 0
            updated_count = 0
            skipped_count = 0

            time_min, time_max = default_fetch_window_years(2)
            with st.spinner("既存イベントを取得中..."):
                events = fetch_all_events(service, calendar_id, time_min, time_max)

            worksheet_to_event: Dict[str, dict] = {}
            for event in events or []:
                wid = extract_worksheet_id_from_description(event.get("description") or "")
                if wid:
                    worksheet_to_event[wid] = event

            total = len(df)

            for i, row in df.iterrows():
                desc_text = safe_get(row, "Description", "")
                worksheet_id = extract_worksheet_id_from_text(desc_text)

                all_day_flag = safe_get(row, "All Day Event", "True")
                private_flag = safe_get(row, "Private", "True")
                start_date_str = safe_get(row, "Start Date", "")
                end_date_str = safe_get(row, "End Date", "")
                start_time_str = safe_get(row, "Start Time", "")
                end_time_str = safe_get(row, "End Time", "")

                event_data = {
                    "summary": safe_get(row, "Subject", ""),
                    "location": safe_get(row, "Location", ""),
                    "description": desc_text,
                    "transparency": "transparent" if private_flag == "True" else "opaque",
                }

                try:
                    if all_day_flag == "True":
                        sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                        ed = datetime.strptime(end_date_str, "%Y/%m/%d").date()
                        event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                        event_data["end"] = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
                    else:
                        sdt = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                        edt = datetime.strptime(f"{end_date_str} {end_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                        event_data["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
                        event_data["end"] = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
                except Exception as e:
                    st.error(f"行 {i} の日時パースに失敗しました: {e}")
                    progress.progress((i + 1) / total)
                    continue

                existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

                try:
                    if existing_event:
                        if is_event_changed(existing_event, event_data):
                            _ = update_event_if_needed(service, calendar_id, existing_event["id"], event_data)
                            updated_count += 1
                        else:
                            skipped_count += 1
                    else:
                        added_event = add_event_to_calendar(service, calendar_id, event_data)
                        if added_event:
                            added_count += 1
                            if worksheet_id:
                                worksheet_to_event[worksheet_id] = added_event
                except Exception as e:
                    st.error(f"イベント '{event_data.get('summary','(無題)')}' の登録/更新に失敗しました: {e}")

                progress.progress((i + 1) / total)

            st.success(f"✅ 登録: {added_count} / 🔧 更新: {updated_count} / ↪ スキップ: {skipped_count}")
