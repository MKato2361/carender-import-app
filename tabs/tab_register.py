import streamlit as st
from typing import List, Optional, Dict
import pandas as pd
from datetime import datetime, timedelta
from state.calendar_state import get_calendar, set_calendar

# === タブ2：イベントの登録 ===
def render_tab_register(service, editable_calendar_options, user_id, current_calendar_name: str):
    st.subheader("イベントを登録・更新")

    # ---- カレンダー選択（サイドバーと同期：追加部分） ----
    if editable_calendar_options:
        calendar_options = list(editable_calendar_options.keys())
        try:
            idx = calendar_options.index(current_calendar_name)
        except:
            idx = 0

        selected_tab_calendar = st.selectbox(
            "登録先カレンダーを選択",
            calendar_options,
            index=idx,
            key="reg_calendar_select_tab"
        )

        if selected_tab_calendar != current_calendar_name:
            set_calendar(user_id, selected_tab_calendar)
            st.session_state["selected_calendar_name"] = selected_tab_calendar
            st.rerun()

        calendar_id = editable_calendar_options[selected_tab_calendar]
    else:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    # ---- ここから下、あなたが貼ったコードをそのまま移植 ----

    from user_settings import (
        get_user_setting,
        set_user_setting,
        save_user_setting_to_firestore,
    )
    from utils.event_utils import (
        check_event_name_columns,
        get_available_columns_for_event_name,
        process_excel_data_for_calendar,
        extract_worksheet_id_from_description,
        extract_worksheet_id_from_text,
        default_fetch_window_years,
        fetch_all_events,
        safe_get,
        is_event_changed,
        update_event_if_needed,
        add_event_to_calendar,
    )
    from utils.timezone import JST  # あなたの環境で必要な場合

    description_columns: List[str] = []
    selected_event_name_col: Optional[str] = None
    add_task_type_to_event_name = False
    all_day_event_override = False
    private_event = True
    fallback_event_name_column: Optional[str] = None

    if not st.session_state.get("uploaded_files") or st.session_state["merged_df_for_selector"].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    saved_calendar_name = get_user_setting(user_id, "selected_calendar_name")
    try:
        default_index = calendar_options.index(saved_calendar_name)
    except Exception:
        default_index = 0

    # ✅ ここは上で同期済みなので削除せず、そのまま残す（挙動変えない）
    selected_calendar_name = selected_tab_calendar

    set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)
    save_user_setting_to_firestore(user_id, "selected_calendar_name", selected_calendar_name)

    description_columns_pool = st.session_state.get("description_columns_pool", [])
    saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
    saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
    saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")
    saved_create_todo_flag = get_user_setting(user_id, "create_todo_checkbox_state")

    expand_event_setting = not bool(saved_description_cols)
    expand_name_setting = not (saved_event_name_col or saved_task_type_flag)
    expand_todo_setting = bool(saved_create_todo_flag)

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

    with st.expander("🧱 イベント名の生成設定", expanded=expand_name_setting):
        has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
        selected_event_name_col = saved_event_name_col
        add_task_type_to_event_name = st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=bool(saved_task_type_flag),
            key=f"add_task_type_checkbox_{user_id}",
        )
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

    st.subheader("✅ ToDoリスト連携設定 (オプション)")
    with st.expander("ToDoリスト作成オプション", expanded=expand_todo_setting):
        create_todo = st.checkbox(
            "このイベントに対応するToDoリストを作成する",
            value=bool(saved_create_todo_flag),
            key="create_todo_checkbox",
        )
        set_user_setting(user_id, "create_todo_checkbox_state", create_todo)
        save_user_setting_to_firestore(user_id, "create_todo_checkbox_state", create_todo)

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

    st.subheader("➡️ イベント登録・更新実行")
    if st.button("Googleカレンダーに登録・更新する"):
        set_user_setting(user_id, "description_columns_selected", description_columns)
        set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
        set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

        save_user_setting_to_firestore(user_id, "description_columns_selected", description_columns)
        save_user_setting_to_firestore(user_id, "event_name_col_selected", selected_event_name_col)
        save_user_setting_to_firestore(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

        from utils.timezone import JST  # 再import念のため

        with st.spinner("イベントデータ処理中..."):
            try:
                df = process_excel_data_for_calendar(
                    st.session_state["uploaded_files"],
                    description_columns,
                    all_day_event_override,
                    private_event,
                    fallback_event_name_column,
                    add_task_type_to_event_name,
                )
            except (ValueError, IOError) as e:
                st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                df = pd.DataFrame()

            if df.empty:
                st.warning("有効なイベントデータがありません。処理を中断しました。")
            else:
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

                    all_day_flag  = safe_get(row, "All Day Event", "True")
                    private_flag  = safe_get(row, "Private", "True")
                    start_date_str = safe_get(row, "Start Date", "")
                    end_date_str   = safe_get(row, "End Date", "")
                    start_time_str = safe_get(row, "Start Time", "")
                    end_time_str   = safe_get(row, "End Time", "")

                    event_data = {
                        "summary":   safe_get(row, "Subject", ""),
                        "location":  safe_get(row, "Location", ""),
                        "description": desc_text,
                        "transparency": "transparent" if private_flag == "True" else "opaque",
                    }

                    try:
                        if all_day_flag == "True":
                            sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                            ed = datetime.strptime(end_date_str, "%Y/%m/%d").date()
                            event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                            event_data["end"]   = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
                        else:
                            sdt = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                            edt = datetime.strptime(f"{end_date_str} {end_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                            event_data["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
                            event_data["end"]   = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
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
