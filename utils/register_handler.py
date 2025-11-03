import streamlit as st
import pandas as pd
from datetime import date, timedelta
from typing import List, Optional



from utils.helpers import default_fetch_window_years
from excel_parser import process_excel_data_for_calendar
from calendar_utils import fetch_all_events


def render_tab2_register(user_id: str, editable_calendar_options: dict, service, tasks_service=None, default_task_list_id=None):
    st.subheader("イベントを登録・更新")

    if not st.session_state.get("uploaded_files") or st.session_state["merged_df_for_selector"].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    description_columns: List[str] = []
    selected_event_name_col: Optional[str] = None
    add_task_type_to_event_name = False
    all_day_event_override = False
    private_event = True
    fallback_event_name_column: Optional[str] = None

    # カレンダー選択
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
    save_user_setting_to_firestore(user_id, "selected_calendar_name", selected_calendar_name)

    # 設定読み込み
    description_columns_pool = st.session_state.get("description_columns_pool", [])
    saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
    saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
    saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")
    saved_create_todo_flag = get_user_setting(user_id, "create_todo_checkbox_state")

    expand_event_setting = not bool(saved_description_cols)
    expand_name_setting = not (saved_event_name_col or saved_task_type_flag)
    expand_todo_setting = bool(saved_create_todo_flag)

    # イベント設定
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

    # イベント名生成設定
    with st.expander("🧱 イベント名の生成設定", expanded=expand_name_setting):
        from excel_parser import check_event_name_columns, get_available_columns_for_event_name
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

    # ToDo連携設定
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

    # 実行ボタン
    st.subheader("➡️ イベント登録・更新実行")
    if st.button("Googleカレンダーに登録・更新する"):

        # 設定保存
        set_user_setting(user_id, "description_columns_selected", description_columns)
        set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
        set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

        save_user_setting_to_firestore(user_id, "description_columns_selected", description_columns)
        save_user_setting_to_firestore(user_id, "event_name_col_selected", selected_event_name_col)
        save_user_setting_to_firestore(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

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

            # イベント候補生成（UI側では抽出処理せず handler に委譲）
            prep = prepare_events(df, description_columns, fallback_event_name_column, add_task_type_to_event_name)
            if prep["errors"]:
                st.error("以下のエラーが発生しました:\n" + "\n".join(prep["errors"]))
                if not prep["events"]:
                    return

            if prep["warnings"]:
                st.warning("以下の警告があります:\n" + "\n".join(prep["warnings"]))

            st.info(f"{len(prep['events'])} 件のイベントを処理します。")
            progress = st.progress(0)

            # 既存イベント取得
            time_min, time_max = default_fetch_window_years(2)
            existing_event_map = fetch_existing_events(service, calendar_id, time_min, time_max)

            results = {"added": 0, "updated": 0, "skipped": 0}
            total = len(prep["events"])

            # handlerで処理
            for idx, event_data in enumerate(prep["events"]):
                partial_res = register_or_update_events(
                    service,
                    calendar_id,
                    [event_data],
                    existing_event_map,
                )
                results["added"] += partial_res["added"]
                results["updated"] += partial_res["updated"]
                results["skipped"] += partial_res["skipped"]

                progress.progress((idx + 1) / total)

            st.success(f"✅ 登録: {results['added']} / 🔧 更新: {results['updated']} / ↪ スキップ: {results['skipped']}")