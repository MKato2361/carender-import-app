"""
tab2_register.py
タブ2：イベント登録 UI（軽い改善版）
"""

from __future__ import annotations
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone, date

from utils.register_handler import (
    prepare_events,
    fetch_existing_events,
    register_or_update_events,
)

from utils.helpers import safe_get, default_fetch_window_years
from excel_parser import process_excel_data_for_calendar
from firebase_auth import get_firebase_user_id
from session_utils import get_user_setting, set_user_setting
from calendar_utils import add_task_to_todo_list, build_tasks_service, fetch_all_events


JST = timezone(timedelta(hours=9))


def render_tab2_register(service, editable_calendar_options):
    user_id = get_firebase_user_id()

    st.subheader("イベントを登録・更新")

    # データ未アップロード時のガード
    if (
        "uploaded_files" not in st.session_state
        or not st.session_state["uploaded_files"]
        or st.session_state.get("merged_df_for_selector", pd.DataFrame()).empty
    ):
        st.info("先に「1. ファイルのアップロード」でファイルを読み込んでください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーがありません。Google認証をご確認ください。")
        return

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
        key="reg_calendar_select",
    )
    calendar_id = editable_calendar_options[selected_calendar_name]

    set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)

    # 設定UI
    df = st.session_state["merged_df_for_selector"]
    description_columns_pool = df.columns.tolist()

    saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
    saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
    saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")

    st.subheader("📝 イベント設定")
    description_columns = st.multiselect(
        "説明欄に含める列（複数選択可）",
        description_columns_pool,
        default=[c for c in saved_description_cols if c in description_columns_pool],
    )

    st.subheader("🧱 イベント名の生成設定")
    add_task_type_to_event_name = st.checkbox(
        "イベント名の先頭に作業タイプを追加する",
        value=bool(saved_task_type_flag),
    )
    event_name_col = st.selectbox(
        "代替イベント名に使用する列（Subjectが空の場合）",
        options=["選択しない"] + description_columns_pool,
        index=(description_columns_pool.index(saved_event_name_col) + 1) if saved_event_name_col in description_columns_pool else 0,
    )
    fallback_event_name_column = None if event_name_col == "選択しない" else event_name_col

    st.subheader("✅ ToDo作成（オプション）")
    create_todo = st.checkbox("このイベントに対応するToDoリストを作成する", value=False)
    deadline_offset = st.slider("ToDo期限（イベント開始日の何日前）", 1, 30, 7, disabled=not create_todo)

    # 登録ボタン
    st.subheader("➡️ イベント登録・更新実行")
    if st.button("Googleカレンダーに登録・更新する"):
        # 設定保存
        set_user_setting(user_id, "description_columns_selected", description_columns)
        set_user_setting(user_id, "event_name_col_selected", fallback_event_name_column)
        set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

        with st.spinner("イベントデータを準備中..."):
            try:
                df_processed = process_excel_data_for_calendar(
                    st.session_state["uploaded_files"],
                    description_columns,
                    False,  # all_day override → 現仕様では使わないためFalse固定
                    True,   # private_event default → True固定（詳細はhandlerで反映）
                    fallback_event_name_column,
                    add_task_type_to_event_name,
                )
            except Exception as e:
                st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                return

        prep = prepare_events(df_processed, description_columns, fallback_event_name_column, add_task_type_to_event_name)

        if prep["errors"]:
            st.error(f"❌ {len(prep['errors'])}件の行でエラーがあり、スキップされます。")
            with st.expander("エラー詳細を表示"):
                for err in prep["errors"]:
                    st.write(f"- {err}")

        if prep["warnings"]:
            st.warning(f"⚠️ {len(prep['warnings'])}件の警告があります。")
            with st.expander("警告の詳細を見る"):
                for warn in prep["warnings"]:
                    st.write(f"- {warn}")

        event_candidates = prep["events"]
        if not event_candidates:
            st.error("イベント候補が生成できなかったため処理を中止します。")
            return

        with st.spinner("既存イベントを取得中..."):
            time_min, time_max = default_fetch_window_years(2)
            existing_map = fetch_existing_events(service, calendar_id, time_min, time_max)

        total = len(event_candidates)
        progress = st.progress(0)

        results = {"added": 0, "updated": 0, "skipped": 0}

        for i, candidate in enumerate(event_candidates):
            # 登録・更新
            r = register_or_update_events(service, calendar_id, [candidate], existing_map)
            for k in results:
                results[k] += r[k]

            # ToDo生成（失敗しても処理継続）
            if create_todo:
                try:
                    event_start = datetime.strptime(candidate["Start Date"], "%Y/%m/%d").date()
                    due_date = event_start - timedelta(days=deadline_offset)
                    title = f"【ToDo】{candidate['Subject']}"
                    add_task_to_todo_list(title, "", due_date)  # 関数は既存のものを流用
                except Exception:
                    pass

            progress.progress((i + 1) / total)

        st.success(f"✅ 登録: {results['added']} / 🔧 更新: {results['updated']} / ↪ スキップ: {results['skipped']}")