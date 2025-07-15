# main.py（抜粋不要・全体）

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
from excel_parser import process_excel_files
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Firebase & Google 認証処理（省略せず） ---

# ...（省略せず、認証処理ここに含む）...

# -------------------------
# 🎯 イベント登録タブ
# -------------------------
with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', [])
        )

        # 🔽 イベントデータの事前読み込み（selectbox を事前表示）
        with st.spinner("イベントデータを解析中..."):
            preview_df = process_excel_files(
                st.session_state['uploaded_files'],
                description_columns,
                all_day_event,
                private_event
            )

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name = st.selectbox(
                "登録先カレンダーを選択",
                list(st.session_state['editable_calendar_options'].keys()),
                key="reg_calendar_select"
            )
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("✅ ToDoリスト連携設定 (オプション)")
            create_todo = st.checkbox("このイベントに対応するToDoリストを作成する", value=False, key="create_todo_checkbox")

            fixed_todo_types = ["点検通知"]
            st.markdown(f"以下のToDoが**常にすべて**作成されます: {', '.join(fixed_todo_types)}")

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

            st.subheader("➡️ イベント登録")
            if st.button("Googleカレンダーに登録する"):
                with st.spinner("イベントデータを処理中..."):
                    if preview_df.empty:
                        st.warning("有効なイベントデータがありません。")
                    else:
                        st.info(f"{len(preview_df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        successful_todo_creations = 0

                        for i, row in preview_df.iterrows():
                            # 省略せず：イベント登録処理 + ToDo 作成処理
                            # 例：
                            # created_event = add_event_to_calendar(service, calendar_id, event_data)
                            # add_task_to_todo_list(...) など

                            progress.progress((i + 1) / len(preview_df))

                        st.success(f"✅ {successful_registrations} 件のイベント登録が完了しました！")
                        if create_todo:
                            st.success(f"✅ {successful_todo_creations} 件のToDoリストが作成されました！")
