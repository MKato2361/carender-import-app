# tabs/tab_export.py
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from state.calendar_state import set_calendar

def render_tab_export(editable_calendar_options, user_id, current_calendar_name: str):
    st.subheader("📤 イベントのExcel出力")

    if not editable_calendar_options:
        st.error("出力可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    try:
        idx = calendar_options.index(current_calendar_name)
    except Exception:
        idx = 0

    selected_tab_calendar = st.selectbox(
        "対象カレンダーを選択",
        calendar_options,
        index=idx,
        key=f"export_calendar_select_tab_{user_id}",
    )
    if selected_tab_calendar != current_calendar_name:
        set_calendar(user_id, selected_tab_calendar)
        st.session_state["selected_calendar_name"] = selected_tab_calendar
        st.rerun()

    # 元データ（アップロード済み）をそのまま出力対象にする
    df = st.session_state.get("merged_df_for_selector")
    if df is None or df.empty:
        st.info("先に「1. ファイルのアップロード」でデータを読み込んでください。")
        return

    with st.expander("プレビュー（先頭50行）", expanded=False):
        st.dataframe(df.head(50), use_container_width=True)

    if st.button("Excelを生成してダウンロード"):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Events")
        buffer.seek(0)
        fname = f"events_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="📥 ダウンロード",
            data=buffer,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("✅ Excelを生成しました。")
