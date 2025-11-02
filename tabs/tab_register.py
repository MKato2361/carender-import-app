import streamlit as st
from state.calendar_state import get_calendar, set_calendar

# === タブ2：イベント登録 ===
def render_tab_register(service, editable_calendar_options, user_id, current_calendar_name: str):
    st.header("📝 イベントの登録")

    # --- P1：タブ最上部にカレンダー選択（サイドバーと同期） ---
    if editable_calendar_options:
        calendar_options = list(editable_calendar_options.keys())

        try:
            idx = calendar_options.index(current_calendar_name)
        except:
            idx = 0

        selected_tab_calendar = st.selectbox(
            "使用するカレンダーを選択",
            calendar_options,
            index=idx,
            key="tab_register_calendar"
        )

        if selected_tab_calendar != current_calendar_name:
            # カレンダー変更 → 全タブに反映
            set_calendar(user_id, selected_tab_calendar)
            st.session_state["selected_calendar_name"] = selected_tab_calendar
            st.rerun()

        # 選択されたカレンダーID
        calendar_id = editable_calendar_options[selected_tab_calendar]
    else:
        st.warning("利用可能な編集権限付きカレンダーが見つかりません。")
        return

    st.write("---")

    # ==========================================================
    # ※※ここから下は「元コードのイベント登録処理」をそのまま配置してください※※
    # あなたの元 main.py のタブ2内部のロジックを丸ごと移植する場所です
    # ----------------------------------------------------------
    # R1要件：「現在の処理はそのまま」＝ロジック改変禁止
    # UI分離のみ行い、内部処理は既存関数・変数そのまま利用
    # ==========================================================

    st.info("✅ ここに既存のイベント登録処理部分を移植してください（本枠は保持）")

    # 例（ダミーUI、削除して構いません）
    st.write("ここに元のイベント登録UIと処理を貼り付けてください。")
