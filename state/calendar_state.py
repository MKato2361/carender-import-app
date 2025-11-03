# state/calendar_state.py
import streamlit as st

# 全タブで共通利用するカレンダー選択の保存/取得

def get_calendar(user_id: str, editable_calendar_options: dict, default_name: str = None) -> str:
    # 1) セッション
    name = st.session_state.get("selected_calendar_name")
    if name and name in editable_calendar_options:
        return name

    # 2) ユーザ設定（存在すれば）
    try:
        from user_settings import get_user_setting
        saved = get_user_setting(user_id, "selected_calendar_name")
        if saved and saved in editable_calendar_options:
            st.session_state["selected_calendar_name"] = saved
            return saved
    except Exception:
        pass

    # 3) デフォルト or 最初の要素
    if default_name and default_name in editable_calendar_options:
        st.session_state["selected_calendar_name"] = default_name
        return default_name

    if editable_calendar_options:
        first = list(editable_calendar_options.keys())[0]
        st.session_state["selected_calendar_name"] = first
        return first

    return None


def set_calendar(user_id: str, calendar_name: str):
    st.session_state["selected_calendar_name"] = calendar_name
    # Firestore等があれば保存
    try:
        from user_settings import set_user_setting, save_user_setting_to_firestore
        set_user_setting(user_id, "selected_calendar_name", calendar_name)
        save_user_setting_to_firestore(user_id, "selected_calendar_name", calendar_name)
    except Exception:
        pass
