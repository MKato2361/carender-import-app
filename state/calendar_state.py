# state/calendar_state.py
"""
全タブで共通して利用する「選択中カレンダー」の状態管理モジュール
・常にタブ間共有
・SessionState + Firestore 両方で保持
"""

import streamlit as st
from session_utils import get_user_setting, set_user_setting
from firebase_admin import firestore

db = firestore.client()

_KEY = "selected_calendar_name"


def get_calendar(user_id: str, default: str = None) -> str | None:
    """
    現在選択中のカレンダー名を取得する。
    優先順：SessionState → Firestore → default引数
    """
    # 1. SessionStateにあれば最優先
    if _KEY in st.session_state:
        return st.session_state[_KEY]

    # 2. Firestore保存分を取得
    saved = get_user_setting(user_id, _KEY)
    if saved:
        st.session_state[_KEY] = saved
        return saved

    # 3. defaultがあれば採用
    if default:
        st.session_state[_KEY] = default
        return default

    return None


def set_calendar(user_id: str, calendar_name: str) -> None:
    """
    選択したカレンダー名を保存（SessionState + Firestore）
    """
    st.session_state[_KEY] = calendar_name
    set_user_setting(user_id, _KEY, calendar_name)

    # Firestoreにも保存（merge=Trueで他設定を壊さない）
    db.collection("user_settings").document(user_id).set(
        {_KEY: calendar_name}, merge=True
    )
