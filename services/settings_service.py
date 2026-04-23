"""
services/settings_service.py
ユーザー設定の読み書き（session_utils.py + user_settings.py を統合）

session_state のアクセスは許可。st.error 等の UI 表示は不可。
"""
from __future__ import annotations
import copy
import streamlit as st
from core.storage.firestore_client import load_settings, save_setting

DEFAULT_SETTINGS: dict = {
    "description_columns_selected": ["内容", "詳細"],
    "event_name_col_selected": "選択しない",
    "event_name_col_selected_update": "選択しない",
    "add_task_type_to_event_name": False,
    "add_task_type_to_event_name_update": False,
    "default_private_event": True,
    "default_allday_event": False,
    "default_create_todo": False,
    "selected_calendar_name": None,
    "share_calendar_selection_across_tabs": True,
    "default_github_logical_names": "",
}


def _ensure_initialized(user_id: str) -> None:
    """セッション上の設定ストアを確保し、未ロードなら Firestore から読み込む。"""
    if "user_settings" not in st.session_state:
        st.session_state["user_settings"] = {}
    if "_settings_loaded" not in st.session_state:
        st.session_state["_settings_loaded"] = set()

    if not user_id:
        return

    if user_id not in st.session_state["user_settings"]:
        st.session_state["user_settings"][user_id] = copy.deepcopy(DEFAULT_SETTINGS)

    if user_id not in st.session_state["_settings_loaded"]:
        saved = load_settings(user_id)
        if saved:
            st.session_state["user_settings"][user_id].update(saved)
        st.session_state["_settings_loaded"].add(user_id)


def get_setting(user_id: str, key: str):
    """設定値を取得する。なければデフォルト値を返す。"""
    _ensure_initialized(user_id)
    if not user_id:
        return DEFAULT_SETTINGS.get(key)
    return st.session_state["user_settings"][user_id].get(key, DEFAULT_SETTINGS.get(key))


def set_setting(user_id: str, key: str, value, persist: bool = True) -> None:
    """設定値をセッションに保存し、オプションで Firestore にも永続化する。"""
    _ensure_initialized(user_id)
    if not user_id:
        return
    st.session_state["user_settings"][user_id][key] = value
    if persist:
        save_setting(user_id, key, value)


def clear_session(user_id: str) -> None:
    """ログアウト時などにセッション上の設定を消去する（Firestore は削除しない）。"""
    ss = st.session_state
    if "user_settings" in ss and user_id in ss["user_settings"]:
        del ss["user_settings"][user_id]
    if "_settings_loaded" in ss:
        ss["_settings_loaded"].discard(user_id)
