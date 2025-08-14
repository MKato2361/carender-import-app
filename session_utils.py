import streamlit as st

# デフォルト設定
DEFAULT_SETTINGS = {
    'description_columns_selected': ["内容", "詳細"],
    'event_name_col_selected': "選択しない",
    'event_name_col_selected_update': "選択しない",
    'add_task_type_to_event_name': False,
    'add_task_type_to_event_name_update': False,
    'delete_with_mng': False,
    'all_day_event_override': False,
    'private_event': False,
    'all_day_event_override_upd': False,
    'private_event_upd': False
}

def initialize_session_state(user_id):
    """ユーザーごとのセッション状態を初期化"""
    if 'user_settings' not in st.session_state:
        st.session_state['user_settings'] = {}
    st.session_state['user_settings'].setdefault(user_id, DEFAULT_SETTINGS.copy())

def get_user_setting(user_id, key):
    """ユーザー設定を取得"""
    initialize_session_state(user_id)
    return st.session_state['user_settings'][user_id].get(key)

def set_user_setting(user_id, key, value):
    """ユーザー設定を保存"""
    initialize_session_state(user_id)
    st.session_state['user_settings'][user_id][key] = value

def get_all_user_settings(user_id):
    """すべてのユーザー設定を取得"""
    initialize_session_state(user_id)
    return st.session_state['user_settings'][user_id]

def clear_user_settings(user_id):
    """ユーザー設定をクリア"""
    if 'user_settings' in st.session_state and user_id in st.session_state['user_settings']:
        del st.session_state['user_settings'][user_id]
