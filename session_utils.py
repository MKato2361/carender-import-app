"""
session_utils.py — 後方互換ラッパー（実体は services/settings_service.py）
直接 import している箇所は services.settings_service に切り替えてください。
"""
from services.settings_service import (
    get_setting as get_user_setting,
    set_setting as set_user_setting,
    clear_session as clear_user_settings,
    _ensure_initialized as initialize_session_state,
    DEFAULT_SETTINGS,
)
from core.storage.firestore_client import (
    load_settings as load_user_settings_from_firestore,
    save_setting as save_user_setting_to_firestore,
)
