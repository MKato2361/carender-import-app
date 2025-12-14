"""user_settings.py

state/calendar_state.py などが参照するための互換ラッパーです。
実体は session_utils.py に集約しています。
"""

from session_utils import (
    get_user_setting,
    set_user_setting,
    save_user_setting_to_firestore,
    load_user_settings_from_firestore,
)
