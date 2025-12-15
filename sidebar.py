from __future__ import annotations

from typing import Dict, Optional, Callable

import streamlit as st

from session_utils import get_user_setting, set_user_setting


# Known widget keys used by tabs (clean state).
_TAB_CALENDAR_WIDGET_KEYS = [
    "reg_calendar_select",
    "reg_calendar_select_outside",
    "del_calendar_select",
    "dup_calendar_select",
    "export_calendar_select",
    "ins_todo_calendar",
    "notice_fax_calendar",
]


def render_sidebar(
    user_id: str,
    editable_calendar_options: Optional[Dict[str, str]],
    save_user_setting_to_firestore: Callable[[str, str, object], None],
) -> None:
    """
    Sidebar rebuilt from scratch to make the "base calendar" stable and
    consistently shared across tabs.

    Key design:
    - Load the stored base calendar first (Firestore via session_utils).
    - Set st.session_state BEFORE widgets are created (no "index + session_state" conflicts).
    - Keep a single source of truth: setting key "selected_calendar_name".
    - When sharing is ON, also push the chosen calendar into known tab widget keys (if they exist).
    """
    with st.sidebar:
        st.subheader("⚙️ 設定")

        if not editable_calendar_options:
            st.info("編集可能なカレンダーが取得できていません。認証状態や権限を確認してください。")
            return

        calendar_names = list(editable_calendar_options.keys())
        if not calendar_names:
            st.info("編集可能なカレンダーが見つかりませんでした。")
            return

        # ----------------------------
        # Share toggle (persisted)
        # ----------------------------
        stored_share = get_user_setting(user_id, "share_calendar_selection_across_tabs")
        if stored_share is None:
            stored_share = True

        if "share_calendar_selection_across_tabs" not in st.session_state:
            st.session_state["share_calendar_selection_across_tabs"] = bool(stored_share)

        share_on = st.checkbox(
            "タブ間で選択を共有",
            key="share_calendar_selection_across_tabs",
            help="ONにすると、基準カレンダー変更時に他タブの選択も可能な範囲で同期します。",
        )

        # Persist if changed
        if bool(share_on) != bool(stored_share):
            set_user_setting(user_id, "share_calendar_selection_across_tabs", bool(share_on))
            save_user_setting_to_firestore(user_id, "share_calendar_selection_across_tabs", bool(share_on))

        st.divider()

        # ----------------------------
        # Base calendar (persisted)
        # ----------------------------
        stored_base = get_user_setting(user_id, "selected_calendar_name")
        if stored_base not in calendar_names:
            stored_base = calendar_names[0]

        # IMPORTANT:
        # Set default value BEFORE creating the widget.
        # Never overwrite on every rerun (that caused "first calendar fixed" issues).
        widget_key = "sb_base_calendar"
        if widget_key not in st.session_state:
            st.session_state[widget_key] = stored_base
        elif st.session_state.get(widget_key) not in calendar_names:
            st.session_state[widget_key] = stored_base

        def _apply_base_calendar_change():
            new_name = st.session_state.get(widget_key)
            if not new_name or new_name not in calendar_names:
                return

            # Global value used by many tabs
            st.session_state["selected_calendar_name"] = new_name

            # Persist
            set_user_setting(user_id, "selected_calendar_name", new_name)
            save_user_setting_to_firestore(user_id, "selected_calendar_name", new_name)

            # If share is ON, also push into existing tab widgets (if present)
            if st.session_state.get("share_calendar_selection_across_tabs", True):
                for k in _TAB_CALENDAR_WIDGET_KEYS:
                    if k in st.session_state:
                        st.session_state[k] = new_name

                # Also sync common session keys if they already exist
                for sk in (
                    "selected_calendar_name_delete",
                    "selected_calendar_name_duplicates",
                    "selected_calendar_name_export",
                    "selected_calendar_name_outside",
                ):
                    if sk in st.session_state:
                        st.session_state[sk] = new_name

        st.selectbox(
            "基準カレンダー",
            calendar_names,
            key=widget_key,
            on_change=_apply_base_calendar_change,
        )

        # Ensure global base calendar is available for tabs rendered later in the script
        if "selected_calendar_name" not in st.session_state or st.session_state["selected_calendar_name"] not in calendar_names:
            st.session_state["selected_calendar_name"] = st.session_state[widget_key]

        # Small status block (useful for debugging)
        with st.expander("🔎 現在の同期状態", expanded=False):
            st.write(f"- user_id: `{user_id}`")
            st.write(f"- stored(selected_calendar_name): `{stored_base}`")
            st.write(f"- session(selected_calendar_name): `{st.session_state.get('selected_calendar_name')}`")
            st.write(f"- share: `{st.session_state.get('share_calendar_selection_across_tabs')}`")
