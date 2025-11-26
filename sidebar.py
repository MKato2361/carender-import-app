from __future__ import annotations
from typing import Dict, Optional, Callable

import streamlit as st

from session_utils import get_user_setting, set_user_setting, clear_user_settings
from github_loader import _headers, GITHUB_OWNER, GITHUB_REPO


def render_sidebar(
    user_id: str,
    editable_calendar_options: Optional[Dict[str, str]],
    save_user_setting_to_firestore: Callable[[str, str, object], None],
) -> None:
    """サイドバー全体を描画する"""

    with st.sidebar:
        st.subheader("⚙️ 設定・管理")

        # 📅 カレンダー設定
        with st.expander("📅 カレンダー設定", expanded=True):
            if editable_calendar_options:
                calendar_options = list(editable_calendar_options.keys())

                # ✅ まずはセッションの値を優先して取得
                saved_calendar = st.session_state.get("selected_calendar_name")
                if not saved_calendar:
                    # セッションになければ Firestore 上の設定を読む
                    saved_calendar = get_user_setting(user_id, "selected_calendar_name")

                try:
                    default_cal_index = (
                        calendar_options.index(saved_calendar)
                        if saved_calendar in calendar_options
                        else 0
                    )
                except ValueError:
                    default_cal_index = 0

                # ★ ここでの選択値が「今の基準カレンダー」
                default_calendar = st.selectbox(
                    "デフォルトカレンダー",
                    calendar_options,
                    index=default_cal_index,
                    key="sidebar_default_calendar",
                )

                # ✅ 毎回、現在の選択をグローバルキーに反映しておく
                #    → 各タブ側は st.session_state["selected_calendar_name"] を見て初期値を決定
                st.session_state["selected_calendar_name"] = default_calendar

                prev_share = st.session_state.get(
                    "share_calendar_selection_across_tabs"
                )
                if prev_share is None:
                    prev_share = True

                share_calendar = st.checkbox(
                    "タブ間で選択を共有",
                    value=prev_share,
                    help="ONにすると、登録タブで選んだカレンダーが他のタブにも自動で反映されます。",
                )

                # 設定変更時の即時反映ロジック
                if share_calendar != prev_share:
                    st.session_state["share_calendar_selection_across_tabs"] = (
                        share_calendar
                    )
                    set_user_setting(
                        user_id, "share_calendar_selection_across_tabs", share_calendar
                    )
                    save_user_setting_to_firestore(
                        user_id, "share_calendar_selection_across_tabs", share_calendar
                    )
                    st.rerun()

                st.divider()

                saved_private = get_user_setting(user_id, "default_private_event")
                default_private = st.checkbox(
                    "標準で「非公開」",
                    value=(saved_private if saved_private is not None else True),
                    key="sidebar_default_private",
                )

                saved_allday = get_user_setting(user_id, "default_allday_event")
                default_allday = st.checkbox(
                    "標準で「終日」",
                    value=(saved_allday if saved_allday is not None else False),
                    key="sidebar_default_allday",
                )
            else:
                # カレンダー未取得時でもエラーにならないように
                saved_private = get_user_setting(user_id, "default_private_event")
                default_private = st.checkbox(
                    "標準で「非公開」",
                    value=(saved_private if saved_private is not None else True),
                    key="sidebar_default_private",
                )

                saved_allday = get_user_setting(user_id, "default_allday_event")
                default_allday = st.checkbox(
                    "標準で「終日」",
                    value=(saved_allday if saved_allday is not None else False),
                    key="sidebar_default_allday",
                )

        # ✅ ToDo設定
        with st.expander("✅ ToDo設定", expanded=False):
            saved_todo = get_user_setting(user_id, "default_create_todo")
            default_todo = st.checkbox(
                "標準で「ToDo作成」",
                value=(saved_todo if saved_todo is not None else False),
                key="sidebar_default_todo",
            )

        # 保存・リセットボタン
        col_save, col_reset = st.columns(2)
        with col_save:
            if st.button("💾 設定保存", use_container_width=True):
                if editable_calendar_options:
                    calendar_options = list(editable_calendar_options.keys())
                    # selectbox の現在値をそのまま使う
                    default_calendar = st.session_state.get(
                        "sidebar_default_calendar", calendar_options[0]
                    )

                    # 共通のデフォルトカレンダー設定（Firestore に保存）
                    set_user_setting(
                        user_id, "selected_calendar_name", default_calendar
                    )
                    save_user_setting_to_firestore(
                        user_id, "selected_calendar_name", default_calendar
                    )
                    st.session_state["selected_calendar_name"] = default_calendar

                    # ★ 全タブへの連携用キーをまとめて更新
                    if st.session_state.get(
                        "share_calendar_selection_across_tabs", True
                    ):
                        # 各タブ専用キー名（tab3 / tab5 / tab7 / tab8 等で使っている suffix）
                        tab_keys_for_share = [
                            "register",
                            "delete",
                            "export",
                            "inspection_todo",
                            "notice_fax",
                            "property_master",
                            "admin",
                        ]
                        for suffix in tab_keys_for_share:
                            st.session_state[
                                f"selected_calendar_name_{suffix}"
                            ] = default_calendar

                # その他設定
                set_user_setting(user_id, "default_private_event", default_private)
                save_user_setting_to_firestore(
                    user_id, "default_private_event", default_private
                )

                set_user_setting(user_id, "default_allday_event", default_allday)
                save_user_setting_to_firestore(
                    user_id, "default_allday_event", default_allday
                )

                set_user_setting(user_id, "default_create_todo", default_todo)
                save_user_setting_to_firestore(
                    user_id, "default_create_todo", default_todo
                )

                st.toast("設定を保存しました", icon="✅")

        with col_reset:
            if st.button("🔄 リセット", use_container_width=True):
                for key in [
                    "default_private_event",
                    "default_allday_event",
                    "default_create_todo",
                ]:
                    set_user_setting(user_id, key, None)
                    save_user_setting_to_firestore(user_id, key, None)
                st.toast("設定をリセットしました", icon="🧹")
                st.rerun()

        st.divider()

        # 📡 ステータス表示（全ての認証項目）
        with st.container(border=True):
            st.caption("📡 接続ステータス")

            # Firebase 認証（user_id が取れていれば OK）
            firebase_ok = bool(user_id)

            # Google API 系は session_state で確認
            calendar_ok = bool(st.session_state.get("calendar_service"))
            tasks_ok = bool(st.session_state.get("tasks_service"))
            sheets_ok = bool(st.session_state.get("sheets_service"))

            # GitHub：トークンの有無＋OWNER/REPO が設定されているかで判定
            token_in_secrets = False
            try:
                token_in_secrets = bool(st.secrets.get("GITHUB_TOKEN", ""))
            except Exception:
                token_in_secrets = False

            token_in_headers = False
            try:
                token_in_headers = bool(_headers.get("Authorization"))
            except Exception:
                token_in_headers = False

            owner_repo_ok = bool(GITHUB_OWNER and GITHUB_REPO)
            github_ok = owner_repo_ok and (token_in_secrets or token_in_headers)

            st.markdown(
                f"""
- **Firebase 認証**: {'✅ ログイン中' if firebase_ok else '⚠️ 未ログイン'}
- **Google Calendar API**: {'✅ 接続中' if calendar_ok else '⚠️ 未接続'}
- **Google Tasks API**: {'✅ 利用可' if tasks_ok else '⛔ 利用不可'}
- **Google Sheets API**: {'✅ 利用可' if sheets_ok else '⛔ 利用不可'}
- **GitHub API**: {'✅ 設定済' if github_ok else '⚠️ 未設定またはエラー'}
"""
            )

        st.divider()

        # 🚪 ログアウト
        if st.button("🚪 ログアウト", type="primary", use_container_width=True):
            if user_id:
                clear_user_settings(user_id)
            for key in list(st.session_state.keys()):
                if not key.startswith("google_auth") and not key.startswith(
                    "firebase_"
                ):
                    del st.session_state[key]
            st.rerun()
