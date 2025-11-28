from __future__ import annotations
from typing import Dict, Optional, Callable

import os
import re
import streamlit as st

from session_utils import get_user_setting, set_user_setting, clear_user_settings
from github_loader import (
    _headers,
    GITHUB_OWNER,
    GITHUB_REPO,
    walk_repo_tree,
    is_supported_file,
)


def _logical_github_name(filename: str) -> str:
    """末尾の数字（日付）を除いた論理名に変換"""
    base, _ext = os.path.splitext(filename)
    base = re.sub(r"\d+$", "", base)
    return base


def render_sidebar(
    user_id: str,
    editable_calendar_options: Optional[Dict[str, str]],
    save_user_setting_to_firestore: Callable[[str, str, object], None],
) -> None:
    """サイドバー全体を描画する"""

    with st.sidebar:
        st.subheader("⚙️ 設定・管理")
        st.caption("まずは下の『現在の設定状況』を確認し、必要に応じて各設定を開いて調整してください。")

        # =========================
        # 🧾 現在の設定状況（サマリー）
        # =========================
        with st.container(border=True):
            st.markdown("**🧾 現在の設定状況**")

            # --- 基準カレンダー ---
            calendar_label = "未設定"
            if editable_calendar_options:
                calendar_options = list(editable_calendar_options.keys())

                # session_state / Firestore から有効なカレンダー名を決定
                cal_from_state = (
                    st.session_state.get("sidebar_default_calendar")
                    or st.session_state.get("selected_calendar_name")
                )
                cal_from_store = get_user_setting(user_id, "selected_calendar_name")

                if cal_from_state in calendar_options:
                    calendar_label = cal_from_state
                elif cal_from_store in calendar_options:
                    calendar_label = cal_from_store
                else:
                    calendar_label = "未設定（カレンダー一覧は取得済み）"
            else:
                calendar_label = "カレンダー未取得"

            st.write(f"- **基準カレンダー**：{calendar_label}")

            # --- 新規イベントのデフォルト ---
            # 非公開
            saved_private = get_user_setting(user_id, "default_private_event")
            private_val = st.session_state.get("sidebar_default_private")
            if private_val is None:
                private_val = saved_private if saved_private is not None else True

            # 終日
            saved_allday = get_user_setting(user_id, "default_allday_event")
            allday_val = st.session_state.get("sidebar_default_allday")
            if allday_val is None:
                allday_val = saved_allday if saved_allday is not None else False

            st.write(
                "- **新規イベント**："
                f"{'非公開' if private_val else '公開'}, "
                f"{'終日' if allday_val else '時間指定'}"
            )

            # --- ToDo設定 ---
            saved_todo = get_user_setting(user_id, "default_create_todo")
            todo_val = st.session_state.get("sidebar_default_todo")
            if todo_val is None:
                todo_val = bool(saved_todo) if saved_todo is not None else False

            st.write(f"- **ToDo作成**：{'あり' if todo_val else 'なし'}")

            # --- GitHubデフォルトファイル（論理名） ---
            gh_text = st.session_state.get("default_github_logical_names")
            if gh_text is None:
                gh_text = get_user_setting(user_id, "default_github_logical_names") or ""
            gh_list = [line.strip() for line in gh_text.splitlines() if line.strip()]

            if not gh_list:
                st.write("- **GitHubデフォルトファイル**：なし")
            else:
                st.write("- **GitHubデフォルトファイル**：")
                for name in gh_list[:5]:
                    st.write(f"  - {name}")
                if len(gh_list) > 5:
                    st.caption(f"　…ほか {len(gh_list) - 5} 件")

        st.divider()

        # ========================
        # 📅 カレンダー設定（折りたたみ）
        # ========================
        with st.expander("📅 カレンダー設定", expanded=False):
            if editable_calendar_options:
                calendar_options = list(editable_calendar_options.keys())

                # Firestore に保存されているカレンダー名
                stored_calendar = get_user_setting(user_id, "selected_calendar_name")
                # 画面での直近の選択状態
                session_calendar = st.session_state.get("sidebar_default_calendar")

                # 有効なカレンダー名を決定（優先順位：画面 > Firestore > 先頭）
                effective_calendar = calendar_options[0]
                if session_calendar in calendar_options:
                    effective_calendar = session_calendar
                elif stored_calendar in calendar_options:
                    effective_calendar = stored_calendar

                # selectbox の state を常に「有効なカレンダー名」に同期
                st.session_state["sidebar_default_calendar"] = effective_calendar

                st.markdown("**基準カレンダー**")

                default_calendar = st.selectbox(
                    "デフォルトカレンダー",
                    calendar_options,
                    key="sidebar_default_calendar",
                )

                # グローバルにも反映（他タブで使う想定）
                st.session_state["selected_calendar_name"] = default_calendar

                # 🔽 共有設定：その下に縦に配置
                prev_share = st.session_state.get("share_calendar_selection_across_tabs")
                if prev_share is None:
                    saved_share = get_user_setting(
                        user_id, "share_calendar_selection_across_tabs"
                    )
                    prev_share = True if saved_share is None else bool(saved_share)
                    st.session_state["share_calendar_selection_across_tabs"] = (
                        prev_share
                    )

                share_calendar = st.checkbox(
                    "タブ間で選択を共有",
                    value=prev_share,
                    help="ONにすると、登録タブで選んだカレンダーが他のタブにも自動で反映されます。",
                )

                if share_calendar != prev_share:
                    st.session_state["share_calendar_selection_across_tabs"] = (
                        share_calendar
                    )
                    set_user_setting(
                        user_id, "share_calendar_selection_across_tabs", share_calendar
                    )
                    save_user_setting_to_firestore(
                        user_id,
                        "share_calendar_selection_across_tabs",
                        share_calendar,
                    )
                    st.rerun()
            else:
                st.info(
                    "編集可能なカレンダーが取得できていません。認証状態や権限を確認してください。"
                )

            st.markdown("---")

            # --- 新規イベントのデフォルト設定（非公開／終日） ---
            st.markdown("**新規イベントのデフォルト**")

            saved_private = get_user_setting(user_id, "default_private_event")
            saved_allday = get_user_setting(user_id, "default_allday_event")

            default_private = st.checkbox(
                "標準で「非公開」",
                value=(saved_private if saved_private is not None else True),
                key="sidebar_default_private",
            )

            default_allday = st.checkbox(
                "標準で「終日」",
                value=(saved_allday if saved_allday is not None else False),
                key="sidebar_default_allday",
            )

        # ========================
        # ✅ ToDo設定（折りたたみ）
        # ========================
        with st.expander("✅ ToDo設定", expanded=False):
            st.caption("新規イベント作成時に、同時にToDoを発行するかどうかを決めます。")
            saved_todo = get_user_setting(user_id, "default_create_todo")
            default_todo = st.checkbox(
                "標準で「ToDo作成」",
                value=(saved_todo if saved_todo is not None else False),
                key="sidebar_default_todo",
            )

        # ===========================
        # 📦 GitHubファイル設定（折りたたみ）
        # ===========================
        with st.expander("📦 GitHubファイル設定", expanded=False):
            st.caption(
                "GitHub上のファイルから、末尾の日付部分を除いた『論理名』単位でデフォルト選択を設定します。\n"
                "ここでチェックした論理名は、アップロードタブ側の初期選択として自動でONになります。"
            )

            # 1) Firestore から文字列を取得 → セッションと同期
            saved_gh_text = get_user_setting(user_id, "default_github_logical_names")
            if saved_gh_text is None:
                saved_gh_text = ""
            if "default_github_logical_names" not in st.session_state:
                st.session_state["default_github_logical_names"] = saved_gh_text

            current_gh_text = st.session_state["default_github_logical_names"]
            saved_gh_set = {
                line.strip()
                for line in current_gh_text.splitlines()
                if line.strip()
            }

            logical_to_files: Dict[str, list[str]] = {}

            # GitHub から対象ファイル一覧を読み込む
            try:
                gh_nodes = walk_repo_tree(base_path="", max_depth=3)
                for node in gh_nodes:
                    if node.get("type") == "file" and is_supported_file(node["name"]):
                        logical = _logical_github_name(node["name"])
                        logical_to_files.setdefault(logical, []).append(node["name"])
            except Exception as e:
                st.warning(f"GitHubファイル一覧の取得に失敗しました: {e}")
                logical_to_files = {}

            if logical_to_files:
                st.write("**デフォルトで選択しておきたい論理名にチェックしてください。**")
                for logical in sorted(logical_to_files.keys()):
                    key = f"sidebar_gh_default::{logical}"
                    # 保存済み設定を初期値とする
                    if key not in st.session_state:
                        st.session_state[key] = logical in saved_gh_set

                    examples = ", ".join(logical_to_files[logical][:3])
                    if len(logical_to_files[logical]) > 3:
                        examples += " など"

                    st.checkbox(
                        logical,
                        key=key,
                        help=f"例: {examples}",
                    )
            else:
                st.info("GitHub上に対象のCSV/Excelファイルが見つかりませんでした。")

        # ===========================
        # 💾 保存・リセットボタン
        # ===========================
        with st.container(border=True):
            st.markdown("**💾 設定の保存／リセット**")
            st.caption("設定を変更したら『設定保存』を押すと次回以降も引き継がれます。")

            if st.button("💾 設定保存", use_container_width=True):
                if editable_calendar_options:
                    calendar_options = list(editable_calendar_options.keys())
                    default_calendar = st.session_state.get(
                        "sidebar_default_calendar", calendar_options[0]
                    )

                    set_user_setting(
                        user_id, "selected_calendar_name", default_calendar
                    )
                    save_user_setting_to_firestore(
                        user_id, "selected_calendar_name", default_calendar
                    )
                    st.session_state["selected_calendar_name"] = default_calendar

                    if st.session_state.get(
                        "share_calendar_selection_across_tabs", True
                    ):
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

                # カレンダー系
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

                # GitHub 論理名デフォルト（チェックされたものだけ保存）
                selected_logicals = []
                for key, val in st.session_state.items():
                    if key.startswith("sidebar_gh_default::") and val:
                        logical = key.split("::", 1)[1]
                        selected_logicals.append(logical)

                selected_logicals = sorted(set(selected_logicals))
                default_gh_text = "\n".join(selected_logicals)

                set_user_setting(
                    user_id, "default_github_logical_names", default_gh_text
                )
                save_user_setting_to_firestore(
                    user_id, "default_github_logical_names", default_gh_text
                )

                # タブ側で使えるようにセッションにも反映
                st.session_state["default_github_logical_names"] = default_gh_text

                st.toast("設定を保存しました", icon="✅")

            if st.button("🔄 リセット", use_container_width=True):
                for key in [
                    "default_private_event",
                    "default_allday_event",
                    "default_create_todo",
                    "default_github_logical_names",
                    "selected_calendar_name",
                    "share_calendar_selection_across_tabs",
                ]:
                    set_user_setting(user_id, key, None)
                    save_user_setting_to_firestore(user_id, key, None)

                # セッション上の GitHub / カレンダー関連キーをクリア
                for k in list(st.session_state.keys()):
                    if k.startswith("sidebar_gh_default::"):
                        del st.session_state[k]
                for k in [
                    "default_github_logical_names",
                    "sidebar_default_calendar",
                    "selected_calendar_name",
                    "share_calendar_selection_across_tabs",
                ]:
                    if k in st.session_state:
                        del st.session_state[k]

                st.toast("設定をリセットしました", icon="🧹")
                st.rerun()

        st.divider()

        # ===========================
        # 📡 接続ステータス（折りたたみ）
        # ===========================
        with st.expander("📡 接続ステータス", expanded=False):
            st.caption("各種APIとの接続状態の確認用です。")

            firebase_ok = bool(user_id)
            calendar_ok = bool(st.session_state.get("calendar_service"))
            tasks_ok = bool(st.session_state.get("tasks_service"))
            sheets_ok = bool(st.session_state.get("sheets_service"))

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
        st.caption("ログアウトすると、次回は再度ログインが必要です。")
        if st.button("🚪 ログアウト", type="primary", use_container_width=True):
            if user_id:
                clear_user_settings(user_id)
            for key in list(st.session_state.keys()):
                if not key.startswith("google_auth") and not key.startswith(
                    "firebase_"
                ):
                    del st.session_state[key]
            st.rerun()