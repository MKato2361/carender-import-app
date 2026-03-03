from __future__ import annotations

import os
import re
import streamlit as st

from typing import Dict, Optional, Callable
from session_utils import get_user_setting, set_user_setting, clear_user_settings
from github_loader import (
    _headers,
    GITHUB_OWNER,
    GITHUB_REPO,
    walk_repo_tree,
    is_supported_file,
)


# ──────────────────────────────────────────────
# ユーティリティ
# ──────────────────────────────────────────────

def _logical_github_name(filename: str) -> str:
    """末尾の数字（日付）を除いた論理名に変換"""
    base, _ = os.path.splitext(filename)
    return re.sub(r"\d+$", "", base)


def _resolve(user_id: str, key: str, default, session_key: str | None = None):
    """セッション → Firestore → デフォルト値 の順で設定値を解決"""
    sk = session_key or key
    val = st.session_state.get(sk)
    if val is None:
        val = get_user_setting(user_id, key)
    return val if val is not None else default


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_github_files() -> Dict[str, list[str]]:
    """GitHub ファイル一覧を5分キャッシュで取得"""
    logical_to_files: Dict[str, list[str]] = {}
    try:
        nodes = walk_repo_tree(base_path="", max_depth=3)
        for node in nodes:
            if node.get("type") == "file" and is_supported_file(node["name"]):
                logical = _logical_github_name(node["name"])
                logical_to_files.setdefault(logical, []).append(node["name"])
    except Exception:
        pass
    return logical_to_files


def _has_unsaved_changes(user_id: str) -> bool:
    """Firestore の保存値とセッション値を比較して未保存変更があるか判定"""
    checks = [
        ("default_private_event", "sidebar_default_private", True),
        ("default_allday_event",  "sidebar_default_allday",  False),
        ("default_create_todo",   "sidebar_default_todo",    False),
    ]
    for store_key, session_key, default in checks:
        stored  = get_user_setting(user_id, store_key)
        current = st.session_state.get(session_key)
        if current is None:
            continue
        stored_val = stored if stored is not None else default
        if current != stored_val:
            return True

    # GitHub 論理名
    saved_text = get_user_setting(user_id, "default_github_logical_names") or ""
    saved_set  = {l.strip() for l in saved_text.splitlines() if l.strip()}
    current_set = {
        k.split("::", 1)[1]
        for k, v in st.session_state.items()
        if k.startswith("sidebar_gh_default::") and v
    }
    if current_set != saved_set:
        return True

    return False


def _do_save(
    user_id: str,
    editable_calendar_options: Dict[str, str],
    save_to_firestore: Callable,
) -> None:
    """設定を Firestore・セッションへ一括保存"""
    calendar_options = list(editable_calendar_options.keys()) if editable_calendar_options else []

    # ── カレンダー ──
    if calendar_options:
        cal = st.session_state.get("sidebar_default_calendar", calendar_options[0])
        set_user_setting(user_id, "selected_calendar_name", cal)
        save_to_firestore(user_id, "selected_calendar_name", cal)
        st.session_state["selected_calendar_name"] = cal
        st.session_state["base_calendar_name"] = cal

        if st.session_state.get("share_calendar_selection_across_tabs", True):
            for suffix in ["register", "delete", "export", "inspection_todo",
                           "notice_fax", "property_master", "admin"]:
                st.session_state[f"selected_calendar_name_{suffix}"] = cal

    # ── イベントデフォルト ──
    for key, session_key, default in [
        ("default_private_event", "sidebar_default_private", True),
        ("default_allday_event",  "sidebar_default_allday",  False),
        ("default_create_todo",   "sidebar_default_todo",    False),
    ]:
        val = st.session_state.get(session_key, default)
        set_user_setting(user_id, key, val)
        save_to_firestore(user_id, key, val)

    # ── GitHub 論理名 ──
    selected = sorted({
        k.split("::", 1)[1]
        for k, v in st.session_state.items()
        if k.startswith("sidebar_gh_default::") and v
    })
    gh_text = "\n".join(selected)
    set_user_setting(user_id, "default_github_logical_names", gh_text)
    save_to_firestore(user_id, "default_github_logical_names", gh_text)
    st.session_state["default_github_logical_names"] = gh_text


def _do_reset(user_id: str, save_to_firestore: Callable) -> None:
    """全設定を Firestore・セッションからリセット"""
    keys = [
        "default_private_event", "default_allday_event", "default_create_todo",
        "default_github_logical_names", "selected_calendar_name",
        "share_calendar_selection_across_tabs",
    ]
    for key in keys:
        set_user_setting(user_id, key, None)
        save_to_firestore(user_id, key, None)

    for k in list(st.session_state.keys()):
        if k.startswith("sidebar_gh_default::"):
            del st.session_state[k]
    for k in ["default_github_logical_names", "sidebar_default_calendar",
              "selected_calendar_name", "share_calendar_selection_across_tabs",
              "_confirm_reset"]:
        st.session_state.pop(k, None)


# ──────────────────────────────────────────────
# メイン描画
# ──────────────────────────────────────────────

def render_sidebar(
    user_id: str,
    editable_calendar_options: Optional[Dict[str, str]],
    save_user_setting_to_firestore: Callable[[str, str, object], None],
) -> None:
    """サイドバー全体を描画する"""

    with st.sidebar:
        st.subheader("⚙️ 設定")

        # ════════════════════════════════
        # 📅 カレンダー設定
        # ════════════════════════════════
        with st.expander("📅 カレンダー設定", expanded=True):
            if editable_calendar_options:
                calendar_options = list(editable_calendar_options.keys())

                # 有効なカレンダーを解決
                stored  = get_user_setting(user_id, "selected_calendar_name")
                session = st.session_state.get("sidebar_default_calendar")
                effective = (
                    session if session in calendar_options else
                    stored  if stored  in calendar_options else
                    calendar_options[0]
                )
                if st.session_state.get("sidebar_default_calendar") not in calendar_options:
                    st.session_state["sidebar_default_calendar"] = effective

                default_calendar = st.selectbox(
                    "基準カレンダー",
                    calendar_options,
                    key="sidebar_default_calendar",
                )
                st.session_state["selected_calendar_name"] = default_calendar
                st.session_state["base_calendar_name"]     = default_calendar

                # カレンダー変更は即時保存（UI上の変更と Firestore を常に同期）
                if stored != default_calendar:
                    set_user_setting(user_id, "selected_calendar_name", default_calendar)
                    save_user_setting_to_firestore(user_id, "selected_calendar_name", default_calendar)

                # タブ間共有チェックボックス
                share_prev = _resolve(
                    user_id, "share_calendar_selection_across_tabs", True,
                    "share_calendar_selection_across_tabs"
                )
                st.session_state.setdefault("share_calendar_selection_across_tabs", share_prev)

                share_calendar = st.checkbox(
                    "タブ間でカレンダー選択を共有",
                    key="share_calendar_selection_across_tabs",
                    help="ONにすると、登録タブで選んだカレンダーが他タブにも反映されます。",
                )
                if share_calendar != share_prev:
                    set_user_setting(user_id, "share_calendar_selection_across_tabs", share_calendar)
                    save_user_setting_to_firestore(user_id, "share_calendar_selection_across_tabs", share_calendar)
                    st.rerun()
            else:
                st.info("編集可能なカレンダーが取得できていません。認証状態や権限を確認してください。")

            st.markdown("---")

            # 新規イベントのデフォルト
            st.markdown("**新規イベントのデフォルト**")
            st.checkbox(
                "標準で「非公開」",
                value=_resolve(user_id, "default_private_event", True),
                key="sidebar_default_private",
            )
            st.checkbox(
                "標準で「終日」",
                value=_resolve(user_id, "default_allday_event", False),
                key="sidebar_default_allday",
            )

        # ════════════════════════════════
        # ✅ ToDo設定
        # ════════════════════════════════
        with st.expander("✅ ToDo設定", expanded=False):
            st.caption("新規イベント作成時に、同時に ToDo を発行するかどうかを設定します。")
            st.checkbox(
                "標準で「ToDo作成」",
                value=_resolve(user_id, "default_create_todo", False),
                key="sidebar_default_todo",
            )

        # ════════════════════════════════
        # 📦 GitHubファイル設定
        # ════════════════════════════════
        with st.expander("📦 GitHubデフォルトファイル", expanded=False):
            st.caption(
                "末尾の日付を除いた『論理名』単位でデフォルト選択を設定します。"
                "チェックしたファイルはアップロードタブの初期選択に反映されます。"
            )

            saved_gh_text = _resolve(user_id, "default_github_logical_names", "")
            st.session_state.setdefault("default_github_logical_names", saved_gh_text)
            saved_gh_set = {l.strip() for l in saved_gh_text.splitlines() if l.strip()}

            logical_to_files = _fetch_github_files()

            if logical_to_files:
                for logical in sorted(logical_to_files.keys()):
                    key = f"sidebar_gh_default::{logical}"
                    st.session_state.setdefault(key, logical in saved_gh_set)
                    examples = ", ".join(logical_to_files[logical][:3])
                    if len(logical_to_files[logical]) > 3:
                        examples += " など"
                    st.checkbox(logical, key=key, help=f"例: {examples}")
            else:
                st.info("対象の CSV/Excel ファイルが GitHub 上に見つかりませんでした。")

        # ════════════════════════════════
        # 💾 保存・リセット
        # ════════════════════════════════
        st.divider()
        unsaved = _has_unsaved_changes(user_id)

        # 未保存バッジ
        if unsaved:
            st.caption("🟡 未保存の変更があります")

        col_save, col_reset = st.columns([3, 2])

        with col_save:
            if st.button(
                "💾 設定保存",
                type="primary" if unsaved else "secondary",
                use_container_width=True,
            ):
                _do_save(user_id, editable_calendar_options or {}, save_user_setting_to_firestore)
                st.session_state.pop("_confirm_reset", None)
                st.toast("設定を保存しました ✅")
                st.rerun()

        with col_reset:
            # リセット：1回目でフラグON → 2回目で実行
            confirm = st.session_state.get("_confirm_reset", False)

            if not confirm:
                if st.button("🔄 リセット", use_container_width=True):
                    st.session_state["_confirm_reset"] = True
                    st.rerun()
            else:
                st.warning("本当にリセットしますか？")
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("✅ はい", use_container_width=True):
                        _do_reset(user_id, save_user_setting_to_firestore)
                        st.toast("設定をリセットしました 🧹")
                        st.rerun()
                with c2:
                    if st.button("❌ いいえ", use_container_width=True):
                        st.session_state.pop("_confirm_reset", None)
                        st.rerun()

        st.divider()

        # ════════════════════════════════
        # 📡 接続ステータス
        # ════════════════════════════════
        with st.expander("📡 接続ステータス", expanded=False):
            st.caption("各種 API との接続状態です。")

            firebase_ok  = bool(user_id)
            calendar_ok  = bool(st.session_state.get("calendar_service"))
            tasks_ok     = bool(st.session_state.get("tasks_service"))
            sheets_ok    = bool(st.session_state.get("sheets_service"))

            token_in_secrets = False
            try:
                token_in_secrets = bool(st.secrets.get("GITHUB_PAT", ""))
            except Exception:
                pass

            token_in_headers = False
            try:
                token_in_headers = bool(_headers().get("Authorization"))
            except Exception:
                pass

            github_ok = bool(GITHUB_OWNER and GITHUB_REPO) and (token_in_secrets or token_in_headers)

            def _icon(ok: bool) -> str:
                return "✅" if ok else "⚠️"

            st.markdown(
                f"- **Firebase**：{_icon(firebase_ok)} {'ログイン中' if firebase_ok else '未ログイン'}\n"
                f"- **Google Calendar**：{_icon(calendar_ok)} {'接続中' if calendar_ok else '未接続'}\n"
                f"- **Google Tasks**：{'✅ 利用可' if tasks_ok else '⛔ 利用不可'}\n"
                f"- **Google Sheets**：{'✅ 利用可' if sheets_ok else '⛔ 利用不可'}\n"
                f"- **GitHub**：{_icon(github_ok)} {'設定済' if github_ok else '未設定またはエラー'}\n"
            )

        st.divider()

        # ════════════════════════════════
        # 🚪 ログアウト
        # ════════════════════════════════
        st.caption("ログアウトすると次回は再ログインが必要です。")
        if st.button("🚪 ログアウト", type="primary", use_container_width=True):
            if user_id:
                clear_user_settings(user_id)
            for key in list(st.session_state.keys()):
                if not key.startswith("google_auth") and not key.startswith("firebase_"):
                    del st.session_state[key]
            st.rerun()
