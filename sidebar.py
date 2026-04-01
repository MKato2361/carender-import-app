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
    """未保存の変更があるか判定"""
    checks = [
        ("default_private_event", "sidebar_default_private", True),
        ("default_allday_event",  "sidebar_default_allday",  False),
        ("default_create_todo",   "sidebar_default_todo",    False),
    ]
    for store_key, session_key, default in checks:
        stored  = get_user_setting(user_id, store_key)
        current = st.session_state.get(session_key)
        if current is None: continue
        stored_val = stored if stored is not None else default
        if current != stored_val: return True

    # GitHub 論理名
    saved_text = get_user_setting(user_id, "default_github_logical_names") or ""
    saved_set  = {l.strip() for l in saved_text.splitlines() if l.strip()}
    current_set = {
        k.split("::", 1)[1]
        for k, v in st.session_state.items()
        if k.startswith("sidebar_gh_default::") and v
    }
    return current_set != saved_set

def _do_save(user_id: str, editable_calendar_options: Dict[str, str], save_to_firestore: Callable) -> None:
    """設定を一括保存"""
    calendar_options = list(editable_calendar_options.keys()) if editable_calendar_options else []
    if calendar_options:
        cal = st.session_state.get("sidebar_default_calendar", calendar_options[0])
        set_user_setting(user_id, "selected_calendar_name", cal)
        save_to_firestore(user_id, "selected_calendar_name", cal)
        st.session_state["selected_calendar_name"] = cal
        st.session_state["base_calendar_name"] = cal
        if st.session_state.get("share_calendar_selection_across_tabs", True):
            for suffix in ["register", "delete", "export", "inspection_todo", "notice_fax", "property_master", "admin"]:
                st.session_state[f"selected_calendar_name_{suffix}"] = cal

    for key, session_key, default in [
        ("default_private_event", "sidebar_default_private", True),
        ("default_allday_event",  "sidebar_default_allday",  False),
        ("default_create_todo",   "sidebar_default_todo",    False),
    ]:
        val = st.session_state.get(session_key, default)
        set_user_setting(user_id, key, val)
        save_to_firestore(user_id, key, val)

    selected = sorted({k.split("::", 1)[1] for k, v in st.session_state.items() if k.startswith("sidebar_gh_default::") and v})
    gh_text = "\n".join(selected)
    set_user_setting(user_id, "default_github_logical_names", gh_text)
    save_to_firestore(user_id, "default_github_logical_names", gh_text)
    st.session_state["default_github_logical_names"] = gh_text

def _do_reset(user_id: str, save_to_firestore: Callable) -> None:
    """全設定をリセット"""
    keys = ["default_private_event", "default_allday_event", "default_create_todo", "default_github_logical_names", "selected_calendar_name", "share_calendar_selection_across_tabs"]
    for key in keys:
        set_user_setting(user_id, key, None)
        save_to_firestore(user_id, key, None)
    for k in list(st.session_state.keys()):
        if k.startswith("sidebar_gh_default::"): st.session_state.pop(k, None)
    for k in ["default_github_logical_names", "sidebar_default_calendar", "selected_calendar_name", "share_calendar_selection_across_tabs", "_confirm_reset"]:
        st.session_state.pop(k, None)

# ──────────────────────────────────────────────
# メイン描画
# ──────────────────────────────────────────────

def render_sidebar(
    user_id: str,
    editable_calendar_options: Optional[Dict[str, str]],
    save_user_setting_to_firestore: Callable[[str, str, object], None],
) -> None:
    """サイドバー全体をモダンに描画する"""

    with st.sidebar:
        st.title("⚙️ App Settings")
        
        # ════════════════════════════════
        # 📅 メイン設定: カレンダー選択 (常に表示)
        # ════════════════════════════════
        st.markdown("### 📅 カレンダー選択")
        if editable_calendar_options:
            calendar_options = list(editable_calendar_options.keys())
            stored = get_user_setting(user_id, "selected_calendar_name")
            session = st.session_state.get("sidebar_default_calendar")
            effective = session if session in calendar_options else stored if stored in calendar_options else calendar_options[0]
            
            if st.session_state.get("sidebar_default_calendar") not in calendar_options:
                st.session_state["sidebar_default_calendar"] = effective

            default_calendar = st.selectbox(
                "基準カレンダー",
                calendar_options,
                key="sidebar_default_calendar",
                label_visibility="collapsed"
            )
            st.session_state["selected_calendar_name"] = default_calendar
            st.session_state["base_calendar_name"] = default_calendar

            if stored != default_calendar:
                set_user_setting(user_id, "selected_calendar_name", default_calendar)
                save_user_setting_to_firestore(user_id, "selected_calendar_name", default_calendar)

            share_prev = _resolve(user_id, "share_calendar_selection_across_tabs", True, "share_calendar_selection_across_tabs")
            st.session_state.setdefault("share_calendar_selection_across_tabs", share_prev)
            share_calendar = st.toggle("全タブで選択を共有", key="share_calendar_selection_across_tabs")
            if share_calendar != share_prev:
                set_user_setting(user_id, "share_calendar_selection_across_tabs", share_calendar)
                save_user_setting_to_firestore(user_id, "share_calendar_selection_across_tabs", share_calendar)
                st.rerun()
        else:
            st.warning("カレンダーが取得できません")

        st.divider()

        # ════════════════════════════════
        # 🛠️ 詳細設定 (折りたたみ)
        # ════════════════════════════════
        with st.expander("🛠️ 詳細設定", expanded=False):
            st.markdown("**イベントの初期値**")
            st.checkbox("標準で「非公開」", value=_resolve(user_id, "default_private_event", True), key="sidebar_default_private")
            st.checkbox("標準で「終日」", value=_resolve(user_id, "default_allday_event", False), key="sidebar_default_allday")
            
            st.divider()
            st.markdown("**ToDo連携**")
            st.checkbox("標準で「ToDo作成」", value=_resolve(user_id, "default_create_todo", False), key="sidebar_default_todo")

        # ════════════════════════════════
        # 📦 GitHub連携 (折りたたみ)
        # ════════════════════════════════
        with st.expander("📦 GitHubファイル連携", expanded=False):
            st.caption("デフォルトで選択するファイル（論理名）を指定します。")
            saved_gh_text = _resolve(user_id, "default_github_logical_names", "")
            st.session_state.setdefault("default_github_logical_names", saved_gh_text)
            saved_gh_set = {l.strip() for l in saved_gh_text.splitlines() if l.strip()}
            
            logical_to_files = _fetch_github_files()
            if logical_to_files:
                for logical in sorted(logical_to_files.keys()):
                    key = f"sidebar_gh_default::{logical}"
                    st.session_state.setdefault(key, logical in saved_gh_set)
                    st.checkbox(logical, key=key)
            else:
                st.info("ファイルが見つかりません")

        # ════════════════════════════════
        # 📡 接続状況 (折りたたみ)
        # ════════════════════════════════
        with st.expander("📡 システム接続状況", expanded=False):
            firebase_ok = bool(user_id)
            calendar_ok = bool(st.session_state.get("calendar_service"))
            tasks_ok = bool(st.session_state.get("tasks_service"))
            sheets_ok = bool(st.session_state.get("sheets_service"))
            
            def _status_badge(ok: bool, label: str):
                icon = "🟢" if ok else "🔴"
                st.markdown(f"{icon} **{label}**")

            _status_badge(firebase_ok, "Firebase")
            _status_badge(calendar_ok, "Google Calendar")
            _status_badge(tasks_ok, "Google Tasks")
            _status_badge(sheets_ok, "Google Sheets")

        # ════════════════════════════════
        # 💾 保存・リセット (下部に固定的な配置)
        # ════════════════════════════════
        st.divider()
        unsaved = _has_unsaved_changes(user_id)
        if unsaved:
            st.warning("⚠️ 未保存の変更があります", icon="🔔")

        col_save, col_reset = st.columns(2)
        with col_save:
            if st.button("💾 保存", type="primary", use_container_width=True):
                _do_save(user_id, editable_calendar_options or {}, save_user_setting_to_firestore)
                st.toast("設定を保存しました ✅")
                st.rerun()
        with col_reset:
            if st.button("🔄 リセット", use_container_width=True):
                st.session_state["_confirm_reset"] = True
                st.rerun()

        if st.session_state.get("_confirm_reset"):
            st.error("設定をリセットしますか？")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("はい", use_container_width=True):
                    _do_reset(user_id, save_user_setting_to_firestore)
                    st.rerun()
            with c2:
                if st.button("いいえ", use_container_width=True):
                    st.session_state.pop("_confirm_reset", None)
                    st.rerun()

        st.divider()
        
        # 🚪 ログアウト
        if st.button("🚪 ログアウト", use_container_width=True, help="セッションを終了します"):
            if user_id: clear_user_settings(user_id)
            for key in list(st.session_state.keys()):
                if not key.startswith("google_auth") and not key.startswith("firebase_"):
                    del st.session_state[key]
            st.rerun()
