from __future__ import annotations
# tabs/tab_admin.py

from datetime import datetime

import pandas as pd
import streamlit as st

from utils.user_roles import (
    list_users,
    set_user_role,
    get_or_create_user,
    get_user_role,
    ROLE_ADMIN,
    ROLE_USER,
)
from github_loader import (
    GITHUB_OWNER,
    GITHUB_REPO,
    list_github_files,
    upload_file_to_github,
    get_dir_commit_dates,
)

# 外部タブ・ユーティリティ
from services.calendar_service import get_events as fetch_all_events
from tabs.tab4_duplicates import render_tab4_duplicates

# ==============================
# 管理者タブ UI 本体 (AuthManager対応版)
# ==============================
def render_tab_admin(manager, current_user_email: str) -> None:
    """
    管理者専用タブ。
    manager: AuthManager インスタンス
    current_user_email: ログイン中のユーザーメール
    """
    if not current_user_email:
        st.error("ログイン情報が取得できません。")
        return

    # 権限チェック
    user_doc = get_or_create_user(current_user_email, None)
    role = user_doc.get("role") or get_user_role(current_user_email)
    if role != ROLE_ADMIN:
        st.error("管理者権限が必要です。")
        return

    st.subheader("🔧 システム管理者設定")

    tab_users, tab_files, tab_dup = st.tabs(
        ["👥 ユーザー管理", "📂 GitHubファイル管理", "🔁 重複イベント管理"]
    )

    # --- 1. ユーザー管理 ---
    with tab_users:
        st.markdown("### ユーザー権限の一覧・編集")
        users = list_users()
        if not users:
            st.info("登録されているユーザーはいません。")
        else:
            df = pd.DataFrame(users)
            cols = ["email", "display_name", "role", "updated_at"]
            df = df[[c for c in cols if c in df.columns]]

            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                column_config={
                    "email": st.column_config.TextColumn("メールアドレス", disabled=True),
                    "display_name": st.column_config.TextColumn("名前", disabled=True),
                    "role": st.column_config.SelectboxColumn("ロール", options=[ROLE_USER, ROLE_ADMIN], required=True),
                    "updated_at": st.column_config.DatetimeColumn("最終更新", disabled=True),
                },
                key="admin_users_editor",
            )

            if st.button("💾 変更を保存", type="primary", key="admin_users_save", use_container_width=True):
                with st.spinner("保存中..."):
                    for _, row in edited_df.iterrows():
                        email = str(row.get("email") or "").strip().lower()
                        role_val = str(row.get("role") or ROLE_USER).strip().lower()
                        if email:
                            set_user_role(email, role_val)
                st.success("ユーザー権限を更新しました。")
                st.rerun()

    # --- 2. GitHub ファイル管理 ---
    with tab_files:
        st.markdown("### 📂 GitHub リポジトリ管理")
        st.caption(f"リポジトリ: `{GITHUB_OWNER}/{GITHUB_REPO}`")

        base_path = st.text_input(
            "対象ディレクトリパス",
            value=st.session_state.get("admin_github_base_path", ""),
            placeholder="例: templates / state",
            key="admin_github_base_input",
        )
        st.session_state["admin_github_base_path"] = base_path

        col_reload, _ = st.columns([1, 4])
        if col_reload.button("🔄 一覧を更新", use_container_width=True):
            st.session_state.pop("admin_github_last_list", None)
            get_dir_commit_dates.clear()
            st.rerun()

        # ファイル一覧の取得
        cache_key = "admin_github_last_list"
        if cache_key not in st.session_state:
            try:
                st.session_state[cache_key] = list_github_files(base_path)
            except Exception as e:
                st.error(f"取得エラー: {e}")
                st.session_state[cache_key] = []
        
        items = st.session_state[cache_key]
        file_items = [it for it in items if it.get("type") == "file"]

        if file_items:
            commit_dates = get_dir_commit_dates(base_path)
            
            # 簡易テーブル表示
            display_data = []
            for item in file_items:
                path = item.get("path", "")
                display_data.append({
                    "ファイル名": path,
                    "SHA": item.get("sha", "")[:7],
                    "最終更新": commit_dates.get(path, "-"),
                    "URL": item.get("html_url", "")
                })
            
            st.dataframe(
                pd.DataFrame(display_data),
                hide_index=True,
                use_container_width=True,
                column_config={
                    "URL": st.column_config.LinkColumn("GitHubで開く")
                }
            )
        else:
            st.info("このディレクトリにはファイルがありません。")

        st.divider()
        st.markdown("#### 📤 ファイルをアップロード")
        uploaded_files = st.file_uploader("ファイルを選択", accept_multiple_files=True, key="admin_github_uploader")
        
        if uploaded_files:
            msg = st.text_input("コミットメッセージ", value=f"Admin update: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            if st.button("▶ アップロード実行", type="primary", use_container_width=True):
                clean_base = base_path.strip().strip("/")
                for f in uploaded_files:
                    target_path = f"{clean_base}/{f.name}" if clean_base else f.name
                    try:
                        upload_file_to_github(target_path, f.getvalue(), msg)
                        st.toast(f"成功: {f.name}")
                    except Exception as e:
                        st.error(f"失敗: {f.name} ({e})")
                st.session_state.pop(cache_key, None)
                st.rerun()

    # --- 3. 重複イベント管理 ---
    with tab_dup:
        st.markdown("### 🔁 重複イベントの検出と一括削除")
        if not st.session_state.get("calendar_service"):
            st.warning("カレンダーサービスが利用できません。")
        else:
            # 既存の render_tab4_duplicates を呼び出し
            # 注: render_tab4_duplicates も将来的に manager 対応するのが望ましい
            render_tab4_duplicates(
                st.session_state.get("calendar_service"),
                st.session_state.get("editable_calendar_options", {}),
                fetch_all_events
            )
