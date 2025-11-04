import streamlit as st
from typing import List
from io import BytesIO

from github_loader import walk_repo_tree, load_file_bytes_from_github, is_supported_file
from utils.file_loader import (
    update_uploaded_files,
    clear_uploaded_files,
    merge_uploaded_files,
    has_merged_data,
)


def render_tab1_upload():
    """タブ1：ファイルのアップロード UI部分"""

    st.subheader("ファイルをアップロード")

    # ===== Session State 初期化 =====
    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []  # 作業指示書一覧用

    if "uploaded_outside_work_file" not in st.session_state:
        st.session_state["uploaded_outside_work_file"] = None  # 作業外予定一覧用（単一ファイル）

    if "description_columns_pool" not in st.session_state:
        st.session_state["description_columns_pool"] = []

    if "merged_df_for_selector" not in st.session_state:
        st.session_state["merged_df_for_selector"] = None

    if "gh_checked" not in st.session_state:
        st.session_state["gh_checked"] = {}

    with st.expander("ℹ️作業手順と補足"):
        st.info(
            """
**☀「作業指示書一覧」または「作業外予定一覧」をアップロードできます（同時選択不可）**

**📌 作業指示書一覧 → 複数ファイルOK + GitHubから選択可**  
**📌 作業外予定一覧 → ローカル1ファイルのみ、GitHub選択不可**
"""
        )

    # --- 状態 ---
    has_work_files = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    disable_work_upload = has_outside_work  # 外予定アップ済 → 作業指示書アップを無効化
    disable_outside_upload = has_work_files  # 作業指示書アップ済 → 外予定アップを無効化

    # ==========================================================
    # ① ローカルアップロード（作業指示書一覧：複数可、GitHub選択可）
    # ==========================================================
    uploaded_work_files = st.file_uploader(
        "📂 作業指示書一覧ファイルを選択（複数可）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        disabled=disable_work_upload,
        help="※ 作業外予定をアップ済みの場合は選択できません",
    )

    # ==========================================================
    # ② ローカルアップロード（作業外予定一覧：1ファイルのみ）
    # ==========================================================
    uploaded_outside_file = st.file_uploader(
        "🗂️ 作業外予定一覧ファイルを選択（1ファイルのみ）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        disabled=disable_outside_upload,
        help="※ 作業指示書一覧をアップ済みの場合は選択できません",
    )

    # ==========================================================
    # GitHub UI（作業指示書のみ表示）
    # ==========================================================
    selected_github_files: List[BytesIO] = []

    if not has_outside_work:  # 外予定アップ時はGitHub UIを非表示
        try:
            gh_nodes = walk_repo_tree(base_path="", max_depth=3)
            st.markdown("📦 **GitHub上のCSV/Excel（作業指示書用）**")

            for node in gh_nodes:
                if node["type"] == "file" and is_supported_file(node["name"]):
                    key = f"gh::{node['path']}"

                    checked = st.checkbox(
                        node["name"],
                        key=key,
                        value=st.session_state["gh_checked"].get(key, False),
                        disabled=disable_work_upload  # 外予定アップ時は操作不可
                    )
                    st.session_state["gh_checked"][key] = checked

                    if checked:
                        try:
                            bio = load_file_bytes_from_github(node["path"])
                            bio.name = node["name"]
                            selected_github_files.append(bio)
                        except Exception as e:
                            st.warning(f"GitHub取得エラー: {e}")

        except Exception as e:
            st.warning(f"GitHubツリーの取得に失敗しました: {e}")

    # ==========================================================
    # アップロードデータの反映
    # ==========================================================
    # 作業外予定一覧
    if uploaded_outside_file and not has_work_files:
        st.session_state["uploaded_outside_work_file"] = uploaded_outside_file
        st.success(f"作業外予定一覧ファイルを読み込みました：{uploaded_outside_file.name}")

    # 作業指示書一覧（ローカル&GitHub）
    new_files = []
    if uploaded_work_files and not has_outside_work:
        new_files.extend(uploaded_work_files)
    if selected_github_files and not has_outside_work:
        new_files.extend(selected_github_files)

    if new_files:
        update_uploaded_files(new_files)
        merge_uploaded_files()

    # ==========================================================
    # 表示ブロック
    # ==========================================================
    # ✅ 作業外予定一覧
    if has_outside_work:
        f = st.session_state["uploaded_outside_work_file"]
        st.info(f"📄 作業外予定ファイル：{f.name}")

    # ✅ 作業指示書一覧
    if has_work_files:
        st.subheader("📄 現在の作業指示書ファイル一覧")
        for f in st.session_state["uploaded_files"]:
            st.write(f"- {getattr(f, 'name', '不明なファイル名')}")

        if has_merged_data():
            df = st.session_state["merged_df_for_selector"]
            st.info(f"📊 データ列数: {len(df.columns)}、行数: {len(df)}")

    # ==========================================================
    # クリアボタン（GitHubのチェックもリセット）
    # ==========================================================
    if st.button("🗑️ すべてのアップロードファイルをクリア", help="登録済みファイルとデータを削除します。"):
        clear_uploaded_files()
        st.session_state["uploaded_outside_work_file"] = None
        st.session_state["gh_checked"] = {}  # ← GitHubのチェック状態リセット追加
        st.success("アップロード済みファイルをクリアしました。")
        st.rerun()
