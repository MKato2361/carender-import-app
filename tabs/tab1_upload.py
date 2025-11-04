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
    st.subheader("ファイルをアップロード")

    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []
    if "uploaded_outside_work_file" not in st.session_state:
        st.session_state["uploaded_outside_work_file"] = None
    if "merged_df_for_selector" not in st.session_state:
        st.session_state["merged_df_for_selector"] = None
    if "description_columns_pool" not in st.session_state:
        st.session_state["description_columns_pool"] = []
    if "gh_checked" not in st.session_state:
        st.session_state["gh_checked"] = {}
    if "upload_version" not in st.session_state:
        st.session_state["upload_version"] = 0
    if "gh_version" not in st.session_state:
        st.session_state["gh_version"] = 0

    with st.expander("ℹ️作業手順と補足"):
        st.info(
            """
「作業指示書一覧」または「作業外予定一覧」をアップロードできます（同時不可）

📌 作業指示書 → 複数ファイルOK + GitHubから選択可  
📌 作業外予定 → ローカル1ファイルのみ、GitHub選択不可
"""
        )

    has_work_files = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    disable_work_upload = has_outside_work
    disable_outside_upload = has_work_files

    uploaded_work_files = st.file_uploader(
        "📂 作業指示書一覧ファイルを選択（複数可）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        disabled=disable_work_upload,
        key=f"work_uploader_{st.session_state['upload_version']}"
    )

    uploaded_outside_file = st.file_uploader(
        "🗂️ 作業外予定一覧ファイルを選択（1ファイルのみ）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        disabled=disable_outside_upload,
        key=f"outside_uploader_{st.session_state['upload_version']}"
    )

    selected_github_files: List[BytesIO] = []

    if not has_outside_work:
        try:
            gh_nodes = walk_repo_tree(base_path="", max_depth=3)
            st.markdown("📦 **GitHub上のCSV/Excel（作業指示書用）**")
            for node in gh_nodes:
                if node["type"] == "file" and is_supported_file(node["name"]):
                    key = f"gh::{st.session_state['gh_version']}::{node['path']}"
                    checked = st.checkbox(
                        node["name"],
                        key=key,
                        disabled=disable_work_upload
                    )
                    if checked:
                        try:
                            bio = load_file_bytes_from_github(node["path"])
                            bio.name = node["name"]
                            selected_github_files.append(bio)
                        except Exception as e:
                            st.warning(f"GitHub取得エラー: {e}")
        except Exception as e:
            st.warning(f"GitHubツリー取得失敗: {e}")

    if uploaded_outside_file and not has_work_files:
        st.session_state["uploaded_outside_work_file"] = uploaded_outside_file
        st.success(f"作業外予定一覧ファイルを読み込みました：{uploaded_outside_file.name}")

    new_files = []
    if uploaded_work_files and not has_outside_work:
        new_files.extend(uploaded_work_files)
    if selected_github_files and not has_outside_work:
        new_files.extend(selected_github_files)

    if new_files:
        update_uploaded_files(new_files)
        merge_uploaded_files()

    if has_outside_work:
        f = st.session_state["uploaded_outside_work_file"]
        st.info(f"📄 作業外予定ファイル：{f.name}")

    if has_work_files:
        st.subheader("📄 現在の作業指示書ファイル一覧")
        for f in st.session_state["uploaded_files"]:
            st.write(f"- {getattr(f, 'name', '不明なファイル名')}")
        if has_merged_data():
            df = st.session_state["merged_df_for_selector"]
            st.info(f"📊 データ列数: {len(df.columns)}、行数: {len(df)}")

    if st.button("🗑️ すべてのアップロードファイルをクリア"):
        clear_uploaded_files()
        st.session_state["uploaded_outside_work_file"] = None
        st.session_state["merged_df_for_selector"] = None

        st.session_state["gh_checked"] = {}
        keys_to_delete = [k for k in list(st.session_state.keys()) if k.startswith("gh::")]
        for k in keys_to_delete:
            st.session_state.pop(k, None)

        st.session_state["upload_version"] += 1
        st.session_state["gh_version"] += 1

        st.success("アップロード済みファイルとGitHub選択をすべてクリアしました。")
        st.rerun()
