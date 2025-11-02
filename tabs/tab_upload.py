# tabs/tab_upload.py
import streamlit as st
import pandas as pd
from typing import List
from io import BytesIO

from github_loader import walk_repo_tree, load_file_bytes_from_github, is_supported_file
from excel_parser import _load_and_merge_dataframes


def render_tab_upload():
    """タブ1：ファイルのアップロード"""
    st.subheader("ファイルをアップロード")

    # Session state 初期化
    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []
    if "description_columns_pool" not in st.session_state:
        st.session_state["description_columns_pool"] = []
    if "merged_df_for_selector" not in st.session_state:
        st.session_state["merged_df_for_selector"] = pd.DataFrame()

    with st.expander("ℹ️作業手順と補足"):
        st.info(
            """
**☀ 検索した作業指示書をダウンロードしてココにアップロードする事でGoogleカレンダーへ登録できます**

**☀ 北海道現場一覧、備考ファイルを選択する事でGoogleカレンダーへ住所、備考他様々な情報を入力することができます**
            """
        )

    # ---- ローカルファイルアップロード ----
    uploaded_files = st.file_uploader(
        "ExcelまたはCSVファイルを選択（複数可）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )

    # ---- GitHubファイル選択 ----
    selected_github_files: List[BytesIO] = []

    try:
        gh_nodes = walk_repo_tree(base_path="", max_depth=3)
        st.markdown("📦 **GitHub上のCSV/Excel（全ツリー）**")

        if "gh_checked" not in st.session_state:
            st.session_state["gh_checked"] = {}

        for node in gh_nodes:
            if node["type"] == "file" and is_supported_file(node["name"]):
                key = f"gh::{node['path']}"
                checked = st.checkbox(node["name"], key=key, value=st.session_state["gh_checked"].get(key, False))
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

    # ---- 新規ファイル追加処理（重複を避ける）----
    new_files = []
    if uploaded_files:
        new_files.extend(uploaded_files)
    if selected_github_files:
        new_files.extend(selected_github_files)

    if new_files:
        existing_names = [getattr(f, "name", None) for f in st.session_state["uploaded_files"]]

        for f in new_files:
            if getattr(f, "name", None) not in existing_names:
                st.session_state["uploaded_files"].append(f)

        # マージデータのロード
        try:
            merged = _load_and_merge_dataframes(st.session_state["uploaded_files"])
            st.session_state["merged_df_for_selector"] = merged
            st.session_state["description_columns_pool"] = merged.columns.tolist()
        except Exception as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")

    # ---- UI表示 ----
    if st.session_state["uploaded_files"]:
        st.subheader("📄 現在の処理対象ファイル一覧")
        for f in st.session_state["uploaded_files"]:
            st.write(f"- {getattr(f, 'name', '不明な名前')}")

        st.info(
            f"📊 データ列数: {len(st.session_state['merged_df_for_selector'].columns)}、"
            f"行数: {len(st.session_state['merged_df_for_selector'])}"
        )

        if st.button("🗑️ アップロード済みファイルをクリア", help="登録済みファイルとデータを削除します。"):
            st.session_state["uploaded_files"] = []
            st.session_state["merged_df_for_selector"] = pd.DataFrame()
            st.session_state["description_columns_pool"] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()
