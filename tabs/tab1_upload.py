import streamlit as st
import os
import re
from typing import List
from io import BytesIO

from github_loader import walk_repo_tree, load_file_bytes_from_github, is_supported_file
from utils.file_loader import (
    update_uploaded_files,
    clear_uploaded_files,
    merge_uploaded_files,
    has_merged_data,
)


def _logical_github_name(filename: str) -> str:
    """
    GitHubファイル名から末尾の連続した数字（例: 日付のバージョン）を取り除いた論理名を返す。
    例: '北海道現場一覧20251127.xlsx' → '北海道現場一覧'
    """
    base, _ext = os.path.splitext(filename)
    # ベース名の末尾に並んでいる数字だけを削除
    base = re.sub(r"\d+$", "", base)
    return base


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
        # 論理名（末尾の日付除去）ごとの選択状態を保持する
        st.session_state["gh_checked"] = {}
    if "upload_version" not in st.session_state:
        st.session_state["upload_version"] = 0
    if "gh_version" not in st.session_state:
        st.session_state["gh_version"] = 0

    # --- GitHub デフォルト論理名（サイドバー設定） ---
    # サイドバーの「📦 GitHubファイル設定」で保存した値を利用する想定
    default_gh_logicals = set()
    default_gh_text = st.session_state.get("default_github_logical_names", "")
    if isinstance(default_gh_text, str):
        default_gh_logicals = {
            line.strip()
            for line in default_gh_text.splitlines()
            if line.strip()
        }

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
        key=f"work_uploader_{st.session_state['upload_version']}",
    )

    uploaded_outside_file = st.file_uploader(
        "🗂️ 作業外予定一覧ファイルを選択（1ファイルのみ）",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        disabled=disable_outside_upload,
        key=f"outside_uploader_{st.session_state['upload_version']}",
    )

    selected_github_files: List[BytesIO] = []

    if not has_outside_work:
        try:
            gh_nodes = walk_repo_tree(base_path="", max_depth=3)
            st.markdown("📦 **GitHub上のCSV/Excel（作業指示書用）**")
            for node in gh_nodes:
                if node["type"] == "file" and is_supported_file(node["name"]):
                    # ファイル名から論理名を作成（末尾の日付部分を無視）
                    logical_key = _logical_github_name(node["name"])
                    widget_key = f"gh::{st.session_state['gh_version']}::{node['path']}"

                    # ① gh_checked に記録されている選択状態をベースに
                    prev_checked_for_logical = st.session_state["gh_checked"].get(
                        logical_key
                    )

                    initial_checked = bool(prev_checked_for_logical)
                    # ② まだ一度も選択されたことがない論理名で、
                    #    サイドバーのデフォルト一覧に含まれていれば初期ON
                    if prev_checked_for_logical is None and logical_key in default_gh_logicals:
                        initial_checked = True

                    # ③ ウィジェットの state がまだ無ければ、ここで初期値を流し込む
                    if widget_key not in st.session_state:
                        st.session_state[widget_key] = initial_checked

                    # ✔ チェックボックス（value は省略・state を使う）
                    checked = st.checkbox(
                        node["name"],
                        key=widget_key,
                        disabled=disable_work_upload,
                    )

                    # 論理名ごとの選択状況を記録（末尾の日付が変わっても維持）
                    st.session_state["gh_checked"][logical_key] = checked

                    if checked and not disable_work_upload:
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
        st.success(
            f"作業外予定一覧ファイルを読み込みました：{uploaded_outside_file.name}"
        )

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

        # GitHub選択状態も完全クリア
        st.session_state["gh_checked"] = {}
        keys_to_delete = [
            k for k in list(st.session_state.keys()) if k.startswith("gh::")
        ]
        for k in keys_to_delete:
            st.session_state.pop(k, None)

        st.session_state["upload_version"] += 1
        st.session_state["gh_version"] += 1

        st.success("アップロード済みファイルとGitHub選択をすべてクリアしました。")
        st.rerun()
