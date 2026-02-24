import streamlit as st
import os
import re
from typing import List
from io import BytesIO

from github_loader import (
    walk_repo_tree_with_dates,  # ★ 更新日付き版に変更
    load_file_bytes_from_github,
    is_supported_file,
)
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
    base = re.sub(r"\d+$", "", base)
    return base


def render_tab1_upload():
    st.subheader("ファイルをアップロード")

    # --- session_state 初期化 ---
    defaults = {
        "uploaded_files": [],
        "uploaded_outside_work_file": None,
        "merged_df_for_selector": None,
        "description_columns_pool": [],
        "gh_checked": {},
        "upload_version": 0,
        "gh_version": 0,
        "gh_defaults_applied": False,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # --- GitHub デフォルト論理名（サイドバー設定） ---
    default_gh_logicals: set = set()
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

    has_work_files   = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    disable_work_upload    = has_outside_work
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

    # 今回のrunで自動デフォルト適用するか判定
    auto_apply_gh_defaults_now = (
        bool(uploaded_work_files)
        and not has_outside_work
        and not st.session_state["gh_defaults_applied"]
        and len(default_gh_logicals) > 0
    )

    selected_github_files: List[BytesIO] = []

    # --- GitHub から作業指示書ファイルを選択 ---
    if not has_outside_work:
        try:
            # ★ walk_repo_tree_with_dates で更新日も一緒に取得（キャッシュ済み）
            gh_nodes = walk_repo_tree_with_dates(base_path="", max_depth=3)
            file_nodes = [
                n for n in gh_nodes
                if n["type"] == "file" and is_supported_file(n["name"])
            ]

            if file_nodes:
                st.markdown("📦 **GitHub上のCSV/Excel（作業指示書用）**")

                for node in file_nodes:
                    logical_key = _logical_github_name(node["name"])
                    widget_key  = f"gh::{st.session_state['gh_version']}::{node['path']}"
                    updated      = node.get("updated", "")  # ★ 更新日

                    # 初期化
                    if widget_key not in st.session_state:
                        st.session_state[widget_key] = False

                    # デフォルト論理名に一致する場合は自動ON
                    if auto_apply_gh_defaults_now and logical_key in default_gh_logicals:
                        st.session_state[widget_key] = True

                    # ★ ファイル名に更新日を付けてチェックボックス表示
                    label = f"{node['name']}　`{updated}`" if updated else node["name"]
                    checked = st.checkbox(
                        label,
                        key=widget_key,
                        disabled=disable_work_upload,
                    )

                    st.session_state["gh_checked"][logical_key] = checked

                    if checked and not disable_work_upload:
                        try:
                            bio = load_file_bytes_from_github(node["path"])
                            bio.name = node["name"]
                            selected_github_files.append(bio)
                        except Exception as e:
                            st.warning(f"GitHub取得エラー: {e}")

            if auto_apply_gh_defaults_now:
                st.session_state["gh_defaults_applied"] = True

        except Exception as e:
            st.warning(f"GitHubツリー取得失敗: {e}")

    # --- 作業外予定ファイル（ローカル1ファイル） ---
    if uploaded_outside_file and not has_work_files:
        st.session_state["uploaded_outside_work_file"] = uploaded_outside_file
        st.success(f"作業外予定一覧ファイルを読み込みました：{uploaded_outside_file.name}")

    # --- 新規ファイルをまとめて処理 ---
    new_files: List = []
    if uploaded_work_files and not has_outside_work:
        new_files.extend(uploaded_work_files)
    if selected_github_files and not has_outside_work:
        new_files.extend(selected_github_files)

    if new_files:
        update_uploaded_files(new_files)
        merge_uploaded_files()

    # --- ステータス表示 ---
    if has_outside_work:
        f = st.session_state["uploaded_outside_work_file"]
        st.info(f"📄 作業外予定ファイル：{f.name}")

    if len(st.session_state["uploaded_files"]) > 0:
        st.subheader("📄 現在の作業指示書ファイル一覧")
        for f in st.session_state["uploaded_files"]:
            st.write(f"- {getattr(f, 'name', '不明なファイル名')}")
        if has_merged_data():
            df = st.session_state["merged_df_for_selector"]
            st.info(f"📊 データ列数: {len(df.columns)}、行数: {len(df)}")

    # --- 全クリアボタン ---
    if st.button("🗑️ すべてのアップロードファイルをクリア"):
        clear_uploaded_files()
        st.session_state["uploaded_outside_work_file"] = None
        st.session_state["merged_df_for_selector"]     = None

        # GitHubチェックボックスの状態をクリア
        for k in [k for k in st.session_state if k.startswith("gh::")]:
            st.session_state.pop(k, None)

        st.session_state["gh_defaults_applied"] = False
        st.session_state["upload_version"] += 1
        st.session_state["gh_version"]     += 1

        st.success("アップロード済みファイルとGitHub選択をすべてクリアしました。")
        st.rerun()
