import streamlit as st
import os
import re
from typing import List
from io import BytesIO

from github_loader import (
    walk_repo_tree_with_dates,
    load_file_bytes_from_github,
    is_supported_file,
    list_dir,
)
from utils.file_loader import (
    update_uploaded_files,
    clear_uploaded_files,
    merge_uploaded_files,
    has_merged_data,
)

def _logical_github_name(filename: str) -> str:
    """
    GitHubファイル名から末尾の連続した数字を取り除いた論理名を返す。
    例: '北海道現場一覧20251127.xlsx' → '北海道現場一覧'
    """
    base, _ext = os.path.splitext(filename)
    base = re.sub(r"\d+$", "", base)
    return base

def _clear_github_cache():
    """GitHub API関連のキャッシュをクリアする。"""
    list_dir.clear()
    walk_repo_tree_with_dates.clear()

def _inject_navigate_js():
    """「2. 登録・削除」タブをJSで自動クリックする"""
    st.markdown("""
<script>
(function() {
    function tryClick() {
        var tabs = window.parent.document.querySelectorAll('[data-baseweb="tab"]');
        for (var i = 0; i < tabs.length; i++) {
            if (tabs[i].innerText.indexOf('登録・削除') !== -1) {
                tabs[i].click();
                return true;
            }
        }
        return false;
    }
    if (!tryClick()) {
        var n = 0;
        var timer = setInterval(function() {
            if (tryClick() || ++n > 20) clearInterval(timer);
        }, 80);
    }
})();
</script>
""", unsafe_allow_html=True)


def render_tab1_upload():
    """
    タブ1: ファイルのアップロードと管理
    AuthManager導入後の main.py から呼び出されることを想定。
    """
    st.subheader("📁 ファイルをアップロード")

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
        "navigate_to_register": False,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # --- タブ自動遷移（確定ボタン押下後） ---
    if st.session_state.get("navigate_to_register"):
        st.session_state["navigate_to_register"] = False
        _inject_navigate_js()

    # --- ヘルプ・ガイド ---
    with st.expander("💡 ファイル形式とアップロードのヒント", expanded=False):
        st.markdown("""
        ### 📌 アップロード可能なファイル
        - **作業指示書一覧**: 複数ファイル可。Excel (.xlsx, .xls) または CSV。
        - **作業外予定一覧**: **1ファイルのみ**。ローカルからのみアップロード可能。
        
        ### 📌 GitHub連携
        - サイドバーの設定で「GitHubデフォルトファイル」を選択しておくと、
          作業指示書アップロード時に自動的にチェックが入ります。
        """)

    # --- アップロード状態の判定 ---
    has_work_files   = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    # 排他制御: 作業指示書と作業外予定は同時には扱わない設計
    disable_work_upload    = has_outside_work
    disable_outside_upload = has_work_files

    # --- ローカルファイルアップローダー ---
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🛠️ 作業指示書 (複数可)")
        uploaded_work_files = st.file_uploader(
            "Excel/CSVを選択",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            disabled=disable_work_upload,
            key=f"work_uploader_{st.session_state['upload_version']}",
            help="複数の現場一覧ファイルをまとめて取り込めます。"
        )

    with col2:
        st.markdown("### 🗓️ 作業外予定 (1件)")
        uploaded_outside_file = st.file_uploader(
            "Excel/CSVを選択",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=False,
            disabled=disable_outside_upload,
            key=f"outside_uploader_{st.session_state['upload_version']}",
            help="個人の予定や作業外のスケジュールを取り込みます。"
        )

    # --- GitHub連携セクション ---
    if not has_outside_work:
        st.divider()
        col_title, col_reload = st.columns([6, 1])
        with col_title:
            st.markdown("### 📦 GitHubリポジトリから選択")
        with col_reload:
            if st.button("🔄", help="GitHubのファイル一覧を更新", disabled=disable_work_upload):
                _clear_github_cache()
                st.session_state["gh_version"] += 1
                st.session_state["gh_defaults_applied"] = False
                st.rerun()

        # GitHubデフォルト論理名の取得
        default_gh_logicals = set()
        default_gh_text = st.session_state.get("default_github_logical_names", "")
        if isinstance(default_gh_text, str):
            default_gh_logicals = {line.strip() for line in default_gh_text.splitlines() if line.strip()}

        auto_apply_gh_defaults_now = (
            bool(uploaded_work_files)
            and not has_outside_work
            and not st.session_state["gh_defaults_applied"]
            and len(default_gh_logicals) > 0
        )

        selected_github_files: List[BytesIO] = []
        try:
            gh_nodes = walk_repo_tree_with_dates(base_path="", max_depth=3)
            file_nodes = [n for n in gh_nodes if n["type"] == "file" and is_supported_file(n["name"])]

            if file_nodes:
                # 2カラムでチェックボックスを表示して省スペース化
                gh_cols = st.columns(2)
                for idx, node in enumerate(file_nodes):
                    logical_key = _logical_github_name(node["name"])
                    widget_key  = f"gh::{st.session_state['gh_version']}::{node['path']}"
                    updated     = node.get("updated", "")

                    if widget_key not in st.session_state:
                        st.session_state[widget_key] = False
                    
                    if auto_apply_gh_defaults_now and logical_key in default_gh_logicals:
                        st.session_state[widget_key] = True

                    label = f"{node['name']} ({updated})" if updated else node["name"]
                    with gh_cols[idx % 2]:
                        checked = st.checkbox(label, key=widget_key, disabled=disable_work_upload)

                    if checked and not disable_work_upload:
                        try:
                            bio = load_file_bytes_from_github(node["path"])
                            bio.name = node["name"]
                            selected_github_files.append(bio)
                        except Exception as e:
                            st.error(f"GitHubファイル '{node['name']}' の取得に失敗: {e}")

                if auto_apply_gh_defaults_now:
                    st.session_state["gh_defaults_applied"] = True
            else:
                st.info("GitHubリポジトリ内に対応するファイルが見つかりませんでした。")
        except Exception as e:
            st.warning(f"GitHub連携エラー: {e}")

    # --- ファイル処理ロジック ---
    if uploaded_outside_file and not has_work_files:
        st.session_state["uploaded_outside_work_file"] = uploaded_outside_file

    new_files: List = []
    if uploaded_work_files and not has_outside_work:
        new_files.extend(uploaded_work_files)
    if selected_github_files and not has_outside_work:
        new_files.extend(selected_github_files)

    if new_files:
        with st.spinner("データを解析中..."):
            update_uploaded_files(new_files)
            merge_uploaded_files()

    # ファイル処理後に状態を再評価（処理前の値は古いため）
    has_work_files   = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    # --- 現在の状態表示 ---
    if has_outside_work or has_work_files:
        st.divider()
        st.markdown("### 📋 現在の取り込み状況")
        
        if has_outside_work:
            f = st.session_state["uploaded_outside_work_file"]
            st.success(f"✅ 作業外予定: **{f.name}**")
        
        if has_work_files:
            st.markdown(f"✅ **作業指示書 ({len(st.session_state['uploaded_files'])} 件)**")
            for f in st.session_state["uploaded_files"]:
                st.caption(f"- {getattr(f, 'name', 'Unknown')}")
            
            if has_merged_data():
                df = st.session_state["merged_df_for_selector"]
                st.info(f"📊 統合データ: {len(df)} 行 / {len(df.columns)} 列")

        # --- 確定セクション ---
        st.divider()
        st.markdown("#### 📋 読み込み内容の確認")

        # 読み込んだファイルのサマリーを表示
        if has_work_files:
            file_names = [getattr(f, "name", "Unknown") for f in st.session_state["uploaded_files"]]
            df_preview = st.session_state.get("merged_df_for_selector")
            row_count = len(df_preview) if df_preview is not None else "—"
            col_count = len(df_preview.columns) if df_preview is not None else "—"

            st.markdown(f"""
<div style="background:var(--color-background-secondary);border:0.5px solid var(--color-border-tertiary);border-radius:10px;padding:14px 18px;margin-bottom:14px;">
  <div style="font-size:12px;color:var(--color-text-secondary);margin-bottom:8px;font-weight:500;">作業指示書 {len(file_names)} ファイル</div>
  {''.join(f'<div style="font-size:13px;color:var(--color-text-primary);padding:2px 0;">• {n}</div>' for n in file_names)}
  <div style="margin-top:10px;padding-top:10px;border-top:0.5px solid var(--color-border-tertiary);font-size:13px;color:var(--color-text-secondary);">
    統合データ：<strong style="color:var(--color-text-primary);">{row_count} 行 / {col_count} 列</strong>
  </div>
</div>
""", unsafe_allow_html=True)

        if has_outside_work:
            f = st.session_state["uploaded_outside_work_file"]
            st.markdown(f"""
<div style="background:var(--color-background-secondary);border:0.5px solid var(--color-border-tertiary);border-radius:10px;padding:14px 18px;margin-bottom:14px;">
  <div style="font-size:12px;color:var(--color-text-secondary);margin-bottom:6px;font-weight:500;">作業外予定</div>
  <div style="font-size:13px;color:var(--color-text-primary);">• {f.name}</div>
</div>
""", unsafe_allow_html=True)

        confirm_col, clear_col = st.columns([3, 1])
        with confirm_col:
            if st.button(
                "✅ この内容で確定してカレンダー登録へ進む →",
                type="primary",
                use_container_width=True,
            ):
                st.session_state["navigate_to_register"] = True
                st.rerun()
        with clear_col:
            pass  # クリアボタンは下に配置

        st.markdown("<div style='margin-top:6px;'></div>", unsafe_allow_html=True)

        # --- クリアボタン ---
        if st.button("🗑️ すべてのアップロードをクリア", type="secondary", use_container_width=True):
            clear_uploaded_files()
            st.session_state["uploaded_outside_work_file"] = None
            st.session_state["merged_df_for_selector"]     = None
            for k in [k for k in st.session_state if k.startswith("gh::")]:
                st.session_state.pop(k, None)
            st.session_state["gh_defaults_applied"] = False
            st.session_state["upload_version"] += 1
            st.session_state["gh_version"]     += 1
            st.rerun()
