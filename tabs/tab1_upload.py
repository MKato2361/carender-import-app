import streamlit as st
import streamlit.components.v1 as components
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
    base, _ext = os.path.splitext(filename)
    base = re.sub(r"\d+$", "", base)
    return base

def _clear_github_cache():
    list_dir.clear()
    walk_repo_tree_with_dates.clear()

def _navigate_to_register_tab():
    """st.components.v1.html で JS を実行してタブを切り替える"""
    components.html("""
<script>
(function() {
    function tryClick() {
        var tabs = window.parent.document.querySelectorAll('[data-baseweb="tab"]');
        for (var i = 0; i < tabs.length; i++) {
            if (tabs[i].innerText.indexOf('\u767b\u9332\u30fb\u64cd\u4f5c') !== -1) {
                tabs[i].click();
                return true;
            }
        }
        return false;
    }
    if (!tryClick()) {
        var n = 0;
        var timer = setInterval(function() {
            if (tryClick() || ++n > 30) clearInterval(timer);
        }, 80);
    }
})();
</script>
""", height=0)


def _render_confirm_bar(has_work_files, has_outside_work):
    """ファイル取込済み時のサマリー + 確定ボタン（2段レイアウト）"""
    files   = st.session_state.get("uploaded_files", [])
    outside = st.session_state.get("uploaded_outside_work_file")
    df      = st.session_state.get("merged_df_for_selector")

    if has_work_files:
        names     = [getattr(f, "name", "Unknown") for f in files]
        row_count = len(df) if df is not None else "—"
        badge     = f"{len(names)} ファイル / {row_count} 行"
        summary   = "、".join(names)
        kind      = "作業指示書"
    else:
        summary  = getattr(outside, "name", "")
        badge    = "1 ファイル"
        kind     = "作業外予定"

    # 行1: ファイル情報 + クリアボタン
    col_info, col_clear = st.columns([9, 1])
    with col_info:
        st.markdown(
            f"**{kind}** &nbsp;"
            f"<span style='background:var(--success-surface);"
            f"color:var(--success);font-size:12px;font-weight:600;"
            f"padding:2px 8px;border-radius:4px;'>{badge}</span>"
            f"&nbsp;&nbsp;<span style='font-size:12px;color:var(--text-2);'>{summary}</span>",
            unsafe_allow_html=True,
        )
    with col_clear:
        if st.button("クリア", help="アップロードをクリア", use_container_width=True):
            clear_uploaded_files()
            st.session_state["uploaded_outside_work_file"] = None
            st.session_state["merged_df_for_selector"]     = None
            for k in [k for k in st.session_state if k.startswith("gh::")]:
                st.session_state.pop(k, None)
            st.session_state["gh_defaults_applied"] = False
            st.session_state["upload_version"] += 1
            st.session_state["gh_version"]     += 1
            st.rerun()

    # 行2: 確定ボタン（全幅プライマリ）
    if st.button(
        "カレンダー登録へ進む",
        type="primary",
        use_container_width=True,
    ):
        st.session_state["navigate_to_register"] = True
        st.rerun()


def render_tab1_upload():
    """タブ1: ファイルのアップロードと管理"""

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
        "_gh_version_at_last_apply": -1,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # gh_version が変わったら gh_defaults_applied をリセット（クリア後の再アップロード対応）
    if st.session_state["gh_version"] != st.session_state["_gh_version_at_last_apply"]:
        st.session_state["gh_defaults_applied"] = False

    # --- タブ自動遷移（確定ボタン押下後） ---
    if st.session_state.get("navigate_to_register"):
        st.session_state["navigate_to_register"] = False
        _navigate_to_register_tab()
        # return しない — tab1のDOMを維持する
        # （Streamlitのタブ切り替えはPython再実行なしのクライアントサイド操作のため、
        #   returnするとtab1が空になり、戻ったときに何も表示されなくなる）

    # --- アップロード状態の判定（初期値） ---
    has_work_files   = len(st.session_state["uploaded_files"]) > 0
    has_outside_work = st.session_state["uploaded_outside_work_file"] is not None

    # ── 取込済みならページ最上部にコンパクトサマリー ──
    if has_work_files or has_outside_work:
        _render_confirm_bar(has_work_files, has_outside_work)
        st.divider()

    # --- ローカルファイルアップローダー ---
    disable_work_upload    = has_outside_work
    disable_outside_upload = has_work_files

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-heading"><span class="mi">build</span>作業指示書</div>', unsafe_allow_html=True)
        uploaded_work_files = st.file_uploader(
            "作業指示書",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            disabled=disable_work_upload,
            key=f"work_uploader_{st.session_state['upload_version']}",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown('<div class="section-heading"><span class="mi">event</span>作業外予定</div>', unsafe_allow_html=True)
        uploaded_outside_file = st.file_uploader(
            "作業外予定",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=False,
            disabled=disable_outside_upload,
            key=f"outside_uploader_{st.session_state['upload_version']}",
            label_visibility="collapsed",
        )

    # --- GitHub連携セクション ---
    if not has_outside_work:
        st.divider()
        col_title, col_reload = st.columns([8, 1])
        with col_title:
            st.markdown('<div class="section-heading"><span class="mi">cloud_download</span>GitHubから選択</div>', unsafe_allow_html=True)
        with col_reload:
            if st.button("更新", help="ファイル一覧を更新", disabled=disable_work_upload):
                _clear_github_cache()
                st.session_state["gh_version"] += 1
                st.session_state["gh_defaults_applied"] = False
                st.rerun()

        default_gh_logicals = set()
        default_gh_text = st.session_state.get("default_github_logical_names", "")
        if isinstance(default_gh_text, str):
            default_gh_logicals = {l.strip() for l in default_gh_text.splitlines() if l.strip()}

        # ローカルファイルが存在する（今選択 or セッション既存）かつ未適用の場合に自動選択
        has_any_work = bool(uploaded_work_files) or has_work_files
        auto_apply_gh_defaults_now = (
            has_any_work
            and not has_outside_work
            and not st.session_state["gh_defaults_applied"]
            and len(default_gh_logicals) > 0
        )

        selected_github_files: List[BytesIO] = []
        try:
            gh_nodes = walk_repo_tree_with_dates(base_path="", max_depth=3)
            file_nodes = [n for n in gh_nodes if n["type"] == "file" and is_supported_file(n["name"])]

            if file_nodes:
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
                        except Exception:
                            st.error(f"'{node['name']}' の取得に失敗しました。")

                if auto_apply_gh_defaults_now:
                    st.session_state["gh_defaults_applied"] = True
                    st.session_state["_gh_version_at_last_apply"] = st.session_state["gh_version"]
            else:
                st.info("GitHubリポジトリに対応ファイルが見つかりませんでした。")
        except Exception:
            st.warning("GitHub連携に失敗しました。ネットワーク接続を確認してください。")

    # --- ファイル処理 ---
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

    # 処理後に再評価してサマリーを出す
    has_work_files_after   = len(st.session_state["uploaded_files"]) > 0
    has_outside_work_after = st.session_state["uploaded_outside_work_file"] is not None

    # 新規にファイルが取り込まれたら rerun → 先頭のサマリーバーが表示される
    if (has_work_files_after or has_outside_work_after) and not (has_work_files or has_outside_work):
        st.rerun()
