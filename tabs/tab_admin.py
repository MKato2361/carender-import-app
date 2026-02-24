# tabs/tab_admin.py
from __future__ import annotations

from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import requests
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
    GITHUB_API_BASE,
    _headers,
)

# ★ 追加: 重複イベントタブを管理者タブ内から呼び出す
from calendar_utils import fetch_all_events
from tabs.tab4_duplicates import render_tab4_duplicates


# ==============================
# GitHub ヘルパー
# ==============================
def list_github_files(path: str = "") -> List[Dict]:
    """
    指定パス配下の GitHub Contents API 一覧を取得。
    ディレクトリとファイルの両方が返るので type を確認して利用。
    """
    clean_path = path.strip().strip("/")
    url_path = clean_path if clean_path else ""

    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{url_path}"
    res = requests.get(url, headers=_headers())
    res.raise_for_status()
    items = res.json()

    # 単一ファイルの場合 dict で返ることがある
    if isinstance(items, dict):
        items = [items]

    items_sorted = sorted(items, key=lambda x: (x.get("type", ""), x.get("path", "")))
    return items_sorted


def upload_file_to_github(target_path: str, content: bytes, message: str) -> Dict:
    """
    GitHub にファイルを新規作成 / 更新する。
    既存の場合は先に GET して sha を取得して PUT に含める。
    """
    import base64

    clean_path = target_path.strip().lstrip("/")
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{clean_path}"

    b64_content = base64.b64encode(content).decode("utf-8")
    payload: Dict[str, object] = {
        "message": message,
        "content": b64_content,
        "branch": "main",
    }

    # 既存ファイルか確認
    get_res = requests.get(url, headers=_headers())
    if get_res.status_code == 200:
        existing = get_res.json()
        if isinstance(existing, dict) and "sha" in existing:
            payload["sha"] = existing["sha"]

    res = requests.put(url, headers=_headers(), json=payload)
    res.raise_for_status()
    return res.json()


def delete_file_from_github(target_path: str, sha: str, message: str) -> Dict:
    """
    GitHub 上のファイルを削除する。
    """
    clean_path = target_path.strip().lstrip("/")
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{clean_path}"

    payload = {
        "message": message,
        "sha": sha,
        "branch": "main",
    }
    res = requests.delete(url, headers=_headers(), json=payload)
    res.raise_for_status()
    return res.json()


def get_dir_commit_dates(base_path: str = "") -> Dict[str, str]:
    """
    指定ディレクトリ直下の各ファイルの最終コミット日を一括取得。
    返り値: { "path/to/file.csv": "2025-01-10", ... }
    """
    clean = base_path.strip().strip("/")
    result: Dict[str, str] = {}

    try:
        items = list_github_files(clean)
        file_paths = [it["path"] for it in items if it.get("type") == "file"]
    except Exception:
        return result

    for path in file_paths:
        try:
            url = (
                f"{GITHUB_API_BASE}/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
                f"/commits?path={path}&per_page=1"
            )
            res = requests.get(url, headers=_headers())
            if res.status_code == 200 and res.json():
                raw = res.json()[0]["commit"]["committer"]["date"]  # ISO8601
                result[path] = raw[:10]  # "YYYY-MM-DD"
            else:
                result[path] = "-"
        except Exception:
            result[path] = "-"

    return result


# ==============================
# 管理者タブ UI 本体
# ==============================
def render_tab_admin(
    current_user_email: str,
    current_user_name: Optional[str] = None,
) -> None:
    """
    管理者専用タブ。
    - current_user_email: Firebase 認証などから取得したユーザーのメールアドレス
    - current_user_name : 表示名(あれば)
    """

    # ログイン情報チェック
    if not current_user_email:
        st.error("ログイン情報が取得できません。再度ログインしてください。")
        return

    # app_users に同期＆ロール取得
    user_doc = get_or_create_user(current_user_email, current_user_name)
    role = user_doc.get("role") or get_user_role(current_user_email)

    if role != ROLE_ADMIN:
        st.error("このページは管理者専用です。権限がありません。")
        return

    st.title("🔧 管理者メニュー")

    tab_users, tab_files, tab_dup = st.tabs(
        ["👥 ユーザー管理", "📂 GitHubファイル管理", "🔁 重複イベントの検出・削除"]
    )

    # --------------------------
    # 👥 ユーザー管理
    # --------------------------
    with tab_users:
        st.subheader("ユーザー一覧 / ロール編集")

        users = list_users()
        if not users:
            st.info("ユーザー情報がまだありません。ユーザーがログインすると自動登録されます。")
        else:
            df = pd.DataFrame(users)

            cols_order = [
                c
                for c in (
                    "email",
                    "display_name",
                    "role",
                    "created_at",
                    "updated_at",
                )
                if c in df.columns
            ]
            df = df[cols_order]

            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                hide_index=True,
                column_config={
                    "email": st.column_config.TextColumn("メールアドレス", disabled=True),
                    "display_name": st.column_config.TextColumn("表示名", disabled=True),
                    "role": st.column_config.SelectboxColumn(
                        "ロール",
                        options=[ROLE_USER, ROLE_ADMIN],
                        required=True,
                    ),
                },
                key="admin_users_editor",
            )

            if st.button("変更を保存", type="primary", key="admin_users_save"):
                for _, row in edited_df.iterrows():
                    email = str(row.get("email") or "").strip().lower()
                    role_val = str(row.get("role") or ROLE_USER).strip().lower()
                    if not email:
                        continue
                    set_user_role(email, role_val)

                st.success("ユーザー情報を保存しました。必要に応じてページを再読み込みしてください。")

        st.markdown("---")
        st.subheader("個別ロール変更（メールアドレス指定）")

        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            target_email = st.text_input(
                "対象メールアドレス",
                key="single_role_email",
                placeholder="user@example.com",
            )
        with col2:
            target_role = st.selectbox(
                "付与するロール",
                [ROLE_USER, ROLE_ADMIN],
                key="single_role_role",
            )
        with col3:
            if st.button("更新", key="single_role_update"):
                if target_email:
                    set_user_role(target_email, target_role)
                    st.success(f"{target_email} のロールを {target_role} に更新しました。")
                else:
                    st.warning("メールアドレスを入力してください。")

    # --------------------------
    # 📂 GitHub ファイル管理
    # --------------------------
    with tab_files:
        st.subheader("📂 GitHub ファイル管理")
        st.caption(
            f"対象リポジトリ: `{GITHUB_OWNER}/{GITHUB_REPO}`  （PAT: secrets の GITHUB_PAT を利用）"
        )

        # ── 対象ディレクトリ ──────────────────────────────────
        base_path = st.text_input(
            "対象ディレクトリ（例: state / templates / 空欄でリポジトリルート）",
            value=st.session_state.get("admin_github_base_path", ""),
            key="admin_github_base_input",
        )
        st.session_state["admin_github_base_path"] = base_path

        st.markdown("---")

        # ── 現在のファイル状況 ────────────────────────────────
        st.markdown("#### 現在のファイル状況")

        col_reload, _ = st.columns([1, 5])
        with col_reload:
            if st.button("🔄 再取得", key="admin_github_reload"):
                st.session_state.pop("admin_github_last_list", None)
                st.session_state.pop("admin_github_commit_dates", None)

        cache_key = "admin_github_last_list"
        if cache_key not in st.session_state:
            try:
                items = list_github_files(base_path)
                st.session_state[cache_key] = items
            except Exception as e:
                st.error(f"ファイル一覧取得中にエラーが発生しました: {e}")
                items = []
        else:
            items = st.session_state[cache_key]

        file_items = [it for it in items if it.get("type") == "file"]

        if file_items:
            # 更新日を一括取得
            date_cache_key = "admin_github_commit_dates"
            if date_cache_key not in st.session_state:
                with st.spinner("更新日時を取得中..."):
                    st.session_state[date_cache_key] = get_dir_commit_dates(base_path)
            commit_dates: Dict[str, str] = st.session_state[date_cache_key]

            # ヘッダー行
            hcol1, hcol2, hcol3 = st.columns([4, 2, 3])
            with hcol1:
                st.caption("ファイル名")
            with hcol2:
                st.caption("SHA")
            with hcol3:
                st.caption("最終更新日")

            for item in file_items:
                path     = item.get("path", "")
                sha      = item.get("sha", "")
                html_url = item.get("html_url", "")
                updated  = commit_dates.get(path, "-")

                c1, c2, c3 = st.columns([4, 2, 3])
                with c1:
                    if html_url:
                        st.markdown(f"[`{path}`]({html_url})")
                    else:
                        st.write(f"`{path}`")
                with c2:
                    st.write(f"`{sha[:7]}`" if sha else "-")
                with c3:
                    st.write(updated)
        else:
            st.info("ファイルが見つかりませんでした。")

        st.markdown("---")

        # ── アップロード ──────────────────────────────────────
        st.markdown("#### ファイルをアップロード")
        st.caption("既存ファイルは自動で上書き、新規ファイルは新規作成されます。")

        col_up1, col_up2 = st.columns([3, 2])
        with col_up1:
            uploaded_files = st.file_uploader(
                "ファイルを選択（複数可）",
                key="admin_github_uploader",
                accept_multiple_files=True,
            )
        with col_up2:
            commit_message = st.text_input(
                "コミットメッセージ",
                value=f"Upload from admin UI ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
                key="admin_github_commit_msg",
            )

        # アップロード予定ファイルのプレビュー
        if uploaded_files:
            existing_names = {it.get("name") for it in file_items}
            for f in uploaded_files:
                if f.name in existing_names:
                    st.warning(f"⚠️ 上書き: `{f.name}` （既存ファイルを更新します）")
                else:
                    st.success(f"✅ 新規: `{f.name}`")

        if st.button("▶ アップロード実行", type="primary", key="admin_github_do_upload"):
            if not uploaded_files:
                st.warning("ファイルを選択してください。")
            else:
                clean_base    = base_path.strip().strip("/")
                success_count = 0
                error_count   = 0

                for f in uploaded_files:
                    target_path = f"{clean_base}/{f.name}" if clean_base else f.name
                    try:
                        upload_file_to_github(
                            target_path=target_path,
                            content=f.getvalue(),
                            message=commit_message,
                        )
                        success_count += 1
                        st.success(f"完了: `{target_path}`")
                    except Exception as e:
                        error_count += 1
                        st.error(f"エラー: `{f.name}` ({e})")

                # キャッシュ削除して一覧を更新
                st.session_state.pop(cache_key, None)
                st.session_state.pop("admin_github_commit_dates", None)

                if error_count == 0:
                    st.info(f"{success_count} 件のアップロードが完了しました。")
                else:
                    st.warning(f"{success_count} 件成功、{error_count} 件でエラーが発生しました。")

                st.rerun()

        st.markdown("---")

        # ── 削除操作（折りたたみ）────────────────────────────
        with st.expander("⚙️ 削除操作（展開して表示）"):
            if not file_items:
                st.info("削除対象のファイルがありません。")
            else:
                delete_all = st.checkbox(
                    "⚠️ このディレクトリ内の全ファイルを削除する",
                    key="admin_github_delete_all",
                )
                st.caption("または行ごとにチェックして個別削除できます。")

                for idx, item in enumerate(file_items):
                    path   = item.get("path", "")
                    sha    = item.get("sha", "")
                    cb_key = f"admin_github_ck_{idx}_{sha}"

                    col_ck, col_name = st.columns([1, 6])
                    with col_ck:
                        st.checkbox("", key=cb_key)
                    with col_name:
                        st.write(f"`{path}`")

                if st.button("🗑️ 選択したファイルを削除", type="primary", key="admin_github_delete_selected"):
                    if delete_all:
                        targets = file_items
                    else:
                        targets = [
                            item for idx, item in enumerate(file_items)
                            if st.session_state.get(f"admin_github_ck_{idx}_{item.get('sha')}")
                        ]

                    if not targets:
                        st.warning("削除対象のファイルが選択されていません。")
                    else:
                        error_count = 0
                        for item in targets:
                            path = item.get("path", "")
                            sha  = item.get("sha", "")
                            if not path or not sha:
                                continue
                            try:
                                delete_file_from_github(
                                    target_path=path,
                                    sha=sha,
                                    message=f"Delete from admin UI ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
                                )
                            except Exception as e:
                                error_count += 1
                                st.error(f"削除エラー: `{path}` ({e})")

                        if error_count == 0:
                            st.success(f"{len(targets)} 件を削除しました。")
                        else:
                            st.warning(f"{len(targets)} 件中 {error_count} 件でエラーが発生しました。")

                        # キャッシュ・チェック状態をリセット
                        st.session_state.pop(cache_key, None)
                        st.session_state.pop("admin_github_commit_dates", None)
                        for idx, item in enumerate(file_items):
                            cb_key = f"admin_github_ck_{idx}_{item.get('sha')}"
                            st.session_state.pop(cb_key, None)
                        st.session_state["admin_github_delete_all"] = False

                        st.rerun()

    # --------------------------
    # 🔁 重複イベントの検出・削除（元タブ4）
    # --------------------------
    with tab_dup:
        st.subheader("🔁 重複イベントの検出・削除（管理者専用）")

        service = st.session_state.get("calendar_service")
        editable_calendar_options = st.session_state.get("editable_calendar_options")

        if not service or not editable_calendar_options:
            st.warning("カレンダーサービスが初期化されていません。トップ画面でGoogle認証を完了してください。")
            return

        render_tab4_duplicates(
            service,
            editable_calendar_options,
            fetch_all_events,
        )
