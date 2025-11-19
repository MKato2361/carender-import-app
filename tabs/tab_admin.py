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


# ==============================
# GitHub ヘルパー
# ==============================
def list_github_files(path: str = "") -> List[Dict]:
    """
    指定パス配下の GitHub Contents API 一覧を取得。
    ディレクトリとファイルの両方が返るので type を確認して利用。
    """
    clean_path = path.strip().strip("/")
    if clean_path:
        url_path = clean_path
    else:
        url_path = ""

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
        # ブランチ指定：必要に応じて変更
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
    - current_user_name : 表示名（あれば）
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

    tab_users, tab_files = st.tabs(["👥 ユーザー管理", "📂 GitHubファイル管理"])

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

            # 表示順
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
                    "display_name": st.column_config.TextColumn(
                        "表示名",
                        disabled=True,
                    ),
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
        st.subheader("GitHub ファイルアップロード / 削除")

        st.caption(
            f"対象リポジトリ: `{GITHUB_OWNER}/{GITHUB_REPO}`  （PAT: secrets の GITHUB_PAT を利用）"
        )

        # ベースディレクトリ
        default_base = st.session_state.get("admin_github_base_path", "")
        base_path = st.text_input(
            "対象ディレクトリ（例: state / templates / 空欄でリポジトリルート）",
            value=default_base,
            key="admin_github_base_input",
        )
        st.session_state["admin_github_base_path"] = base_path

        col_up1, col_up2 = st.columns([3, 1])
        with col_up1:
            uploaded_file = st.file_uploader(
                "アップロードするファイル",
                key="admin_github_uploader",
            )
        with col_up2:
            commit_message = st.text_input(
                "コミットメッセージ",
                value=f"Upload from admin UI ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
                key="admin_github_commit_msg",
            )

        if st.button("アップロード実行", type="primary", key="admin_github_do_upload"):
            if not uploaded_file:
                st.warning("ファイルを選択してください。")
            else:
                clean_base = base_path.strip().strip("/")
                if clean_base:
                    target_path = f"{clean_base}/{uploaded_file.name}"
                else:
                    target_path = uploaded_file.name

                try:
                    res = upload_file_to_github(
                        target_path=target_path,
                        content=uploaded_file.getvalue(),
                        message=commit_message,
                    )
                    st.success(f"アップロード完了: `{target_path}`")
                    with st.expander("GitHub API レスポンス", expanded=False):
                        st.json(res)
                    # 再取得のためキャッシュ削除
                    st.session_state.pop("admin_github_last_list", None)
                except Exception as e:
                    st.error(f"アップロード中にエラーが発生しました: {e}")

        st.markdown("---")
        st.subheader("ディレクトリ内のファイル一覧 / 削除")

        if st.button("一覧を再取得", key="admin_github_reload"):
            st.session_state.pop("admin_github_last_list", None)

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

        if not items:
            st.info("ファイルが見つかりませんでした。パスや権限を確認してください。")
            return

        st.caption("※ type=file のみ削除対象です。削除すると GitHub 上から即時消えるため慎重に操作してください。")

        for item in items:
            if item.get("type") != "file":
                continue

            path = item.get("path")
            sha = item.get("sha")
            size = item.get("size")
            html_url = item.get("html_url")

            col_f1, col_f2, col_f3, col_f4 = st.columns([4, 2, 2, 2])
            with col_f1:
                if html_url:
                    st.markdown(f"[`{path}`]({html_url})")
                else:
                    st.write(f"`{path}`")
            with col_f2:
                st.write(f"SHA: `{sha[:7]}`" if sha else "-")
            with col_f3:
                st.write(f"{size} bytes" if size is not None else "")
            with col_f4:
                btn_key = f"del_{sha}"
                if st.button("削除", key=btn_key):
                    try:
                        delete_file_from_github(
                            target_path=path,
                            sha=sha,
                            message=f"Delete from admin UI ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
                        )
                        st.success(f"削除完了: `{path}`")
                        # 再取得のためキャッシュ削除
                        st.session_state.pop(cache_key, None)
                        st.experimental_rerun()
                    except Exception as e:
                        st.error(f"削除中にエラーが発生しました: {e}")