from __future__ import annotations
"""
firebase_auth.py — 後方互換ラッパー

実体は以下に移行済み:
  core/auth/firebase_client.py  … 認証ロジック・セッション getter
  ui/auth_forms.py              … login_form (UI)
  core/storage/firestore_client.py … トークン保存

既存の import が壊れないようエイリアスを公開する。
新規コードでは直接 core/auth/firebase_client を import すること。
"""
from core.auth.firebase_client import (
    initialize_firebase,                    # noqa: F401
    sign_in   as authenticate_user,         # noqa: F401 — 後方互換名
    sign_up   as create_user_account,       # noqa: F401 — 後方互換名
    get_user_id   as get_firebase_user_id,  # noqa: F401
    get_user_email as get_firebase_user_email, # noqa: F401
    get_id_token  as get_firebase_id_token, # noqa: F401
    is_authenticated as is_user_authenticated, # noqa: F401
)
from core.storage.firestore_client import (
    save_token as save_tokens_to_firestore, # noqa: F401
    load_token as safe_load_tokens_from_firestore, # noqa: F401
)
# UI フォームは ui/auth_forms.py の login_form が正規実装
# firebase_auth_form は後方互換エイリアスとして残す
from ui.auth_forms import login_form as firebase_auth_form  # noqa: F401


def process_tokens_safely(user_id: str) -> None:
    """後方互換: トークンのキー一覧を表示する（デバッグ用）。"""
    import streamlit as st
    from core.storage.firestore_client import load_token
    tokens = load_token(user_id)
    if not tokens:
        st.info("利用可能なトークンがありません。")
        return
    st.caption(f"トークンキー: {list(tokens.keys())}")
