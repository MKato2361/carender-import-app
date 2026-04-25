from __future__ import annotations
"""
core/storage/firestore_client.py
Firestore CRUD（st.* 禁止）

session_utils.py / firebase_auth.py から Firestore 操作を抽出。
UI 依存なし・純粋な読み書き関数のみ。
"""
import logging
from typing import Any, Optional

logger = logging.getLogger(__name__)


def _db():
    from firebase_admin import firestore
    return firestore.client()


# ── ユーザー設定 ──

SETTINGS_COLLECTION = "user_settings"

def load_settings(user_id: str) -> dict:
    """Firestore からユーザー設定を取得。存在しなければ空 dict。"""
    try:
        snap = _db().collection(SETTINGS_COLLECTION).document(user_id).get()
        return snap.to_dict() if getattr(snap, "exists", False) else {}
    except Exception as e:
        logger.warning("load_settings 失敗 user=%s: %s", user_id, e)
        return {}

def save_setting(user_id: str, key: str, value: Any) -> None:
    """Firestore にユーザー設定を 1 キーだけ保存（merge）。"""
    from firebase_admin import firestore as _fs
    try:
        _db().collection(SETTINGS_COLLECTION).document(user_id).set(
            {key: value, "updated_at": _fs.SERVER_TIMESTAMP}, merge=True
        )
    except Exception as e:
        logger.warning("save_setting 失敗 user=%s key=%s: %s", user_id, key, e)


# ── Google OAuth トークン ──

TOKEN_COLLECTION = "google_tokens"

def load_token(user_id: str) -> Optional[dict]:
    """Firestore から Google OAuth トークンを取得。存在しなければ None。"""
    try:
        doc = _db().collection(TOKEN_COLLECTION).document(user_id).get()
        return doc.to_dict() if doc.exists else None
    except Exception as e:
        logger.warning("load_token 失敗 user=%s: %s", user_id, e)
        return None

def save_token(user_id: str, token_dict: dict) -> None:
    """Google OAuth トークンを Firestore に保存。"""
    try:
        _db().collection(TOKEN_COLLECTION).document(user_id).set(token_dict)
    except Exception as e:
        logger.warning("save_token 失敗 user=%s: %s", user_id, e)

def delete_token(user_id: str) -> None:
    """Google OAuth トークンを Firestore から削除。"""
    try:
        _db().collection(TOKEN_COLLECTION).document(user_id).delete()
    except Exception as e:
        logger.warning("delete_token 失敗 user=%s: %s", user_id, e)
