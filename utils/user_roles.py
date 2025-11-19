# utils/user_roles.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Optional

import firebase_admin
from firebase_admin import firestore


# Firestore コレクション名
USER_COLLECTION = "app_users"

# ロール定義
ROLE_USER = "user"
ROLE_ADMIN = "admin"


@dataclass
class AppUser:
    email: str
    display_name: Optional[str]
    role: str
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    @classmethod
    def from_doc(cls, doc) -> "AppUser":
        data = doc.to_dict() or {}
        return cls(
            email=data.get("email", doc.id),
            display_name=data.get("display_name"),
            role=data.get("role", ROLE_USER),
            created_at=data.get("created_at"),
            updated_at=data.get("updated_at"),
        )

    def to_dict(self) -> Dict:
        return {
            "email": self.email,
            "display_name": self.display_name,
            "role": self.role,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


def _get_db():
    # 既に初期化済みなら再初期化しない
    if not firebase_admin._apps:
        firebase_admin.initialize_app()
    return firestore.client()


def get_or_create_user(email: str, display_name: Optional[str] = None) -> Dict:
    """
    ログイン時に呼び出して、ユーザー情報を Firestore に作成/更新する。
    戻り値は dict（AppUser.to_dict）で返します。
    """
    if not email:
        raise ValueError("email is required")

    email = email.strip().lower()
    db = _get_db()
    doc_ref = db.collection(USER_COLLECTION).document(email)

    now_iso = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    snapshot = doc_ref.get()

    if snapshot.exists:
        data = snapshot.to_dict() or {}
        # 表示名が変わっていたら更新
        update_data = {}
        if display_name and data.get("display_name") != display_name:
            update_data["display_name"] = display_name
        update_data["updated_at"] = now_iso
        if update_data:
            doc_ref.update(update_data)
            data.update(update_data)
        return data

    # 新規作成（デフォルトは一般ユーザー）
    user = AppUser(
        email=email,
        display_name=display_name,
        role=ROLE_USER,
        created_at=now_iso,
        updated_at=now_iso,
    )
    doc_ref.set(user.to_dict())
    return user.to_dict()


def get_user_role(email: str) -> str:
    """
    メールアドレスからロールを取得。存在しなければ一般ユーザーとして作成。
    """
    if not email:
        return ROLE_USER
    email = email.strip().lower()
    db = _get_db()
    doc_ref = db.collection(USER_COLLECTION).document(email)
    snapshot = doc_ref.get()
    if not snapshot.exists:
        data = get_or_create_user(email)
        return data.get("role", ROLE_USER)
    data = snapshot.to_dict() or {}
    return data.get("role", ROLE_USER)


def set_user_role(email: str, role: str) -> None:
    """
    ユーザーのロールを更新（admin/user など）。
    """
    if not email:
        raise ValueError("email is required")
    role = role.strip().lower()
    if role not in (ROLE_USER, ROLE_ADMIN):
        raise ValueError(f"invalid role: {role}")

    email = email.strip().lower()
    db = _get_db()
    doc_ref = db.collection(USER_COLLECTION).document(email)

    now_iso = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    snapshot = doc_ref.get()
    if snapshot.exists:
        doc_ref.update({"role": role, "updated_at": now_iso})
    else:
        # 存在しない場合は新規作成
        user = AppUser(
            email=email,
            display_name=None,
            role=role,
            created_at=now_iso,
            updated_at=now_iso,
        )
        doc_ref.set(user.to_dict())


def list_users() -> List[Dict]:
    """
    Firestore 上の app_users コレクションをすべて取得。
    """
    db = _get_db()
    docs = db.collection(USER_COLLECTION).order_by("created_at").stream()
    users: List[Dict] = []
    for doc in docs:
        users.append(AppUser.from_doc(doc).to_dict())
    return users
