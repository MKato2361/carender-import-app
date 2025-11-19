# utils/user_roles.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional

import firebase_admin
from firebase_admin import firestore


# Firestore コレクション名
USER_COLLECTION = "app_users"

# ロール
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
    """firebase_admin が初期化されていなければ初期化して Firestore クライアントを返す。"""
    if not firebase_admin._apps:
        firebase_admin.initialize_app()
    return firestore.client()


def get_or_create_user(email: str, display_name: Optional[str] = None) -> Dict:
    """
    ログイン時などに呼び出し：
    - app_users/{email} があれば updated_at を更新
    - なければ role=user で新規作成
    戻り値は dict 形式。
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
        update_data: Dict[str, object] = {"updated_at": now_iso}
        if display_name and data.get("display_name") != display_name:
            update_data["display_name"] = display_name
        doc_ref.update(update_data)
        data.update(update_data)
        return data

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
    メールアドレスからロールを取得。
    ドキュメントが存在しない場合は user として作成して user を返す。
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
    ユーザーのロールを admin / user に更新。
    ドキュメントが存在しない場合は新規作成。
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
    app_users コレクションのユーザー一覧を created_at 昇順で取得。
    """
    db = _get_db()
    docs = db.collection(USER_COLLECTION).order_by("created_at").stream()
    users: List[Dict] = []
    for doc in docs:
        users.append(AppUser.from_doc(doc).to_dict())
    return users