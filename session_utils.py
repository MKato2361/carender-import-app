import streamlit as st
from typing import Any, Optional

# Firestore is optional. If firebase_admin isn't configured/initialized yet,
# we gracefully fall back to session_state-only storage.
try:
    from firebase_admin import firestore  # type: ignore
except Exception:  # pragma: no cover
    firestore = None  # type: ignore


_COLLECTION = "user_settings"


def _get_db():
    if firestore is None:
        return None
    try:
        return firestore.client()
    except Exception:
        return None


def get_user_setting(user_id: str, key: str, default: Any = None) -> Any:
    """Get a per-user setting.
    Preference order:
      1) Firestore user_settings/{user_id}
      2) st.session_state cache
      3) default
    """
    if not user_id or not key:
        return default

    cache_key = f"__setting_cache__::{user_id}::{key}"
    if cache_key in st.session_state:
        return st.session_state.get(cache_key, default)

    db = _get_db()
    if db is not None:
        try:
            doc = db.collection(_COLLECTION).document(user_id).get()
            data = doc.to_dict() if doc and doc.exists else None
            if isinstance(data, dict) and key in data:
                st.session_state[cache_key] = data.get(key)
                return data.get(key)
        except Exception:
            # Do not mark as cached on failure; allow next rerun to retry.
            return default

    # fallback
    return st.session_state.get(cache_key, default) if cache_key in st.session_state else default


def set_user_setting(user_id: str, key: str, value: Any) -> None:
    """Set a per-user setting and persist to Firestore if available."""
    if not user_id or not key:
        return

    cache_key = f"__setting_cache__::{user_id}::{key}"
    st.session_state[cache_key] = value

    db = _get_db()
    if db is None:
        return

    try:
        db.collection(_COLLECTION).document(user_id).set({key: value}, merge=True)
    except Exception:
        # Ignore persistence errors; keep in session.
        return
