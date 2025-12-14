import copy
import streamlit as st

# Firestore 永続化（Firebase Admin が初期化されていない環境でも落ちないようにする）
try:
    from firebase_admin import firestore
except Exception:  # pragma: no cover
    firestore = None

# デフォルト設定
DEFAULT_SETTINGS = {
    "description_columns_selected": ["内容", "詳細"],
    "event_name_col_selected": "選択しない",
    "event_name_col_selected_update": "選択しない",
    "add_task_type_to_event_name": False,
    "add_task_type_to_event_name_update": False,
    # 今後増えてもここに追加
}

# Firestore コレクション名
SETTINGS_COLLECTION = "user_settings"


def _get_db():
    if firestore is None:
        return None
    try:
        return firestore.client()
    except Exception:
        return None


def load_user_settings_from_firestore(user_id: str) -> dict:
    """Firestore からユーザー設定を取得（存在しなければ空dict）。"""
    db = _get_db()
    if not db or not user_id:
        return {}
    try:
        snap = db.collection(SETTINGS_COLLECTION).document(user_id).get()
        return snap.to_dict() if getattr(snap, "exists", False) else {}
    except Exception:
        return {}


def save_user_setting_to_firestore(user_id: str, key: str, value):
    """Firestore に 1キーだけ保存（merge）。"""
    db = _get_db()
    if not db or not user_id or not key:
        return
    try:
        doc_ref = db.collection(SETTINGS_COLLECTION).document(user_id)
        doc_ref.set({key: value, "updated_at": firestore.SERVER_TIMESTAMP}, merge=True)
    except Exception:
        # 永続化失敗してもセッション内は動作させる
        pass


def initialize_session_state(user_id: str):
    """ユーザーごとのセッション状態を初期化（初回のみ Firestore からもロード）。"""
    if "user_settings" not in st.session_state:
        st.session_state["user_settings"] = {}
    if "_user_settings_loaded" not in st.session_state:
        st.session_state["_user_settings_loaded"] = set()

    if not user_id:
        return

    if user_id not in st.session_state["user_settings"]:
        st.session_state["user_settings"][user_id] = copy.deepcopy(DEFAULT_SETTINGS)

    # 1セッション中に Firestore から何度もロードしない
    if user_id not in st.session_state["_user_settings_loaded"]:
        saved = load_user_settings_from_firestore(user_id)
        if isinstance(saved, dict) and saved:
            # DEFAULT に上書き（未知キーも保持）
            st.session_state["user_settings"][user_id].update(saved)
        st.session_state["_user_settings_loaded"].add(user_id)


def set_user_setting(user_id: str, key: str, value, persist: bool = True):
    """指定されたユーザーの設定を保存（デフォルトでFirestoreにも永続化）。"""
    initialize_session_state(user_id)
    if not user_id:
        return
    st.session_state["user_settings"][user_id][key] = value
    if persist:
        save_user_setting_to_firestore(user_id, key, value)


def get_user_setting(user_id: str, key: str):
    """指定されたユーザーの設定を取得（なければ DEFAULT を返す）。"""
    initialize_session_state(user_id)
    if not user_id:
        return DEFAULT_SETTINGS.get(key)
    return st.session_state["user_settings"][user_id].get(key, DEFAULT_SETTINGS.get(key))


def get_all_user_settings(user_id: str):
    """ユーザーの全設定を取得"""
    initialize_session_state(user_id)
    if not user_id:
        return copy.deepcopy(DEFAULT_SETTINGS)
    return st.session_state["user_settings"][user_id]


def clear_user_settings(user_id: str):
    """ユーザーのセッション設定をクリア（ログアウト時など）。Firestoreは削除しない。"""
    if "user_settings" in st.session_state and user_id in st.session_state["user_settings"]:
        del st.session_state["user_settings"][user_id]
    if "_user_settings_loaded" in st.session_state and user_id in st.session_state["_user_settings_loaded"]:
        try:
            st.session_state["_user_settings_loaded"].remove(user_id)
        except Exception:
            pass
