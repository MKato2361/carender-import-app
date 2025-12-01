import streamlit as st
from firebase_admin import firestore, _apps
from typing import Optional, Any, Dict

# --- Firestore Client Caching ---

@st.cache_resource
def get_firestore_client():
    """
    Firestoreクライアントをキャッシュして返す。
    firebase_auth.py で initialize_firebase() が実行されている必要がある。
    """
    # Firebase Admin SDKが初期化されているか確認
    if not _apps:
        # アプリケーションのメインファイル（例: firebase_auth.pyのmain関数）で
        # initialize_firebase()が実行されている必要があります。
        # ここで直接initialize_firebase()を呼ぶことはできませんが、
        # 初期化が前提となります。
        return None
    try:
        # firestore.client() をキャッシュし、再実行を防ぐ
        return firestore.client()
    except Exception:
        # 初期化が不十分な場合はNoneを返す
        return None

# --- Default Settings ---

DEFAULT_SETTINGS = {
    # tab2_register.py で使用されるカレンダー選択の設定を追加
    'selected_calendar_name': None,
    'selected_calendar_name_outside': None,
    # 既存の設定
    'description_columns_selected': ["内容", "詳細"],
    'event_name_col_selected': "選択しない",
    'event_name_col_selected_update': "選択しない",
    'add_task_type_to_event_name': False,
    'add_task_type_to_event_name_update': False
}

# --- Core Functions ---

def initialize_cache(user_id: str):
    """ユーザーごとの設定キャッシュを初期化"""
    if 'user_settings_cache' not in st.session_state:
        st.session_state['user_settings_cache'] = {}
    if user_id not in st.session_state['user_settings_cache']:
        # キャッシュには初期データとして空の辞書をセット (Firestoreから読み込むため)
        st.session_state['user_settings_cache'][user_id] = {}


def get_user_setting(user_id: str, key: str) -> Any:
    """
    指定されたユーザーの設定を取得 (キャッシュ > Firestore > デフォルト値 の順に試行)
    """
    initialize_cache(user_id)
    
    # 1. セッションキャッシュから取得を試みる
    if key in st.session_state['user_settings_cache'][user_id]:
        return st.session_state['user_settings_cache'][user_id][key]

    db = get_firestore_client()
    if db:
        # 2. Firestoreから取得を試みる
        try:
            doc_ref = db.collection("user_settings").document(user_id)
            doc = doc_ref.get()
            if doc.exists:
                data = doc.to_dict() or {}
                # 読み込んだデータをキャッシュに保存
                st.session_state['user_settings_cache'][user_id] = data
                # 取得した値、またはFirestoreドキュメント内にキーがない場合はデフォルト値を返す
                return data.get(key, DEFAULT_SETTINGS.get(key))
        except Exception:
            # Firestore読み込みエラーは無視し、デフォルトに進む
            pass
            
    # 3. どちらにもない場合はデフォルト値を返す
    return DEFAULT_SETTINGS.get(key)


def set_user_setting(user_id: str, key: str, value: Any):
    """
    指定されたユーザーの設定を Firestore に保存し、キャッシュを更新する (永続化)
    """
    initialize_cache(user_id)
    
    # 1. セッションキャッシュを更新
    st.session_state['user_settings_cache'][user_id][key] = value
    
    # 2. Firestoreに書き込む
    db = get_firestore_client()
    if db:
        try:
            doc_ref = db.collection("user_settings").document(user_id)
            # merge=True で、他の設定を上書きせずにこのキーだけを更新
            doc_ref.set({key: value}, merge=True)
        except Exception as e:
            st.error(f"🚨 設定の保存（Firestore）に失敗しました: {e}")


def get_all_user_settings(user_id: str) -> Dict[str, Any]:
    """ユーザーの全設定を取得 (Firestoreを優先)"""
    db = get_firestore_client()
    if db:
        try:
            doc_ref = db.collection("user_settings").document(user_id)
            doc = doc_ref.get()
            if doc.exists:
                data = DEFAULT_SETTINGS.copy()
                data.update(doc.to_dict() or {})
                return data
        except Exception:
            pass
            
    initialize_cache(user_id)
    current_settings = DEFAULT_SETTINGS.copy()
    current_settings.update(st.session_state['user_settings_cache'].get(user_id, {}))
    return current_settings


def clear_user_settings(user_id: str):
    """ユーザーのセッション設定と永続設定をクリア"""
    # 1. セッションキャッシュをクリア
    if 'user_settings_cache' in st.session_state and user_id in st.session_state['user_settings_cache']:
        del st.session_state['user_settings_cache'][user_id]
        
    # 2. Firestoreドキュメントをクリア
    db = get_firestore_client()
    if db:
        try:
            db.collection("user_settings").document(user_id).delete()
            st.toast("✅ 永続化された設定をクリアしました", icon="🗑️")
        except Exception as e:
            st.error(f"🚨 永続設定のクリアに失敗しました: {e}")
