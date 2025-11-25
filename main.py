from __future__ import annotations
import re
import unicodedata
from datetime import datetime, date, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from firebase_admin import firestore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from streamlit.runtime.uploaded_file_manager import UploadedFile

# ---- Utils & Helpers ----
from utils.helpers import safe_get, to_utc_range, default_fetch_window_years
from utils.parsers import extract_worksheet_id_from_text
from utils.user_roles import get_or_create_user, get_user_role, ROLE_ADMIN
from github_loader import walk_repo_tree, load_file_bytes_from_github, is_supported_file
from github_loader import _headers, GITHUB_OWNER, GITHUB_REPO

# ---- Tab Modules ----
from tabs.tab1_upload import render_tab1_upload
from tabs.tab2_register import render_tab2_register
from tabs.tab3_delete import render_tab3_delete
from tabs.tab4_duplicates import render_tab4_duplicates
from tabs.tab5_export import render_tab5_export
from tabs.tab_admin import render_tab_admin
from tabs.tab6_property_master import render_tab6_property_master
from tabs.tab7_inspection_todo import render_tab7_inspection_todo
from tabs.tab8_notice_fax import render_tab8_notice_fax

# ---- Auth & Logic ----
from calendar_utils import (
    authenticate_google,
    fetch_all_events,
    build_tasks_service,
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from session_utils import (
    initialize_session_state,
    get_user_setting,
    set_user_setting,
    get_all_user_settings,
    clear_user_settings,
)

# ==================================================
# 0) ページ設定 & デザイン (CSS)
# ==================================================
st.set_page_config(
    page_title="Googleカレンダー管理ツール",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# モダンなカスタムCSS
st.markdown(
    """
    <style>
    /* 全体のフォントと背景 */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* ヘッダー装飾 */
    .main-header {
        background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }
    .main-header h1 {
        color: white !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 600;
    }
    .main-header p {
        color: #e0e0e0;
        margin: 5px 0 0 0;
    }

    /* タブのスタイル調整 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #ffffff;
        border-radius: 8px 8px 0 0;
        border: 1px solid #ddd;
        border-bottom: none;
        padding: 0 20px;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #4b6cb7;
        background-color: #f0f4ff;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        border-top: 3px solid #4b6cb7 !important;
        color: #4b6cb7 !important;
        font-weight: bold;
    }

    /* ボタンのスタイル */
    .stButton > button {
        border-radius: 6px;
        font-weight: 500;
    }
    
    /* サイドバー */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #f0f0f0;
    }
    
    /* フッター隠し */
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ヘッダー表示
st.markdown(
    """
    <div class="main-header">
        <h1>📅 Googleカレンダー 一括管理システム</h1>
        <p>イベント登録・削除・ToDo連携・帳票出力を統合管理</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ==================================================
# 1) クラス定義 & 共通ユーティリティ
# ==================================================
JST = timezone(timedelta(hours=9))

# バイトデータ変換用クラス
class GitHubUploadedFile:
    def __init__(self, file_bytes: bytes, name: str, type: str = None):
        self._file_bytes = file_bytes
        self.name = name
        self.type = type or "application/octet-stream"

    def read(self):
        return self._file_bytes

    def getvalue(self):
        return self._file_bytes

def convert_bytes_to_uploadedfile(file_bytes: bytes, filename: str, mime_type: str = None):
    return GitHubUploadedFile(
        file_bytes=file_bytes,
        name=filename,
        type=mime_type or "application/octet-stream",
    )

# サービス初期化ヘルパー
def build_calendar_service(creds):
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        # 編集権限があるカレンダーのみ抽出 (reader以外)
        editable = {
            cal["summary"]: cal["id"] 
            for cal in calendar_list.get("items", []) 
            if cal.get("accessRole") in ["owner", "writer"]
        }
        return service, editable
    except Exception as e:
        st.error(f"カレンダーサービスの初期化エラー: {e}")
    return None, None

def build_tasks_service_safe(creds):
    try:
        ts = build_tasks_service(creds)
        if not ts: return None, None
        tl = ts.tasklists().list().execute()
        items = tl.get("items", [])
        default_id = next((i["id"] for i in items if i.get("title") == "My Tasks"), None)
        if not default_id and items:
            default_id = items[0]["id"]
        return ts, default_id
    except Exception as e:
        st.warning(f"ToDoリスト初期化警告: {e}")
    return None, None

def ensure_services(creds):
    if "calendar_service" not in st.session_state or not st.session_state["calendar_service"]:
        s, e = build_calendar_service(creds)
        if not s:
            st.error("Googleカレンダーに接続できません。認証を再試行してください。")
            st.stop()
        st.session_state["calendar_service"] = s
        st.session_state["editable_calendar_options"] = e
        
    if "tasks_service" not in st.session_state:
        ts, tid = build_tasks_service_safe(creds)
        st.session_state["tasks_service"] = ts
        st.session_state["default_task_list_id"] = tid

    return st.session_state["calendar_service"], st.session_state["editable_calendar_options"]

# ==================================================
# 2) 認証プロセス (Firebase -> Google)
# ==================================================
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    # 未ログイン時の中央配置フォーム
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            st.subheader("ログイン")
            firebase_auth_form()
    st.stop()

# ユーザー設定ロード
def load_user_settings(uid):
    initialize_session_state(uid)
    doc = db.collection("user_settings").document(uid).get()
    if doc.exists:
        for k, v in doc.to_dict().items():
            set_user_setting(uid, k, v)

def save_setting(uid, key, val):
    if not uid: return
    db.collection("user_settings").document(uid).set({key: val}, merge=True)

load_user_settings(user_id)

# 権限チェック
user_doc = get_or_create_user(user_id, None)
current_role = user_doc.get("role") or get_user_role(user_id)
is_admin = (current_role == ROLE_ADMIN)

# Google認証
creds = authenticate_google()
if not creds:
    st.warning("左上の認証ボタン、または認証フローを完了してください。")
    st.stop()

service, editable_calendar_options = ensure_services(creds)
tasks_service = st.session_state.get("tasks_service")
default_task_list_id = st.session_state.get("default_task_list_id")

# Sheets Service
try:
    sheets_service = build("sheets", "v4", credentials=creds)
except:
    sheets_service = None

# ==================================================
# 3) サイドバー (設定・情報)
# ==================================================
with st.sidebar:
    st.title("⚙️ 設定メニュー")
    
    # --- ユーザー情報 ---
    with st.expander("👤 アカウント情報", expanded=True):
        st.caption(f"ID: {user_id}")
        st.caption(f"Role: {current_role}")
        if st.button("ログアウト", type="primary", use_container_width=True):
            clear_user_settings(user_id)
            st.session_state.clear()
            st.rerun()

    # --- デフォルト設定 ---
    with st.expander("📅 デフォルト設定", expanded=False):
        if editable_calendar_options:
            cal_names = list(editable_calendar_options.keys())
            saved_cal = get_user_setting(user_id, "selected_calendar_name")
            def_idx = cal_names.index(saved_cal) if saved_cal in cal_names else 0
            
            sel_cal = st.selectbox("標準カレンダー", cal_names, index=def_idx)
            
            is_shared = st.toggle("タブ間でカレンダー選択を共有", value=st.session_state.get("share_calendar_selection_across_tabs", True))
            
            st.divider()
            def_priv = st.checkbox("標準で「非公開」", value=get_user_setting(user_id, "default_private_event", True))
            def_all = st.checkbox("標準で「終日」", value=get_user_setting(user_id, "default_allday_event", False))
            def_todo = st.checkbox("標準で「ToDo作成」", value=get_user_setting(user_id, "default_create_todo", False))
            
            if st.button("設定を保存", use_container_width=True):
                save_setting(user_id, "selected_calendar_name", sel_cal)
                save_setting(user_id, "share_calendar_selection_across_tabs", is_shared)
                save_setting(user_id, "default_private_event", def_priv)
                save_setting(user_id, "default_allday_event", def_all)
                save_setting(user_id, "default_create_todo", def_todo)
                
                # Session State反映
                st.session_state["selected_calendar_name"] = sel_cal
                st.session_state["share_calendar_selection_across_tabs"] = is_shared
                st.toast("設定を保存しました！", icon="✅")

    # --- ステータス ---
    st.markdown("---")
    st.subheader("システム状態")
    
    st.markdown(
        f"""
        <div style='background-color: white; padding: 10px; border-radius: 5px; font-size: 0.9em; border: 1px solid #eee;'>
            <div>🔥 <b>Firebase:</b> <span style='color:green'>Online</span></div>
            <div>📅 <b>Calendar:</b> <span style='color:{'green' if service else 'red'}'>{'Connected' if service else 'Error'}</span></div>
            <div>✅ <b>Tasks:</b> <span style='color:{'green' if tasks_service else 'orange'}'>{'Active' if tasks_service else 'Inactive'}</span></div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.caption(f"Uploaded Files: {len(st.session_state.get('uploaded_files', []))}")


# ==================================================
# 4) メインコンテンツ (Tabs)
# ==================================================
if "uploaded_files" not in st.session_state:
    st.session_state["uploaded_files"] = []

# 管理者かどうかでタブ構成を変更
tabs_list = [
    "📂 ファイル取込",
    "📝 登録・編集・削除",
    "📤 データ出力",
    "🏢 物件マスタ",
]
if is_admin:
    tabs_list.append("🛠 管理者")

# タブ描画
tabs = st.tabs(tabs_list)

# --- Tab 1: Upload ---
with tabs[0]:
    with st.container(border=True):
        st.markdown("### ファイルのアップロード")
        st.info("Excelファイル、またはCSVファイルをアップロードしてください。")
        render_tab1_upload()

# --- Tab 2: Operations (Register/Delete/ToDo/Fax) ---
with tabs[1]:
    # サブタブのデザイン
    st.markdown("#### イベント操作メニュー")
    sub_tabs = st.tabs(["📥 登録", "🗑 削除", "✅ 点検ToDo", "📄 貼り紙・FAX"])
    
    with sub_tabs[0]:
        with st.container(border=True):
            render_tab2_register(user_id, editable_calendar_options, service)
            
    with sub_tabs[1]:
        with st.container(border=True):
            render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id)
            
    with sub_tabs[2]:
        with st.container(border=True):
            render_tab7_inspection_todo(
                service=service,
                editable_calendar_options=editable_calendar_options,
                tasks_service=tasks_service,
                default_task_list_id=default_task_list_id,
                sheets_service=sheets_service,
                current_user_email=user_id,
            )
            
    with sub_tabs[3]:
        with st.container(border=True):
            render_tab8_notice_fax(
                service=service,
                editable_calendar_options=editable_calendar_options,
                sheets_service=sheets_service,
                current_user_email=user_id,
            )

# --- Tab 3: Export ---
with tabs[2]:
    with st.container(border=True):
        st.markdown("### カレンダーデータの出力")
        render_tab5_export(editable_calendar_options, service, fetch_all_events)

# --- Tab 4: Property Master ---
with tabs[3]:
    with st.container(border=True):
        st.markdown("### 物件マスタ管理")
        render_tab6_property_master(
            sheets_service=sheets_service,
            default_spreadsheet_id=st.secrets.get("PROPERTY_MASTER_SHEET_ID", ""),
            basic_sheet_title="物件基本情報",
            master_sheet_title="物件マスタ",
            current_user_email=user_id,
        )

# --- Tab 5: Admin (Optional) ---
if is_admin:
    with tabs[4]:
        with st.container(border=True):
            st.markdown("### 管理者機能")
            render_tab_admin(
                current_user_email=user_id,
                current_user_name=None,
            )
