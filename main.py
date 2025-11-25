from __future__ import annotations
import re
import unicodedata
from datetime import datetime, date, timedelta, timezone
from io import BytesIO
from pathlib import Path
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

from sidebar import render_sidebar  # ★ 追加
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
# 0) ページ設定 & スタイル (カラーは変更せず維持)
# ==================================================
st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")

def load_custom_css() -> None:
    try:
        with open("custom_sidebar.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass

load_custom_css()

# 元のCSSを維持（ダークモード/ライトモード対応）
st.markdown(
    """
<style>
@media (prefers-color-scheme: light) {
    .header-bar { background-color: rgba(249, 249, 249, 0.95); color: #333; border-bottom: 1px solid #ccc; }
}
@media (prefers-color-scheme: dark) {
    .header-bar { background-color: rgba(30, 30, 30, 0.9); color: #eee; border-bottom: 1px solid #444; }
}
.header-bar { position: sticky; top: 0; width: 100%; text-align: center; font-weight: 500;
    font-size: 14px; padding: 8px 0; z-index: 20; backdrop-filter: blur(6px); }
div[data-testid="stTabs"] { position: sticky; top: 42px; z-index: 15; background-color: inherit;
    border-bottom: 1px solid rgba(128, 128, 128, 0.3); padding-top: 4px; padding-bottom: 4px;
    backdrop-filter: blur(6px); }
.block-container, section[data-testid="stMainBlockContainer"], main {
    padding-top: 0 !important; padding-bottom: 0 !important; margin-bottom: 0 !important;
    height: auto !important; min-height: 100vh !重要; overflow: visible !重要; }
footer, div[data-testid="stBottomBlockContainer"] { display: none !重要; height: 0 !重要; margin: 0 !重要; padding: 0 !重要; }
html, body, #root { height: auto !重要; min-height: 100% !重要; margin: 0 !重要; padding: 0 !重要;
    overflow-x: hidden !重要; overflow-y: auto !重要; overscroll-behavior: none !重要; -webkit-overflow-scrolling: touch !重要; }
</style>
<div class="header-bar">📅 Googleカレンダー一括イベント登録・削除</div>
""",
    unsafe_allow_html=True,
)

# ==================================================
# 1) 共通ユーティリティ & クラス定義
# ==================================================
JST = timezone(timedelta(hours=9))

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
    """GitHub等から取得したバイトデータをStreamlitのUploadedFile互換に変換"""
    return GitHubUploadedFile(
        file_bytes=file_bytes,
        name=filename,
        type=mime_type or "application/octet-stream",
    )

def build_calendar_service(creds):
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        editable = {cal["summary"]: cal["id"] for cal in calendar_list.get("items", []) if cal.get("accessRole") != "reader"}
        return service, editable
    except HttpError as e:
        st.error(f"カレンダーサービスの初期化に失敗しました (HTTP): {e}")
    except Exception as e:
        st.error(f"カレンダーサービスの初期化に失敗しました: {e}")
    return None, None

def build_tasks_service_safe(creds):
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
        task_lists = tasks_service.tasklists().list().execute()
        default_id = None
        for item in task_lists.get("items", []):
            if item.get("title") == "My Tasks":
                default_id = item["id"]
                break
        if not default_id and task_lists.get("items"):
            default_id = task_lists["items"][0]["id"]
        return tasks_service, default_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTP): {e}")
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
    return None, None

def ensure_services(creds):
    if "calendar_service" not in st.session_state or not st.session_state["calendar_service"]:
        service, editable = build_calendar_service(creds)
        if not service:
            st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
            st.stop()
        st.session_state["calendar_service"] = service
        st.session_state["editable_calendar_options"] = editable
    if "tasks_service" not in st.session_state or not st.session_state.get("tasks_service"):
        tasks_service, default_task_list_id = build_tasks_service_safe(creds)
        st.session_state["tasks_service"] = tasks_service
        st.session_state["default_task_list_id"] = default_task_list_id
        if not tasks_service:
            st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
    return st.session_state["calendar_service"], st.session_state["editable_calendar_options"]

# ==================================================
# 2) Firebase 認証
# ==================================================
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()
if not user_id:
    # ログイン画面を少し中央寄せで見やすく
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("利用を開始するにはログインしてください")
        firebase_auth_form()
    st.stop()

def load_user_settings_from_firestore(user_id: str) -> None:
    if not user_id:
        return
    initialize_session_state(user_id)
    doc = db.collection("user_settings").document(user_id).get()
    if doc.exists:
        for key, value in doc.to_dict().items():
            set_user_setting(user_id, key, value)

def save_user_setting_to_firestore(user_id: str, setting_key: str, setting_value) -> None:
    if not user_id:
        return
    try:
        db.collection("user_settings").document(user_id).set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

load_user_settings_from_firestore(user_id)

# ==================================================
# 2-b) ユーザー情報 / 権限
# ==================================================
current_user_email = user_id
current_user_name: Optional[str] = None

user_doc = get_or_create_user(current_user_email, current_user_name)
current_role = user_doc.get("role") or get_user_role(current_user_email)
is_admin = current_role == ROLE_ADMIN

# ==================================================
# 3) Google 認証
# ==================================================
google_auth_placeholder = st.empty()
with google_auth_placeholder.container():
    creds = authenticate_google()
    if not creds:
        st.warning("🔐 Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()

service, editable_calendar_options = ensure_services(creds)
tasks_service = st.session_state.get("tasks_service")
default_task_list_id = st.session_state.get("default_task_list_id")

try:
    sheets_service = build("sheets", "v4", credentials=creds)
except Exception as e:
    st.warning(f"Googleスプレッドシートサービスの初期化に失敗しました: {e}")
    sheets_service = None

st.session_state["sheets_service"] = sheets_service

# ==================================================
# 4) メインコンテンツ (UI改善版)
# ==================================================
# タブのコンテナ
st.markdown('<div class="fixed-tabs">', unsafe_allow_html=True)

tab_labels = [
    "1. ファイル取込",
    "2. 登録・削除",
    "3. 出力",
    "4. 物件マスタ",
]
if is_admin:
    tab_labels.append("5. 管理者")

tabs = st.tabs(tab_labels)
st.markdown("</div>", unsafe_allow_html=True)

if "uploaded_files" not in st.session_state:
    st.session_state["uploaded_files"] = []
    st.session_state["description_columns_pool"] = []
    st.session_state["merged_df_for_selector"] = pd.DataFrame()

# --- Tab 1: Upload ---
with tabs[0]:
    # コンテナで囲んで視認性を向上（色は変えず枠線と余白のみ）
    with st.container(border=True):
        st.caption("ExcelまたはCSVファイルをアップロードしてください")
        render_tab1_upload()

# --- Tab 2: Operations ---
with tabs[1]:
    sub_tab_reg, sub_tab_del, sub_tab_todo, sub_tab_notice_fax = st.tabs(
        ["📥 イベント登録", "🗑 イベント削除", "✅ 点検連絡ToDo", "📄 貼り紙・FAX"]
    )

    with sub_tab_reg:
        with st.container(border=True):
            render_tab2_register(user_id, editable_calendar_options, service)

    with sub_tab_del:
        with st.container(border=True):
            render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id)

    with sub_tab_todo:
        with st.container(border=True):
            render_tab7_inspection_todo(
                service=service,
                editable_calendar_options=editable_calendar_options,
                tasks_service=tasks_service,
                default_task_list_id=default_task_list_id,
                sheets_service=sheets_service,
                current_user_email=current_user_email,
            )

    with sub_tab_notice_fax:
        with st.container(border=True):
            render_tab8_notice_fax(
                service=service,
                editable_calendar_options=editable_calendar_options,
                sheets_service=sheets_service,
                current_user_email=current_user_email,
            )

# --- Tab 3: Export ---
with tabs[2]:
    with st.container(border=True):
        render_tab5_export(editable_calendar_options, service, fetch_all_events)

# --- Tab 4: Property Master ---
with tabs[3]:
    with st.container(border=True):
        render_tab6_property_master(
            sheets_service=sheets_service,
            default_spreadsheet_id=st.secrets.get("PROPERTY_MASTER_SHEET_ID", ""),
            basic_sheet_title="物件基本情報",
            master_sheet_title="物件マスタ",
            current_user_email=current_user_email,
        )

# --- Tab 5: Admin ---
if is_admin:
    with tabs[4]:
        with st.container(border=True):
            render_tab_admin(
                current_user_email=current_user_email,
                current_user_name=current_user_name,
            )

# ==================================================
# 5) サイドバー（別ファイル）
# ==================================================
render_sidebar(
    user_id=user_id,
    editable_calendar_options=editable_calendar_options,
    save_user_setting_to_firestore=save_user_setting_to_firestore,
)

