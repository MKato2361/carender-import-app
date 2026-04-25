from __future__ import annotations
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone

# ---- 認証・サービス管理 (AuthManager導入) ----
from auth_manager import get_auth_manager, AuthManager

# ---- Utils & Helpers ----
from utils.user_roles import get_or_create_user, ROLE_ADMIN
from sidebar import render_sidebar
from ui.auth_forms import login_form as firebase_auth_form

# ---- Tab Modules ----
from tabs.tab1_upload import render_tab1_upload
from tabs.tab2_register import render_tab2_register
from tabs.tab3_delete import render_tab3_delete
from tabs.tab5_export import render_tab5_export
from tabs.tab_admin import render_tab_admin
from tabs.tab6_property_master import render_tab6_property_master
from tabs.tab7_inspection_todo import render_tab7_inspection_todo
from tabs.tab8_notice_fax import render_tab8_notice_fax

# ==================================================
# 0) ページ設定 & スタイル
# ==================================================
st.set_page_config(
    page_title="G-Cal Pro | Googleカレンダー一括管理",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

def apply_custom_styles():
    """
    UIの最適化: Stickyヘッダーの改善と階層の整理
    """
    st.markdown("""
    <style>
    /* メインヘッダーの装飾 */
    .main-header {
        font-size: 24px;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 10px;
        padding: 10px 0;
        border-bottom: 2px solid #f0f2f6;
    }
    /* タブの余白調整 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
        position: sticky;
        top: 0px;
        background-color: white;
        z-index: 1000;
        padding: 5px 0;
    }
    /* ダークモード対応 */
    @media (prefers-color-scheme: dark) {
        .stTabs [data-baseweb="tab-list"] {
            background-color: #0e1117;
        }
    }
    </style>
    <div class="main-header">📅 Googleカレンダー一括管理システム</div>
    """, unsafe_allow_html=True)

apply_custom_styles()

# ==================================================
# メインアプリケーションロジック
# ==================================================
def main():
    # AuthManagerの取得 (シングルトン)
    manager: AuthManager = get_auth_manager()
    
    # 1. Firebase 認証チェック
    user_id = manager.sync_with_session()
    if not user_id:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            # ステップ進捗表示
            st.markdown("""
<div style="display:flex;gap:8px;align-items:center;margin-bottom:16px">
  <span style="background:#1E88E5;color:white;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:600">1</span>
  <span style="font-weight:600;color:#1E88E5">アカウントにログイン</span>
  <span style="color:#ccc;margin:0 4px">→</span>
  <span style="background:#e0e0e0;color:#999;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:600">2</span>
  <span style="color:#999">Googleカレンダーと連携</span>
</div>
""", unsafe_allow_html=True)
            firebase_auth_form()
        st.stop()

    # 2. Google 認証 & サービス初期化
    if not manager.ensure_google_services():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("""
<div style="display:flex;gap:8px;align-items:center;margin-bottom:16px">
  <span style="background:#4caf50;color:white;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:600">✓</span>
  <span style="color:#4caf50">ログイン済み</span>
  <span style="color:#ccc;margin:0 4px">→</span>
  <span style="background:#1E88E5;color:white;border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:600">2</span>
  <span style="font-weight:600;color:#1E88E5">Googleカレンダーと連携</span>
</div>
""", unsafe_allow_html=True)
        # authenticate_google() が内部で URL を表示し st.stop() するため、ここでの明示的な stop は不要な場合が多い
        st.stop()

    # 3. ユーザー情報 / 権限
    current_user_email = st.session_state.get("user_email") or user_id
    user_doc = get_or_create_user(current_user_email, None)
    is_admin = user_doc.get("role") == ROLE_ADMIN

    # 4. サイドバー
    render_sidebar(
        user_id=user_id,
        editable_calendar_options=manager.editable_calendar_options,
        save_user_setting_to_firestore=manager.save_user_setting,
    )

    # 5. メインコンテンツ (タブ構成)
    tab_labels = ["1. ファイル取込", "2. 登録・削除", "3. 出力", "4. 物件マスタ"]
    if is_admin:
        tab_labels.append("5. 管理者")

    tabs = st.tabs(tab_labels)

    # --- Tab 1: Upload ---
    with tabs[0]:
        with st.container(border=True):
            render_tab1_upload()

    # --- Tab 2: Operations ---
    with tabs[1]:
        sub_tab_reg, sub_tab_del, sub_tab_todo, sub_tab_notice_fax = st.tabs(
            ["📥 イベント登録", "🗑 イベント削除", "✅ 点検連絡ToDo", "📄 貼り紙・FAX"]
        )

        with sub_tab_reg:
            with st.container(border=True):
                # 修正ポイント: 引数を manager に集約
                render_tab2_register(user_id, manager)

        with sub_tab_del:
            with st.container(border=True):
                # 他のタブも順次 manager 1つに修正することを想定。
                # 現時点では互換性のために manager から個別に取り出して渡す。
                render_tab3_delete(
                    manager.editable_calendar_options, 
                    manager.calendar_service, 
                    manager.tasks_service, 
                    manager.default_task_list_id
                )

        with sub_tab_todo:
            with st.container(border=True):
                render_tab7_inspection_todo(
                    service=manager.calendar_service,
                    editable_calendar_options=manager.editable_calendar_options,
                    tasks_service=manager.tasks_service,
                    default_task_list_id=manager.default_task_list_id,
                    sheets_service=manager.sheets_service,
                    current_user_email=current_user_email,
                )

        with sub_tab_notice_fax:
            with st.container(border=True):
                render_tab8_notice_fax(manager, current_user_email)

    # --- Tab 3: Export ---
    with tabs[2]:
        with st.container(border=True):
            render_tab5_export(manager)

    # --- Tab 4: Property Master ---
    with tabs[3]:
        with st.container(border=True):
            render_tab6_property_master(
                sheets_service=manager.sheets_service,
                default_spreadsheet_id=st.secrets.get("PROPERTY_MASTER_SHEET_ID", ""),
                basic_sheet_title="物件基本情報",
                master_sheet_title="物件マスタ",
                current_user_email=current_user_email,
            )

    # --- Tab 5: Admin ---
    if is_admin:
        with tabs[4]:
            with st.container(border=True):
                render_tab_admin(manager, current_user_email)

if __name__ == "__main__":
    main()
