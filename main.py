from __future__ import annotations
import streamlit as st
from auth_manager import get_auth_manager, AuthManager
from utils.user_roles import get_or_create_user, ROLE_ADMIN
from sidebar import render_sidebar
from ui.auth_forms import login_form as firebase_auth_form
from tabs.tab1_upload import render_tab1_upload
from tabs.tab2_register import render_tab2_register
from tabs.tab3_delete import render_tab3_delete
from tabs.tab5_export import render_tab5_export
from tabs.tab_admin import render_tab_admin
from tabs.tab6_property_master import render_tab6_property_master
from tabs.tab7_inspection_todo import render_tab7_inspection_todo
from tabs.tab8_notice_fax import render_tab8_notice_fax

# ──────────────────────────────────────────────
# ページ設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Googleカレンダー一括管理システム",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# グローバルスタイル
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20,300,0,0&display=block');

/* ── Material Icons ── */
.mi {
    font-family: 'Material Symbols Outlined';
    font-variation-settings: 'FILL' 0, 'wght' 300, 'GRAD' 0, 'opsz' 20;
    font-size: 16px;
    line-height: 1;
    vertical-align: middle;
    display: inline-block;
    color: inherit;
    margin-right: 4px;
    position: relative;
    top: -1px;
}
.section-heading .mi { font-size: 15px; color: #9d9b96; }
.cal-card-icon { font-size: 14px; }

/* ── ベース ── */
html, body, [class*="css"] {
    font-family: 'Noto Sans JP', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* ── ページ背景 ── */
.stApp {
    background: #f8f7f5;
}
@media (prefers-color-scheme: dark) {
    .stApp { background: #0f0f0f; }
}

/* ── メインコンテンツエリア ── */
section[data-testid="stMain"] > div {
    padding-top: 1.5rem;
}

/* ── タブバー ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    background: transparent;
    border-bottom: 1px solid #e5e3df;
    padding: 0;
    position: sticky;
    top: 0;
    z-index: 100;
    background-color: #f8f7f5;
}
@media (prefers-color-scheme: dark) {
    .stTabs [data-baseweb="tab-list"] { background-color: #0f0f0f; border-color: #2a2a2a; }
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
    font-weight: 500;
    padding: 10px 20px;
    border-radius: 0;
    color: #888;
    letter-spacing: .01em;
    border-bottom: 2px solid transparent;
    transition: color .15s, border-color .15s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #4f46e5;
    background: transparent;
}
.stTabs [aria-selected="true"] {
    background: transparent !important;
    color: #4f46e5 !important;
    font-weight: 700;
    border-bottom: 2px solid #4f46e5 !important;
}

/* ── セクション見出し ── */
.section-heading {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 11px;
    font-weight: 700;
    color: #9d9b96;
    letter-spacing: .08em;
    text-transform: uppercase;
    margin: 22px 0 10px;
    padding-bottom: 0;
    border-bottom: none;
    display: flex;
    align-items: center;
    gap: 6px;
}
.section-heading::after {
    content: '';
    flex: 1;
    height: 1px;
    background: #e5e3df;
}

/* ── アプリ名 ── */
.app-subtitle {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
    font-weight: 500;
    color: #9d9b96;
    margin-bottom: 18px;
    letter-spacing: .02em;
}

/* ── カレンダー選択カード ── */
div[style*="border:2px solid #1E88E5"] {
    border: 1.5px solid #c7d2fe !important;
    background: #eef2ff !important;
    border-radius: 12px !important;
}

/* ── プライマリボタン ── */
.stButton > button[kind="primary"] {
    background: #4f46e5;
    border: none;
    border-radius: 8px;
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
    font-weight: 600;
    letter-spacing: .02em;
    padding: 0 20px;
    height: 38px;
    box-shadow: 0 1px 3px rgba(79,70,229,.25);
    transition: background .15s, box-shadow .15s, transform .1s;
}
.stButton > button[kind="primary"]:hover {
    background: #4338ca;
    box-shadow: 0 4px 12px rgba(79,70,229,.35);
    transform: translateY(-1px);
}
.stButton > button[kind="primary"]:active {
    transform: translateY(0);
    box-shadow: 0 1px 2px rgba(79,70,229,.2);
}

/* ── セカンダリボタン ── */
.stButton > button[kind="secondary"] {
    border-radius: 8px;
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
    border: 1px solid #e5e3df;
    background: white;
    transition: border-color .15s, background .15s;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #c7d2fe;
    background: #f5f3ff;
}

/* ── リンクボタン ── */
.stLinkButton a {
    background: #4f46e5 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'Noto Sans JP', sans-serif !important;
}

/* ── divider ── */
hr {
    border: none !important;
    border-top: 1px solid #e5e3df !important;
    margin: 16px 0 !important;
}

/* ── st.info / success / warning / error ── */
[data-testid="stAlert"] {
    border-radius: 10px;
    border: none;
    padding: 12px 16px;
}

/* ── 入力フィールド ── */
.stTextInput > div > div > input,
.stSelectbox > div > div,
.stMultiSelect > div > div {
    border-radius: 8px;
    border-color: #e5e3df;
    font-family: 'Noto Sans JP', sans-serif;
}
.stTextInput > div > div > input:focus {
    border-color: #4f46e5;
    box-shadow: 0 0 0 3px rgba(79,70,229,.12);
}

/* ── expander ── */
.stExpander {
    border: 1px solid #e5e3df !important;
    border-radius: 10px !important;
    background: white;
}
.stExpander > details > summary {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
    font-weight: 600;
    color: #3d3b37;
}

/* ── progress bar ── */
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, #4f46e5, #818cf8);
}

/* ── スピナー ── */
.stSpinner > div {
    border-top-color: #4f46e5 !important;
}

/* ── checkbox ── */
.stCheckbox > label > span[data-testid="stCheckboxLabel"] {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 13px;
}

/* ── step indicator (認証フロー) ── */
.step-dot-active {
    background: #4f46e5;
    color: white;
    border-radius: 50%;
    width: 22px; height: 22px;
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 700;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────
def main():
    manager: AuthManager = get_auth_manager()

    # ── 1. Firebase 認証 ──
    user_id = manager.sync_with_session()
    if not user_id:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>', unsafe_allow_html=True)
            st.markdown("""
<div style="display:flex;gap:8px;align-items:center;margin-bottom:20px;">
  <span style="background:#4f46e5;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">1</span>
  <span style="font-weight:600;color:#4f46e5;font-size:13px;">アカウントにログイン</span>
  <span style="color:#ddd;margin:0 2px;">→</span>
  <span style="background:#f1f0ee;color:#bbb;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">2</span>
  <span style="color:#bbb;font-size:13px;">Googleカレンダーと連携</span>
</div>
""", unsafe_allow_html=True)
            firebase_auth_form()
        st.stop()

    # ── 2. Google 認証 ──
    if not manager.ensure_google_services():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>', unsafe_allow_html=True)
            st.markdown("""
<div style="display:flex;gap:8px;align-items:center;margin-bottom:20px;">
  <span style="background:#16a34a;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">✓</span>
  <span style="font-weight:600;color:#16a34a;font-size:13px;">ログイン済み</span>
  <span style="color:#ddd;margin:0 2px;">→</span>
  <span style="background:#4f46e5;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">2</span>
  <span style="font-weight:600;color:#4f46e5;font-size:13px;">Googleカレンダーと連携</span>
</div>
""", unsafe_allow_html=True)
        st.stop()

    # ── 3. ユーザー情報 ──
    current_user_email = st.session_state.get("user_email") or user_id
    user_doc  = get_or_create_user(current_user_email, None)
    is_admin  = user_doc.get("role") == ROLE_ADMIN

    # ── 4. サイドバー ──
    render_sidebar(
        user_id=user_id,
        editable_calendar_options=manager.editable_calendar_options,
        save_user_setting_to_firestore=manager.save_user_setting,
    )

    # ── 5. ページタイトル（サイドバーが開いているとき用） ──
    st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>', unsafe_allow_html=True)

    # ── 6. タブ ──
    tab_labels = ["ファイル取込", "登録・操作", "出力", "物件マスタ"]
    if is_admin:
        tab_labels.append("管理者")

    tabs = st.tabs(tab_labels)

    with tabs[0]:
        render_tab1_upload()

    with tabs[1]:
        sub_tabs = st.tabs(["イベント登録", "イベント削除", "点検ToDo", "貼り紙・FAX"])
        with sub_tabs[0]:
            render_tab2_register(user_id, manager)
        with sub_tabs[1]:
            render_tab3_delete(
                manager.editable_calendar_options,
                manager.calendar_service,
                manager.tasks_service,
                manager.default_task_list_id,
            )
        with sub_tabs[2]:
            render_tab7_inspection_todo(
                service=manager.calendar_service,
                editable_calendar_options=manager.editable_calendar_options,
                tasks_service=manager.tasks_service,
                default_task_list_id=manager.default_task_list_id,
                sheets_service=manager.sheets_service,
                current_user_email=current_user_email,
            )
        with sub_tabs[3]:
            render_tab8_notice_fax(manager, current_user_email)

    with tabs[2]:
        render_tab5_export(manager)

    with tabs[3]:
        render_tab6_property_master(
            sheets_service=manager.sheets_service,
            default_spreadsheet_id=st.secrets.get("PROPERTY_MASTER_SHEET_ID", ""),
            basic_sheet_title="物件基本情報",
            master_sheet_title="物件マスタ",
            current_user_email=current_user_email,
        )

    if is_admin:
        with tabs[4]:
            render_tab_admin(manager, current_user_email)


if __name__ == "__main__":
    main()
