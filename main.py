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
/* ── タブバー ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: transparent;
    border-bottom: 1.5px solid var(--color-border-tertiary, #e8eaf0);
    padding: 0;
    position: sticky;
    top: 0;
    z-index: 100;
    background-color: var(--background-color, white);
}
.stTabs [data-baseweb="tab"] {
    font-size: 14px;
    font-weight: 500;
    padding: 10px 18px;
    border-radius: 8px 8px 0 0;
    color: var(--text-color-secondary, #666);
}
.stTabs [aria-selected="true"] {
    background: #EFF6FF !important;
    color: #1E88E5 !important;
    font-weight: 700;
    border-bottom: 2.5px solid #1E88E5;
}
/* ダークモード */
@media (prefers-color-scheme: dark) {
    .stTabs [data-baseweb="tab-list"] { background-color: #0e1117; }
    .stTabs [aria-selected="true"] { background: #1a2744 !important; }
}
/* ── セクション見出し共通 ── */
.section-heading {
    font-size: 13px;
    font-weight: 700;
    color: var(--text-color-secondary, #666);
    letter-spacing: .04em;
    text-transform: uppercase;
    margin: 20px 0 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--color-border-tertiary, #e8eaf0);
}
/* ── ページタイトル ── */
.app-subtitle {
    font-size: 12px;
    color: var(--text-color-secondary, #888);
    margin-bottom: 16px;
}
/* ── カード・境界線をスッキリ ── */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlock"] > div[class*="st-emotion-cache"] {
    border-radius: 10px;
}
/* ── ボタン ── */
.stButton > button[kind="primary"] {
    border-radius: 8px;
    font-weight: 600;
}
/* ── divider を細く ── */
hr {
    border-color: var(--color-border-tertiary, #e8eaf0) !important;
    margin: 12px 0 !important;
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
  <span style="background:#1E88E5;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">1</span>
  <span style="font-weight:600;color:#1E88E5;font-size:13px;">アカウントにログイン</span>
  <span style="color:#ddd;margin:0 2px;">→</span>
  <span style="background:#e8eaf0;color:#bbb;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">2</span>
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
  <span style="background:#43a047;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">✓</span>
  <span style="font-weight:600;color:#43a047;font-size:13px;">ログイン済み</span>
  <span style="color:#ddd;margin:0 2px;">→</span>
  <span style="background:#1E88E5;color:white;border-radius:50%;width:22px;height:22px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;">2</span>
  <span style="font-weight:600;color:#1E88E5;font-size:13px;">Googleカレンダーと連携</span>
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
    tab_labels = ["📁 ファイル取込", "📥 登録・操作", "📤 出力", "🏠 物件マスタ"]
    if is_admin:
        tab_labels.append("⚙️ 管理者")

    tabs = st.tabs(tab_labels)

    with tabs[0]:
        render_tab1_upload()

    with tabs[1]:
        sub_tabs = st.tabs(["📥 イベント登録", "🗑️ イベント削除", "✅ 点検ToDo", "📄 貼り紙・FAX"])
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
