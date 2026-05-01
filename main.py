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

# ── ページ設定 ──
st.set_page_config(
    page_title="Googleカレンダー一括管理システム",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── グローバルスタイル ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;600;700&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20,300,0,0&display=block');

/* ── Material Icons ── */
.mi {
    font-family: 'Material Symbols Outlined';
    font-variation-settings: 'FILL' 0, 'wght' 300, 'GRAD' 0, 'opsz' 20;
    font-size: 16px; line-height: 1; vertical-align: middle;
    display: inline-block; color: inherit; margin-right: 4px;
    position: relative; top: -1px;
}
.section-heading .mi { font-size: 15px; color: var(--text-3); }

/* ── iOS準拠カラー変数（ライトモード） ── */
:root {
  --app-bg:          #F2F2F7;
  --surface:         #FFFFFF;
  --surface-2:       #F2F2F7;
  --surface-3:       #E5E5EA;
  --border:          #C6C6C8;
  --border-strong:   #AEAEB2;
  --separator:       rgba(60,60,67,.36);
  --text-1:          #000000;
  --text-2:          rgba(60,60,67,.6);
  --text-3:          rgba(60,60,67,.3);
  --text-4:          rgba(60,60,67,.18);
  --accent:          #5856D6;
  --accent-hover:    #4644BC;
  --accent-surface:  #EDEDFA;
  --accent-border:   #C0BFF2;
  --accent-text:     #1D1C6E;
  --accent-label:    #5856D6;
  --success:         #34C759;
  --success-surface: #E8FAF0;
  --danger:          #FF3B30;
  --danger-surface:  #FFF0EF;
  --warning:         #FF9500;
}

/* ── iOS準拠カラー変数（ダークモード） ── */
@media (prefers-color-scheme: dark) {
  :root {
    --app-bg:          #000000;
    --surface:         #1C1C1E;
    --surface-2:       #2C2C2E;
    --surface-3:       #3A3A3C;
    --border:          #38383A;
    --border-strong:   #48484A;
    --separator:       rgba(84,84,88,.6);
    --text-1:          #FFFFFF;
    --text-2:          rgba(235,235,245,.7);
    --text-3:          rgba(235,235,245,.45);
    --text-4:          rgba(235,235,245,.2);
    --accent:          #5E5CE6;
    --accent-hover:    #7472EC;
    --accent-surface:  #1C1438;
    --accent-border:   #3D3A7A;
    --accent-text:     #A5A3F0;
    --accent-label:    #8381EB;
    --success:         #30D158;
    --success-surface: #0D2C1A;
    --danger:          #FF453A;
    --danger-surface:  #2C0F0E;
    --warning:         #FF9F0A;
  }
}

/* ── ベース ── */
html, body, [class*="css"] {
    font-family: 'Noto Sans JP', -apple-system, BlinkMacSystemFont, sans-serif;
}
.stApp { background: var(--app-bg) !important; }
section[data-testid="stMain"] > div { padding-top: 1.5rem; }

/* ── タブバー ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 0; background: var(--app-bg) !important;
    border-bottom: 1px solid var(--border);
    padding: 0; position: sticky; top: 0; z-index: 100;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Noto Sans JP', sans-serif; font-size: 13px; font-weight: 500;
    padding: 10px 20px; border-radius: 0; color: var(--text-3);
    letter-spacing: .01em; border-bottom: 2px solid transparent;
    transition: color .15s, border-color .15s;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--accent); background: transparent; }
.stTabs [aria-selected="true"] {
    background: transparent !important;
    color: var(--accent) !important; font-weight: 700;
    border-bottom: 2px solid var(--accent) !important;
}

/* ── セクション見出し ── */
.section-heading {
    font-size: 11px; font-weight: 700; color: var(--text-3);
    letter-spacing: .08em; text-transform: uppercase;
    margin: 22px 0 10px; display: flex; align-items: center; gap: 6px;
}
.section-heading::after { content: ''; flex: 1; height: 1px; background: var(--border); }

/* ── アプリ名 ── */
.app-subtitle {
    font-size: 13px; font-weight: 500; color: var(--text-3);
    margin-bottom: 18px; letter-spacing: .02em;
}

/* ── プライマリボタン ── */
.stButton > button[kind="primary"] {
    background: var(--accent); border: none; border-radius: 8px;
    font-family: 'Noto Sans JP', sans-serif; font-size: 13px;
    font-weight: 600; letter-spacing: .02em; padding: 0 20px; height: 38px;
    box-shadow: 0 1px 3px rgba(88,86,214,.25);
    transition: background .15s, box-shadow .15s, transform .1s;
}
.stButton > button[kind="primary"]:hover {
    background: var(--accent-hover) !important;
    box-shadow: 0 4px 12px rgba(88,86,214,.3); transform: translateY(-1px);
}
.stButton > button[kind="primary"]:active { transform: translateY(0); }

/* ── セカンダリボタン ── */
.stButton > button[kind="secondary"] {
    border-radius: 8px; font-family: 'Noto Sans JP', sans-serif; font-size: 13px;
    border: 1px solid var(--border); background: var(--surface); color: var(--text-1);
    transition: border-color .15s, background .15s;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--accent-border); background: var(--accent-surface);
}

/* ── リンクボタン ── */
.stLinkButton a {
    background: var(--accent) !important; border-radius: 8px !important;
    font-weight: 600 !important;
}

/* ── divider ── */
hr { border: none !important; border-top: 1px solid var(--border) !important; margin: 16px 0 !important; }

/* ── Alert ── */
[data-testid="stAlert"] { border-radius: 10px; border: none; padding: 12px 16px; }

/* ── 入力フィールド ── */
.stTextInput > div > div > input,
.stSelectbox > div > div,
.stMultiSelect > div > div {
    border-radius: 8px; border-color: var(--border);
    font-family: 'Noto Sans JP', sans-serif;
    background: var(--surface); color: var(--text-1);
}
.stTextInput > div > div > input:focus {
    border-color: var(--accent); box-shadow: 0 0 0 3px rgba(94,92,230,.15);
}

/* ── expander ── */
.stExpander {
    border: 1px solid var(--border) !important; border-radius: 10px !important;
    background: var(--surface) !important;
}
.stExpander > details > summary {
    font-family: 'Noto Sans JP', sans-serif; font-size: 13px;
    font-weight: 600; color: var(--text-1);
}

/* ── progress / spinner ── */
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent-label));
}
.stSpinner > div { border-top-color: var(--accent) !important; }

/* ── checkbox ── */
.stCheckbox > label > span[data-testid="stCheckboxLabel"] {
    font-family: 'Noto Sans JP', sans-serif; font-size: 13px; color: var(--text-1);
}

/* ── dataframe ── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background: var(--surface) !important; border-right: 1px solid var(--border);
}
</style>
""", unsafe_allow_html=True)


def _step_indicator(step1_done: bool) -> None:
    """認証ステップインジケーターを表示する。"""
    if step1_done:
        s1_bg, s1_text, s1_label_color = "var(--success)", "white", "var(--success)"
        s1_icon = "✓"
        s2_bg, s2_text, s2_label_color = "var(--accent)", "white", "var(--accent)"
    else:
        s1_bg, s1_text, s1_label_color = "var(--accent)", "white", "var(--accent)"
        s1_icon = "1"
        s2_bg, s2_text, s2_label_color = "var(--surface-3)", "var(--text-4)", "var(--text-4)"

    s1_label = "ログイン済み" if step1_done else "アカウントにログイン"
    s2_label = "Googleカレンダーと連携"

    st.markdown(f"""
<div style="display:flex;gap:8px;align-items:center;margin-bottom:20px;">
  <span style="background:{s1_bg};color:{s1_text};border-radius:50%;width:22px;height:22px;
    display:inline-flex;align-items:center;justify-content:center;
    font-size:11px;font-weight:700;">{s1_icon}</span>
  <span style="font-weight:600;color:{s1_label_color};font-size:13px;">{s1_label}</span>
  <span style="color:var(--border-strong);margin:0 2px;">→</span>
  <span style="background:{s2_bg};color:{s2_text};border-radius:50%;width:22px;height:22px;
    display:inline-flex;align-items:center;justify-content:center;
    font-size:11px;font-weight:700;">2</span>
  <span style="font-weight:600;color:{s2_label_color};font-size:13px;">{s2_label}</span>
</div>
""", unsafe_allow_html=True)


def main():
    manager: AuthManager = get_auth_manager()

    # ── 1. Firebase 認証 ──
    user_id = manager.sync_with_session()
    if not user_id:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>',
                        unsafe_allow_html=True)
            _step_indicator(step1_done=False)
            firebase_auth_form()
        st.stop()

    # ── 2. Google 認証 ──
    if not manager.ensure_google_services():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>',
                        unsafe_allow_html=True)
            _step_indicator(step1_done=True)
        st.stop()

    # ── 3. サービス取得 ──
    calendar_service          = st.session_state.get("calendar_service")
    tasks_service             = st.session_state.get("tasks_service")
    sheets_service            = st.session_state.get("sheets_service")
    editable_calendar_options = st.session_state.get("editable_calendar_options", {})
    default_task_list_id      = st.session_state.get("default_task_list_id")

    # ── 4. ユーザー情報 ──
    current_user_email = st.session_state.get("user_email") or user_id
    user_doc  = get_or_create_user(current_user_email, None)
    is_admin  = user_doc.get("role") == ROLE_ADMIN

    # ── 5. サイドバー ──
    render_sidebar(
        user_id=user_id,
        editable_calendar_options=editable_calendar_options,
        save_user_setting_to_firestore=manager.save_user_setting,
    )

    # ── 6. アプリタイトル ──
    st.markdown('<div class="app-subtitle">Googleカレンダー一括管理システム</div>',
                unsafe_allow_html=True)

    # ── 7. タブ ──
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
                editable_calendar_options,
                calendar_service,
                tasks_service,
                default_task_list_id,
            )
        with sub_tabs[2]:
            render_tab7_inspection_todo(
                service=calendar_service,
                editable_calendar_options=editable_calendar_options,
                tasks_service=tasks_service,
                default_task_list_id=default_task_list_id,
                sheets_service=sheets_service,
                current_user_email=current_user_email,
            )
        with sub_tabs[3]:
            render_tab8_notice_fax(manager, current_user_email)

    with tabs[2]:
        render_tab5_export(manager)

    with tabs[3]:
        render_tab6_property_master(
            sheets_service=sheets_service,
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
