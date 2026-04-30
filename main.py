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


def main():
    manager: AuthManager = get_auth_manager()

    # ── 1. Firebase 認証 ──
    user_id = manager.sync_with_session()
    if not user_id:
        firebase_auth_form()
        st.stop()

    # ── 2. Google 認証 ──
    if not manager.ensure_google_services():
        st.stop()

    # ★ ここ重要：session_stateから取得
    calendar_service = st.session_state.get("calendar_service")
    tasks_service = st.session_state.get("tasks_service")
    sheets_service = st.session_state.get("sheets_service")
    editable_calendar_options = st.session_state.get("editable_calendar_options", {})
    default_task_list_id = st.session_state.get("default_task_list_id")

    # ── 3. ユーザー情報 ──
    current_user_email = st.session_state.get("user_email") or user_id
    user_doc = get_or_create_user(current_user_email, None)
    is_admin = user_doc.get("role") == ROLE_ADMIN

    # ── 4. サイドバー ──
    render_sidebar(
        user_id=user_id,
        editable_calendar_options=editable_calendar_options,
        save_user_setting_to_firestore=manager.save_user_setting,
    )

    # ── 5. タブ ──
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
