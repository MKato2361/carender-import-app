with tabs[1]:
    sub_tab_reg, sub_tab_del, sub_tab_todo, sub_tab_notice_fax = st.tabs(
        ["📥 イベント登録", "🗑 イベント削除", "✅ 点検連絡ToDo自動作成", "📄 貼り紙自動作成"]
    )

    with sub_tab_reg:
        render_tab2_register(user_id, editable_calendar_options, service)

    with sub_tab_del:
        render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id)

    with sub_tab_todo:
        render_tab7_inspection_todo(
            service=service,
            editable_calendar_options=editable_calendar_options,
            tasks_service=tasks_service,
            default_task_list_id=default_task_list_id,
            sheets_service=sheets_service,
            current_user_email=current_user_email,
        )

    with sub_tab_notice_fax:
        render_tab8_notice_fax(
            service=service,
            editable_calendar_options=editable_calendar_options,
            sheets_service=sheets_service,
            current_user_email=current_user_email,
            # もし以前 tasks_service / default_task_list_id を渡していても
            # **kwargs で無視されるようにしてあります
        )
