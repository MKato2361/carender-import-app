import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_from_calendar, get_existing_calendar_events, update_event_in_calendar, reconcile_events
from googleapiclient.discovery import build
import re
import io

st.set_page_config(page_title="Googleカレンダー登録・更新・削除ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・更新・削除")

# --- Google Calendar Authentication ---
google_auth_placeholder = st.empty()

with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    creds = authenticate_google()

    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()
        st.sidebar.success("✅ Googleカレンダーに認証済みです！")

if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    try:
        service = build("calendar", "v3", credentials=creds)
        st.session_state['calendar_service'] = service
        calendar_list = service.calendarList().list().execute()
        editable_calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') != 'reader'
        }
        st.session_state['editable_calendar_options'] = editable_calendar_options
    except Exception as e:
        st.error(f"カレンダーサービスの取得またはカレンダーリストの取得に失敗しました: {e}")
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop()
else:
    service = st.session_state['calendar_service']


# --- Tabs ---
tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの新規登録", "3. イベントの更新（作業指示書番号基準）", "4. イベントの削除"])

with tabs[0]: # 1. ファイルのアップロード
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        description_columns_pool = set()
        for file in uploaded_files:
            try:
                df_temp = pd.read_excel(file, engine="openpyxl")
                df_temp.columns = [str(c).strip() for c in df_temp.columns]
                description_columns_pool.update(df_temp.columns)
            except Exception as e:
                st.warning(f"{file.name} の読み込みに失敗しました: {e}")
        st.session_state['description_columns_pool'] = list(description_columns_pool)
    elif 'uploaded_files' not in st.session_state:
        st.session_state['uploaded_files'] = []
        st.session_state['description_columns_pool'] = []

    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みのファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")


with tabs[1]: # 2. イベントの新規登録 (既存の登録機能)
    st.header("イベントを新規登録")
    st.info("アップロードされたExcelファイルのすべてのイベントを新規にGoogleカレンダーに登録します。既存のイベントとの重複チェックは行われません。")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event_reg = st.checkbox("終日イベントとして登録", value=False, key="reg_all_day_event")
        private_event_reg = st.checkbox("非公開イベントとして登録", value=True, key="reg_private_event")

        description_columns_reg = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', []),
            key="reg_description_columns"
        )

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name_reg = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id_reg = st.session_state['editable_calendar_options'][selected_calendar_name_reg]

            st.subheader("👀 登録イベントのプレビュー")
            if st.button("登録プレビューを生成", key="generate_register_preview_button"):
                # strict_work_order_match=False: 作業指示書番号の有無にかかわらず全行を処理
                preview_df = process_excel_files(
                    st.session_state['uploaded_files'],
                    description_columns_reg,
                    all_day_event_reg,
                    private_event_reg,
                    strict_work_order_match=False # 新規登録タブではWO番号がなくても登録対象
                )
                if not preview_df.empty:
                    display_df = preview_df.copy()
                    display_df = display_df.rename(columns={
                        'WorkOrderNumber': '作業指示書番号',
                        'Subject': 'イベント名',
                        'Start Date': '開始日',
                        'Start Time': '開始時刻',
                        'End Date': '終了日',
                        'End Time': '終了時刻',
                        'Location': '場所',
                        'Description': '説明',
                        'All Day Event': '終日',
                        'Private': '非公開'
                    })
                    columns_to_display = ['作業指示書番号', 'イベント名', '開始日', '開始時刻', '終了日', '終了時刻', '場所', '説明', '終日', '非公開']
                    display_df = display_df[[col for col in columns_to_display if col in display_df.columns]]
                    st.info(f"{len(display_df)} 件のイベントが以下のように登録されます。")
                    st.dataframe(display_df, use_container_width=True)
                    st.session_state['preview_register_df'] = preview_df
                else:
                    st.warning("プレビューする有効なイベントデータがありません。")
                    st.session_state['preview_register_df'] = pd.DataFrame()

            st.subheader("➡️ イベント登録実行")
            if st.session_state.get('preview_register_df') is not None and not st.session_state['preview_register_df'].empty:
                if st.button("Googleカレンダーに一括登録する", key="execute_register_button"):
                    df_to_register = st.session_state['preview_register_df']
                    with st.spinner("イベントデータを登録中..."):
                        progress = st.progress(0)
                        successful_registrations = 0
                        for i, row in df_to_register.iterrows():
                            try:
                                if row['All Day Event'] == "True":
                                    start_date_str = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                                    end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date() # 既にexcel_parserで翌日になっている
                                    end_date_str = end_date_obj.strftime("%Y-%m-%d")
                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_str},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                else:
                                    # 日付と時刻が既にISO形式でなくても対応できるように
                                    start = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M").isoformat()
                                    end = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M").isoformat()
                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'],
                                        'description': row['Description'],
                                        'start': {'dateTime': start, 'timeZone': 'Asia/Tokyo'},
                                        'end': {'dateTime': end, 'timeZone': 'Asia/Tokyo'},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                add_event_to_calendar(service, calendar_id_reg, event_data)
                                successful_registrations += 1
                            except Exception as e:
                                st.error(f"{row['Subject']} の登録に失敗しました: {e}")
                            progress.progress((i + 1) / len(df_to_register))

                        st.success(f"✅ {successful_registrations} 件のイベント登録が完了しました！")
                        st.session_state['preview_register_df'] = pd.DataFrame()
            else:
                st.info("プレビューを生成してから登録を実行してください。")


with tabs[2]: # 3. イベントの更新（作業指示書番号基準）
    st.header("イベントの更新（作業指示書番号基準）")
    st.info("アップロードされたExcelファイル内の「作業指示書番号」をキーとして、既存のGoogleカレンダーイベントを更新します。変更がない場合はスキップし、Googleカレンダーに存在しない作業指示書番号のイベントは新規登録されます。")
    st.warning("Excelファイルに「作業指示書番号」がない、または空欄の行は、この機能の処理対象外となります。それらを新規登録したい場合は「2. イベントの新規登録」タブをご利用ください。")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
            st.stop()

        selected_calendar_name_update = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
        calendar_id_update = st.session_state['editable_calendar_options'][selected_calendar_name_update]

        st.subheader("🔍 期間と設定")
        if st.session_state.get('uploaded_files'):
            min_date_excel = date.today() - timedelta(days=30)
            max_date_excel = date.today() + timedelta(days=30)
        else:
            min_date_excel = date.today() - timedelta(days=30)
            max_date_excel = date.today() + timedelta(days=30)

        update_search_start_date = st.date_input("カレンダー検索開始日", value=min_date_excel - timedelta(days=7), key="update_start_search_date")
        update_search_end_date = st.date_input("カレンダー検索終了日", value=max_date_excel + timedelta(days=7), key="update_end_search_date")

        all_day_event_upd = st.checkbox("終日イベントとして更新（Excelの値で上書き）", value=False, key="upd_all_day_event")
        private_event_upd = st.checkbox("非公開イベントとして更新（Excelの値で上書き）", value=True, key="upd_private_event")
        description_columns_upd = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', []),
            key="upd_description_columns"
        )

        st.subheader("🔄 変更プレビュー")
        if st.button("変更をプレビュー", key="preview_update_button"):
            with st.spinner("既存イベントとExcelデータを比較中..."):
                # strict_work_order_match=True: WorkOrderNumberがない行はここで除外される
                excel_df_for_update = process_excel_files(
                    st.session_state['uploaded_files'],
                    description_columns_upd,
                    all_day_event_upd,
                    private_event_upd,
                    strict_work_order_match=True
                )

                if excel_df_for_update.empty:
                    st.warning("プレビューする有効なExcelデータがありません。作業指示書番号が特定できる行がないか、ファイルが空です。")
                    st.session_state['events_to_add_update'] = []
                    st.session_state['events_to_update_update'] = []
                    st.session_state['events_to_skip_update'] = []
                else:
                    # 日付オブジェクトをJSTのdatetimeオブジェクトに変換してから渡す
                    jst = timezone(timedelta(hours=9)) # JSTのタイムゾーンオブジェクトを作成
                    # 修正: .localize() の代わりに .replace(tzinfo=jst) を使用
                    start_dt_search = datetime.combine(update_search_start_date, datetime.min.time()).replace(tzinfo=jst)
                    end_dt_search = datetime.combine(update_search_end_date, datetime.max.time()).replace(tzinfo=jst)

                    existing_gcal_events = get_existing_calendar_events(
                        service, calendar_id_update,
                        start_dt_search,
                        end_dt_search
                    )

                    events_to_add_to_gcal, events_to_update_in_gcal, events_to_skip_due_to_no_change = reconcile_events(excel_df_for_update, existing_gcal_events)
                    
                    st.session_state['events_to_add_update'] = events_to_add_to_gcal
                    st.session_state['events_to_update_update'] = events_to_update_in_gcal
                    st.session_state['events_to_skip_update'] = events_to_skip_due_to_no_change

                    st.markdown("---")
                    st.success(f"結果: 新規登録 {len(events_to_add_to_gcal)} 件, 更新 {len(events_to_update_in_gcal)} 件, スキップ {len(st.session_state['events_to_skip_update'])} 件")

                    if events_to_add_to_gcal:
                        st.subheader("➕ 新規登録されるイベント")
                        display_add_df = pd.DataFrame({
                            '作業指示書番号': [e.get('WorkOrderNumber', '') for e in events_to_add_to_gcal],
                            'イベント名': [e['summary'] for e in events_to_add_to_gcal],
                            '開始': [e['start'].get('dateTime', e['start'].get('date')) for e in events_to_add_to_gcal],
                            '終了': [e['end'].get('dateTime', e['end'].get('date')) for e in events_to_add_to_gcal],
                            '場所': [e.get('location', '') for e in events_to_add_to_gcal],
                            '説明': [e.get('description', '') for e in events_to_add_to_gcal]
                        })
                        st.dataframe(display_add_df, use_container_width=True)

                    if events_to_update_in_gcal:
                        st.subheader("✏️ 更新されるイベント")
                        update_display_data = []
                        for e_upd in events_to_update_in_gcal:
                            new_data = e_upd['new_data']
                            old_summary = e_upd['old_summary']
                            
                            # 更新されるイベントの作業指示書番号も表示 (既存のsummaryから抽出)
                            wo_match_old = re.match(r"^(\d+)\s", old_summary) 
                            wo_number_old_display = wo_match_old.group(1) if wo_match_old else "N/A"

                            update_display_data.append({
                                '作業指示書番号': wo_number_old_display,
                                '既存イベント名': old_summary,
                                '新しいイベント名': new_data['summary'],
                                '新しい開始': new_data['start'].get('dateTime', new_data['start'].get('date')),
                                '新しい終了': new_data['end'].get('dateTime', new_data['end'].get('date')),
                                '新しい場所': new_data.get('location', ''),
                                '新しい説明': new_data.get('description', '')
                            })
                        st.dataframe(pd.DataFrame(update_display_data), use_container_width=True)
                    
                    if st.session_state['events_to_skip_update']:
                        st.subheader("➡️ スキップされるイベント（変更なし）")
                        display_skip_df = pd.DataFrame(st.session_state['events_to_skip_update'])
                        display_skip_df = display_skip_df.rename(columns={
                            'WorkOrderNumber': '作業指示書番号',
                            'Subject': 'イベント名',
                            'Start Date': '開始日',
                            'Start Time': '開始時刻',
                            'End Date': '終了日',
                            'End Time': '終了時刻',
                            'Location': '場所',
                            'Description': '説明',
                            'All Day Event': '終日',
                            'Private': '非公開'
                        })
                        columns_to_display = ['作業指示書番号', 'イベント名', '開始日', '開始時刻', '終了日', '終了時刻', '場所', '説明', '終日', '非公開']
                        display_skip_df = display_skip_df[[col for col in columns_to_display if col in display_skip_df.columns]]
                        st.dataframe(display_skip_df, use_container_width=True)


                    if not events_to_add_to_gcal and not events_to_update_in_gcal and not st.session_state['events_to_skip_update']:
                        st.info("変更・新規登録が必要なイベントは見つかりませんでした。")

        st.subheader("🚀 イベント更新実行")
        # プレビューデータが存在する場合のみ実行ボタンを表示
        if st.session_state.get('events_to_add_update') is not None and \
           st.session_state.get('events_to_update_update') is not None and \
           (st.session_state['events_to_add_update'] or st.session_state['events_to_update_update']):

            if st.button("Googleカレンダーに変更を反映する", key="execute_update_button"):
                with st.spinner("変更をGoogleカレンダーに反映中..."):
                    added_count = 0
                    updated_count = 0
                    
                    # 新規登録イベントの処理
                    if st.session_state['events_to_add_update']:
                        st.info(f"新規イベントを登録中 ({len(st.session_state['events_to_add_update'])} 件)...")
                        add_progress = st.progress(0, text="新規登録イベント...")
                        for i, event_data in enumerate(st.session_state['events_to_add_update']):
                            try:
                                # WorkOrderNumberはExcelから渡されるデータフレームにのみ存在するので、APIリクエストから除外
                                event_data_for_api = {k: v for k, v in event_data.items() if k != 'WorkOrderNumber'}
                                add_event_to_calendar(service, calendar_id_update, event_data_for_api)
                                added_count += 1
                            except Exception as e:
                                st.error(f"新規登録に失敗しました ({event_data.get('summary', '無題')}): {e}")
                            add_progress.progress((i + 1) / len(st.session_state['events_to_add_update']))
                        add_progress.empty()

                    # 更新イベントの処理
                    if st.session_state['events_to_update_update']:
                        st.info(f"既存イベントを更新中 ({len(st.session_state['events_to_update_update'])} 件)...")
                        update_progress = st.progress(0, text="既存イベント更新中...")
                        for i, update_item in enumerate(st.session_state['events_to_update_update']):
                            try:
                                # WorkOrderNumberはAPIリクエストから除外
                                new_data_for_api = {k: v for k, v in update_item['new_data'].items() if k != 'WorkOrderNumber'}
                                update_event_in_calendar(service, calendar_id_update, update_item['id'], new_data_for_api)
                                updated_count += 1
                            except Exception as e:
                                st.error(f"更新に失敗しました ({update_item['new_data'].get('summary', '無題')}): {e}")
                            update_progress.progress((i + 1) / len(st.session_state['events_to_update_update']))
                        update_progress.empty()

                    st.success(f"✅ Googleカレンダーへの変更が完了しました！ (新規登録: {added_count} 件, 更新: {updated_count} 件)")
                    # 処理完了後にプレビューデータをクリア
                    st.session_state['events_to_add_update'] = []
                    st.session_state['events_to_update_update'] = []
                    st.session_state['events_to_skip_update'] = []
                    st.rerun()

        else:
            st.info("「変更をプレビュー」ボタンを押して、変更内容を確認してください。")

with tabs[3]: # 4. イベントの削除
    st.header("イベントを削除")

    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダー認証を完了しているか、Googleカレンダーの設定を確認してください。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today = date.today()
        default_start_date = today - timedelta(days=30)
        default_end_date = today

        delete_start_date = st.date_input("削除開始日", value=default_start_date, key="del_start_date")
        delete_end_date = st.date_input("削除終了日", value=default_end_date, key="del_end_date")

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("👀 削除対象イベントのプレビュー")
            if st.button("削除対象をプレビュー", key="generate_delete_preview_button"):
                jst = timezone(timedelta(hours=9)) # JSTのタイムゾーンオブジェクトを作成
                # 修正: .localize() の代わりに .replace(tzinfo=jst) を使用
                start_dt_search = datetime.combine(delete_start_date, datetime.min.time()).replace(tzinfo=jst)
                end_dt_search = datetime.combine(delete_end_date, datetime.max.time()).replace(tzinfo=jst)

                events_to_delete_preview = get_existing_calendar_events(
                    service, calendar_id_del,
                    start_dt_search,
                    end_dt_search
                )

                if events_to_delete_preview:
                    st.info(f"以下の {len(events_to_delete_preview)} 件のイベントが削除されます。")
                    display_events = []
                    for event in events_to_delete_preview:
                        summary = event.get('summary', 'タイトルなし')
                        start_info = event['start'].get('dateTime', event['start'].get('date'))
                        end_info = event['end'].get('dateTime', event['end'].get('date'))
                        
                        # Descriptionから作業指示書番号を抽出して表示 (数字のみ)
                        description = event.get('description', '')
                        wo_match = re.match(r"^作業指示書:(\d+)\s*/?\s*", description) # 数字のみを抽出
                        wo_number_display = wo_match.group(1) if wo_match else "N/A"

                        display_events.append({
                            '作業指示書番号': wo_number_display,
                            'イベント名': summary,
                            '開始日時': start_info,
                            '終了日時': end_info,
                            '場所': event.get('location', '')
                        })
                    st.dataframe(pd.DataFrame(display_events), use_container_width=True)
                    st.session_state['events_to_delete_confirm'] = events_to_delete_preview
                else:
                    st.info("指定された期間内に削除するイベントは見つかりませんでした。")
                    st.session_state['events_to_delete_confirm'] = []

            st.subheader("🗑️ 削除実行")

            if 'show_delete_confirmation' not in st.session_state:
                st.session_state.show_delete_confirmation = False
            if 'last_deleted_count' not in st.session_state:
                st.session_state.last_deleted_count = None

            if st.session_state.get('events_to_delete_confirm') and len(st.session_state['events_to_delete_confirm']) > 0 and not st.session_state.show_delete_confirmation:
                if st.button("上記のイベントを削除する", key="initiate_delete_button"):
                    st.session_state.show_delete_confirmation = True
                    st.session_state.last_deleted_count = None # 削除前にリセット
                    st.rerun() # 確認ダイアログを表示するために再実行

            if st.session_state.show_delete_confirmation:
                st.warning(f"「{selected_calendar_name_del}」カレンダーから {delete_start_date.strftime('%Y年%m月%d日')}から{delete_end_date.strftime('%Y年%m月%d日')}までの全てのイベントを削除します。この操作は元に戻せません。よろしいですか？")

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("はい、削除を実行します", key="confirm_delete_button_final"):
                        # 日付オブジェクトをJSTのdatetimeオブジェクトに変換してから渡す
                        jst = timezone(timedelta(hours=9)) # JSTのタイムゾーンオブジェクトを作成
                        start_dt_delete = datetime.combine(delete_start_date, datetime.min.time()).replace(tzinfo=jst)
                        end_dt_delete = datetime.combine(delete_end_date, datetime.max.time()).replace(tzinfo=jst)
                        
                        deleted_count = delete_events_from_calendar(
                            service, calendar_id_del,
                            start_dt_delete,
                            end_dt_delete
                        )
                        st.session_state.last_deleted_count = deleted_count
                        st.session_state.show_delete_confirmation = False
                        st.session_state['events_to_delete_confirm'] = [] # 削除後はプレビューをクリア
                        st.rerun()
                with col2:
                    if st.button("いいえ、キャンセルします", key="cancel_delete_button"):
                        st.info("削除はキャンセルされました。")
                        st.session_state.show_delete_confirmation = False
                        st.session_state.last_deleted_count = None
                        st.session_state['events_to_delete_confirm'] = [] # キャンセル時もプレビューをクリア
                        st.rerun()

            # 削除完了メッセージを表示
            if not st.session_state.show_delete_confirmation and st.session_state.last_deleted_count is not None:
                if st.session_state.last_deleted_count > 0:
                    st.success(f"✅ {st.session_state.last_deleted_count} 件のイベントが削除されました。")
                else:
                    st.info("指定された期間内に削除されたイベントはありませんでした。")
