import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from io import BytesIO
import re
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from firebase_admin import firestore
from excel_parser import (
    process_excel_data_for_calendar,
    _load_and_merge_dataframes,
    get_available_columns_for_event_name,
    check_event_name_columns,
    format_worksheet_value
)
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id,
    delete_event_from_calendar
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from session_utils import (
    initialize_session_state,
    get_user_setting,
    set_user_setting,
    get_all_user_settings,
    clear_user_settings
)

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# --- 初期化と認証 ---
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    firebase_auth_form()
    st.stop()

def load_user_settings_from_firestore():
    """Firestoreからユーザー設定を読み込み、セッションに同期"""
    initialize_session_state(user_id)
    doc = db.collection('user_settings').document(user_id).get()
    if doc.exists:
        for key, value in doc.to_dict().items():
            set_user_setting(user_id, key, value)

def save_user_setting_to_firestore(setting_key, setting_value):
    """Firestoreにユーザー設定を保存"""
    try:
        db.collection('user_settings').document(user_id).set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定 '{setting_key}' の保存に失敗しました: {e}")

load_user_settings_from_firestore()

if 'creds' not in st.session_state:
    st.session_state.creds = authenticate_google()

if not st.session_state.creds:
    st.warning("Googleカレンダー認証を完了してください。")
    st.stop()
else:
    st.sidebar.success("✅ Googleカレンダーに認証済みです！")

@st.cache_resource(ttl=3600)
def get_calendar_service(creds):
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        editable_calendars = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') in ['owner', 'writer']
        }
        return service, editable_calendars
    except HttpError as e:
        st.error(f"カレンダーサービスの初期化に失敗しました: {e}")
        return None, None

@st.cache_resource(ttl=3600)
def get_tasks_service(creds):
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
        task_lists = tasks_service.tasklists().list().execute()
        default_list = next(
            (tl for tl in task_lists.get('items', []) if tl.get('title') == 'My Tasks'),
            task_lists.get('items', [{}])[0]
        )
        return tasks_service, default_list.get('id') if default_list else None
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
        return None, None

st.session_state.service, st.session_state.editable_calendar_options = get_calendar_service(st.session_state.creds)
st.session_state.tasks_service, st.session_state.default_task_list_id = get_tasks_service(st.session_state.creds)

if not st.session_state.service:
    st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
    st.stop()

if not st.session_state.tasks_service:
    st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")

tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの登録", "3. イベントの削除", "4. イベントの更新"])
selected_calendar_summary = st.sidebar.selectbox(
    "カレンダーを選択",
    list(st.session_state.editable_calendar_options.keys()),
    key="calendar_select"
)
calendar_id = st.session_state.editable_calendar_options[selected_calendar_summary]

with tabs[0]:
    st.header("1. Excelファイルのアップロード")
    uploaded_files = st.file_uploader(
        "Excelファイルをアップロードしてください",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    if uploaded_files:
        st.session_state.uploaded_files = uploaded_files
        try:
            merged_df = _load_and_merge_dataframes(uploaded_files)
            has_mng_data, has_name_data = check_event_name_columns(merged_df)
            
            st.session_state.merged_df = merged_df
            st.session_state.available_columns = merged_df.columns.tolist()
            st.session_state.has_mng_data = has_mng_data
            st.session_state.has_name_data = has_name_data
            
            st.success("ファイルの読み込みと統合が完了しました。")
            st.dataframe(merged_df.head())
        except Exception as e:
            st.error(f"ファイルの処理中にエラーが発生しました: {e}")
    else:
        st.session_state.uploaded_files = []
        st.session_state.pop("merged_df", None)

with tabs[1]:
    st.header("2. イベント登録設定")
    if 'merged_df' in st.session_state:
        df = st.session_state.merged_df
        available_columns = st.session_state.available_columns
        
        col1, col2 = st.columns(2)
        with col1:
            description_options = st.multiselect(
                "説明文に含める列を選択",
                options=available_columns,
                default=get_user_setting(user_id, 'description_columns_selected')
            )
            set_user_setting(user_id, 'description_columns_selected', description_options)
        
        with col2:
            event_name_options = ["選択しない"] + get_available_columns_for_event_name(df)
            event_name_col = st.selectbox(
                "イベント名に追加する列を選択",
                options=event_name_options,
                index=event_name_options.index(get_user_setting(user_id, 'event_name_col_selected')),
                help="『管理番号』と『物件名』に加えて、イベント名に含めたい列を選択できます。"
            )
            set_user_setting(user_id, 'event_name_col_selected', event_name_col)
        
        col_checkbox1, col_checkbox2 = st.columns(2)
        with col_checkbox1:
            all_day_event_override = st.checkbox(
                "全てのイベントを終日イベントとして登録する",
                value=st.session_state.get('all_day_event_override', False)
            )
            st.session_state.all_day_event_override = all_day_event_override
        with col_checkbox2:
            private_event = st.checkbox(
                "全てのイベントを非公開として登録する",
                value=st.session_state.get('private_event', False)
            )
            st.session_state.private_event = private_event

        add_task_type = st.checkbox(
            "イベント名に作業タイプを追加する（作業タイプ列がある場合）",
            value=get_user_setting(user_id, 'add_task_type_to_event_name')
        )
        set_user_setting(user_id, 'add_task_type_to_event_name', add_task_type)
        
        if st.button("イベントを登録", type="primary"):
            if not uploaded_files:
                st.warning("イベントを登録するには、まずファイルをアップロードしてください。")
            else:
                try:
                    df_to_add = process_excel_data_for_calendar(
                        st.session_state.uploaded_files,
                        description_columns=description_options,
                        all_day_event_override=all_day_event_override,
                        private_event=private_event,
                        fallback_event_name_column=event_name_col if event_name_col != "選択しない" else None,
                        add_task_type_to_event_name=add_task_type
                    )
                    st.session_state.df_to_add = df_to_add
                    st.success(f"{len(df_to_add)}件のイベントを登録準備完了しました。プレビューを確認し、問題なければ登録ボタンを押してください。")
                    
                    st.dataframe(df_to_add)

                    if st.button("上記の内容でカレンダーに登録"):
                        add_count = 0
                        progress_bar = st.progress(0)
                        
                        for i, row in df_to_add.iterrows():
                            add_event_to_calendar(st.session_state.service, calendar_id, row)
                            add_count += 1
                            progress_bar.progress((i + 1) / len(df_to_add))
                        
                        st.success(f"✅ {add_count} 件のイベントを登録しました。")
                except ValueError as e:
                    st.error(f"データの処理に失敗しました: {e}")
                except Exception as e:
                    st.error(f"予期せぬエラーが発生しました: {e}")
    else:
        st.info("まず「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")

with tabs[2]:
    st.header("3. イベント削除設定")
    delete_with_mng = st.checkbox("管理番号でイベントを削除", value=st.session_state.get('delete_with_mng', False))
    st.session_state.delete_with_mng = delete_with_mng
    
    if st.button("イベントを削除", type="primary"):
        if 'merged_df' not in st.session_state or 'uploaded_files' not in st.session_state:
            st.warning("ファイルをアップロードしてから削除してください。")
        else:
            try:
                df = st.session_state.merged_df
                if "管理番号" not in df.columns:
                    st.error("アップロードされたファイルに『管理番号』列が見つかりません。")
                else:
                    st.info("カレンダーからイベントを検索しています...")
                    worksheet_ids_to_delete = df['管理番号'].dropna().tolist()
                    
                    all_events = fetch_all_events(st.session_state.service, calendar_id, datetime.now(timezone.utc) - timedelta(days=90))
                    
                    events_to_delete = [
                        event for event in all_events
                        if any(
                            wid.lower() in event.get('summary', '').lower() for wid in worksheet_ids_to_delete
                        )
                    ]
                    
                    if not events_to_delete:
                        st.warning("削除対象のイベントが見つかりませんでした。")
                    else:
                        st.warning(f"以下の {len(events_to_delete)} 件のイベントを削除します。よろしいですか？")
                        for event in events_to_delete:
                            st.write(f"- {event['summary']} (開始日時: {event['start'].get('dateTime') or event['start'].get('date')})")

                        if st.button("はい、削除します"):
                            delete_count = 0
                            progress_bar_del = st.progress(0)
                            for i, event in enumerate(events_to_delete):
                                delete_event_from_calendar(st.session_state.service, calendar_id, event['id'])
                                delete_count += 1
                                progress_bar_del.progress((i + 1) / len(events_to_delete))
                            st.success(f"✅ {delete_count} 件のイベントを削除しました。")
            except Exception as e:
                st.error(f"イベント削除中にエラーが発生しました: {e}")

with tabs[3]:
    st.header("4. イベント更新設定")
    if 'merged_df' in st.session_state:
        df = st.session_state.merged_df
        available_columns_upd = st.session_state.available_columns
        
        col1_upd, col2_upd = st.columns(2)
        with col1_upd:
            description_options_upd = st.multiselect(
                "更新後の説明文に含める列を選択",
                options=available_columns_upd,
                default=get_user_setting(user_id, 'description_columns_selected')
            )
            set_user_setting(user_id, 'description_columns_selected_update', description_options_upd)
        
        with col2_upd:
            event_name_options_upd = ["選択しない"] + get_available_columns_for_event_name(df)
            event_name_col_upd = st.selectbox(
                "更新後のイベント名に追加する列を選択",
                options=event_name_options_upd,
                index=event_name_options_upd.index(get_user_setting(user_id, 'event_name_col_selected_update')),
                key="update_event_name_col",
                help="『管理番号』と『物件名』に加えて、イベント名に含めたい列を選択できます。"
            )
            set_user_setting(user_id, 'event_name_col_selected_update', event_name_col_upd)
        
        col_checkbox1_upd, col_checkbox2_upd = st.columns(2)
        with col_checkbox1_upd:
            all_day_event_override_upd = st.checkbox(
                "全てのイベントを終日イベントとして更新する",
                value=st.session_state.get('all_day_event_override_upd', False),
                key="update_all_day_event"
            )
            st.session_state.all_day_event_override_upd = all_day_event_override_upd
        with col_checkbox2_upd:
            private_event_upd = st.checkbox(
                "全てのイベントを非公開として更新する",
                value=st.session_state.get('private_event_upd', False),
                key="update_private_event"
            )
            st.session_state.private_event_upd = private_event_upd
            
        add_task_type_upd = st.checkbox(
            "イベント名に作業タイプを追加する（作業タイプ列がある場合）",
            value=get_user_setting(user_id, 'add_task_type_to_event_name_update'),
            key="update_add_task_type"
        )
        set_user_setting(user_id, 'add_task_type_to_event_name_update', add_task_type_upd)

        calendar_id_upd = st.sidebar.selectbox(
            "更新対象カレンダーを選択",
            list(st.session_state.editable_calendar_options.keys()),
            key="update_calendar_select"
        )
        calendar_id_upd = st.session_state.editable_calendar_options[calendar_id_upd]
        
        if st.button("イベントを更新", type="primary"):
            if 'merged_df' not in st.session_state:
                st.warning("ファイルをアップロードしてから更新してください。")
            elif "管理番号" not in st.session_state.merged_df.columns:
                st.error("アップロードされたファイルに『管理番号』列が見つかりません。")
            else:
                try:
                    df = process_excel_data_for_calendar(
                        st.session_state.uploaded_files,
                        description_columns=description_options_upd,
                        all_day_event_override=all_day_event_override_upd,
                        private_event=private_event_upd,
                        fallback_event_name_column=event_name_col_upd if event_name_col_upd != "選択しない" else None,
                        add_task_type_to_event_name=add_task_type_upd
                    )
                    
                    st.info("カレンダーからイベントを検索しています...")
                    service = st.session_state.service
                    all_events = fetch_all_events(service, calendar_id_upd, datetime.now(timezone.utc) - timedelta(days=90))
                    
                    df['worksheet_id'] = df['Description'].str.extract(r'作業指示書:\s*([^\n/]+)')
                    df = df[df['worksheet_id'].notna()]
                    
                    update_count = 0
                    progress_bar = st.progress(0)
                    
                    for i, row in df.iterrows():
                        worksheet_id = row['worksheet_id']
                        matched_event = next(
                            (event for event in all_events if worksheet_id.lower() in event.get('summary', '').lower()),
                            None
                        )

                        if matched_event:
                            start_dt_str = f"{row['Start Date']}T{row['Start Time']}"
                            end_dt_str = f"{row['End Date']}T{row['End Time']}"
                            start_dt_obj = datetime.strptime(start_dt_str, "%Y/%m/%dT%H:%M:%S").astimezone(timezone(timedelta(hours=9)))
                            end_dt_obj = datetime.strptime(end_dt_str, "%Y/%m/%dT%H:%M:%S").astimezone(timezone(timedelta(hours=9)))
                            
                            event_data = {
                                'summary': row['Subject'],
                                'description': row['Description'],
                                'location': row['Location'],
                            }

                            if row['All Day Event'] == "True":
                                event_data['start'] = {'date': start_dt_obj.strftime("%Y-%m-%d")}
                                event_data['end'] = {'date': end_dt_obj.strftime("%Y-%m-%d")}
                            else:
                                event_data['start'] = {'dateTime': start_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
                                event_data['end'] = {'dateTime': end_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

                            try:
                                if update_event_if_needed(service, calendar_id_upd, matched_event['id'], event_data):
                                    update_count += 1
                            except Exception as e:
                                st.error(f"イベント '{row['Subject']}' (作業指示書: {worksheet_id}) の更新に失敗しました: {e}")
                            
                            progress_bar.progress((i + 1) / len(df))

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")

                except Exception as e:
                    st.error(f"イベント更新中にエラーが発生しました: {e}")
    else:
        st.info("まず「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")

with st.sidebar:
    st.header("🔐 認証状態")
    if user_id:
        st.success("✅ Firebase認証済み")
    else:
        st.warning("⚠️ Firebase認証が未完了です")

    if st.session_state.get('service'):
        st.success("✅ Googleカレンダー認証済み")
    else:
        st.warning("⚠️ Googleカレンダー認証が未完了です")
    
    if st.session_state.get('tasks_service'):
        st.success("✅ ToDoリスト利用可能")
    else:
        st.warning("⚠️ ToDoリスト利用不可")
    
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    st.metric("アップロード済みファイル", uploaded_count)
    
    if st.button("🚪 ログアウト", type="secondary"):
        if user_id:
            clear_user_settings(user_id)
        st.session_state.clear()
        st.rerun()
