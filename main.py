import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
import re
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
    find_and_delete_tasks_by_event_id
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from session_utils import (
    initialize_session_state,
    get_user_setting,
    set_user_setting,
    get_all_user_settings,
    clear_user_settings
)
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from firebase_admin import firestore
import os
from pathlib import Path
from io import BytesIO

# 【追加】JSTタイムゾーンを定義
jst = timezone(timedelta(hours=9))

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    # ユーザーが認証されていない場合はフォームを表示
    firebase_auth_form(db)
    st.stop()
    
initialize_session_state()

# Google認証
creds = authenticate_google()
if creds:
    try:
        service = build('calendar', 'v3', credentials=creds)
        st.session_state['calendar_service'] = service
        st.session_state['tasks_service'] = build_tasks_service(creds)
    except Exception as e:
        st.error(f"カレンダーサービスのビルドに失敗しました: {e}")
        st.session_state['calendar_service'] = None
else:
    st.warning("Googleカレンダーにログインしてください。")
    st.stop()

# ==============================
# Streamlit UI
# ==============================

CALENDAR_ID = get_user_setting(user_id, 'calendar_id')
TASK_LIST_ID = get_user_setting(user_id, 'task_list_id')

tabs = st.tabs([
    "1. 設定", 
    "2. イベントの登録", 
    "3. イベントの削除", 
    "4. イベントの更新", 
    "5. カレンダーイベントの確認", 
    "6. 設定のリセット"
])

# ------------------------------
# 1. 設定
# ------------------------------
with tabs[0]:
    st.header("🗓️ カレンダー設定")
    
    # 利用可能なカレンダーの取得
    if st.session_state.get('calendar_service'):
        try:
            calendar_list = st.session_state['calendar_service'].calendarList().list().execute()
            calendars = calendar_list.get('items', [])
            calendar_map = {c.get('summary'): c.get('id') for c in calendars}
            calendar_names = list(calendar_map.keys())
            
            if calendar_names:
                default_calendar_name = next((name for name, id in calendar_map.items() if id == CALENDAR_ID), calendar_names[0])
                
                selected_calendar_name = st.selectbox(
                    "利用するカレンダーを選択してください:", 
                    calendar_names,
                    index=calendar_names.index(default_calendar_name) if default_calendar_name in calendar_names else 0
                )
                
                new_calendar_id = calendar_map[selected_calendar_name]
                if new_calendar_id != CALENDAR_ID:
                    set_user_setting(user_id, 'calendar_id', new_calendar_id)
                    st.success(f"カレンダーIDを {new_calendar_id} に設定しました。")
                    CALENDAR_ID = new_calendar_id
            else:
                st.warning("利用可能なカレンダーが見つかりませんでした。")
                CALENDAR_ID = None
        except HttpError as e:
            st.error(f"カレンダーリストの取得中にエラーが発生しました: {e}")
            CALENDAR_ID = None
    
    st.markdown("---")
    st.header("📝 ToDoリスト設定")
    if st.session_state.get('tasks_service'):
        try:
            task_lists_result = st.session_state['tasks_service'].tasklists().list().execute()
            task_lists = task_lists_result.get('items', [])
            task_list_map = {tl.get('title'): tl.get('id') for tl in task_lists}
            task_list_titles = list(task_list_map.keys())
            
            if task_list_titles:
                default_list_title = next((title for title, id in task_list_map.items() if id == TASK_LIST_ID), task_list_titles[0])
                
                selected_list_title = st.selectbox(
                    "利用するToDoリストを選択してください:", 
                    task_list_titles,
                    index=task_list_titles.index(default_list_title) if default_list_title in task_list_titles else 0
                )
                
                new_task_list_id = task_list_map[selected_list_title]
                if new_task_list_id != TASK_LIST_ID:
                    set_user_setting(user_id, 'task_list_id', new_task_list_id)
                    st.success(f"ToDoリストIDを {new_task_list_id} に設定しました。")
                    TASK_LIST_ID = new_task_list_id
            else:
                st.warning("利用可能なToDoリストが見つかりませんでした。")
                TASK_LIST_ID = None
        except HttpError as e:
            st.error(f"ToDoリストの取得中にエラーが発生しました: {e}")
            TASK_LIST_ID = None


# ------------------------------
# 2. イベントの登録
# ------------------------------
with tabs[1]:
    st.header("📤 Excelファイルからイベントを登録")
    st.caption("既存のイベントと重複する場合、作業指示書IDをキーに**更新**されます。")
    
    uploaded_files_reg = st.file_uploader(
        "Excelファイルをアップロードしてください", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True,
        key="reg_uploader"
    )

    if uploaded_files_reg and CALENDAR_ID:
        try:
            merged_df = _load_and_merge_dataframes(uploaded_files_reg)
            
            col1, col2 = st.columns(2)
            
            # イベント名に含める列の選択
            available_cols = get_available_columns_for_event_name(merged_df.columns)
            default_cols = [c for c in ['Subject', '担当者'] if c in available_cols]
            
            selected_name_columns = col1.multiselect(
                "イベント名に含める列を選択 (左側の列が優先されます):",
                available_cols,
                default=default_cols
            )

            # 説明欄に含める列の選択
            available_desc_cols = [c for c in merged_df.columns if c not in selected_name_columns]
            default_desc_cols = [c for c in ['作業内容', '備考'] if c in available_desc_cols]

            selected_desc_columns = col2.multiselect(
                "説明欄に含める列を選択:",
                available_desc_cols,
                default=default_desc_cols
            )

            if selected_name_columns:
                st.subheader("🗓️ 登録内容プレビュー")
                
                # イベントデータ整形
                df_filtered = process_excel_data_for_calendar(merged_df, selected_name_columns, selected_desc_columns)
                st.dataframe(df_filtered)
                
                st.markdown("---")
                
                # 既存イベントの取得 (更新/重複チェック用)
                st.subheader("重複チェック用の既存イベントの取得")
                events = fetch_all_events(service, CALENDAR_ID)
                
                # 作業指示書IDとイベントをマッピング
                worksheet_to_event = {}
                for event in events:
                    desc = event.get('description', '')
                    match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                    if match:
                        worksheet_id = match.group(1)
                        worksheet_to_event[worksheet_id] = event

                st.info(f"既存カレンダーイベントから {len(worksheet_to_event)} 件の作業指示書IDを抽出しました。")

                if st.button("🚀 イベント登録/更新を実行", type="primary"):
                    st.session_state['registration_in_progress'] = True
                    progress_bar = st.progress(0, text="登録/更新中...")
                    
                    added_count = 0
                    updated_count = 0
                    task_added_count = 0

                    for i, row in df_filtered.iterrows():
                        # イベントデータ構築
                        event_data = {
                            'summary': row['Subject'],
                            'location': row.get('Location', ''),
                            'description': row.get('Description', ''),
                            'reminders': {
                                'useDefault': False,
                                'overrides': [
                                    {'method': 'popup', 'minutes': 30},
                                ],
                            }
                        }
                        
                        # 日時情報
                        if row['All Day Event'] == "True":
                            event_data['start'] = {'date': row['Start Date']}
                            event_data['end'] = {'date': row['End Date']}
                        else:
                            start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                            end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")

                            # 【修正箇所 1/2】JSTタイムゾーンを付与し、isoformat()でタイムゾーン付き文字列を生成
                            start_dt_jst = start_dt_obj.replace(tzinfo=jst)
                            end_dt_jst = end_dt_obj.replace(tzinfo=jst)
                            
                            event_data['start'] = {'dateTime': start_dt_jst.isoformat(), 'timeZone': 'Asia/Tokyo'}
                            event_data['end'] = {'dateTime': end_dt_jst.isoformat(), 'timeZone': 'Asia/Tokyo'}

                        # 作業指示書IDの抽出 (更新チェック用)
                        worksheet_id = None
                        match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
                        if match:
                            worksheet_id = match.group(1)
                        
                        # イベントの重複チェックと更新/登録
                        matched_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None
                        
                        if matched_event:
                            # 既存イベントを更新
                            updated_event = update_event_if_needed(
                                service, 
                                CALENDAR_ID, 
                                matched_event['id'], 
                                event_data
                            )
                            if updated_event and updated_event['id'] != matched_event['id']: # idが同じなら更新済み、異なるとは考えにくいが念のため
                                updated_count += 1
                                st.code(f"✅ 更新: {row['Subject']} (ID: {matched_event['id']})")
                            elif updated_event:
                                pass # 更新不要
                            else:
                                st.error(f"❌ 更新失敗: {row['Subject']}")
                        else:
                            # 新規イベントを登録
                            new_event = add_event_to_calendar(service, CALENDAR_ID, event_data)
                            if new_event:
                                added_count += 1
                                st.code(f"➕ 登録: {row['Subject']} (ID: {new_event['id']})")
                                
                                # ToDoリストにも追加
                                if TASK_LIST_ID and st.session_state.get('tasks_service'):
                                    task_data = {
                                        'title': f"[カレンダー] {row['Subject']}",
                                        'notes': f"イベントID: {new_event['id']}\n作業指示書: {worksheet_id if worksheet_id else 'N/A'}\n場所: {row.get('Location', '')}",
                                        'due': (start_dt_jst.isoformat() if row['All Day Event'] == "False" else None),
                                        'status': 'needsAction'
                                    }
                                    if add_task_to_todo_list(st.session_state['tasks_service'], TASK_LIST_ID, task_data):
                                        task_added_count += 1
                                
                            else:
                                st.error(f"❌ 登録失敗: {row['Subject']}")

                        progress_bar.progress((i + 1) / len(df_filtered))
                    
                    progress_bar.empty()
                    st.success(f"イベント登録/更新が完了しました！")
                    st.info(f"新規登録: {added_count} 件, 更新: {updated_count} 件")
                    if TASK_LIST_ID:
                        st.info(f"ToDoリストにタスク追加: {task_added_count} 件")
                    st.session_state['registration_in_progress'] = False

            else:
                st.warning("イベント名に含める列を選択してください。")
        
        except Exception as e:
            st.error(f"イベントデータの処理中にエラーが発生しました: {e}")
            st.session_state['registration_in_progress'] = False

# ------------------------------
# 3. イベントの削除
# ------------------------------
with tabs[2]:
    st.header("🗑️ イベントの削除")
    st.caption("Excelファイルで指定された作業指示書IDに対応するイベントを削除します。")
    
    uploaded_files_del = st.file_uploader(
        "削除対象の作業指示書IDを含むExcelファイルをアップロードしてください", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True,
        key="del_uploader"
    )

    if uploaded_files_del and CALENDAR_ID:
        try:
            merged_df = _load_and_merge_dataframes(uploaded_files_del)
            
            # 作業指示書IDの列を特定
            worksheet_col = merged_df.columns[0] # ここは適切なロジックが必要（例: excel_parser.pyのfind_closest_column）
            
            # 既存イベントの取得
            st.subheader("重複チェック用の既存イベントの取得")
            events = fetch_all_events(service, CALENDAR_ID)
            
            worksheet_to_event = {}
            for event in events:
                desc = event.get('description', '')
                match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                if match:
                    worksheet_id = match.group(1)
                    worksheet_to_event[worksheet_id] = event

            st.info(f"既存カレンダーイベントから {len(worksheet_to_event)} 件の作業指示書IDを抽出しました。")
            
            # Excelから削除対象の作業指示書IDリストを抽出
            worksheet_ids_to_delete = set()
            for index, row in merged_df.iterrows():
                # ここは、Excelから作業指示書IDを抽出する適切なロジックに置き換えてください
                # 仮に、最初の列を作業指示書IDとして扱う
                ws_id = format_worksheet_value(row[worksheet_col]) 
                if ws_id and ws_id.isdigit():
                    worksheet_ids_to_delete.add(ws_id)
            
            st.warning(f"Excelから {len(worksheet_ids_to_delete)} 件の削除対象作業指示書IDを抽出しました。")

            if st.button("💥 イベント削除を実行", type="primary"):
                progress_bar = st.progress(0, text="削除中...")
                deleted_count = 0
                
                for i, ws_id in enumerate(list(worksheet_ids_to_delete)):
                    if ws_id in worksheet_to_event:
                        event_to_delete = worksheet_to_event[ws_id]
                        event_id = event_to_delete['id']
                        
                        # イベントを削除
                        try:
                            service.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute()
                            st.code(f"🗑️ 削除成功: 作業指示書ID {ws_id} (イベント: {event_to_delete.get('summary', 'N/A')})")
                            deleted_count += 1
                            
                            # ToDoリストからも削除
                            if TASK_LIST_ID and st.session_state.get('tasks_service'):
                                find_and_delete_tasks_by_event_id(st.session_state['tasks_service'], TASK_LIST_ID, event_id)

                        except HttpError as e:
                            st.error(f"❌ 削除失敗 (HTTPエラー): 作業指示書ID {ws_id}, イベントID {event_id}: {e}")
                        except Exception as e:
                            st.error(f"❌ 削除失敗: 作業指示書ID {ws_id}, イベントID {event_id}: {e}")
                            
                    progress_bar.progress((i + 1) / len(worksheet_ids_to_delete))

                progress_bar.empty()
                st.success(f"イベント削除が完了しました！ 削除件数: {deleted_count} 件")
        
        except Exception as e:
            st.error(f"削除データの処理中にエラーが発生しました: {e}")

# ------------------------------
# 4. イベントの更新
# ------------------------------
with tabs[3]:
    st.header("🔄 Excelファイルから既存イベントを更新")
    st.caption("作業指示書IDをキーに、既存のカレンダーイベントを更新します。")
    
    uploaded_files_upd = st.file_uploader(
        "更新情報を含むExcelファイルをアップロードしてください", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True,
        key="upd_uploader"
    )

    if uploaded_files_upd and CALENDAR_ID:
        try:
            merged_df = _load_and_merge_dataframes(uploaded_files_upd)
            
            col1, col2 = st.columns(2)
            
            available_cols = get_available_columns_for_event_name(merged_df.columns)
            default_cols = [c for c in ['Subject', '担当者'] if c in available_cols]
            
            selected_name_columns = col1.multiselect(
                "イベント名に含める列を選択 (左側の列が優先されます):",
                available_cols,
                default=default_cols,
                key="upd_name_cols"
            )

            available_desc_cols = [c for c in merged_df.columns if c not in selected_name_columns]
            default_desc_cols = [c for c in ['作業内容', '備考'] if c in available_desc_cols]

            selected_desc_columns = col2.multiselect(
                "説明欄に含める列を選択:",
                available_desc_cols,
                default=default_desc_cols,
                key="upd_desc_cols"
            )

            if selected_name_columns:
                st.subheader("🗓️ 更新内容プレビュー")
                
                df_filtered = process_excel_data_for_calendar(merged_df, selected_name_columns, selected_desc_columns)
                st.dataframe(df_filtered)
                
                st.markdown("---")
                
                # 既存イベントの取得 (更新対象チェック用)
                st.subheader("更新対象チェック用の既存イベントの取得")
                events = fetch_all_events(service, CALENDAR_ID)
                
                worksheet_to_event = {}
                for event in events:
                    desc = event.get('description', '')
                    match = re.search(r"作業指示書[：:]\s*(\d+)", desc)
                    if match:
                        worksheet_id = match.group(1)
                        worksheet_to_event[worksheet_id] = event

                st.info(f"既存カレンダーイベントから {len(worksheet_to_event)} 件の作業指示書IDを抽出しました。")

                if st.button("🔄 イベント更新を実行", type="primary"):
                    st.session_state['update_in_progress'] = True
                    progress_bar = st.progress(0, text="更新中...")
                    
                    updated_count = 0
                    skipped_count = 0

                    for i, row in df_filtered.iterrows():
                        # イベントデータ構築 (登録時と同じロジック)
                        event_data = {
                            'summary': row['Subject'],
                            'location': row.get('Location', ''),
                            'description': row.get('Description', ''),
                            'reminders': {
                                'useDefault': False,
                                'overrides': [
                                    {'method': 'popup', 'minutes': 30},
                                ],
                            }
                        }
                        
                        # 日時情報
                        if row['All Day Event'] == "True":
                            event_data['start'] = {'date': row['Start Date']}
                            event_data['end'] = {'date': row['End Date']}
                        else:
                            start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                            end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")

                            # 【修正箇所 2/2】JSTタイムゾーンを付与し、isoformat()でタイムゾーン付き文字列を生成
                            start_dt_jst = start_dt_obj.replace(tzinfo=jst)
                            end_dt_jst = end_dt_obj.replace(tzinfo=jst)
                            
                            event_data['start'] = {'dateTime': start_dt_jst.isoformat(), 'timeZone': 'Asia/Tokyo'}
                            event_data['end'] = {'dateTime': end_dt_jst.isoformat(), 'timeZone': 'Asia/Tokyo'}

                        # 作業指示書IDの抽出
                        worksheet_id = None
                        match = re.search(r"作業指示書[：:]\s*(\d+)", row['Description'])
                        if match:
                            worksheet_id = match.group(1)
                        
                        # イベントの照合
                        matched_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None
                        
                        if not matched_event:
                            skipped_count += 1
                            # st.warning(f"スキップ: 作業指示書ID {worksheet_id} の既存イベントが見つかりませんでした。")
                            progress_bar.progress((i + 1) / len(df_filtered))
                            continue # 更新対象が見つからない場合はスキップ
                        
                        # 既存イベントを更新
                        updated_event = update_event_if_needed(
                            service, 
                            CALENDAR_ID, 
                            matched_event['id'], 
                            event_data
                        )
                        
                        if updated_event and updated_event['id'] != matched_event['id']:
                            updated_count += 1
                            st.code(f"✅ 更新: {row['Subject']} (ID: {matched_event['id']})")
                        elif updated_event:
                            pass # 更新不要
                        else:
                            st.error(f"❌ 更新失敗: {row['Subject']}")

                        progress_bar.progress((i + 1) / len(df_filtered))
                    
                    progress_bar.empty()
                    st.success(f"イベント更新が完了しました！")
                    st.info(f"更新: {updated_count} 件, スキップ (未発見): {skipped_count} 件")
                    st.session_state['update_in_progress'] = False

            else:
                st.warning("イベント名に含める列を選択してください。")
        
        except Exception as e:
            st.error(f"イベントデータの処理中にエラーが発生しました: {e}")
            st.session_state['update_in_progress'] = False

# ------------------------------
# 5. カレンダーイベントの確認
# ------------------------------
with tabs[4]:
    st.header("🔎 カレンダーイベントの確認とダウンロード")
    
    if CALENDAR_ID:
        st.subheader(f"カレンダーID: {CALENDAR_ID}")
        
        # 期間選択
        today = date.today()
        default_start = today - timedelta(days=30)
        default_end = today + timedelta(days=90)
        
        start_date = st.date_input("開始日", value=default_start)
        end_date = st.date_input("終了日", value=default_end)
        
        if st.button("イベントを読み込む", type="primary"):
            with st.spinner("イベントを読み込み中..."):
                try:
                    # タイムゾーンを考慮してdatetimeに変換
                    time_min = datetime.combine(start_date, datetime.min.time(), tzinfo=jst).isoformat()
                    time_max = datetime.combine(end_date, datetime.max.time(), tzinfo=jst).isoformat()
                    
                    events = service.events().list(
                        calendarId=CALENDAR_ID, 
                        timeMin=time_min,
                        timeMax=time_max,
                        singleEvents=True,
                        orderBy='startTime'
                    ).execute().get('items', [])
                    
                    if not events:
                        st.info("指定された期間にイベントは見つかりませんでした。")
                        
                    output_records = []
                    for event in events:
                        event_id = event.get('id')
                        summary = event.get('summary', 'タイトルなし')
                        location = event.get('location', '')
                        description = event.get('description', '')
                        
                        # 日時情報の処理
                        start = event['start']
                        end = event['end']
                        
                        is_all_day = 'date' in start
                        
                        if is_all_day:
                            start_str = start.get('date')
                            end_str = end.get('date')
                            start_dt = datetime.strptime(start_str, "%Y-%m-%d").date()
                            # 終日イベントの終了日は翌日の日付が格納されているため、1日戻す
                            end_dt = datetime.strptime(end_str, "%Y-%m-%d").date() - timedelta(days=1)
                            
                            start_date_display = start_dt.strftime("%Y/%m/%d")
                            end_date_display = end_dt.strftime("%Y/%m/%d")
                            start_time_display = ""
                            end_time_display = ""
                        else:
                            start_dt_str = start.get('dateTime')
                            end_dt_str = end.get('dateTime')
                            
                            # タイムゾーン情報を含めてパース
                            start_dt = datetime.fromisoformat(start_dt_str).astimezone(jst)
                            end_dt = datetime.fromisoformat(end_dt_str).astimezone(jst)
                            
                            start_date_display = start_dt.strftime("%Y/%m/%d")
                            end_date_display = end_dt.strftime("%Y/%m/%d")
                            start_time_display = start_dt.strftime("%H:%M")
                            end_time_display = end_dt.strftime("%H:%M")
                            
                        # 作業指示書IDの抽出
                        worksheet_id = None
                        match = re.search(r"作業指示書[：:]\s*(\d+)", description)
                        if match:
                            worksheet_id = match.group(1)

                        output_records.append({
                            "ID": event_id,
                            "作業指示書ID": worksheet_id,
                            "Subject": summary,
                            "Start Date": start_date_display,
                            "Start Time": start_time_display,
                            "End Date": end_date_display,
                            "End Time": end_time_display,
                            "All Day Event": "True" if is_all_day else "False",
                            "Location": location,
                            "Description": description
                        })
                    
                    if output_records:
                        output_df = pd.DataFrame(output_records)
                        st.dataframe(output_df)

                        # ダウンロードボタン
                        buffer = BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            output_df.to_excel(writer, sheet_name='CalendarEvents', index=False)
                        buffer.seek(0)
                        
                        st.download_button(
                            label="✅ Excelファイルとしてダウンロード",
                            data=buffer,
                            file_name="Googleカレンダー_イベントリスト.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"{len(output_df)} 件のイベントを読み込みました。")
                
                except Exception as e:
                    st.error(f"イベントの読み込み中にエラーが発生しました: {e}")
                    
# ------------------------------
# 6. 設定のリセット
# ------------------------------
with tabs[5]:
    st.header("⚙️ ユーザー設定と認証のリセット")
    
    st.warning("この操作を行うと、保存されているカレンダーID、ToDoリストID、Google認証情報がすべて削除されます。再利用には再度認証が必要です。")
    
    if st.button("全てのユーザー設定と認証をリセット", type="secondary"):
        if user_id:
            clear_user_settings(user_id)
        
        # Streamlitセッションステートもクリア
        keys_to_delete = [
            'credentials', 
            'calendar_service', 
            'tasks_service', 
            'calendar_id', 
            'task_list_id',
            'registration_in_progress',
            'update_in_progress'
        ]
        for key in keys_to_delete:
            if key in st.session_state:
                del st.session_state[key]
        
        st.success("全てのユーザー設定と認証情報がリセットされました。ページをリロードしてください。")
        st.experimental_rerun()


# ==============================
# サイドバー
# ==============================
with st.sidebar:
    st.header("🔐 認証状態")
    # Firebase認証の状態表示 (get_firebase_user_id()がNoneでなければ成功)
    if get_firebase_user_id():
        st.success("✅ Firebase認証済み")
    else:
        st.warning("⚠️ Firebase認証が未完了です")
    
    if st.session_state.get('calendar_service'):
        st.success("✅ Googleカレンダー認証済み")
    else:
        st.warning("⚠️ Googleカレンダー認証が未完了です")
    
    if st.session_state.get('tasks_service'):
        st.success("✅ ToDoリスト利用可能")
    else:
        st.warning("⚠️ ToDoリスト利用不可")
    
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    # st.metric("アップロード済みファイル", uploaded_count) # この部分は元のコードに合わせます
    
    if st.button("🚪 ログアウト", type="secondary"):
        if user_id:
            clear_user_settings(user_id)
        for key in list(st.session_state.keys()):
            if not key.startswith("google_auth") and not key.startswith("firebase_"):
                del st.session_state[key]
        st.experimental_rerun()
