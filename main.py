import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
from excel_parser import process_excel_files
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    delete_events_from_calendar,
    fetch_all_events,
    update_event_if_needed
)
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

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

tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

with tabs[0]:
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

with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードすると、イベント登録機能が利用可能になります。")
    else:
        st.subheader("📝 イベント設定")
        all_day_event = st.checkbox("終日イベントとして登録", value=False)
        private_event = st.checkbox("非公開イベントとして登録", value=True)
        
        # ToDoオプションを追加
        include_todo = st.checkbox("作業ToDoリストを追加", value=True, help="☐ 点検通知（FAX）, ☐ 点検通知（電話）, ☐ 貼紙")

        description_columns = st.multiselect(
            "説明欄に含める列（複数選択可）",
            st.session_state.get('description_columns_pool', [])
        )

        if not st.session_state['editable_calendar_options']:
            st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        else:
            selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
            calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

            st.subheader("➡️ イベント登録")
            if st.button("Googleカレンダーに登録する"):
                with st.spinner("イベントデータを処理中..."):
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event, include_todo)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                    else:
                        st.info(f"{len(df)} 件のイベントを登録します。")
                        progress = st.progress(0)
                        successful_registrations = 0
                        for i, row in df.iterrows():
                            try:
                                if row['All Day Event'] == "True":
                                    start_date_str = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                                    end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d") + timedelta(days=1)
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
                                add_event_to_calendar(service, calendar_id, event_data)
                                successful_registrations += 1
                            except Exception as e:
                                st.error(f"{row['Subject']} の登録に失敗しました: {e}")
                            progress.progress((i + 1) / len(df))

                        st.success(f"✅ {successful_registrations} 件のイベント登録が完了しました！")

with tabs[2]:
    st.header("イベントを削除")

    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today = date.today()
        delete_start_date = st.date_input("削除開始日", value=today - timedelta(days=30))
        delete_end_date = st.date_input("削除終了日", value=today)

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")
            if st.button("選択期間のイベントを削除する"):
                deleted_count = delete_events_from_calendar(
                    service, calendar_id_del,
                    datetime.combine(delete_start_date, datetime.min.time()),
                    datetime.combine(delete_end_date, datetime.max.time())
                )
                if deleted_count > 0:
                    st.success(f"{deleted_count} 件のイベントが削除されました。")
                else:
                    st.info("指定期間内に削除するイベントはありませんでした。")

with tabs[3]:
    st.header("イベントを更新")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        all_day_event = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")
        
        # 更新タブにもToDoオプションを追加
        include_todo_update = st.checkbox("作業ToDoリストを追加", value=True, key="update_todo", help="☐ 点検通知（FAX）, ☐ 点検通知（電話）, ☐ 貼紙")
        
        description_columns = st.multiselect("説明欄に含める列", st.session_state['description_columns_pool'], key="update_desc_cols")

        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name_upd = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
            calendar_id_upd = st.session_state['editable_calendar_options'][selected_calendar_name_upd]

            if st.button("イベントを照合・更新"):
                with st.spinner("イベントを処理中..."):
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event, include_todo_update)
                    if df.empty:
                        st.warning("有効なイベントデータがありません。")
                        st.stop()

                    today = datetime.now()
                    time_min = (today - timedelta(days=180)).isoformat() + 'Z'
                    time_max = (today + timedelta(days=180)).isoformat() + 'Z'
                    events = fetch_all_events(service, calendar_id_upd, time_min, time_max)

                    worksheet_to_event = {}
                    for event in events:
                        desc = event.get('description', '')
                        match = re.search(r"作業指示書：(\d+)", desc)
                        if match:
                            worksheet_to_event[match.group(1)] = event

                    update_count = 0
                    for i, row in df.iterrows():
                        match = re.search(r"作業指示書：(\d+)", row['Description'])
                        if not match:
                            continue
                        worksheet_id = match.group(1)
                        matched_event = worksheet_to_event.get(worksheet_id)
                        if not matched_event:
                            continue

                        if row['All Day Event'] == "True":
                            start_date = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                            end_date = (datetime.strptime(row['End Date'], "%Y/%m/%d") + timedelta(days=1)).strftime("%Y-%m-%d")
                            event_data = {
                                'start': {'date': start_date},
                                'end': {'date': end_date}
                            }
                        else:
                            start_dt = f"{row['Start Date']} {row['Start Time']}"
                            end_dt = f"{row['End Date']} {row['End Time']}"
                            event_data = {
                                'start': {'dateTime': datetime.strptime(start_dt, "%Y/%m/%d %H:%M").isoformat(), 'timeZone': 'Asia/Tokyo'},
                                'end': {'dateTime': datetime.strptime(end_dt, "%Y/%m/%d %H:%M").isoformat(), 'timeZone': 'Asia/Tokyo'}
                            }

                        try:
                            if update_event_if_needed(service, calendar_id_upd, matched_event, event_data):
                                update_count += 1
                        except Exception as e:
                            st.error(f"{row['Subject']} の更新に失敗: {e}")

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")

                    st.success(f"✅ {update_count} 件のイベントを更新しました。")


