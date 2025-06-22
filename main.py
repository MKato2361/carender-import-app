
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_from_calendar
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー一括管理ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除・更新")

# Google認証
st.subheader("🔐 Google認証")
creds = authenticate_google()
if not creds:
    st.warning("Google認証を完了してください。")
    st.stop()

if 'calendar_service' not in st.session_state:
    service = build("calendar", "v3", credentials=creds)
    st.session_state['calendar_service'] = service
    calendar_list = service.calendarList().list().execute()
    st.session_state['editable_calendar_options'] = {
        cal['summary']: cal['id']
        for cal in calendar_list['items']
        if cal.get('accessRole') != 'reader'
    }
else:
    service = st.session_state['calendar_service']

tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの登録", "3. イベントの削除", "4. イベントの更新"])

# アップロードタブ
with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)
    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        desc_cols = set()
        for file in uploaded_files:
            try:
                df = pd.read_excel(file, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                desc_cols.update(df.columns)
            except Exception as e:
                st.warning(f"{file.name} の読み込みに失敗: {e}")
        st.session_state['description_columns_pool'] = list(desc_cols)
    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")

# 登録タブ
with tabs[1]:
    st.header("イベントを登録")
    if not st.session_state.get('uploaded_files'):
        st.info("先にファイルをアップロードしてください。")
        st.stop()
    all_day_event = st.checkbox("終日イベントとして登録", value=False)
    private_event = st.checkbox("非公開イベントとして登録", value=True)
    description_columns = st.multiselect("説明欄に含める列", st.session_state.get('description_columns_pool', []))
    calendar_id = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar")
    calendar_id = st.session_state['editable_calendar_options'][calendar_id]

    if st.button("Googleカレンダーに登録する"):
        df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)
        for i, row in df.iterrows():
            try:
                if row["All Day Event"] == "True":
                    start = datetime.strptime(row["Start Date"], "%Y/%m/%d").date()
                    end = datetime.strptime(row["End Date"], "%Y/%m/%d").date() + timedelta(days=1)
                    event_data = {
                        "summary": row["Subject"],
                        "location": row["Location"],
                        "description": row["Description"],
                        "start": {"date": start.isoformat()},
                        "end": {"date": end.isoformat()},
                        "transparency": "transparent" if row["Private"] == "True" else "opaque"
                    }
                else:
                    start = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                    end = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                    event_data = {
                        "summary": row["Subject"],
                        "location": row["Location"],
                        "description": row["Description"],
                        "start": {"dateTime": start.isoformat(), "timeZone": "Asia/Tokyo"},
                        "end": {"dateTime": end.isoformat(), "timeZone": "Asia/Tokyo"},
                        "transparency": "transparent" if row["Private"] == "True" else "opaque"
                    }
                add_event_to_calendar(service, calendar_id, event_data)
            except Exception as e:
                st.error(f"{row['Subject']} の登録に失敗: {e}")
        st.success("イベントの登録が完了しました。")

# 削除タブ
with tabs[2]:
    st.header("イベントを削除")
    calendar_id = st.selectbox("削除対象カレンダー", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar")
    calendar_id = st.session_state['editable_calendar_options'][calendar_id]
    start_date = st.date_input("削除開始日", value=date.today() - timedelta(days=30))
    end_date = st.date_input("削除終了日", value=date.today())

    if start_date > end_date:
        st.error("日付の範囲が正しくありません。")
    elif st.button("イベント削除を実行"):
        count = delete_events_from_calendar(
            service,
            calendar_id,
            datetime.combine(start_date, datetime.min.time()),
            datetime.combine(end_date, datetime.max.time())
        )
        st.success(f"{count} 件のイベントを削除しました。")

# 更新タブ（省略していたコード再挿入済み）

# main.py（更新タブを含む全体構成の一部として）
# これは追記分のみの例で、他のコードは既存のままです

with st.tabs(["1. ファイルのアップロード", "2. イベントの登録", "3. イベントの削除", "4. イベントの更新"])[3]:
    st.header("📤 イベントを更新")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        st.stop()

    all_day_event = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day_event")
    private_event = st.checkbox("非公開イベントとして扱う", value=True, key="update_private_event")
    description_columns = st.multiselect("説明欄に含める列（複数選択可）", st.session_state['description_columns_pool'], key="update_description_columns")

    calendar_id_update = st.selectbox("更新対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="update_calendar_select")
    calendar_id_update = st.session_state['editable_calendar_options'][calendar_id_update]

    df_new = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)

    st.subheader("📅 カレンダーイベントと突合・差分検出中...")

    existing_events = []
    page_token = None
    while True:
        events_result = service.events().list(
            calendarId=calendar_id_update,
            maxResults=2500,
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token
        ).execute()
        existing_events.extend(events_result.get('items', []))
        page_token = events_result.get('nextPageToken')
        if not page_token:
            break

    matched_updates = []

    for index, row in df_new.iterrows():
        mng_num = row['Subject'][:7]
        for evt in existing_events:
            desc = evt.get("description", "")
            if desc.startswith(f"作業指示書：{mng_num}/"):
                start_dt_new = f"{row['Start Date']} {row['Start Time']}"
                end_dt_new = f"{row['End Date']} {row['End Time']}"
                new_start = datetime.strptime(start_dt_new, "%Y/%m/%d %H:%M")
                new_end = datetime.strptime(end_dt_new, "%Y/%m/%d %H:%M")

                if 'dateTime' in evt['start']:
                    cal_start = datetime.fromisoformat(evt['start']['dateTime'])
                    cal_end = datetime.fromisoformat(evt['end']['dateTime'])
                    if abs((cal_start - new_start).total_seconds()) > 60 or abs((cal_end - new_end).total_seconds()) > 60:
                        matched_updates.append((evt, row))
                break

    if not matched_updates:
        st.success("変更が必要なイベントはありませんでした。")
    else:
        st.info(f"{len(matched_updates)} 件のイベントに更新が必要です。")
        for evt, row in matched_updates:
            st.write(f"🔁 {row['Subject']}")
            st.write(f"📍 現在: {evt['start']['dateTime']} - {evt['end']['dateTime']}")
            st.write(f"📍 更新: {row['Start Date']} {row['Start Time']} - {row['End Date']} {row['End Time']}")
            st.markdown("---")

        if st.button("上記イベントをすべて更新する"):
            count_updated = 0
            for evt, row in matched_updates:
                try:
                    event_id = evt['id']
                    new_start = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M").isoformat()
                    new_end = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M").isoformat()

                    evt['start'] = {'dateTime': new_start, 'timeZone': 'Asia/Tokyo'}
                    evt['end'] = {'dateTime': new_end, 'timeZone': 'Asia/Tokyo'}
                    evt['description'] = row['Description']
                    evt['location'] = row['Location']
                    evt['summary'] = row['Subject']

                    service.events().update(calendarId=calendar_id_update, eventId=event_id, body=evt).execute()
                    count_updated += 1
                except Exception as e:
                    st.error(f"{row['Subject']} の更新に失敗: {e}")

            st.success(f"✅ {count_updated} 件のイベントを更新しました！")
