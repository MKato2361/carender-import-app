import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_from_calendar
from googleapiclient.discovery import build
# from firebase_utils import initialize_firebase, login_user, logout_user, get_current_user # Firebaseを導入する場合にコメントアウトを外す

st.set_page_config(page_title="Googleカレンダー登録・削除ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# --- Firebase 認証セクション (Firebaseを導入する場合に有効化) ---
# if 'logged_in' not in st.session_state:
#     st.session_state['logged_in'] = False
#     st.session_state['user_email'] = None

# if not initialize_firebase():
#     st.error("アプリケーションの初期化に失敗しました。")
#     st.stop()

# if not st.session_state['logged_in']:
#     st.sidebar.title("ログイン")
#     email = st.sidebar.text_input("メールアドレス")
#     password = st.sidebar.text_input("パスワード", type="password")

#     if st.sidebar.button("ログイン"):
#         if login_user(email, password):
#             st.rerun()
#     st.sidebar.markdown("---")
#     st.sidebar.info("デモ用: 新規ユーザー登録")
#     new_email = st.sidebar.text_input("新規メールアドレス")
#     new_password = st.sidebar.text_input("新規パスワード", type="password", help="6文字以上")
#     if st.sidebar.button("ユーザー登録"):
#         try:
#             user = auth.create_user(email=new_email, password=new_password)
#             st.sidebar.success(f"ユーザー '{user.email}' を登録しました。ログインしてください。")
#         except Exception as e:
#             st.sidebar.error(f"ユーザー登録に失敗しました: {e}")

#     st.stop()
# else:
#     st.sidebar.success(f"ようこそ、{get_current_user()} さん！")
#     if st.sidebar.button("ログアウト"):
#         logout_user()
#         st.rerun()

# --- ここからGoogleカレンダー認証セクションの変更 ---

# Streamlitのplaceholderを使って、認証セクションの内容を動的に変更
google_auth_placeholder = st.empty()

with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    creds = authenticate_google() # 認証プロセスを実行

    # 認証が完了していない場合はここで警告を表示し、停止
    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        # 認証が完了したら、placeholerの内容をクリアし、認証済みメッセージを表示
        google_auth_placeholder.empty() # コンテンツを削除
        st.sidebar.success("✅ Googleカレンダーに認証済みです！") # サイドバーに認証済みメッセージを表示

# 認証が完了したらサービスをビルドし、セッションステートに保存
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

# --- ここからファイルアップロードとイベント設定、イベント削除のタブ (変更なし) ---
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"  # ← 追加
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
                    df = process_excel_files(st.session_state['uploaded_files'], description_columns, all_day_event, private_event)
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
                                    end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date() + timedelta(days=1)
                                    end_date_str = end_date_obj.strftime("%Y-%m-%d")

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'] if pd.notna(row['Location']) else '',
                                        'description': row['Description'] if pd.notna(row['Description']) else '',
                                        'start': {'date': start_date_str},
                                        'end': {'date': end_date_str},
                                        'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                                    }
                                else:
                                    start_dt_str = f"{row['Start Date']} {row['Start Time']}"
                                    end_dt_str = f"{row['End Date']} {row['End Time']}"

                                    start = datetime.strptime(start_dt_str, "%Y/%m/%d %H:%M").isoformat()
                                    end = datetime.strptime(end_dt_str, "%Y/%m/%d %H:%M").isoformat()

                                    event_data = {
                                        'summary': row['Subject'],
                                        'location': row['Location'] if pd.notna(row['Location']) else '',
                                        'description': row['Description'] if pd.notna(row['Description']) else '',
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
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダー認証を完了しているか、Googleカレンダーの設定を確認してください。")
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today = date.today()
        default_start_date = today - timedelta(days=30)
        default_end_date = today

        delete_start_date = st.date_input("削除開始日", value=default_start_date)
        delete_end_date = st.date_input("削除終了日", value=default_end_date)

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")

            if 'show_delete_confirmation' not in st.session_state:
                st.session_state.show_delete_confirmation = False
            if 'last_deleted_count' not in st.session_state:
                st.session_state.last_deleted_count = None

            if st.button("選択期間のイベントを削除する", key="delete_events_button"):
                st.session_state.show_delete_confirmation = True
                st.session_state.last_deleted_count = None
                st.rerun()

            if st.session_state.show_delete_confirmation:
                st.warning(f"「{selected_calendar_name_del}」カレンダーから {delete_start_date.strftime('%Y年%m月%d日')}から{delete_end_date.strftime('%Y年%m月%d日')}までの全てのイベントを削除します。この操作は元に戻せません。よろしいですか？")

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("はい、削除を実行します", key="confirm_delete_button_final"):
                        deleted_count = delete_events_from_calendar(
                            service, calendar_id_del,
                            datetime.combine(delete_start_date, datetime.min.time()),
                            datetime.combine(delete_end_date, datetime.max.time())
                        )
                        st.session_state.last_deleted_count = deleted_count
                        st.session_state.show_delete_confirmation = False
                        st.rerun()
                with col2:
                    if st.button("いいえ、キャンセルします", key="cancel_delete_button"):
                        st.info("削除はキャンセルされました。")
                        st.session_state.show_delete_confirmation = False
                        st.session_state.last_deleted_count = None
                        st.rerun()

            if not st.session_state.show_delete_confirmation and st.session_state.last_deleted_count is not None:
                if st.session_state.last_deleted_count > 0:
                    st.success(f"✅ {st.session_state.last_deleted_count} 件のイベントが削除されました。")
                else:
                    st.info("指定された期間内に削除するイベントは見つかりませんでした。")

                    
# --- 追加タブ: イベント更新 ---
with tabs[3]:
    st.header("イベントを更新")

    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        all_day_event = st.checkbox("終日イベントとして扱う", value=False, key="update_all_day")
        private_event = st.checkbox("非公開イベントとして扱う", value=True, key="update_private")
        description_columns = st.multiselect(
            "説明欄に含める列",
            st.session_state['description_columns_pool'],
            key="update_desc_cols"
        )

        if not st.session_state['editable_calendar_options']:
            st.error("更新可能なカレンダーが見つかりません。")
        else:
            selected_calendar_name_upd = st.selectbox(
                "更新対象カレンダーを選択",
                list(st.session_state['editable_calendar_options'].keys()),
                key="update_calendar_select"
            )
            calendar_id_upd = st.session_state['editable_calendar_options'][selected_calendar_name_upd]

            if st.button("イベントを照合・更新"):
                with st.spinner("イベントを処理中..."):
                    df = process_excel_files(
                        st.session_state['uploaded_files'],
                        description_columns,
                        all_day_event,
                        private_event
                    )

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


