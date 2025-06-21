import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from excel_parser import process_excel_files
from calendar_utils import authenticate_google, add_event_to_calendar, delete_events_from_calendar
from googleapiclient.discovery import build

st.set_page_config(page_title="Googleカレンダー登録・削除ツール", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# Google認証を最初に実施
st.subheader("🔐 Google認証")
creds = authenticate_google()

# 認証が完了していない場合はここで処理を停止
if not creds:
    st.warning("Google認証を完了してください。")
    st.stop()

# 認証が完了したらサービスをビルドし、セッションステートに保存
# これにより、ユーザーがページをリロードしても認証が維持される
if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    try:
        service = build("calendar", "v3", credentials=creds)
        st.session_state['calendar_service'] = service
        calendar_list = service.calendarList().list().execute()

        # 読み取り専用カレンダーを除外
        editable_calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') != 'reader'
        }
        st.session_state['editable_calendar_options'] = editable_calendar_options

    except Exception as e:
        st.error(f"カレンダーサービスの取得またはカレンダーリストの取得に失敗しました: {e}")
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop() # エラーが発生した場合はそれ以降の処理を停止
else:
    service = st.session_state['calendar_service'] # 既にサービスがあればそれを使用

# ファイルアップロードとイベント設定、イベント削除のタブを作成
# タブの順序は自由に変更できますが、Excelに依存しない削除を独立させるためにタブで区切ります。
tabs = st.tabs(["1. ファイルのアップロード", "2. イベントの登録", "3. イベントの削除"])

with tabs[0]:
    st.header("ファイルをアップロード")
    uploaded_files = st.file_uploader("Excelファイルを選択（複数可）", type=["xlsx"], accept_multiple_files=True)

    # 新しいファイルがアップロードされた場合、またはセッションにファイルがない場合
    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files # アップロードされたファイルをセッションステートに保存
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
        # 初回アクセス時など、ファイルがアップロードされていない場合の初期化
        st.session_state['uploaded_files'] = []
        st.session_state['description_columns_pool'] = []

    # 以前アップロードされたファイルがあればそれを表示
    if st.session_state.get('uploaded_files'):
        st.subheader("アップロード済みのファイル:")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")


with tabs[1]:
    st.header("イベントを登録")
    # アップロードファイルがセッションステートにない場合、処理を停止
    if not st.session_state.get('uploaded_files'):
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
        st.stop() # ここで停止することで、Excelに依存するイベント登録はExcelアップロード後のみ可能

    # イベント設定
    st.subheader("📝 イベント設定")
    all_day_event = st.checkbox("終日イベントとして登録", value=False)
    private_event = st.checkbox("非公開イベントとして登録", value=True)

    # セッションステートからdescription_columns_poolを取得
    description_columns = st.multiselect(
        "説明欄に含める列（複数選択可）",
        st.session_state.get('description_columns_pool', [])
    )

    if not st.session_state['editable_calendar_options']:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        st.stop()

    selected_calendar_name = st.selectbox("登録先カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="reg_calendar_select")
    calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

    # データ処理と登録
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
                            # 終日イベントの場合、日付のみを使用
                            start_date_str = datetime.strptime(row['Start Date'], "%Y/%m/%d").strftime("%Y-%m-%d")
                            # 終日イベントの場合、終了日は開始日の翌日を設定する必要がある (Google Calendar APIの仕様)
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
                            # 時間指定イベントの場合、日付と時間を使用
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

    # カレンダーオプションが利用可能かチェックし、なければエラーメッセージを表示
    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。Google認証を完了しているか、Googleカレンダーの設定を確認してください。")
        # st.stop() # ここではstopせず、メッセージ表示に留めることで、他のタブには影響しない
    else:
        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", list(st.session_state['editable_calendar_options'].keys()), key="del_calendar_select")
        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today = date.today()
        # デフォルトで過去30日間のイベントを対象にする
        default_start_date = today - timedelta(days=30)
        default_end_date = today

        delete_start_date = st.date_input("削除開始日", value=default_start_date)
        delete_end_date = st.date_input("削除終了日", value=default_end_date)

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")

            # 初期化
            if 'show_delete_confirmation' not in st.session_state:
                st.session_state.show_delete_confirmation = False
            if 'last_deleted_count' not in st.session_state:
                st.session_state.last_deleted_count = None

            # 「選択期間のイベントを削除する」ボタン
            if st.button("選択期間のイベントを削除する", key="delete_events_button"):
                st.session_state.show_delete_confirmation = True
                st.session_state.last_deleted_count = None
                st.rerun()

            # 確認フラグがTrueの場合にのみ確認メッセージと「はい」/「いいえ」ボタンを表示
            if st.session_state.show_delete_confirmation:
                st.warning(f"「{selected_calendar_name_del}」カレンダーから {delete_start_date.strftime('%Y年%m月%d日')}から{delete_end_date.strftime('%Y年%m%d日')}までの全てのイベントを削除します。この操作は元に戻せません。よろしいですか？")

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

            # 削除完了メッセージを表示
            if not st.session_state.show_delete_confirmation and st.session_state.last_deleted_count is not None:
                if st.session_state.last_deleted_count > 0:
                    st.success(f"✅ {st.session_state.last_deleted_count} 件のイベントが削除されました。")
                else:
                    st.info("指定された期間内に削除するイベントは見つかりませんでした。")
