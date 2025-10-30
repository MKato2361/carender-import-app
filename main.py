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
import unicodedata

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")

st.markdown("""
<style>
/* ===== カラーテーマ別（ライト／ダーク） ===== */
@media (prefers-color-scheme: light) {
    .header-bar {
        background-color: rgba(249, 249, 249, 0.95);
        color: #333;
        border-bottom: 1px solid #ccc;
    }
}
@media (prefers-color-scheme: dark) {
    .header-bar {
        background-color: rgba(30, 30, 30, 0.9);
        color: #eee;
        border-bottom: 1px solid #444;
    }
}

/* ===== ヘッダー ===== */
.header-bar {
    position: sticky;
    top: 0;
    width: 100%;
    text-align: center;
    font-weight: 500;
    font-size: 14px;
    padding: 8px 0;
    z-index: 20;
    backdrop-filter: blur(6px);
}

/* ===== タブバー（sticky風） ===== */
div[data-testid="stTabs"] {
    position: sticky;
    top: 42px; /* ヘッダー高さ分 */
    z-index: 15;
    background-color: inherit;
    border-bottom: 1px solid rgba(128, 128, 128, 0.3);
    padding-top: 4px;
    padding-bottom: 4px;
    backdrop-filter: blur(6px);
}

/* ===== Streamlit レイアウト調整 ===== */
.block-container,
section[data-testid="stMainBlockContainer"],
main {
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    margin-bottom: 0 !important;
    height: auto !important;
    min-height: 100vh !important;
    overflow: visible !important;
}

/* ===== 不要なフッターや透明ブロックを削除 ===== */
footer, div[data-testid="stBottomBlockContainer"] {
    display: none !important;
    height: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
}

/* ===== ページ全体の高さ／スクロール制御 ===== */
html, body, #root {
    height: auto !important;
    min-height: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    overflow-x: hidden !important;
    overflow-y: auto !important;
    overscroll-behavior: none !important;
    -webkit-overflow-scrolling: touch !important;
}

/* ===== 最後のブロックの余白除去 ===== */
div[data-testid="stVerticalBlock"] > div:last-child {
    margin-bottom: 0 !important;
    padding-bottom: 0 !important;
}

/* ===== iOS専用: Stickyによる余白対策 ===== */
@supports (-webkit-touch-callout: none) {
    html, body {
        height: 100% !important;
        min-height: 100% !important;
        overflow-x: hidden !important;
        overflow-y: auto !important;
        overscroll-behavior-y: contain !important;
        -webkit-overflow-scrolling: touch !important;
    }

    .header-bar, div[data-testid="stTabs"] {
        position: static !important; /* iOSで固定解除 */
        top: auto !important;
    }

    main, section[data-testid="stMainBlockContainer"], .block-container {
        height: auto !important;
        min-height: auto !important;
        padding-bottom: 0 !important;
        margin-bottom: 0 !important;
    }

    footer, div[data-testid="stBottomBlockContainer"] {
        display: none !important;
        height: 0 !important;
    }

    body {
        padding-bottom: env(safe-area-inset-bottom, 0px);
        background-color: transparent !important;
    }
}
</style>

<div class="header-bar">
📅 Googleカレンダー一括イベント登録・削除
</div>
""", unsafe_allow_html=True)

# ==================================================
# Firebase 初期化・認証処理（元コードそのまま）
# ==================================================
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()

if not user_id:
    firebase_auth_form()
    st.stop()


def load_user_settings_from_firestore(user_id):
    """Firestoreからユーザー設定を読み込み、セッションに同期"""
    if not user_id:
        return
    initialize_session_state(user_id)
    doc_ref = db.collection('user_settings').document(user_id)
    doc = doc_ref.get()
    if doc.exists:
        settings = doc.to_dict()
        for key, value in settings.items():
            set_user_setting(user_id, key, value)


def save_user_setting_to_firestore(user_id, setting_key, setting_value):
    """Firestoreにユーザー設定を保存"""
    if not user_id:
        return
    doc_ref = db.collection('user_settings').document(user_id)
    try:
        doc_ref.set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")


# ユーザー設定の読み込み
load_user_settings_from_firestore(user_id)

google_auth_placeholder = st.empty()

with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    creds = authenticate_google()

    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()


def initialize_calendar_service():
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        editable_calendar_options = {
            cal['summary']: cal['id']
            for cal in calendar_list['items']
            if cal.get('accessRole') != 'reader'
        }
        return service, editable_calendar_options
    except HttpError as e:
        st.error(f"カレンダーサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.error(f"カレンダーサービスの初期化に失敗しました: {e}")
        return None, None


def initialize_tasks_service_wrapper():
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        for task_list in task_lists.get('items', []):
            if task_list.get('title') == 'My Tasks':
                default_task_list_id = task_list['id']
                break
        if not default_task_list_id and task_lists.get('items'):
            default_task_list_id = task_lists['items'][0]['id']
        return tasks_service, default_task_list_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
        return None, None


if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    service, editable_calendar_options = initialize_calendar_service()
    if not service:
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop()
    st.session_state['calendar_service'] = service
    st.session_state['editable_calendar_options'] = editable_calendar_options
else:
    service = st.session_state['calendar_service']
    _, st.session_state['editable_calendar_options'] = initialize_calendar_service()

if 'tasks_service' not in st.session_state or not st.session_state.get('tasks_service'):
    tasks_service, default_task_list_id = initialize_tasks_service_wrapper()
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
    if not tasks_service:
        st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
else:
    tasks_service = st.session_state['tasks_service']


# ==================================================
# タブ部分（固定化・半透明デザイン付き）
# ==================================================
st.markdown('<div class="fixed-tabs">', unsafe_allow_html=True)

tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. 重複イベントの検出・削除",
    "5. イベントのExcel出力"
])



st.markdown('</div>', unsafe_allow_html=True)


# ==================================================
# 以降：元のコード（機能・処理は一切変更なし）
# ==================================================

# ↓↓↓ あなたのオリジナルの全処理（アップロード・登録・削除・更新・出力・サイドバーなど）をそのまま残してください ↓↓↓
# （ここ以降のロジック・UI要素・API処理はすべてオリジナルのままで動作します）



if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
    st.session_state['description_columns_pool'] = []
    st.session_state['merged_df_for_selector'] = pd.DataFrame()

with tabs[0]:
    st.subheader("ファイルをアップロード")
    with st.expander("ℹ️作業手順と補足"):
        st.info("""
    **☀作業指示書一覧をアップロードすると管理番号+物件名をイベント名として任意のカレンダーに登録します。**
    
    **☀イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。（複数選択可能,次回から同じ項目が選択されます）**
    
    **☀イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。**
    
    **☀作業外予定の一覧をアップロードすると、イベント名を選択することができます。**

    **☀ToDoリストを作成すると、点検通知のリマインドが可能です（ToDoとしてイベント登録されます）**
    """)

    def get_local_excel_files():
        current_dir = Path(__file__).parent
        return [f for f in current_dir.glob("*") if f.suffix.lower() in [".xlsx", ".xls", ".csv"]]

    uploaded_files = st.file_uploader("ExcelまたはCSVファイルを選択（複数可）",type=["xlsx", "xls", "csv"],accept_multiple_files=True)


    local_excel_files = get_local_excel_files()
    selected_local_files = []
    if local_excel_files:
        st.markdown("📁 サーバーにあるExcelファイル")
        local_file_names = [f.name for f in local_excel_files]
        selected_names = st.multiselect(
            "以下のファイルを処理対象に含める（アップロードと同様に扱われます）",
            local_file_names
        )
        for name in selected_names:
            full_path = next((f for f in local_excel_files if f.name == name), None)
            if full_path:
                with open(full_path, "rb") as f:
                    file_bytes = f.read()
                    file_obj = BytesIO(file_bytes)
                    file_obj.name = name
                    selected_local_files.append(file_obj)

    all_files = []
    if uploaded_files:
        all_files.extend(uploaded_files)
    if selected_local_files:
        all_files.extend(selected_local_files)

    if all_files:
        st.session_state['uploaded_files'] = all_files
        try:
            st.session_state['merged_df_for_selector'] = _load_and_merge_dataframes(all_files)
            st.session_state['description_columns_pool'] = st.session_state['merged_df_for_selector'].columns.tolist()
            if st.session_state['merged_df_for_selector'].empty:
                st.warning("読み込まれたファイルに有効なデータがありませんでした。")
        except (ValueError, IOError) as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []

    if st.session_state.get('uploaded_files'):
        st.subheader("📄 処理対象ファイル一覧")
        for f in st.session_state['uploaded_files']:
            st.write(f"- {f.name}")
        if not st.session_state['merged_df_for_selector'].empty:
            st.info(f"📊 データ列数: {len(st.session_state['merged_df_for_selector'].columns)}、行数: {len(st.session_state['merged_df_for_selector'])}")

        if st.button("🗑️ アップロード済みファイルをクリア", help="選択中のファイルとデータを削除します。"):
            st.session_state['uploaded_files'] = []
            st.session_state['merged_df_for_selector'] = pd.DataFrame()
            st.session_state['description_columns_pool'] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()

with tabs[1]:
    st.subheader("イベントを登録・更新")

    # 🛠️ 安全のため初期化
    description_columns = []
    selected_event_name_col = None
    add_task_type_to_event_name = False
    all_day_event_override = False
    private_event = True
    fallback_event_name_column = None

    if not st.session_state.get('uploaded_files') or st.session_state['merged_df_for_selector'].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")

    elif not st.session_state['editable_calendar_options']:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")

    else:
        # ---------------------------------------------------------
        # カレンダー選択
        # ---------------------------------------------------------
        calendar_options = list(st.session_state['editable_calendar_options'].keys())
        
        # ✅ ユーザー設定から前回の選択を取得
        saved_calendar_name = get_user_setting(user_id, 'selected_calendar_name')
        
        # ✅ デフォルトインデックスを計算
        try:
            default_index = calendar_options.index(saved_calendar_name)
        except ValueError:
            # 前回の設定がリストにない場合（カレンダーが削除されたなど）、最初の項目をデフォルトとする
            default_index = 0
            
           selected_calendar_name = st.selectbox(
            "登録先カレンダーを選択",
            calendar_options,
            index=default_index,  # 👈 デフォルトインデックスを適用
            key="reg_calendar_select"
        )
        calendar_id = st.session_state['editable_calendar_options'][selected_calendar_name]

        # ✅ 選択されたカレンダーをユーザー設定に保存
        # セレクトボックスが変更された時点で設定を更新(Streamlitの再実行で実行される)
        if selected_calendar_name != saved_calendar_name:
            set_user_setting(user_id, 'selected_calendar_name', selected_calendar_name)
            save_user_setting_to_firestore(user_id, 'selected_calendar_name', selected_calendar_name)
            
        # ---------------------------------------------------------
        # 設定を取得して、展開状態を制御
        # ---------------------------------------------------------
        description_columns_pool = st.session_state.get('description_columns_pool', [])
        saved_description_cols = get_user_setting(user_id, 'description_columns_selected')
        saved_event_name_col = get_user_setting(user_id, 'event_name_col_selected')
        saved_task_type_flag = get_user_setting(user_id, 'add_task_type_to_event_name')
        saved_create_todo_flag = get_user_setting(user_id, 'create_todo_checkbox_state')

        # ✅ 未設定なら展開、設定済みなら閉じる
        expand_event_setting = not bool(saved_description_cols)
        expand_name_setting = not (saved_event_name_col or saved_task_type_flag)
        expand_todo_setting = bool(saved_create_todo_flag)

        # ---------------------------------------------------------
        # 🧩 イベント設定（折りたたみ）
        # ---------------------------------------------------------
        with st.expander("📝 イベント設定", expanded=expand_event_setting):
            all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
            private_event = st.checkbox("非公開イベントとして登録", value=True)

            default_selection = [col for col in (saved_description_cols or []) if col in description_columns_pool]
            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                description_columns_pool,
                default=default_selection,
                key=f"description_selector_register_{user_id}"
            )

        # ---------------------------------------------------------
        # 🧩 イベント名の生成設定（折りたたみ）
        # ---------------------------------------------------------
        with st.expander("🧱 イベント名の生成設定", expanded=expand_name_setting):
            has_mng_data, has_name_data = check_event_name_columns(st.session_state['merged_df_for_selector'])
            selected_event_name_col = saved_event_name_col
            add_task_type_to_event_name = st.checkbox(
                "イベント名の先頭に作業タイプを追加する",
                value=saved_task_type_flag,
                key=f"add_task_type_checkbox_{user_id}"
            )

            if not (has_mng_data and has_name_data):
                available_event_name_cols = get_available_columns_for_event_name(st.session_state['merged_df_for_selector'])
                event_name_options = ["選択しない"] + available_event_name_cols
                default_index = event_name_options.index(selected_event_name_col) if selected_event_name_col in event_name_options else 0

                selected_event_name_col = st.selectbox(
                    "イベント名として使用する代替列を選択してください:",
                    options=event_name_options,
                    index=default_index,
                    key=f"event_name_selector_register_{user_id}"
                )

                if selected_event_name_col != "選択しない":
                    fallback_event_name_column = selected_event_name_col
            else:
                st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

        # ---------------------------------------------------------
        # 🧩 ToDoリスト連携設定（折りたたみ）
        # ---------------------------------------------------------
        st.subheader("✅ ToDoリスト連携設定 (オプション)")
        with st.expander("ToDoリスト作成オプション", expanded=expand_todo_setting):
            create_todo = st.checkbox(
                "このイベントに対応するToDoリストを作成する", 
                value=saved_create_todo_flag if saved_create_todo_flag is not None else False, 
                key="create_todo_checkbox"
            )

            set_user_setting(user_id, 'create_todo_checkbox_state', create_todo)
            save_user_setting_to_firestore(user_id, 'create_todo_checkbox_state', create_todo)


            fixed_todo_types = ["点検通知"]
            if create_todo:
                st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
            else:
                st.markdown(f"ToDoリストの作成は無効です。")

            deadline_offset_options = {
                "2週間前": 14,
                "10日前": 10,
                "1週間前": 7,
                "カスタム日数前": None
            }
            selected_offset_key = st.selectbox(
                "ToDoリストの期限をイベント開始日の何日前に設定しますか？",
                list(deadline_offset_options.keys()),
                disabled=not create_todo,
                key="deadline_offset_select"
            )

            custom_offset_days = None
            if selected_offset_key == "カスタム日数前":
                custom_offset_days = st.number_input(
                    "何日前に設定しますか？ (日数)",
                    min_value=0,
                    value=3,
                    disabled=not create_todo,
                    key="custom_offset_input"
                )
        # ---------------------------------------------------------
        # 登録ボタン
        # ---------------------------------------------------------
        st.subheader("➡️ イベント登録・更新実行")

        if st.button("Googleカレンダーに登録・更新する"):
            # ⚙️ ユーザー設定の保存
            set_user_setting(user_id, 'description_columns_selected', description_columns)
            set_user_setting(user_id, 'event_name_col_selected', selected_event_name_col)
            set_user_setting(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)

            save_user_setting_to_firestore(user_id, 'description_columns_selected', description_columns)
            save_user_setting_to_firestore(user_id, 'event_name_col_selected', selected_event_name_col)
            save_user_setting_to_firestore(user_id, 'add_task_type_to_event_name', add_task_type_to_event_name)

            with st.spinner("イベントデータを処理中..."):
                try:
                    df = process_excel_data_for_calendar(
                        st.session_state['uploaded_files'], 
                        description_columns, 
                        all_day_event_override,
                        private_event, 
                        fallback_event_name_column,
                        add_task_type_to_event_name
                    )
                except (ValueError, IOError) as e:
                    st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                    df = pd.DataFrame()

                if df.empty:
                    st.warning("有効なイベントデータがありません。処理を中断しました。")
                else:
                    st.info(f"{len(df)} 件のイベントを処理します。")
                    progress = st.progress(0)
                    successful_operations = 0

                    worksheet_to_event = {}
                    time_min = (datetime.now(timezone.utc) - timedelta(days=365*2)).isoformat()
                    time_max = (datetime.now(timezone.utc) + timedelta(days=365*2)).isoformat()
                    events = fetch_all_events(service, calendar_id, time_min, time_max)

                    for event in events:
                        desc = event.get('description', '')
                        match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", desc)
                        if match:
                            worksheet_id = match.group(1)
                            worksheet_to_event[worksheet_id] = event

                    for i, row in df.iterrows():
                        match = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", row['Description'])
                        event_data = {
                            'summary': row['Subject'],
                            'location': row['Location'],
                            'description': row['Description'],
                            'transparency': 'transparent' if row['Private'] == "True" else 'opaque'
                        }

                        if row['All Day Event'] == "True":
                            start_date_obj = datetime.strptime(row['Start Date'], "%Y/%m/%d").date()
                            end_date_obj = datetime.strptime(row['End Date'], "%Y/%m/%d").date()
                            event_data['start'] = {'date': start_date_obj.strftime("%Y-%m-%d")}
                            event_data['end'] = {'date': (end_date_obj + timedelta(days=1)).strftime("%Y-%m-%d")}
                        else:
                            start_dt_obj = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", "%Y/%m/%d %H:%M")
                            end_dt_obj = datetime.strptime(f"{row['End Date']} {row['End Time']}", "%Y/%m/%d %H:%M")
                            event_data['start'] = {'dateTime': start_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
                            event_data['end'] = {'dateTime': end_dt_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

                        worksheet_id = match.group(1) if match else None
                        existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

                        try:
                            if existing_event:
                                updated_event = update_event_if_needed(service, calendar_id, existing_event['id'], event_data)
                                if updated_event:
                                    successful_operations += 1
                            else:
                                added_event = add_event_to_calendar(service, calendar_id, event_data)
                                if added_event:
                                    successful_operations += 1
                                    if worksheet_id:
                                        worksheet_to_event[worksheet_id] = added_event
                        except Exception as e:
                            st.error(f"イベント '{row['Subject']}' の登録または更新に失敗しました: {e}")

                        progress.progress((i + 1) / len(df))

                    st.success(f"✅ {successful_operations} 件のイベントが処理されました。")

with tabs[2]:
    st.subheader("イベントを削除")

    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
    else:
        # 📹 カレンダー名一覧を取得
        calendar_names = list(st.session_state['editable_calendar_options'].keys())

        # 📹 デフォルトインデックスを設定(前回の選択を反映)
        saved_calendar_name = get_user_setting(user_id, 'selected_calendar_name')
        default_index = 0
        if saved_calendar_name and saved_calendar_name in calendar_names:
            default_index = calendar_names.index(saved_calendar_name)

        # 📹 セレクトボックス(デフォルト選択付き)
        selected_calendar_name_del = st.selectbox(
            "削除対象カレンダーを選択",
            calendar_names,
            index=default_index,
            key="del_calendar_select"
        )

        # 📹 選択されたカレンダーをユーザー設定に保存
        if selected_calendar_name_del != saved_calendar_name:
            set_user_setting(user_id, 'selected_calendar_name', selected_calendar_name_del)
            save_user_setting_to_firestore(user_id, 'selected_calendar_name', selected_calendar_name_del)

        calendar_id_del = st.session_state['editable_calendar_options'][selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today_date = date.today()
        delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
        delete_end_date = st.date_input("削除終了日", value=today_date)
        delete_related_todos = st.checkbox(
            "関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)",
            value=False
        )

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")

            # 削除確認フラグの初期化
            if 'confirm_delete' not in st.session_state:
                st.session_state['confirm_delete'] = False

            # 最初のボタン: 削除確認を表示
            if not st.session_state['confirm_delete']:
                if st.button("選択期間のイベントを削除する", type="primary"):
                    st.session_state['confirm_delete'] = True
                    st.rerun()

            # 確認メッセージと実行/キャンセルボタン
            if st.session_state['confirm_delete']:
                st.warning(f"""
                ⚠️ **削除確認**
                
                以下のイベントを削除します:
                - **カレンダー名**: {selected_calendar_name_del}
                - **期間**: {delete_start_date.strftime('%Y年%m月%d日')} ～ {delete_end_date.strftime('%Y年%m月%d日')}
                - **ToDoリストも削除**: {'はい' if delete_related_todos else 'いいえ'}
                
                この操作は取り消せません。本当に削除しますか？
                """)

                col1, col2 = st.columns([1, 1])

                with col1:
                    if st.button("✅ 実行", type="primary", use_container_width=True):
                        st.session_state['confirm_delete'] = False
                        calendar_service = st.session_state['calendar_service']
                        tasks_service = st.session_state['tasks_service']
                        default_task_list_id = st.session_state.get('default_task_list_id')

                        start_dt_utc = datetime.combine(
                            delete_start_date, datetime.min.time(),
                            tzinfo=datetime.now().astimezone().tzinfo
                        ).astimezone(timezone.utc)

                        end_dt_utc = datetime.combine(
                            delete_end_date, datetime.max.time(),
                            tzinfo=datetime.now().astimezone().tzinfo
                        ).astimezone(timezone.utc)

                        time_min_utc = start_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                        time_max_utc = end_dt_utc.isoformat(timespec='microseconds').replace('+00:00', 'Z')

                        events_to_delete = fetch_all_events(calendar_service, calendar_id_del, time_min_utc, time_max_utc)

                        if not events_to_delete:
                            st.info("指定期間内に削除するイベントはありませんでした。")

                        deleted_events_count = 0
                        deleted_todos_count = 0
                        total_events = len(events_to_delete)

                        if total_events > 0:
                            progress_bar = st.progress(0)
                            status_text = st.empty()

                            for i, event in enumerate(events_to_delete):
                                event_summary = event.get('summary', '不明なイベント')
                                event_id = event['id']

                                status_text.text(f"イベント '{event_summary}' を削除中... ({i+1}/{total_events})")

                                try:
                                    if delete_related_todos and tasks_service and default_task_list_id:
                                        deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                            tasks_service,
                                            default_task_list_id,
                                            event_id
                                        )
                                        deleted_todos_count += deleted_task_count_for_event

                                    calendar_service.events().delete(
                                        calendarId=calendar_id_del,
                                        eventId=event_id
                                    ).execute()

                                    deleted_events_count += 1
                                except Exception as e:
                                    st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")

                                progress_bar.progress((i + 1) / total_events)

                            status_text.empty()

                            if deleted_events_count > 0:
                                st.success(f"✅ {deleted_events_count} 件のイベントが削除されました。")
                                if delete_related_todos:
                                    if deleted_todos_count > 0:
                                        st.success(f"✅ {deleted_todos_count} 件の関連ToDoタスクが削除されました。")
                                    else:
                                        st.info("関連するToDoタスクは見つからなかったか、すでに削除されていました。")
                            else:
                                st.info("指定期間内に削除するイベントはありませんでした。")
                        else:
                            st.info("指定期間内に削除するイベントはありませんでした。")

                with col2:
                    if st.button("❌ キャンセル", use_container_width=True):
                        st.session_state['confirm_delete'] = False
                        st.rerun()

                        

with tabs[3]:
    st.subheader("🔍 重複イベントの検出・削除")

    # ----------------------------------------------------
    # セッションステートに保存されたメッセージを表示し、クリアする
    if 'last_dup_message' in st.session_state and st.session_state['last_dup_message']:
        msg_type, msg_text = st.session_state['last_dup_message']
        
        if msg_type == "success":
            st.success(msg_text)
        elif msg_type == "error":
            st.error(msg_text)
        elif msg_type == "info":
            st.info(msg_text)
        
        st.session_state['last_dup_message'] = None 
    # ----------------------------------------------------

    calendar_options = list(st.session_state['editable_calendar_options'].keys())
    
    # ✅ ユーザー設定から前回の選択を取得
    saved_calendar_name = get_user_setting(user_id, 'selected_calendar_name')
    default_index = 0
    if saved_calendar_name and saved_calendar_name in calendar_options:
        default_index = calendar_options.index(saved_calendar_name)
    
    selected_calendar = st.selectbox(
        "対象カレンダーを選択", 
        calendar_options, 
        index=default_index,
        key="dup_calendar_select"
    )
    
    # ✅ 選択されたカレンダーをユーザー設定に保存
    if selected_calendar != saved_calendar_name:
        set_user_setting(user_id, 'selected_calendar_name', selected_calendar)
        save_user_setting_to_firestore(user_id, 'selected_calendar_name', selected_calendar)
    
    calendar_id = st.session_state['editable_calendar_options'][selected_calendar]

    # 👇 削除モード選択
    delete_mode = st.radio(
        "削除モードを選択",
        ["手動で選択して削除", "古い方を自動削除", "新しい方を自動削除"],
        horizontal=True,
        key="dup_delete_mode"
    )

    # ----------------------------------------------------
    # ステート初期化
    if 'dup_df' not in st.session_state:
        st.session_state['dup_df'] = pd.DataFrame()
    if 'auto_delete_ids' not in st.session_state:
        st.session_state['auto_delete_ids'] = []
    if 'last_dup_message' not in st.session_state:
        st.session_state['last_dup_message'] = None
    # ----------------------------------------------------
    
    # helper function for parsing created time
    def parse_created(dt_str):
        try:
            return datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
        except Exception:
            return datetime.min

    if st.button("重複イベントをチェック", key="run_dup_check"):
        # ----------------------------------------------------
        # データ取得・計算ロジック
        # ----------------------------------------------------
        with st.spinner("カレンダー内のイベントを取得中..."):
            time_min = (datetime.now(timezone.utc) - timedelta(days=365*2)).isoformat()
            time_max = (datetime.now(timezone.utc) + timedelta(days=365*2)).isoformat()
            events = fetch_all_events(
                st.session_state['calendar_service'],
                calendar_id,
                time_min,
                time_max
            )

        if not events:
            st.session_state['last_dup_message'] = ("info", "イベントが見つかりませんでした。")
            st.session_state['dup_df'] = pd.DataFrame()
            st.session_state['auto_delete_ids'] = []
            st.session_state['current_delete_mode'] = delete_mode
            st.rerun()
            
        st.success(f"{len(events)} 件のイベントを取得しました。")

        # イベントデータの前処理
        # re.DOTALL: . が改行文字にもマッチ | re.IGNORECASE: 大文字/小文字を無視
        pattern = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", re.DOTALL | re.IGNORECASE)
        rows = []
        for e in events:
            desc = e.get("description", "")
            
            desc_stripped = desc.strip()
            match = pattern.search(desc_stripped)
            worksheet_id = match.group(1) if match else None

            # 💡 修正点 2: 抽出した作業指示書番号を正規化
            if worksheet_id:
                # NFKCで全角数字や特殊文字を半角に変換し、両端の空白を削除
                worksheet_id = unicodedata.normalize('NFKC', worksheet_id).strip()
            
            start_time = e["start"].get("dateTime", e["start"].get("date"))
            end_time = e["end"].get("dateTime", e["end"].get("date"))
            rows.append({
                "id": e["id"],
                "summary": e.get("summary", ""),
                "worksheet_id": worksheet_id, # 正規化されたID
                "created": e.get("created", None),
                "start": start_time,
                "end": end_time,
            })

        df = pd.DataFrame(rows)
        # 作業指示書番号が抽出できたイベントのみを対象とする
        df_valid = df[df["worksheet_id"].notna()].copy()
        
        # 💡 修正点 3: 重複検出のキーを ['worksheet_id'] のみに戻す (正規化により揺らぎを吸収)
        dup_mask = df_valid.duplicated(subset=["worksheet_id"], keep=False)
        dup_df = df_valid[dup_mask].sort_values(["worksheet_id", "created"])

        st.session_state['dup_df'] = dup_df 

        if dup_df.empty:
            st.session_state['last_dup_message'] = ("info", "重複している作業指示書番号は見つかりませんでした。")
            st.session_state['auto_delete_ids'] = []
            st.session_state['current_delete_mode'] = delete_mode
            st.rerun()

        # ----------------------------------------------------
        # 自動削除IDの計算（作成日時ベース）
        # ----------------------------------------------------
        if delete_mode != "手動で選択して削除":
            auto_delete_ids = []

            # 💡 修正点 4: グループ化キーを ['worksheet_id'] のみに戻す
            for _, group in dup_df.groupby("worksheet_id"): 
                
                # 作成日時とイベントID（第2キー）でソートし、順序を確定
                group_sorted = group.sort_values(
                    ["created", "id"],
                    key=lambda s: s.map(parse_created) if s.name == "created" else s,
                    ascending=True
                )
                
                if len(group_sorted) <= 1: continue
                    
                if delete_mode == "古い方を自動削除":
                    # 最新のイベント（最後の行）を残し、それ以前のすべてを削除対象とする
                    delete_targets = group_sorted.iloc[:-1] 
                elif delete_mode == "新しい方を自動削除":
                    # 最古のイベント（最初の行）を残し、それ以降のすべてを削除対象とする
                    delete_targets = group_sorted.iloc[1:]
                else: continue
                    
                auto_delete_ids.extend(delete_targets["id"].tolist())

            st.session_state['auto_delete_ids'] = auto_delete_ids 
            st.session_state['current_delete_mode'] = delete_mode
        else:
            st.session_state['auto_delete_ids'] = []
            st.session_state['current_delete_mode'] = delete_mode
            
        st.rerun()
        # ----------------------------------------------------
        
    # ----------------------------------------------------
    # UI表示ロジック
    # ----------------------------------------------------
    if not st.session_state['dup_df'].empty:
        dup_df = st.session_state['dup_df']
        current_mode = st.session_state.get('current_delete_mode', "手動で選択して削除")
        
        st.warning(f"⚠️ {dup_df['worksheet_id'].nunique()} 種類の重複作業指示書が見つかりました。（合計 {len(dup_df)} イベント）")
        st.dataframe(
            dup_df[["worksheet_id", "summary", "created", "start", "end", "id"]],
            use_container_width=True
        )

        service = st.session_state['calendar_service']

        # ==============================
        # 🧩 手動削除モード
        # ==============================
        if current_mode == "手動で選択して削除":
            delete_ids = st.multiselect(
                "削除するイベントを選択してください（イベントIDで指定）",
                dup_df["id"].tolist(),
                key="manual_delete_ids"
            )

            confirm = st.checkbox("削除操作を確認しました", value=False, key="manual_del_confirm")

            if st.button("🗑️ 選択したイベントを削除", type="primary", disabled=not confirm, key="run_manual_delete"):
                deleted_count = 0
                errors = []
                for eid in delete_ids:
                    try:
                        service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                        deleted_count += 1
                    except Exception as e:
                        errors.append(f"イベントID {eid} の削除に失敗: {e}")
                
                if deleted_count > 0:
                    st.session_state['last_dup_message'] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")
                
                if errors:
                    error_message = "\n".join(errors)
                    st.error(f"以下のイベントの削除に失敗しました:\n{error_message}")
                    if deleted_count == 0:
                         st.session_state['last_dup_message'] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")
                         
                st.session_state['dup_df'] = pd.DataFrame() 
                st.rerun()


        # ==============================
        # 🧩 自動削除モード
        # ==============================
        else:
            auto_delete_ids = st.session_state['auto_delete_ids']
            
            if not auto_delete_ids:
                st.info("削除対象のイベントが見つかりませんでした。")
            else:
                st.warning(f"以下のモードで {len(auto_delete_ids)} 件のイベントを自動削除します: **{current_mode}**")
                st.write(auto_delete_ids)

                confirm = st.checkbox("削除操作を確認しました", value=False, key="auto_del_confirm_final")

                if st.button("🗑️ 自動削除を実行", type="primary", disabled=not confirm, key="run_auto_delete"):
                    deleted_count = 0
                    errors = []
                    for eid in auto_delete_ids:
                        try:
                            service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                            deleted_count += 1
                        except Exception as e:
                            errors.append(f"イベントID {eid} の削除に失敗: {e}")
                            
                    if deleted_count > 0:
                        st.session_state['last_dup_message'] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")
                    
                    if errors:
                        error_message = "\n".join(errors)
                        st.error(f"以下のイベントの削除に失敗しました:\n{error_message}")
                        if deleted_count == 0:
                            st.session_state['last_dup_message'] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                    st.session_state['dup_df'] = pd.DataFrame() 
                    st.rerun()
        
with tabs[4]:  # tabs[4]は新しいタブに対応
    st.subheader("カレンダーイベントをExcelに出力")
    if 'editable_calendar_options' not in st.session_state or not st.session_state['editable_calendar_options']:
        st.error("利用可能なカレンダーが見つかりません。")
    else:
        # ✅ ユーザー設定から前回の選択を取得
        calendar_options = list(st.session_state['editable_calendar_options'].keys())
        saved_calendar_name = get_user_setting(user_id, 'selected_calendar_name')
        default_index = 0
        if saved_calendar_name and saved_calendar_name in calendar_options:
            default_index = calendar_options.index(saved_calendar_name)
        
        selected_calendar_name_export = st.selectbox(
            "出力対象カレンダーを選択", 
            calendar_options,
            index=default_index,
            key="export_calendar_select"
        )
        
        # ✅ 選択されたカレンダーをユーザー設定に保存
        if selected_calendar_name_export != saved_calendar_name:
            set_user_setting(user_id, 'selected_calendar_name', selected_calendar_name_export)
            save_user_setting_to_firestore(user_id, 'selected_calendar_name', selected_calendar_name_export)
        
        calendar_id_export = st.session_state['editable_calendar_options'][selected_calendar_name_export]
        
        st.subheader("🗓️ 出力期間の選択")
        today_date_export = date.today()
        export_start_date = st.date_input("出力開始日", value=today_date_export - timedelta(days=30))
        export_end_date = st.date_input("出力終了日", value=today_date_export)
        
        export_format = st.radio("出力形式を選択", ("CSV", "Excel"), index=0)

        if export_start_date > export_end_date:
            st.error("出力開始日は終了日より前に設定してください。")
        else:
            if st.button("指定期間のイベントを読み込む"):
                with st.spinner("イベントを読み込み中..."):
                    try:
                        calendar_service = st.session_state['calendar_service']
                        
                        start_dt_utc_export = datetime.combine(export_start_date, datetime.min.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                        end_dt_utc_export = datetime.combine(export_end_date, datetime.max.time(), tzinfo=datetime.now().astimezone().tzinfo).astimezone(timezone.utc)
                        
                        time_min_utc_export = start_dt_utc_export.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                        time_max_utc_export = end_dt_utc_export.isoformat(timespec='microseconds').replace('+00:00', 'Z')
                        
                        events_to_export = fetch_all_events(calendar_service, calendar_id_export, time_min_utc_export, time_max_utc_export)
                        
                        if not events_to_export:
                            st.info("指定期間内にイベントは見つかりませんでした。")
                        else:
                            extracted_data = []
                            
                            # 正規表現パターン定義 (最終版)
                            wonum_pattern = re.compile(r"\[作業指示書[：:]\s*(.*?)\]")
                            assetnum_pattern = re.compile(r"\[管理番号[：:]\s*(.*?)\]")
                            worktype_pattern = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
                            title_pattern = re.compile(r"\[タイトル[：:]\s*(.*?)\]") # DESCRIPTION用

                            for event in events_to_export:
                                description_text = event.get('description', '')
                                
                                # 説明フィールドからの抽出
                                wonum_match = wonum_pattern.search(description_text)
                                assetnum_match = assetnum_pattern.search(description_text)
                                worktype_match = worktype_pattern.search(description_text)
                                title_match = title_pattern.search(description_text)
                                
                                # 抽出したグループ1 (値) の前後の空白を削除
                                wonum = wonum_match.group(1).strip() if wonum_match else ""
                                assetnum = assetnum_match.group(1).strip() if assetnum_match else ""
                                worktype = worktype_match.group(1).strip() if worktype_match else ""
                                
                                # 🚨 DESCRIPTION (タイトル) の抽出ロジックを修正
                                if title_match:
                                    description_val = title_match.group(1).strip()
                                else:
                                    # [タイトル: ...] が見つからなかった場合は空欄にする (ご要望通り)
                                    description_val = "" 
                                
                                # SCHEDSTART/SCHEDFINISHの処理（ISO 8601形式で出力）
                                start_time_key = 'date' if 'date' in event.get('start', {}) else 'dateTime'
                                end_time_key = 'date' if 'date' in event.get('end', {}) else 'dateTime'
                                
                                schedstart = event['start'].get(start_time_key, '')
                                schedfinish = event['end'].get(end_time_key, '')
                                
                                # 'dateTime'形式の場合、タイムゾーン付きISO 8601形式 (+09:00) で再フォーマット
                                if start_time_key == 'dateTime':
                                    try:
                                        dt_obj = datetime.fromisoformat(schedstart.replace('Z', '+00:00'))
                                        jst = timezone(timedelta(hours=9))
                                        schedstart = dt_obj.astimezone(jst).isoformat(timespec='seconds')
                                    except ValueError:
                                        pass

                                if end_time_key == 'dateTime':
                                    try:
                                        dt_obj = datetime.fromisoformat(schedfinish.replace('Z', '+00:00'))
                                        jst = timezone(timedelta(hours=9))
                                        schedfinish = dt_obj.astimezone(jst).isoformat(timespec='seconds')
                                    except ValueError:
                                        pass
                                
                                extracted_data.append({
                                    "WONUM": wonum,
                                    "DESCRIPTION": description_val, # 抽出したタイトルまたは空欄
                                    "ASSETNUM": assetnum,
                                    "WORKTYPE": worktype,
                                    "SCHEDSTART": schedstart,
                                    "SCHEDFINISH": schedfinish,
                                    "LEAD": "",
                                    "JESSCHEDFIXED": "",
                                    "SITEID": "JES"
                                })
                            
                            output_df = pd.DataFrame(extracted_data)
                            st.dataframe(output_df)
                            
                            # ダウンロードボタン
                            if export_format == "CSV":
                                csv_buffer = output_df.to_csv(index=False).encode('utf-8-sig') 
                                st.download_button(
                                    label="✅ CSVファイルとしてダウンロード",
                                    data=csv_buffer,
                                    file_name="Googleカレンダー_イベントリスト.csv",
                                    mime="text/csv"
                                )
                            else:
                                buffer = BytesIO()
                                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                    output_df.to_excel(writer, index=False, sheet_name='カレンダーイベント')
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
                        
with st.sidebar:
    # ユーザー設定管理
    st.header("⚙️ ユーザー設定")
    with st.expander("デフォルト設定の管理", expanded=False):
        st.subheader("📅 カレンダー設定")
        
        # デフォルトカレンダーの設定
        if st.session_state.get('editable_calendar_options'):
            calendar_options = list(st.session_state['editable_calendar_options'].keys())
            saved_calendar = get_user_setting(user_id, 'selected_calendar_name')
            
            try:
                default_cal_index = calendar_options.index(saved_calendar) if saved_calendar else 0
            except ValueError:
                default_cal_index = 0
            
            default_calendar = st.selectbox(
                "デフォルトカレンダー",
                calendar_options,
                index=default_cal_index,
                key="sidebar_default_calendar",
                help="イベント登録時に最初から選択されるカレンダー"
            )
            
            # 非公開設定
            saved_private = get_user_setting(user_id, 'default_private_event')
            default_private = st.checkbox(
                "デフォルトで非公開イベント",
                value=saved_private if saved_private is not None else True,
                key="sidebar_default_private",
                help="イベント登録時に非公開が初期選択される"
            )
            
            # 終日イベント設定
            saved_allday = get_user_setting(user_id, 'default_allday_event')
            default_allday = st.checkbox(
                "デフォルトで終日イベント",
                value=saved_allday if saved_allday is not None else False,
                key="sidebar_default_allday",
                help="イベント登録時に終日イベントが初期選択される"
            )
        
        st.subheader("✅ ToDo設定")
        saved_todo = get_user_setting(user_id, 'default_create_todo')
        default_todo = st.checkbox(
            "デフォルトでToDo作成",
            value=saved_todo if saved_todo is not None else False,
            key="sidebar_default_todo",
            help="イベント登録時にToDo作成が初期選択される"
        )
        
        # 設定保存ボタン
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 保存", use_container_width=True, type="primary"):
                if st.session_state.get('editable_calendar_options'):
                    set_user_setting(user_id, 'selected_calendar_name', default_calendar)
                    save_user_setting_to_firestore(user_id, 'selected_calendar_name', default_calendar)
                
                set_user_setting(user_id, 'default_private_event', default_private)
                save_user_setting_to_firestore(user_id, 'default_private_event', default_private)
                
                set_user_setting(user_id, 'default_allday_event', default_allday)
                save_user_setting_to_firestore(user_id, 'default_allday_event', default_allday)
                
                set_user_setting(user_id, 'default_create_todo', default_todo)
                save_user_setting_to_firestore(user_id, 'default_create_todo', default_todo)
                
                st.success("✅ 設定を保存しました")
                st.rerun()
        
        with col2:
            if st.button("🔄 リセット", use_container_width=True):
                # デフォルト設定をクリア
                set_user_setting(user_id, 'default_private_event', None)
                set_user_setting(user_id, 'default_allday_event', None)
                set_user_setting(user_id, 'default_create_todo', None)
                
                save_user_setting_to_firestore(user_id, 'default_private_event', None)
                save_user_setting_to_firestore(user_id, 'default_allday_event', None)
                save_user_setting_to_firestore(user_id, 'default_create_todo', None)
                
                st.info("🔄 設定をリセットしました")
                st.rerun()
        
        # 現在の保存済み設定の表示
        st.divider()
        st.caption("📋 保存済み設定一覧")
        all_settings = get_all_user_settings(user_id)
        if all_settings:
            settings_to_display = {
                'selected_calendar_name': 'デフォルトカレンダー',
                'description_columns_selected': '説明欄の列',
                'event_name_col_selected': 'イベント名の列',
                'add_task_type_to_event_name': '作業タイプ追加',
                'create_todo_checkbox_state': 'ToDo作成',
                'default_private_event': '非公開設定',
                'default_allday_event': '終日イベント',
                'default_create_todo': 'デフォルトToDo'
            }
            
            for key, label in settings_to_display.items():
                if key in all_settings and all_settings[key] is not None:
                    value = all_settings[key]
                    if isinstance(value, bool):
                        value = "✅" if value else "❌"
                    elif isinstance(value, list):
                        value = f"{len(value)}項目"
                    st.text(f"• {label}: {value}")
    
    st.divider()
    
    # 認証状態（コンパクト表示）
    with st.expander("🔐 認証状態", expanded=False):
        st.caption("Firebase: ✅ 認証済み")
        
        if st.session_state.get('calendar_service'):
            st.caption("カレンダー: ✅ 接続中")
        else:
            st.caption("カレンダー: ⚠️ 未接続")
        
        if st.session_state.get('tasks_service'):
            st.caption("ToDo: ✅ 利用可能")
        else:
            st.caption("ToDo: ⚠️ 利用不可")
    
    st.divider()
    
    if st.button("🚪 ログアウト", type="secondary", use_container_width=True):
        if user_id:
            clear_user_settings(user_id)
        for key in list(st.session_state.keys()):
            if not key.startswith("google_auth") and not key.startswith("firebase_"):
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
    
    st.divider()
    
    # 統計情報（一番下に移動）
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    st.metric("アップロード済みファイル", uploaded_count)
