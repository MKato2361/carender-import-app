import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
from excel_parser import process_excel_files
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
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")
st.title("📅 Googleカレンダー一括イベント登録・削除")

# Firebaseの初期化
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

# Firebase認証フォームの表示とユーザーIDの取得
user_id = get_firebase_user_id()

if not user_id:
    # ユーザーがログインしていない場合、認証フォームを表示して停止
    firebase_auth_form()
    st.stop()

# --- ここから下の処理は、Firebase認証が完了した場合にのみ実行されます ---

google_auth_placeholder = st.empty()

with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    # Firebase認証後に、Google認証に進む
    creds = authenticate_google()

    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()
        st.sidebar.success("✅ Googleカレンダーに認証済みです！")

# カレンダーサービスの初期化
def initialize_calendar_service():
    """カレンダーサービスを初期化する"""
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

# タスクサービスの初期化
def initialize_tasks_service():
    """タスクサービスを初期化する"""
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
            
        task_lists = tasks_service.tasklists().list().execute()
        default_task_list_id = None
        
        # デフォルトのタスクリストを探す
        for task_list in task_lists.get('items', []):
            if task_list.get('title') == 'My Tasks':
                default_task_list_id = task_list['id']
                break
                
        # 'My Tasks'が見つからない場合、最初のタスクリストを使用
        if not default_task_list_id and task_lists.get('items'):
            default_task_list_id = task_lists['items'][0]['id']
        
        return tasks_service, default_task_list_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTPエラー): {e}")
        return None, None
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
        return None, None

# カレンダーサービスの初期化または取得
if 'calendar_service' not in st.session_state or not st.session_state['calendar_service']:
    service, editable_calendar_options = initialize_calendar_service()
    
    if not service:
        st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
        st.stop()
    
    st.session_state['calendar_service'] = service
    st.session_state['editable_calendar_options'] = editable_calendar_options
else:
    service = st.session_state['calendar_service']

# タスクサービスの初期化または取得
if 'tasks_service' not in st.session_state:
    tasks_service, default_task_list_id = initialize_tasks_service()
    
    st.session_state['tasks_service'] = tasks_service
    st.session_state['default_task_list_id'] = default_task_list_id
    
    if not tasks_service:
        st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
else:
    tasks_service = st.session_state['tasks_service']

# メインタブの作成
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. イベントの更新"
])

with tabs[0]:
    st.header("📁 ファイルをアップロード")
    
    # ファイルアップロード
    uploaded_files = st.file_uploader(
        "Excelファイルを選択（複数可）", 
        type=["xlsx"], 
        accept_multiple_files=True,
        help="複数のExcelファイルを同時にアップロードできます"
    )
    
    # ファイル処理
    if uploaded_files:
        try:
            st.session_state['uploaded_files'] = uploaded_files
            description_columns_pool = set()
            
            # 各ファイルの列名を収集
            for file in uploaded_files:
                try:
                    df_temp = pd.read_excel(file, engine="openpyxl")
                    df_temp.columns = [str(c).strip() for c in df_temp.columns]
                    description_columns_pool.update(df_temp.columns)
                except Exception as e:
                    st.warning(f"⚠️ {file.name} の読み込みに失敗しました: {e}")
            
            st.session_state['description_columns_pool'] = list(description_columns_pool)
            
            # アップロード済みファイルの表示
            st.success(f"✅ {len(uploaded_files)}個のファイルをアップロードしました")
            
        except Exception as e:
            st.error(f"ファイル処理中にエラーが発生しました: {e}")
            
    elif 'uploaded_files' not in st.session_state:
        st.session_state['uploaded_files'] = []
        st.session_state['description_columns_pool'] = []

    # アップロード済みファイルの表示
    if st.session_state.get('uploaded_files'):
        st.subheader("📋 アップロード済みのファイル:")
        for i, f in enumerate(st.session_state['uploaded_files'], 1):
            st.write(f"{i}. {f.name}")
        
        # ファイルクリアボタン
        if st.button("🗑️ ファイルをクリア", type="secondary"):
            st.session_state['uploaded_files'] = []
            st.session_state['description_columns_pool'] = []
            st.success("ファイルをクリアしました")
            st.rerun()

with tabs[1]:
    st.header("➕ イベントを登録")
    
    # アップロードされたファイルの確認
    if not st.session_state.get('uploaded_files'):
        st.warning("⚠️ まず「ファイルのアップロード」タブでExcelファイルをアップロードしてください。")
    else:
        st.info("📝 イベント登録機能を実装してください")
        # TODO: イベント登録のロジックを実装

with tabs[2]:
    st.header("🗑️ イベントを削除")
    
    st.info("📝 イベント削除機能を実装してください")
    # TODO: イベント削除のロジックを実装

with tabs[3]:
    st.header("🔄 イベントを更新")
    
    st.info("📝 イベント更新機能を実装してください")
    # TODO: イベント更新のロジックを実装

# サイドバーに認証情報表示
with st.sidebar:
    st.header("🔐 認証状態")
    st.success("✅ Firebase認証済み")
    st.success("✅ Google認証済み")
    
    if st.session_state.get('tasks_service'):
        st.success("✅ ToDoリスト利用可能")
    else:
        st.warning("⚠️ ToDoリスト利用不可")
    
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get('uploaded_files', []))
    st.metric("アップロード済みファイル", uploaded_count)
    
    # ログアウトボタン
    if st.button("🚪 ログアウト", type="secondary"):
        # セッション状態をクリア
        for key in list(st.session_state.keys()):
            # 全てのセッション情報をクリアして完全にログアウトする
            if key in st.session_state:
                del st.session_state[key]
        st.success("ログアウトしました")
        st.rerun()
