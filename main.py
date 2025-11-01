# -*- coding: utf-8 -*-
# main.py — 完成版（差分更新 + 集計 + 最小ログ）
# Python 3.11 / Streamlit

from __future__ import annotations
from utils.helpers import safe_get, to_utc_range, default_fetch_window_years
from utils.parsers import extract_worksheet_id_from_text

import re
import unicodedata
from datetime import datetime, date, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from firebase_admin import firestore
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ---- アプリ固有モジュール ----
from excel_parser import (
    process_excel_data_for_calendar,
    _load_and_merge_dataframes,
    get_available_columns_for_event_name,
    check_event_name_columns,
    format_worksheet_value,
)
from calendar_utils import (
    authenticate_google,
    add_event_to_calendar,
    fetch_all_events,
    update_event_if_needed,   # ← calendar_utils.py を差分版に差し替え済み
    build_tasks_service,
    add_task_to_todo_list,
    find_and_delete_tasks_by_event_id,
)
from firebase_auth import initialize_firebase, firebase_auth_form, get_firebase_user_id
from session_utils import (
    initialize_session_state,
    get_user_setting,
    set_user_setting,
    get_all_user_settings,
    clear_user_settings,
)

# ==================================================
# 0) スタイル
# ==================================================
st.set_page_config(page_title="Googleカレンダー一括イベント登録・削除", layout="wide")

def load_custom_css() -> None:
    try:
        with open("custom_sidebar.css", "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass

load_custom_css()

st.markdown(
    """
<style>
@media (prefers-color-scheme: light) {
    .header-bar { background-color: rgba(249, 249, 249, 0.95); color: #333; border-bottom: 1px solid #ccc; }
}
@media (prefers-color-scheme: dark) {
    .header-bar { background-color: rgba(30, 30, 30, 0.9); color: #eee; border-bottom: 1px solid #444; }
}
.header-bar { position: sticky; top: 0; width: 100%; text-align: center; font-weight: 500;
    font-size: 14px; padding: 8px 0; z-index: 20; backdrop-filter: blur(6px); }
div[data-testid="stTabs"] { position: sticky; top: 42px; z-index: 15; background-color: inherit;
    border-bottom: 1px solid rgba(128, 128, 128, 0.3); padding-top: 4px; padding-bottom: 4px;
    backdrop-filter: blur(6px); }
.block-container, section[data-testid="stMainBlockContainer"], main {
    padding-top: 0 !important; padding-bottom: 0 !important; margin-bottom: 0 !important;
    height: auto !important; min-height: 100vh !重要; overflow: visible !重要; }
footer, div[data-testid="stBottomBlockContainer"] { display: none !重要; height: 0 !重要; margin: 0 !重要; padding: 0 !重要; }
html, body, #root { height: auto !重要; min-height: 100% !重要; margin: 0 !重要; padding: 0 !重要;
    overflow-x: hidden !重要; overflow-y: auto !重要; overscroll-behavior: none !重要; -webkit-overflow-scrolling: touch !重要; }
</style>
<div class="header-bar">📅 Googleカレンダー一括イベント登録・削除</div>
""",
    unsafe_allow_html=True,
)

# ==================================================
# 1) 共通ユーティリティ
# ==================================================
JST = timezone(timedelta(hours=9))

# 正規表現（事前コンパイル）
RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")
RE_WONUM      = re.compile(r"\[作業指示書[：:]\s*(.*?)\]")
RE_ASSETNUM   = re.compile(r"\[管理番号[：:]\s*(.*?)\]")
RE_WORKTYPE   = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
RE_TITLE      = re.compile(r"\[タイトル[：:]\s*(.*?)\]")

# --- 差分更新ユーティリティ ---
def normalize_worksheet_id(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    return unicodedata.normalize("NFKC", s).strip()

def extract_worksheet_id_from_description(desc: str) -> str | None:
    """Description内の [作業指示書: 123456] からIDを抽出（全角→半角）"""
    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return normalize_worksheet_id(m.group(1))

def is_event_changed(existing_event: dict, new_event_data: dict) -> bool:
    """
    1) summary（タイトル）
    2) start（終日/時間/TimeZone含む）
    3) end   （終日/時間/TimeZone含む）
    4) description（説明）
    5) transparency（非公開/公開）
    ※ Location は比較しない
    """
    nz = lambda v: (v or "")
    # 1) summary
    if nz(existing_event.get("summary")) != nz(new_event_data.get("summary")):
        return True
    # 4) description
    if nz(existing_event.get("description")) != nz(new_event_data.get("description")):
        return True
    # 5) transparency
    if nz(existing_event.get("transparency")) != nz(new_event_data.get("transparency")):
        return True
    # 2) start
    if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
        return True
    # 3) end
    if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
        return True
    return False

def to_utc_range(d1: date, d2: date) -> Tuple[str, str]:
    start_dt_utc = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
    end_dt_utc   = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
    return (
        start_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
        end_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
    )

def default_fetch_window_years(years: int = 2) -> Tuple[str, str]:
    now_utc = datetime.now(timezone.utc)
    return (now_utc - timedelta(days=365 * years)).isoformat(), (now_utc + timedelta(days=365 * years)).isoformat()


def build_calendar_service(creds):
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()
        editable = {cal["summary"]: cal["id"] for cal in calendar_list.get("items", []) if cal.get("accessRole") != "reader"}
        return service, editable
    except HttpError as e:
        st.error(f"カレンダーサービスの初期化に失敗しました (HTTP): {e}")
    except Exception as e:
        st.error(f"カレンダーサービスの初期化に失敗しました: {e}")
    return None, None

def build_tasks_service_safe(creds):
    try:
        tasks_service = build_tasks_service(creds)
        if not tasks_service:
            return None, None
        task_lists = tasks_service.tasklists().list().execute()
        default_id = None
        for item in task_lists.get("items", []):
            if item.get("title") == "My Tasks":
                default_id = item["id"]
                break
        if not default_id and task_lists.get("items"):
            default_id = task_lists["items"][0]["id"]
        return tasks_service, default_id
    except HttpError as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました (HTTP): {e}")
    except Exception as e:
        st.warning(f"Google ToDoリストサービスの初期化に失敗しました: {e}")
    return None, None

def ensure_services(creds):
    if "calendar_service" not in st.session_state or not st.session_state["calendar_service"]:
        service, editable = build_calendar_service(creds)
        if not service:
            st.warning("Google認証の状態を確認するか、ページをリロードしてください。")
            st.stop()
        st.session_state["calendar_service"] = service
        st.session_state["editable_calendar_options"] = editable
    if "tasks_service" not in st.session_state or not st.session_state.get("tasks_service"):
        tasks_service, default_task_list_id = build_tasks_service_safe(creds)
        st.session_state["tasks_service"] = tasks_service
        st.session_state["default_task_list_id"] = default_task_list_id
        if not tasks_service:
            st.info("ToDoリスト機能は利用できませんが、カレンダー機能は引き続き使用できます。")
    return st.session_state["calendar_service"], st.session_state["editable_calendar_options"]

# ==================================================
# 2) Firebase 認証
# ==================================================
if not initialize_firebase():
    st.error("Firebaseの初期化に失敗しました。")
    st.stop()

db = firestore.client()
user_id = get_firebase_user_id()
if not user_id:
    firebase_auth_form()
    st.stop()

def load_user_settings_from_firestore(user_id: str) -> None:
    if not user_id:
        return
    initialize_session_state(user_id)
    doc = db.collection("user_settings").document(user_id).get()
    if doc.exists:
        for key, value in doc.to_dict().items():
            set_user_setting(user_id, key, value)

def save_user_setting_to_firestore(user_id: str, setting_key: str, setting_value) -> None:
    if not user_id:
        return
    try:
        db.collection("user_settings").document(user_id).set({setting_key: setting_value}, merge=True)
    except Exception as e:
        st.error(f"設定の保存に失敗しました: {e}")

load_user_settings_from_firestore(user_id)

# ==================================================
# 3) Google 認証
# ==================================================
google_auth_placeholder = st.empty()
with google_auth_placeholder.container():
    st.subheader("🔐 Googleカレンダー認証")
    creds = authenticate_google()
    if not creds:
        st.warning("Googleカレンダー認証を完了してください。")
        st.stop()
    else:
        google_auth_placeholder.empty()

service, editable_calendar_options = ensure_services(creds)

# ==================================================
# 4) UI（Tabs）
# ==================================================
st.markdown('<div class="fixed-tabs">', unsafe_allow_html=True)
tabs = st.tabs([
    "1. ファイルのアップロード",
    "2. イベントの登録",
    "3. イベントの削除",
    "4. 重複イベントの検出・削除",
    "5. イベントのExcel出力",
])
st.markdown("</div>", unsafe_allow_html=True)

if "uploaded_files" not in st.session_state:
    st.session_state["uploaded_files"] = []
    st.session_state["description_columns_pool"] = []
    st.session_state["merged_df_for_selector"] = pd.DataFrame()

# ==================================================
# 5) タブ1: ファイルのアップロード
# ==================================================
with tabs[0]:
    st.subheader("ファイルをアップロード")
    with st.expander("ℹ️作業手順と補足"):
        st.info(
            """
**☀作業指示書一覧をアップロードすると管理番号+物件名をイベント名として任意のカレンダーに登録します。**
**☀イベントの説明欄に含めたい情報はドロップダウンリストから選択してください。（複数選択可,次回から同じ項目が選択されます）**
**☀イベントに住所を追加したい場合は、物件一覧のファイルを作業指示書一覧と一緒にアップロードしてください。**
**☀作業外予定の一覧をアップロードすると、イベント名を選択することができます。**
**☀ToDoリストを作成すると、点検通知のリマインドが可能です（ToDoとしてイベント登録されます）**
"""
        )

    def get_local_excel_files() -> List[Path]:
        current_dir = Path(__file__).parent
        return [f for f in current_dir.glob("*") if f.suffix.lower() in [".xlsx", ".xls", ".csv"]]

    uploaded_files = st.file_uploader(
        "ExcelまたはCSVファイルを選択（複数可）", type=["xlsx", "xls", "csv"], accept_multiple_files=True
    )

    local_excel_files = get_local_excel_files()
    selected_local_files: List[BytesIO] = []
    if local_excel_files:
        st.markdown("📁 サーバーにあるExcelファイル")
        local_file_names = [f.name for f in local_excel_files]
        selected_names = st.multiselect("以下のファイルを処理対象に含める（アップロードと同様に扱われます）", local_file_names)
        for name in selected_names:
            full_path = next((f for f in local_excel_files if f.name == name), None)
            if full_path:
                with open(full_path, "rb") as f:
                    file_obj = BytesIO(f.read())
                    file_obj.name = name
                    selected_local_files.append(file_obj)

    all_files: List = []
    if uploaded_files:
        all_files.extend(uploaded_files)
    if selected_local_files:
        all_files.extend(selected_local_files)

    if all_files:
        st.session_state["uploaded_files"] = all_files
        try:
            merged = _load_and_merge_dataframes(all_files)
            st.session_state["merged_df_for_selector"] = merged
            st.session_state["description_columns_pool"] = merged.columns.tolist()
            if merged.empty:
                st.warning("読み込まれたファイルに有効なデータがありませんでした。")
        except (ValueError, IOError) as e:
            st.error(f"ファイルの読み込みに失敗しました: {e}")
            st.session_state["uploaded_files"] = []
            st.session_state["merged_df_for_selector"] = pd.DataFrame()
            st.session_state["description_columns_pool"] = []

    if st.session_state.get("uploaded_files"):
        st.subheader("📄 処理対象ファイル一覧")
        for f in st.session_state["uploaded_files"]:
            st.write(f"- {getattr(f, 'name', '不明な名前のファイル')}")
        if not st.session_state["merged_df_for_selector"].empty:
            st.info(
                f"📊 データ列数: {len(st.session_state['merged_df_for_selector'].columns)}、"
                f"行数: {len(st.session_state['merged_df_for_selector'])}"
            )
        if st.button("🗑️ アップロード済みファイルをクリア", help="選択中のファイルとデータを削除します。"):
            st.session_state["uploaded_files"] = []
            st.session_state["merged_df_for_selector"] = pd.DataFrame()
            st.session_state["description_columns_pool"] = []
            st.success("すべてのファイル情報をクリアしました。")
            st.rerun()

# ==================================================
# 6) タブ2: イベントの登録・更新（差分更新＋集計）
# ==================================================
with tabs[1]:
    st.subheader("イベントを登録・更新")

    description_columns: List[str] = []
    selected_event_name_col: Optional[str] = None
    add_task_type_to_event_name = False
    all_day_event_override = False
    private_event = True
    fallback_event_name_column: Optional[str] = None

    if not st.session_state.get("uploaded_files") or st.session_state["merged_df_for_selector"].empty:
        st.info("先に「1. ファイルのアップロード」タブでExcelファイルをアップロードしてください。")

    elif not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")

    else:
        calendar_options = list(editable_calendar_options.keys())
        saved_calendar_name = get_user_setting(user_id, "selected_calendar_name")
        try:
            default_index = calendar_options.index(saved_calendar_name)
        except Exception:
            default_index = 0

        selected_calendar_name = st.selectbox(
            "登録先カレンダーを選択", calendar_options, index=default_index, key="reg_calendar_select"
        )
        calendar_id = editable_calendar_options[selected_calendar_name]

        set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)
        save_user_setting_to_firestore(user_id, "selected_calendar_name", selected_calendar_name)

        description_columns_pool = st.session_state.get("description_columns_pool", [])
        saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
        saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
        saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")
        saved_create_todo_flag = get_user_setting(user_id, "create_todo_checkbox_state")

        expand_event_setting = not bool(saved_description_cols)
        expand_name_setting = not (saved_event_name_col or saved_task_type_flag)
        expand_todo_setting = bool(saved_create_todo_flag)

        with st.expander("📝 イベント設定", expanded=expand_event_setting):
            all_day_event_override = st.checkbox("終日イベントとして登録", value=False)
            private_event = st.checkbox("非公開イベントとして登録", value=True)
            default_selection = [col for col in saved_description_cols if col in description_columns_pool]
            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                description_columns_pool,
                default=default_selection,
                key=f"description_selector_register_{user_id}",
            )

        with st.expander("🧱 イベント名の生成設定", expanded=expand_name_setting):
            has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
            selected_event_name_col = saved_event_name_col
            add_task_type_to_event_name = st.checkbox(
                "イベント名の先頭に作業タイプを追加する",
                value=bool(saved_task_type_flag),
                key=f"add_task_type_checkbox_{user_id}",
            )
            if not (has_mng_data and has_name_data):
                available_event_name_cols = get_available_columns_for_event_name(
                    st.session_state["merged_df_for_selector"]
                )
                event_name_options = ["選択しない"] + available_event_name_cols
                try:
                    name_index = event_name_options.index(selected_event_name_col) if selected_event_name_col else 0
                except Exception:
                    name_index = 0
                selected_event_name_col = st.selectbox(
                    "イベント名として使用する代替列を選択してください:",
                    options=event_name_options,
                    index=name_index,
                    key=f"event_name_selector_register_{user_id}",
                )
                if selected_event_name_col != "選択しない":
                    fallback_event_name_column = selected_event_name_col
            else:
                st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

        st.subheader("✅ ToDoリスト連携設定 (オプション)")
        with st.expander("ToDoリスト作成オプション", expanded=expand_todo_setting):
            create_todo = st.checkbox(
                "このイベントに対応するToDoリストを作成する",
                value=bool(saved_create_todo_flag),
                key="create_todo_checkbox",
            )
            set_user_setting(user_id, "create_todo_checkbox_state", create_todo)
            save_user_setting_to_firestore(user_id, "create_todo_checkbox_state", create_todo)

            fixed_todo_types = ["点検通知"]
            if create_todo:
                st.markdown(f"以下のToDoが**常にすべて**作成されます: `{', '.join(fixed_todo_types)}`")
            else:
                st.markdown("ToDoリストの作成は無効です。")

            deadline_offset_options = {"2週間前": 14, "10日前": 10, "1週間前": 7, "カスタム日数前": None}
            selected_offset_key = st.selectbox(
                "ToDoリストの期限をイベント開始日の何日前に設定しますか？",
                list(deadline_offset_options.keys()),
                disabled=not create_todo,
                key="deadline_offset_select",
            )
            custom_offset_days = None
            if selected_offset_key == "カスタム日数前":
                custom_offset_days = st.number_input(
                    "何日前に設定しますか？ (日数)",
                    min_value=0,
                    value=3,
                    disabled=not create_todo,
                    key="custom_offset_input",
                )

        st.subheader("➡️ イベント登録・更新実行")
        if st.button("Googleカレンダーに登録・更新する"):
            set_user_setting(user_id, "description_columns_selected", description_columns)
            set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
            set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

            save_user_setting_to_firestore(user_id, "description_columns_selected", description_columns)
            save_user_setting_to_firestore(user_id, "event_name_col_selected", selected_event_name_col)
            save_user_setting_to_firestore(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

            with st.spinner("イベントデータを処理中..."):
                try:
                    df = process_excel_data_for_calendar(
                        st.session_state["uploaded_files"],
                        description_columns,
                        all_day_event_override,
                        private_event,
                        fallback_event_name_column,
                        add_task_type_to_event_name,
                    )
                except (ValueError, IOError) as e:
                    st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
                    df = pd.DataFrame()

                if df.empty:
                    st.warning("有効なイベントデータがありません。処理を中断しました。")
                else:
                    st.info(f"{len(df)} 件のイベントを処理します。")
                    progress = st.progress(0)

                    # 集計カウンタ
                    added_count = 0
                    updated_count = 0
                    skipped_count = 0

                    # 既存イベント（期間=±2年）取得 → {worksheet_id: event}
                    time_min, time_max = default_fetch_window_years(2)
                    with st.spinner("既存イベントを取得中..."):
                        events = fetch_all_events(service, calendar_id, time_min, time_max)

                    worksheet_to_event: Dict[str, dict] = {}
                    for event in events or []:
                        wid = extract_worksheet_id_from_description(event.get("description") or "")
                        if wid:
                            worksheet_to_event[wid] = event

                    total = len(df)
                    for i, row in df.iterrows():
                        desc_text = safe_get(row, "Description", "")
                        worksheet_id = extract_worksheet_id_from_text(desc_text)
                        
                        all_day_flag  = safe_get(row, "All Day Event", "True")
                        private_flag  = safe_get(row, "Private", "True")
                        start_date_str = safe_get(row, "Start Date", "")
                        end_date_str   = safe_get(row, "End Date", "")
                        start_time_str = safe_get(row, "Start Time", "")
                        end_time_str   = safe_get(row, "End Time", "")
                        
                        event_data = {
                            "summary":   safe_get(row, "Subject", ""),
                            "location":  safe_get(row, "Location", ""),  # 保存はするが比較対象外
                            "description": desc_text,
                            "transparency": "transparent" if private_flag == "True" else "opaque",
                        }


                        try:
                            if all_day_flag == "True":
                                sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                                ed = datetime.strptime(end_date_str, "%Y/%m/%d").date()
                                event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                                event_data["end"]   = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
                            else:
                                sdt = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                                edt = datetime.strptime(f"{end_date_str} {end_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                                event_data["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
                                event_data["end"]   = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
                        except Exception as e:
                            st.error(f"行 {i} の日時パースに失敗しました: {e}")
                            progress.progress((i + 1) / total)
                            continue

                        existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

                        try:
                            if existing_event:
                                # 事前に差分チェック → 差分あるときのみ更新 & カウント
                                if is_event_changed(existing_event, event_data):
                                    _ = update_event_if_needed(service, calendar_id, existing_event["id"], event_data)
                                    updated_count += 1
                                else:
                                    skipped_count += 1
                            else:
                                added_event = add_event_to_calendar(service, calendar_id, event_data)
                                if added_event:
                                    added_count += 1
                                    if worksheet_id:
                                        worksheet_to_event[worksheet_id] = added_event
                        except Exception as e:
                            st.error(f"イベント '{event_data.get('summary','(無題)')}' の登録/更新に失敗しました: {e}")

                        progress.progress((i + 1) / total)

                    st.success(f"✅ 登録: {added_count} / 🔧 更新: {updated_count} / ↪ スキップ: {skipped_count}")

# ==================================================
# 7) タブ3: イベントの削除（仕様変更なし）
# ==================================================
with tabs[2]:
    st.subheader("イベントを削除")

    if not editable_calendar_options:
        st.error("削除可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
    else:
        calendar_names = list(editable_calendar_options.keys())
        default_index = 0
        saved_name = st.session_state.get("selected_calendar_name")
        if saved_name and saved_name in calendar_names:
            default_index = calendar_names.index(saved_name)

        selected_calendar_name_del = st.selectbox("削除対象カレンダーを選択", calendar_names, index=default_index, key="del_calendar_select")
        st.session_state["selected_calendar_name"] = selected_calendar_name_del
        calendar_id_del = editable_calendar_options[selected_calendar_name_del]

        st.subheader("🗓️ 削除期間の選択")
        today_date = date.today()
        delete_start_date = st.date_input("削除開始日", value=today_date - timedelta(days=30))
        delete_end_date   = st.date_input("削除終了日", value=today_date)
        delete_related_todos = st.checkbox("関連するToDoリストも削除する (イベント詳細にIDが記載されている場合)", value=False)

        if delete_start_date > delete_end_date:
            st.error("削除開始日は終了日より前に設定してください。")
        else:
            st.subheader("🗑️ 削除実行")
            if "confirm_delete" not in st.session_state:
                st.session_state["confirm_delete"] = False
            if not st.session_state["confirm_delete"]:
                if st.button("選択期間のイベントを削除する", type="primary"):
                    st.session_state["confirm_delete"] = True
                    st.rerun()

            if st.session_state["confirm_delete"]:
                st.warning(
                    f"""
⚠️ **削除確認**

以下のイベントを削除します:
- **カレンダー名**: {selected_calendar_name_del}
- **期間**: {delete_start_date.strftime('%Y年%m月%d日')} ～ {delete_end_date.strftime('%Y年%m月%d日')}
- **ToDoリストも削除**: {'はい' if delete_related_todos else 'いいえ'}

この操作は取り消せません。本当に削除しますか？
"""
                )
                col1, col2 = st.columns([1, 1])

                def to_utc_range_btn(d1: date, d2: date):
                    sdt = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
                    edt = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
                    return (
                        sdt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
                        edt.isoformat(timespec="microseconds").replace("+00:00", "Z"),
                    )

                with col1:
                    if st.button("✅ 実行", type="primary", use_container_width=True):
                        st.session_state["confirm_delete"] = False
                        time_min_utc, time_max_utc = to_utc_range_btn(delete_start_date, delete_end_date)
                        events_to_delete = fetch_all_events(service, calendar_id_del, time_min_utc, time_max_utc)

                        if not events_to_delete:
                            st.info("指定期間内に削除するイベントはありませんでした。")

                        deleted_events_count = 0
                        deleted_todos_count = 0
                        total_events = len(events_to_delete or [])

                        if total_events > 0:
                            progress_bar = st.progress(0)
                            status_text = st.empty()

                            for i, event in enumerate(events_to_delete, start=1):
                                event_summary = event.get("summary", "不明なイベント")
                                event_id = event["id"]
                                status_text.text(f"イベント '{event_summary}' を削除中... ({i}/{total_events})")
                                try:
                                    if delete_related_todos and st.session_state.get("tasks_service") and st.session_state.get("default_task_list_id"):
                                        deleted_task_count_for_event = find_and_delete_tasks_by_event_id(
                                            st.session_state["tasks_service"],
                                            st.session_state["default_task_list_id"],
                                            event_id,
                                        )
                                        deleted_todos_count += deleted_task_count_for_event
                                    service.events().delete(calendarId=calendar_id_del, eventId=event_id).execute()
                                    deleted_events_count += 1
                                except Exception as e:
                                    st.error(f"イベント '{event_summary}' (ID: {event_id}) の削除に失敗しました: {e}")
                                progress_bar.progress(i / total_events)

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
                        st.session_state["confirm_delete"] = False
                        st.rerun()

# ==================================================
# 8) タブ4: 重複イベントの検出・削除（現行踏襲）
# ==================================================
with tabs[3]:
    st.subheader("🔍 重複イベントの検出・削除")

    if "last_dup_message" in st.session_state and st.session_state["last_dup_message"]:
        msg_type, msg_text = st.session_state["last_dup_message"]
        if msg_type in {"success", "error", "info", "warning"}:
            getattr(st, msg_type)(msg_text)
        else:
            st.info(msg_text)
        st.session_state["last_dup_message"] = None

    calendar_options = list(editable_calendar_options.keys())
    selected_calendar = st.selectbox("対象カレンダーを選択", calendar_options, key="dup_calendar_select")
    calendar_id = editable_calendar_options[selected_calendar]

    delete_mode = st.radio("削除モードを選択", ["手動で選択して削除", "古い方を自動削除", "新しい方を自動削除"], horizontal=True, key="dup_delete_mode")

    if "dup_df" not in st.session_state:
        st.session_state["dup_df"] = pd.DataFrame()
    if "auto_delete_ids" not in st.session_state:
        st.session_state["auto_delete_ids"] = []
    if "last_dup_message" not in st.session_state:
        st.session_state["last_dup_message"] = None

    def parse_created(dt_str: Optional[str]) -> datetime:
        try:
            if dt_str:
                return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        except Exception:
            pass
        return datetime.min.replace(tzinfo=timezone.utc)

    if st.button("重複イベントをチェック", key="run_dup_check"):
        with st.spinner("カレンダー内のイベントを取得中..."):
            time_min, time_max = default_fetch_window_years(2)
            events = fetch_all_events(service, calendar_id, time_min, time_max)

        if not events:
            st.session_state["last_dup_message"] = ("info", "イベントが見つかりませんでした。")
            st.session_state["dup_df"] = pd.DataFrame()
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode
            st.rerun()

        st.success(f"{len(events)} 件のイベントを取得しました。")

        rows = []
        for e in events:
            desc = (e.get("description") or "").strip()
            m = RE_WORKSHEET_ID.search(desc)
            worksheet_id = normalize_worksheet_id(m.group(1)) if m else None
            start_time = e["start"].get("dateTime", e["start"].get("date"))
            end_time   = e["end"].get("dateTime", e["end"].get("date"))
            rows.append({
                "id": e["id"],
                "summary": e.get("summary", ""),
                "worksheet_id": worksheet_id,
                "created": e.get("created"),
                "start": start_time,
                "end": end_time,
            })

        df = pd.DataFrame(rows)
        df_valid = df[df["worksheet_id"].notna()].copy()
        dup_mask = df_valid.duplicated(subset=["worksheet_id"], keep=False)
        dup_df = df_valid[dup_mask].sort_values(["worksheet_id", "created"])

        st.session_state["dup_df"] = dup_df
        if dup_df.empty:
            st.session_state["last_dup_message"] = ("info", "重複している作業指示書番号は見つかりませんでした。")
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode
            st.rerun()

        if delete_mode != "手動で選択して削除":
            auto_delete_ids: List[str] = []
            for _, group in dup_df.groupby("worksheet_id"):
                group_sorted = group.sort_values(
                    ["created", "id"],
                    key=lambda s: s.map(parse_created) if s.name == "created" else s,
                    ascending=True,
                )
                if len(group_sorted) <= 1:
                    continue
                if delete_mode == "古い方を自動削除":
                    delete_targets = group_sorted.iloc[:-1]
                elif delete_mode == "新しい方を自動削除":
                    delete_targets = group_sorted.iloc[1:]
                else:
                    continue
                auto_delete_ids.extend(delete_targets["id"].tolist())

            st.session_state["auto_delete_ids"] = auto_delete_ids
            st.session_state["current_delete_mode"] = delete_mode
        else:
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode

        st.rerun()

    if not st.session_state["dup_df"].empty:
        dup_df = st.session_state["dup_df"]
        current_mode = st.session_state.get("current_delete_mode", "手動で選択して削除")

        st.warning(f"⚠️ {dup_df['worksheet_id'].nunique()} 種類の重複作業指示書が見つかりました。（合計 {len(dup_df)} イベント）")
        st.dataframe(dup_df[["worksheet_id", "summary", "created", "start", "end", "id"]], use_container_width=True)

        if current_mode == "手動で選択して削除":
            delete_ids = st.multiselect("削除するイベントを選択してください（イベントIDで指定）", dup_df["id"].tolist(), key="manual_delete_ids")
            confirm = st.checkbox("削除操作を確認しました", value=False, key="manual_del_confirm")

            if st.button("🗑️ 選択したイベントを削除", type="primary", disabled=not confirm, key="run_manual_delete"):
                deleted_count = 0
                errors: List[str] = []
                for eid in delete_ids:
                    try:
                        service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                        deleted_count += 1
                    except Exception as e:
                        errors.append(f"イベントID {eid} の削除に失敗: {e}")

                if deleted_count > 0:
                    st.session_state["last_dup_message"] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")

                if errors:
                    st.error("以下のイベントの削除に失敗しました:\n" + "\n".join(errors))
                    if deleted_count == 0:
                        st.session_state["last_dup_message"] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                st.session_state["dup_df"] = pd.DataFrame()
                st.rerun()

        else:
            auto_delete_ids = st.session_state["auto_delete_ids"]
            if not auto_delete_ids:
                st.info("削除対象のイベントが見つかりませんでした。")
            else:
                st.warning(f"以下のモードで {len(auto_delete_ids)} 件のイベントを自動削除します: **{current_mode}**")
                st.write(auto_delete_ids)

                confirm = st.checkbox("削除操作を確認しました", value=False, key="auto_del_confirm_final")
                if st.button("🗑️ 自動削除を実行", type="primary", disabled=not confirm, key="run_auto_delete"):
                    deleted_count = 0
                    errors: List[str] = []
                    for eid in auto_delete_ids:
                        try:
                            service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                            deleted_count += 1
                        except Exception as e:
                            errors.append(f"イベントID {eid} の削除に失敗: {e}")

                    if deleted_count > 0:
                        st.session_state["last_dup_message"] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")

                    if errors:
                        st.error("以下のイベントの削除に失敗しました:\n" + "\n".join(errors))
                        if deleted_count == 0:
                            st.session_state["last_dup_message"] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                    st.session_state["dup_df"] = pd.DataFrame()
                    st.rerun()

# ==================================================
# 9) タブ5: カレンダーイベントをExcel/CSVへ出力
# ==================================================
with tabs[4]:
    st.subheader("カレンダーイベントをExcelに出力")

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。")
    else:
        selected_calendar_name_export = st.selectbox("出力対象カレンダーを選択", list(editable_calendar_options.keys()), key="export_calendar_select")
        calendar_id_export = editable_calendar_options[selected_calendar_name_export]

        st.subheader("🗓️ 出力期間の選択")
        today_date_export = date.today()
        export_start_date = st.date_input("出力開始日", value=today_date_export - timedelta(days=30))
        export_end_date   = st.date_input("出力終了日", value=today_date_export)
        export_format = st.radio("出力形式を選択", ("CSV", "Excel"), index=0)

        if export_start_date > export_end_date:
            st.error("出力開始日は終了日より前に設定してください。")
        else:
            if st.button("指定期間のイベントを読み込む"):
                with st.spinner("イベントを読み込み中..."):
                    try:
                        time_min_utc, time_max_utc = to_utc_range(export_start_date, export_end_date)
                        events_to_export = fetch_all_events(service, calendar_id_export, time_min_utc, time_max_utc)

                        if not events_to_export:
                            st.info("指定期間内にイベントは見つかりませんでした。")
                        else:
                            extracted_data: List[dict] = []
                            for event in events_to_export:
                                description_text = event.get("description", "") or ""
                                wonum_match   = RE_WONUM.search(description_text)
                                assetnum_match= RE_ASSETNUM.search(description_text)
                                worktype_match= RE_WORKTYPE.search(description_text)
                                title_match   = RE_TITLE.search(description_text)

                                wonum = (wonum_match.group(1).strip() if wonum_match else "") or ""
                                assetnum = (assetnum_match.group(1).strip() if assetnum_match else "") or ""
                                worktype = (worktype_match.group(1).strip() if worktype_match else "") or ""
                                description_val = title_match.group(1).strip() if title_match else ""

                                start_time = event["start"].get("dateTime") or event["start"].get("date") or ""
                                end_time   = event["end"].get("dateTime") or event["end"].get("date") or ""

                                def to_jst_iso(s: str) -> str:
                                    try:
                                        if "T" in s and ("+" in s or s.endswith("Z")):
                                            dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
                                            return dt.isoformat(timespec="seconds")
                                    except Exception:
                                        pass
                                    return s

                                schedstart  = to_jst_iso(start_time)
                                schedfinish = to_jst_iso(end_time)

                                extracted_data.append({
                                    "WONUM": wonum,
                                    "DESCRIPTION": description_val,
                                    "ASSETNUM": assetnum,
                                    "WORKTYPE": worktype,
                                    "SCHEDSTART": schedstart,
                                    "SCHEDFINISH": schedfinish,
                                    "LEAD": "",
                                    "JESSCHEDFIXED": "",
                                    "SITEID": "JES",
                                })

                            output_df = pd.DataFrame(extracted_data)
                            st.dataframe(output_df)

                            if export_format == "CSV":
                                csv_buffer = output_df.to_csv(index=False).encode("utf-8-sig")
                                st.download_button(
                                    label="✅ CSVファイルとしてダウンロード",
                                    data=csv_buffer,
                                    file_name="Googleカレンダー_イベントリスト.csv",
                                    mime="text/csv",
                                )
                            else:
                                buffer = BytesIO()
                                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                                    output_df.to_excel(writer, index=False, sheet_name="カレンダーイベント")
                                buffer.seek(0)
                                st.download_button(
                                    label="✅ Excelファイルとしてダウンロード",
                                    data=buffer,
                                    file_name="Googleカレンダー_イベントリスト.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )

                            st.success(f"{len(output_df)} 件のイベントを読み込みました。")
                    except Exception as e:
                        st.error(f"イベントの読み込み中にエラーが発生しました: {e}")

# ==================================================
# 10) サイドバー
# ==================================================
with st.sidebar:
    with st.expander("⚙ デフォルト設定の管理", expanded=False):
        st.subheader("📅 カレンダー設定")
        if editable_calendar_options:
            calendar_options = list(editable_calendar_options.keys())
            saved_calendar = get_user_setting(user_id, "selected_calendar_name")
            try:
                default_cal_index = calendar_options.index(saved_calendar) if saved_calendar else 0
            except ValueError:
                default_cal_index = 0

            default_calendar = st.selectbox("デフォルトカレンダー", calendar_options, index=default_cal_index, key="sidebar_default_calendar")

            prev_share = st.session_state.get("share_calendar_selection_across_tabs", True)
            share_calendar = st.checkbox(
                "カレンダー選択をタブ間で共有する",
                value=prev_share,
                help="ON: 登録タブで選んだカレンダーが他タブに自動反映 / OFF: タブごとに独立",
            )
            if share_calendar != prev_share:
                st.session_state["share_calendar_selection_across_tabs"] = share_calendar
                set_user_setting(user_id, "share_calendar_selection_across_tabs", share_calendar)
                save_user_setting_to_firestore(user_id, "share_calendar_selection_across_tabs", share_calendar)
                st.success("🔄 共有設定を保存しました（更新します）")
                st.rerun()

            saved_private = get_user_setting(user_id, "default_private_event")
            default_private = st.checkbox("デフォルトで非公開イベント", value=(saved_private if saved_private is not None else True), key="sidebar_default_private")

            saved_allday = get_user_setting(user_id, "default_allday_event")
            default_allday = st.checkbox("デフォルトで終日イベント", value=(saved_allday if saved_allday is not None else False), key="sidebar_default_allday")

        st.subheader("✅ ToDo設定")
        saved_todo = get_user_setting(user_id, "default_create_todo")
        default_todo = st.checkbox("デフォルトでToDo作成", value=(saved_todo if saved_todo is not None else False), key="sidebar_default_todo")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 保存", use_container_width=True):
                if editable_calendar_options:
                    set_user_setting(user_id, "selected_calendar_name", default_calendar)
                    save_user_setting_to_firestore(user_id, "selected_calendar_name", default_calendar)
                    st.session_state["selected_calendar_name"] = default_calendar
                    if st.session_state.get("share_calendar_selection_across_tabs", True):
                        for k in ["register", "delete", "dup", "export"]:
                            st.session_state[f"selected_calendar_name_{k}"] = default_calendar

                set_user_setting(user_id, "default_private_event", default_private)
                save_user_setting_to_firestore(user_id, "default_private_event", default_private)

                set_user_setting(user_id, "default_allday_event", default_allday)
                save_user_setting_to_firestore(user_id, "default_allday_event", default_allday)

                set_user_setting(user_id, "default_create_todo", default_todo)
                save_user_setting_to_firestore(user_id, "default_create_todo", default_todo)

                st.success("✅ 設定を保存しました")
                st.rerun()

        with col2:
            if st.button("🔄 リセット", use_container_width=True):
                for key in ["default_private_event", "default_allday_event", "default_create_todo"]:
                    set_user_setting(user_id, key, None)
                    save_user_setting_to_firestore(user_id, key, None)
                st.info("🧹 設定をリセットしました")
                st.rerun()

        st.divider()
        st.caption("📋 保存済み設定")
        all_settings = get_all_user_settings(user_id)
        if all_settings:
            labels = {
                "selected_calendar_name": "デフォルトカレンダー（共有ON時）",
                "default_private_event": "非公開設定",
                "default_allday_event": "終日設定",
                "default_create_todo": "デフォルトToDo",
                "share_calendar_selection_across_tabs": "タブ間共有",
            }
            for k, label in labels.items():
                if k in all_settings and all_settings[k] is not None:
                    v = all_settings[k]
                    if isinstance(v, bool):
                        v = "✅" if v else "❌"
                    st.text(f"• {label}: {v}")

    st.divider()
    with st.expander("🔐 認証状態", expanded=False):
        st.caption("Firebase: ✅ 認証済み")
        st.caption("カレンダー: ✅ 接続中" if st.session_state.get("calendar_service") else "カレンダー: ⚠️ 未接続")
        st.caption("ToDo: ✅ 利用可能" if st.session_state.get("tasks_service") else "ToDo: ⚠️ 利用不可")

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
    st.header("📊 統計情報")
    uploaded_count = len(st.session_state.get("uploaded_files", []))
    st.metric("アップロード済みファイル", uploaded_count)
