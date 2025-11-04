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
from github_loader import walk_repo_tree, load_file_bytes_from_github, is_supported_file
from github_loader import _headers, GITHUB_OWNER, GITHUB_REPO
from io import BytesIO
from streamlit.runtime.uploaded_file_manager import UploadedFile

def convert_bytes_to_uploadedfile(file_bytes: bytes, filename: str, mime_type: str = None):
    """GitHub等から取得したバイトデータをStreamlitのUploadedFile互換に変換"""
    return UploadedFile(
        name=filename,
        type=mime_type or "application/octet-stream",
        data=file_bytes,
    )
import streamlit as st
import pandas as pd
from io import BytesIO
# ←このあたりの import 群の直下に追加してください。
from tabs.tab1_upload import render_tab1_upload
class GitHubUploadedFile:
    def __init__(self, file_bytes: bytes, name: str, type: str = None):
        self._file_bytes = file_bytes
        self.name = name
        self.type = type or "application/octet-stream"

    def read(self):
        return self._file_bytes

    def getvalue(self):
        return self._file_bytes


def convert_bytes_to_uploadedfile(file_bytes: bytes, filename: str, mime_type: str = None):
    return GitHubUploadedFile(
        file_bytes=file_bytes,
        name=filename,
        type=mime_type or "application/octet-stream",
    )

from tabs.tab2_register import render_tab2_register
from tabs.tab3_delete import render_tab3_delete
from tabs.tab4_duplicates import render_tab4_duplicates
from calendar_utils import fetch_all_events



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
tasks_service = st.session_state.get("tasks_service")
default_task_list_id = st.session_state.get("default_task_list_id")

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
# 5) タブ1: ファイルのアップロード（修正版）
# ==================================================
with tabs[0]:
    render_tab1_upload()

# ==================================================
# 6) タブ2: イベントの登録・更新（差分更新＋集計）
# ==================================================
with tabs[1]:
    render_tab2_register(user_id, editable_calendar_options, service, tasks_service, default_task_list_id)


# ==================================================
# 7) タブ3: イベントの削除（仕様変更なし）
# ==================================================
with tabs[2]:
    render_tab3_delete(editable_calendar_options, service, tasks_service, default_task_list_id)

# ==================================================
# 8) タブ4: 重複イベントの検出・削除（現行踏襲）
# ==================================================
with tabs[3]:
    render_tab4_duplicates(service, editable_calendar_options, fetch_all_events)



# ==================================================
# 9) タブ5: カレンダーイベントをExcel/CSVへ出力（安全ファイル名版）
# ==================================================
with tabs[4]:
    st.subheader("カレンダーイベントをExcelに出力")

    import re
    import unicodedata

    def safe_filename(name: str) -> str:
        """日本語保持・全角→半角・禁止文字除去の安全ファイル名生成"""
        name = unicodedata.normalize("NFKC", name)  # 全角→半角
        name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', '', name)  # 禁止文字除去
        name = name.strip(" .")  # 先頭末尾 . と空白除去
        return name or "output"

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。")
    else:
        selected_calendar_name_export = st.selectbox(
            "出力対象カレンダーを選択",
            list(editable_calendar_options.keys()),
            key="export_calendar_select"
        )
        calendar_id_export = editable_calendar_options[selected_calendar_name_export]

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
                        time_min_utc, time_max_utc = to_utc_range(export_start_date, export_end_date)
                        events_to_export = fetch_all_events(service, calendar_id_export, time_min_utc, time_max_utc)

                        if not events_to_export:
                            st.info("指定期間内にイベントは見つかりませんでした。")
                        else:
                            extracted_data: List[dict] = []
                            for event in events_to_export:
                                description_text = event.get("description", "") or ""
                                wonum_match = RE_WONUM.search(description_text)
                                assetnum_match = RE_ASSETNUM.search(description_text)
                                worktype_match = RE_WORKTYPE.search(description_text)
                                title_match = RE_TITLE.search(description_text)

                                wonum = (wonum_match.group(1).strip() if wonum_match else "") or ""
                                assetnum = (assetnum_match.group(1).strip() if assetnum_match else "") or ""
                                worktype = (worktype_match.group(1).strip() if worktype_match else "") or ""
                                description_val = title_match.group(1).strip() if title_match else ""

                                start_time = event["start"].get("dateTime") or event["start"].get("date") or ""
                                end_time = event["end"].get("dateTime") or event["end"].get("date") or ""

                                def to_jst_iso(s: str) -> str:
                                    try:
                                        if "T" in s and ("+" in s or s.endswith("Z")):
                                            dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
                                            return dt.isoformat(timespec="seconds")
                                    except Exception:
                                        pass
                                    return s

                                schedstart = to_jst_iso(start_time)
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

                            # 🔥 安全ファイル名生成
                            start_str = export_start_date.strftime("%Y%m%d")
                            end_str = export_end_date.strftime("%m%d")
                            safe_cal_name = safe_filename(selected_calendar_name_export)
                            file_base_name = f"{safe_cal_name}_{start_str}_{end_str}"

                            if export_format == "CSV":
                                csv_buffer = output_df.to_csv(index=False).encode("utf-8-sig")
                                st.download_button(
                                    label="✅ CSVファイルとしてダウンロード",
                                    data=csv_buffer,
                                    file_name=f"{file_base_name}.csv",
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
                                    file_name=f"{file_base_name}.xlsx",
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
