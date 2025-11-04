import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Dict, List, Optional
from utils.helpers import safe_get
from utils.parsers import extract_worksheet_id_from_text
from excel_parser import (
    process_excel_data_for_calendar,
    get_available_columns_for_event_name,
    check_event_name_columns,
)
from calendar_utils import (
    fetch_all_events,
    add_event_to_calendar,
    update_event_if_needed,
)
from session_utils import (
    get_user_setting,
    set_user_setting,
)
from firebase_admin import firestore

JST = ZoneInfo("Asia/Tokyo")


def is_event_changed(existing_event: dict, new_event_data: dict) -> bool:
    nz = lambda v: (v or "")
    if nz(existing_event.get("summary")) != nz(new_event_data.get("summary")):
        return True
    if nz(existing_event.get("description")) != nz(new_event_data.get("description")):
        return True
    if nz(existing_event.get("transparency")) != nz(new_event_data.get("transparency")):
        return True
    if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
        return True
    if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
        return True
    return False


def default_fetch_window_years(years: int = 2):
    now_utc = datetime.now(tz=ZoneInfo("UTC"))
    return (
        (now_utc - timedelta(days=365 * years)).isoformat(),
        (now_utc + timedelta(days=365 * years)).isoformat(),
    )


def extract_worksheet_id_from_description(desc: str) -> str | None:
    import re
    import unicodedata
    RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")
    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return unicodedata.normalize("NFKC", m.group(1)).strip()


def _to_dt(val: str) -> Optional[datetime]:
    if not val or not isinstance(val, str):
        return None
    s = val.strip().replace("T", " ").replace(".", " ").replace("　", " ")
    s = s.replace("/", "-")

    # Try strict ISO parse first
    try:
        ts = pd.to_datetime(s, utc=True)
        return ts.tz_convert(JST).to_pydatetime()
    except Exception:
        pass

    # Try non-UTC parse
    try:
        ts = pd.to_datetime(s)
        if ts.tzinfo is None:
            ts = ts.tz_localize(JST)
        else:
            ts = ts.tz_convert(JST)
        return ts.to_pydatetime()
    except Exception:
        pass

    # Try defined formats
    fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
    ]
    for fmt in fmts:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.replace(tzinfo=JST)
        except Exception:
            continue
    return None


def _split_dt_cell(val: str) -> tuple[str, str]:
    if isinstance(val, datetime):
    dt = val.astimezone(JST) if val.tzinfo else val.replace(tzinfo=JST)
else:
    dt = _to_dt(val)

    if not dt:
        return "", ""
    return dt.strftime("%Y/%m/%d"), dt.strftime("%H:%M")


def _normalize_minute_str(dt_like: datetime | str) -> str:
    if isinstance(dt_like, str):
        d = _to_dt(dt_like)
    else:
        d = dt_like
    if not d:
        return ""
    d = d.astimezone(JST)
    return d.strftime("%Y-%m-%dT%H:%M")


def _normalize_event_times_to_key(start_dict: dict, end_dict: dict) -> tuple[str, str]:
    def norm_one(d: dict) -> str:
        if not d:
            return ""
        if "dateTime" in d and d["dateTime"]:
            return _normalize_minute_str(d["dateTime"])
        if "date" in d and d["date"]:
            try:
                sd = datetime.strptime(d["date"], "%Y-%m-%d").replace(tzinfo=JST)
                return sd.strftime("%Y-%m-%d")
            except Exception:
                return d["date"]
        return ""
    return norm_one(start_dict), norm_one(end_dict)


def _normalize_row_times_to_key(row: dict, all_day_flag: str) -> tuple[str, str]:
    if all_day_flag == "True":
        try:
            sd = datetime.strptime(row.get("Start Date", ""), "%Y/%m/%d").date().strftime("%Y-%m-%d")
            ed = datetime.strptime(row.get("End Date", "") or row.get("Start Date", ""), "%Y/%m/%d").date().strftime("%Y-%m-%d")
            return sd, ed
        except Exception:
            return row.get("Start Date", ""), row.get("End Date", "") or row.get("Start Date", "")

    try:
        sdt = datetime.strptime(f"{row.get('Start Date', '')} {row.get('Start Time', '')}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
        edt = datetime.strptime(f"{row.get('End Date', '') or row.get('Start Date', '')} {row.get('End Time', '') or row.get('Start Time', '')}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
        return sdt.strftime("%Y-%m-%dT%H:%M"), edt.strftime("%Y-%m-%dT%H:%M")
    except Exception:
        return row.get("Start Date", ""), row.get("End Date", "") or row.get("Start Date", "")


def _strip_outside_suffix(subject: str) -> str:
    s = subject or ""
    suf = " [作業外予定]"
    return s[:-len(suf)].rstrip() if s.endswith(suf) else s


def _read_outside_file_to_df(file_obj) -> pd.DataFrame:
    name = getattr(file_obj, "name", "")
    
    # 1) Excelはdtype=objectで読み込み、datetime型を保持
    if name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj, dtype=object)
    else:
        # 2) CSVは文字化け対策しつつdtype=objectで読み込み
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                df = pd.read_csv(file_obj, dtype=object, encoding=enc, errors="ignore")
                break
            except Exception:
                df = None
        if df is None:
            raise ValueError("CSVの読み込みに失敗しました（対応エンコーディング不明）。")

    # ❗ ここがポイント：datetime列を文字列化しない
    # fillna("") は object列のみに限定し、datetime列は触らない
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("")

    return df

def _build_calendar_df_from_outside(df_raw: pd.DataFrame, private_event: bool, all_day_override: bool) -> pd.DataFrame:
    if "備考" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『備考』列が見つかりません。")
    if "理由コード" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『理由コード』列が見つかりません。")

    col_start_dt = "開始日時" if "開始日時" in df_raw.columns else None
    col_end_dt = "終了日時" if "終了日時" in df_raw.columns else None

    start_date_candidates = ["開始日", "日付", "Start Date", "Date"]
    end_date_candidates = ["終了日", "End Date", "Date2"]
    start_time_candidates = ["開始時刻", "開始時間", "Start Time"]
    end_time_candidates = ["終了時刻", "終了時間", "End Time"]

    def pick(col_names):
        for c in col_names:
            if c in df_raw.columns:
                return c
        return None

    c_sd = pick(start_date_candidates)
    c_ed = pick(end_date_candidates)
    c_st = pick(start_time_candidates)
    c_et = pick(end_time_candidates)

    def fix_hhmm(t: str) -> str:
        t = (t or "").strip().replace(".", ":")
        if t.isdigit() and len(t) in (3, 4):
            t = t.zfill(4)
            return f"{t[:2]}:{t[2:]}"
        return t

    rows = []
    for _, r in df_raw.iterrows():
        subject = f"{str(r['備考']).strip()} [作業外予定]".strip()
        description = str(r["理由コード"]).strip()

        if col_start_dt and col_end_dt:
            sd, stime = _split_dt_cell(str(r[col_start_dt]))
            ed, etime = _split_dt_cell(str(r[col_end_dt]))
        else:
            def get(c): return (str(r[c]).strip() if c and c in r and pd.notna(r[c]) else "")
            sd = get(c_sd).replace("-", "/")
            ed = get(c_ed).replace("-", "/") or sd
            stime = fix_hhmm(get(c_st))
            etime = fix_hhmm(get(c_et))

        all_day = "True" if all_day_override else "False"

        if all_day != "True":
            if not sd:
                all_day = "True"
            elif not stime and not etime:
                all_day = "True"
            else:
                if stime and not etime:
                    try:
                        dt = datetime.strptime(stime, "%H:%M")
                        etime = (dt + timedelta(hours=1)).strftime("%H:%M")
                    except Exception:
                        all_day = "True"
                if etime and not stime:
                    try:
                        dt = datetime.strptime(etime, "%H:%M")
                        stime = (dt - timedelta(hours=1)).strftime("%H:%M")
                    except Exception:
                        all_day = "True"

        rows.append(
            {
                "Subject": subject,
                "Description": description,
                "All Day Event": all_day,
                "Private": "True" if private_event else "False",
                "Start Date": sd or "",
                "End Date": (ed or sd or ""),
                "Start Time": stime or "",
                "End Time": etime or "",
                "Location": "",
            }
        )

    return pd.DataFrame(rows)


def render_tab2_register(user_id: str, editable_calendar_options: dict, service, tasks_service=None, default_task_list_id=None):
    st.subheader("イベントを登録・更新")

    work_files = st.session_state.get("uploaded_files") or []
    has_work = bool(work_files) and st.session_state.get("merged_df_for_selector") is not None and not st.session_state["merged_df_for_selector"].empty

    outside_file = st.session_state.get("uploaded_outside_work_file")
    outside_mode = bool(outside_file) and not has_work

    if not has_work and not outside_mode:
        st.info("先に「1. ファイルのアップロード」タブでファイルをアップロードしてください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    if outside_mode:
        saved_calendar_name = get_user_setting(user_id, "selected_calendar_name_outside")
    else:
        saved_calendar_name = get_user_setting(user_id, "selected_calendar_name")

    try:
        default_index = calendar_options.index(saved_calendar_name)
    except Exception:
        default_index = 0

    selected_calendar_name = st.selectbox(
        "登録先カレンダーを選択" + ("（作業外予定）" if outside_mode else "（作業指示書）"),
        calendar_options,
        index=default_index,
        key="reg_calendar_select_outside" if outside_mode else "reg_calendar_select",
    )
    calendar_id = editable_calendar_options[selected_calendar_name]

    if outside_mode:
        set_user_setting(user_id, "selected_calendar_name_outside", selected_calendar_name)
    else:
        set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)

    with st.expander("📝 イベント設定", expanded=not outside_mode):
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False, key=f"all_day_override_{'outside' if outside_mode else 'work'}")
        private_event = st.checkbox("非公開イベントとして登録", value=True, key=f"private_event_{'outside' if outside_mode else 'work'}")

        if outside_mode:
            description_columns = []
        else:
            description_columns_pool = st.session_state.get("description_columns_pool", [])
            saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
            default_selection = [col for col in saved_description_cols if col in description_columns_pool]

            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                description_columns_pool,
                default=default_selection,
                key=f"description_selector_register_{user_id}",
            )
            set_user_setting(user_id, "description_columns_selected", description_columns)

    if outside_mode:
        st.info("イベント名は『備考 + [作業外予定]』で登録します。")
        add_task_type_to_event_name = False
        fallback_event_name_column = None
    else:
        with st.expander("🧱 イベント名の生成設定", expanded=True):
            has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
            saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
            saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")

            add_task_type_to_event_name = st.checkbox(
                "イベント名の先頭に作業タイプを追加する",
                value=bool(saved_task_type_flag),
                key=f"add_task_type_checkbox_{user_id}",
            )

            fallback_event_name_column = None
            if not (has_mng_data and has_name_data):
                available_event_name_cols = get_available_columns_for_event_name(st.session_state["merged_df_for_selector"])
                event_name_options = ["選択しない"] + available_event_name_cols

                try:
                    name_index = event_name_options.index(saved_event_name_col) if saved_event_name_col else 0
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

                set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
            else:
                st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

            set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

    if not outside_mode:
        st.subheader("✅ ToDoリスト連携設定 (オプション)")
        with st.expander("ToDoリスト作成オプション", expanded=False):
            create_todo = st.checkbox(
                "このイベントに対応するToDoリストを作成する",
                value=bool(get_user_setting(user_id, "create_todo_checkbox_state")),
                key="create_todo_checkbox",
            )
            set_user_setting(user_id, "create_todo_checkbox_state", create_todo)

    st.subheader("➡️ イベント登録・更新実行")
    if not st.button("Googleカレンダーに登録・更新する"):
        return

    try:
        if outside_mode:
            raw_df = _read_outside_file_to_df(outside_file)
            df = _build_calendar_df_from_outside(raw_df, private_event=private_event, all_day_override=all_day_event_override)
        else:
            df = process_excel_data_for_calendar(
                st.session_state["uploaded_files"],
                description_columns,
                all_day_event_override,
                private_event,
                fallback_event_name_column,
                add_task_type_to_event_name,
            )
    except Exception as e:
        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
        return

    if df.empty:
        st.warning("有効なイベントデータがありません。処理を中断しました。")
        return

    st.info(f"{len(df)} 件のイベントを処理します。")
    progress = st.progress(0)

    added_count = 0
    updated_count = 0
    skipped_count = 0

    time_min, time_max = default_fetch_window_years(2)
    with st.spinner("既存イベントを取得中..."):
        events = fetch_all_events(service, calendar_id, time_min, time_max)

    worksheet_to_event: Dict[str, dict] = {}
    for event in events or []:
        wid = extract_worksheet_id_from_description(event.get("description") or "")
        if wid:
            worksheet_to_event[wid] = event

    outside_key_to_event: Dict[str, dict] = {}
    if outside_mode:
        for ev in events or []:
            summ = ev.get("summary") or ""
            core = _strip_outside_suffix(summ)
            if not core:
                continue
            s_key, e_key = _normalize_event_times_to_key(ev.get("start") or {}, ev.get("end") or {})
            if not s_key or not e_key:
                continue
            key = f"{core}|{s_key}|{e_key}"
            outside_key_to_event[key] = ev
    total = len(df)

    for i, row in df.iterrows():
        desc_text = safe_get(row, "Description", "")
        subject = safe_get(row, "Subject", "")
        all_day_flag = safe_get(row, "All Day Event", "True" if outside_mode else "True")
        private_flag = safe_get(row, "Private", "True")

        start_date_str = safe_get(row, "Start Date", "")
        end_date_str = safe_get(row, "End Date", "")
        start_time_str = safe_get(row, "Start Time", "")
        end_time_str = safe_get(row, "End Time", "")

        event_data = {
            "summary": subject,
            "location": safe_get(row, "Location", ""),
            "description": desc_text,
            "transparency": "transparent" if private_flag == "True" else "opaque",
        }

        try:
            if all_day_flag == "True":
                sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                ed = datetime.strptime(end_date_str or start_date_str, "%Y/%m/%d").date()
                event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                event_data["end"] = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
            else:
                sdt = datetime.strptime(
                    f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M"
                ).replace(tzinfo=JST)
                edt = datetime.strptime(
                    f"{end_date_str or start_date_str} {end_time_str or start_time_str}", "%Y/%m/%d %H:%M"
                ).replace(tzinfo=JST)

                event_data["start"] = {
                    "dateTime": sdt.isoformat(),
                    "timeZone": "Asia/Tokyo",
                }
                event_data["end"] = {
                    "dateTime": edt.isoformat(),
                    "timeZone": "Asia/Tokyo",
                }
        except Exception as e:
            st.error(f"行 {i} の日時パースに失敗しました: {e}")
            progress.progress((i + 1) / total)
            continue

        if outside_mode:
            core = _strip_outside_suffix(subject)
            row_s_key, row_e_key = _normalize_row_times_to_key(
                {
                    "Start Date": start_date_str,
                    "End Date": end_date_str,
                    "Start Time": start_time_str,
                    "End Time": end_time_str,
                },
                all_day_flag,
            )
            key = f"{core}|{row_s_key}|{row_e_key}"
            existing_event = outside_key_to_event.get(key)
        else:
            worksheet_id = extract_worksheet_id_from_text(desc_text)
            existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

        try:
            if existing_event:
                if is_event_changed(existing_event, event_data):
                    _ = update_event_if_needed(service, calendar_id, existing_event["id"], event_data)
                    updated_count += 1
                else:
                    skipped_count += 1
            else:
                added_event = add_event_to_calendar(service, calendar_id, event_data)
                if added_event:
                    added_count += 1
                    if outside_mode:
                        s_key, e_key = _normalize_event_times_to_key(
                            added_event.get("start") or {}, added_event.get("end") or {}
                        )
                        outside_key_to_event[f"{core}|{s_key}|{e_key}"] = added_event
                    else:
                        worksheet_id = extract_worksheet_id_from_text(desc_text)
                        if worksheet_id:
                            worksheet_to_event[worksheet_id] = added_event
        except Exception as e:
            st.error(f"イベント '{event_data.get('summary', '(無題)')}' の登録/更新に失敗しました: {e}")

        progress.progress((i + 1) / total)

    st.success(
        f"✅ 登録: {added_count} 件 / 🔧 更新: {updated_count} 件 / ↪ スキップ: {skipped_count} 件 処理完了！"
    )
