import re
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Dict, Optional

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

JST = ZoneInfo("Asia/Tokyo")

# ---- 定数 ----
_ALL_DAY_TRUE = "True"
_PRIVATE_TRUE = "True"


# ============================================================
# イベント比較
# ============================================================

def _normalize_time_dict(d: dict) -> str:
    """start/end 辞書を分単位の文字列に正規化（比較用）"""
    if not d:
        return ""
    if d.get("dateTime"):
        return _normalize_minute_str(d["dateTime"])
    if d.get("date"):
        try:
            return datetime.strptime(d["date"], "%Y-%m-%d").replace(tzinfo=JST).strftime("%Y-%m-%d")
        except Exception:
            return d["date"]
    return ""


def is_event_changed(existing_event: dict, new_event_data: dict) -> bool:
    nz = lambda v: (v or "")
    for field in ("summary", "description", "location", "visibility", "transparency"):
        if nz(existing_event.get(field)) != nz(new_event_data.get(field)):
            return True
    # 正規化して比較（タイムゾーン表記のゆらぎを吸収）
    if _normalize_time_dict(existing_event.get("start") or {}) != _normalize_time_dict(new_event_data.get("start") or {}):
        return True
    if _normalize_time_dict(existing_event.get("end") or {}) != _normalize_time_dict(new_event_data.get("end") or {}):
        return True
    return False


# ============================================================
# フェッチ期間計算
# ============================================================

def default_fetch_window_years(years: int = 2):
    now_utc = datetime.now(tz=ZoneInfo("UTC"))
    return (
        (now_utc - timedelta(days=365 * years)).isoformat(),
        (now_utc + timedelta(days=365 * years)).isoformat(),
    )


def compute_fetch_window_from_df(df: pd.DataFrame, buffer_days: int = 30):
    """DFのStart/End Date列からイベント取得範囲を最小化する。"""
    if df is None or df.empty:
        return None
    try:
        s = pd.to_datetime(df.get("Start Date"), format="%Y/%m/%d", errors="coerce")
        e = pd.to_datetime(df.get("End Date"), format="%Y/%m/%d", errors="coerce")
        e = e.fillna(s)
        s_min, e_max = s.min(), e.max()
        if pd.isna(s_min) or pd.isna(e_max):
            return None

        min_date = s_min.date() - timedelta(days=buffer_days)
        max_date = e_max.date() + timedelta(days=buffer_days)
        time_min_dt = datetime.combine(min_date, datetime.min.time()).replace(tzinfo=JST)
        time_max_dt = datetime.combine(max_date + timedelta(days=1), datetime.min.time()).replace(tzinfo=JST)
        return (time_min_dt.isoformat(), time_max_dt.isoformat())
    except Exception as ex:
        st.warning(f"イベント取得期間の計算に失敗しました（デフォルト範囲を使用します）: {ex}")
        return None


# ============================================================
# 日時ユーティリティ
# ============================================================

def extract_worksheet_id_from_description(desc: str) -> Optional[str]:
    import unicodedata
    RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")
    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return unicodedata.normalize("NFKC", m.group(1)).strip()


def _to_dt(val: str) -> Optional[datetime]:
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None

    s = s.replace("T", " ").replace("　", " ").replace("/", "-").replace(".", " ")
    tz_suffix = bool(re.search(r'(Z|[+-]\d{2}:?\d{2})$', s))

    if tz_suffix:
        try:
            ts = pd.to_datetime(s, utc=True, errors="raise")
            return ts.tz_convert(JST).to_pydatetime()
        except Exception:
            pass

    try:
        ts = pd.to_datetime(s, errors="raise")
        if ts.tzinfo is None:
            ts = ts.tz_localize(JST)
        else:
            ts = ts.tz_convert(JST)
        return ts.to_pydatetime()
    except Exception:
        pass

    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=JST)
        except Exception:
            continue
    return None


def _split_dt_cell(val) -> tuple:
    if isinstance(val, datetime):
        dt = val.astimezone(JST) if val.tzinfo else val.replace(tzinfo=JST)
    else:
        dt = _to_dt(val)
    if not dt:
        return "", ""
    return dt.strftime("%Y/%m/%d"), dt.strftime("%H:%M")


def _normalize_minute_str(dt_like) -> str:
    d = _to_dt(dt_like) if isinstance(dt_like, str) else dt_like
    if not d:
        return ""
    return d.astimezone(JST).strftime("%Y-%m-%dT%H:%M")


def _normalize_event_times_to_key(start_dict: dict, end_dict: dict) -> tuple:
    return _normalize_time_dict(start_dict), _normalize_time_dict(end_dict)


def _normalize_row_times_to_key(row: dict, all_day_flag: str) -> tuple:
    if all_day_flag == _ALL_DAY_TRUE:
        try:
            sd = datetime.strptime(row.get("Start Date", ""), "%Y/%m/%d").date().strftime("%Y-%m-%d")
            ed = datetime.strptime(
                row.get("End Date", "") or row.get("Start Date", ""), "%Y/%m/%d"
            ).date().strftime("%Y-%m-%d")
            return sd, ed
        except Exception:
            return row.get("Start Date", ""), row.get("End Date", "") or row.get("Start Date", "")

    try:
        sdt = datetime.strptime(
            f"{row.get('Start Date', '')} {row.get('Start Time', '')}", "%Y/%m/%d %H:%M"
        ).replace(tzinfo=JST)
        edt = datetime.strptime(
            f"{row.get('End Date', '') or row.get('Start Date', '')} {row.get('End Time', '') or row.get('Start Time', '')}",
            "%Y/%m/%d %H:%M",
        ).replace(tzinfo=JST)
        return sdt.strftime("%Y-%m-%dT%H:%M"), edt.strftime("%Y-%m-%dT%H:%M")
    except Exception:
        return row.get("Start Date", ""), row.get("End Date", "") or row.get("Start Date", "")


def _strip_outside_suffix(subject: str) -> str:
    s = subject or ""
    suf = " [作業外予定]"
    return s[: -len(suf)].rstrip() if s.endswith(suf) else s


# ============================================================
# 作業外予定ファイル読み込み
# ============================================================

def _pick_column(df: pd.DataFrame, candidates: list) -> Optional[str]:
    """候補列名リストから最初にヒットした列名を返す"""
    return next((c for c in candidates if c in df.columns), None)


def _read_outside_file_to_df(file_obj) -> pd.DataFrame:
    name = getattr(file_obj, "name", "")
    if name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj, dtype=object)
    else:
        df = None
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                df = pd.read_csv(file_obj, dtype=object, encoding=enc, errors="ignore")
                break
            except Exception:
                pass
        if df is None:
            raise ValueError("CSV読み込み失敗")

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("")
    return df


def _build_calendar_df_from_outside(df_raw: pd.DataFrame, private_event: bool, all_day_override: bool) -> pd.DataFrame:
    for required in ("備考", "理由コード"):
        if required not in df_raw.columns:
            raise ValueError(f"作業外予定ファイルに『{required}』列が見つかりません。")

    col_start_dt = "開始日時" if "開始日時" in df_raw.columns else None
    col_end_dt = "終了日時" if "終了日時" in df_raw.columns else None

    c_sd = _pick_column(df_raw, ["開始日", "日付", "Start Date", "Date"])
    c_ed = _pick_column(df_raw, ["終了日", "End Date", "Date2"])
    c_st = _pick_column(df_raw, ["開始時刻", "開始時間", "Start Time"])
    c_et = _pick_column(df_raw, ["終了時刻", "終了時間", "End Time"])

    def fix_hhmm(t: str) -> str:
        """'900' → '09:00' のような数字のみ時刻文字列を HH:MM 形式に変換する"""
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
            sd, stime = _split_dt_cell(r[col_start_dt])
            ed, etime = _split_dt_cell(r[col_end_dt])
        else:
            def get(c):
                return str(r[c]).strip().replace("-", "/") if c and c in r and pd.notna(r[c]) else ""
            sd = get(c_sd)
            ed = get(c_ed) or sd
            stime = fix_hhmm(get(c_st))
            etime = fix_hhmm(get(c_et))

        all_day = _ALL_DAY_TRUE if all_day_override else "False"
        if all_day != _ALL_DAY_TRUE:
            if not sd or (not stime and not etime):
                all_day = _ALL_DAY_TRUE
            else:
                if stime and not etime:
                    try:
                        etime = (datetime.strptime(stime, "%H:%M") + timedelta(hours=1)).strftime("%H:%M")
                    except Exception:
                        all_day = _ALL_DAY_TRUE
                elif etime and not stime:
                    try:
                        stime = (datetime.strptime(etime, "%H:%M") - timedelta(hours=1)).strftime("%H:%M")
                    except Exception:
                        all_day = _ALL_DAY_TRUE

        rows.append({
            "Subject": subject,
            "Description": description,
            "All Day Event": all_day,
            "Private": _PRIVATE_TRUE if private_event else "False",
            "Start Date": sd or "",
            "End Date": ed or sd or "",
            "Start Time": stime or "",
            "End Time": etime or "",
            "Location": "",
        })

    return pd.DataFrame(rows)


# ============================================================
# 設定保存コールバック
# ============================================================

def _save_calendar_selection(user_id: str, outside_mode: bool):
    key = "reg_calendar_select_outside" if outside_mode else "reg_calendar_select"
    setting_key = "selected_calendar_name_outside" if outside_mode else "selected_calendar_name"
    if key in st.session_state:
        set_user_setting(user_id, setting_key, st.session_state[key])
        st.toast("✅ カレンダー選択を保存しました", icon="📅")


def _save_description_settings(user_id: str):
    desc_key = f"description_selector_register_{user_id}"
    desc_order_key = f"description_order_register_{user_id}"

    if desc_key not in st.session_state:
        return

    description_columns_pool = st.session_state.get("description_columns_pool", [])
    # valid_selected を以降の処理の基準として使用
    valid_selected = [col for col in st.session_state[desc_key] if col in description_columns_pool]

    current_order = st.session_state.get(desc_order_key, [])
    current_order = [c for c in current_order if c in valid_selected]
    for c in valid_selected:
        if c not in current_order:
            current_order.append(c)
    st.session_state[desc_order_key] = current_order

    set_user_setting(user_id, "description_columns_selected", current_order)
    st.toast("✅ 説明欄の設定を保存しました", icon="💾")


def _save_event_name_settings(user_id: str):
    chk_key = f"add_task_type_checkbox_{user_id}"
    if chk_key in st.session_state:
        set_user_setting(user_id, "add_task_type_to_event_name", st.session_state[chk_key])

    sel_key = f"event_name_selector_register_{user_id}"
    if sel_key in st.session_state:
        selected = st.session_state[sel_key]
        set_user_setting(user_id, "event_name_col_selected", None if selected == "選択しない" else selected)

    st.toast("✅ イベント名の生成設定を保存しました", icon="💾")


# ============================================================
# UI サブコンポーネント
# ============================================================

def _render_calendar_selector(
    user_id: str,
    calendar_options: list,
    base_calendar: str,
    outside_mode: bool,
) -> str:
    """登録先カレンダー選択UIを描画し、選択されたカレンダー名を返す"""
    select_key = "reg_calendar_select_outside" if outside_mode else "reg_calendar_select"
    setting_key = "selected_calendar_name_outside" if outside_mode else "selected_calendar_name"

    if (select_key not in st.session_state) or (st.session_state.get(select_key) not in calendar_options):
        saved = get_user_setting(user_id, setting_key)
        st.session_state[select_key] = saved if saved in calendar_options else base_calendar

    st.selectbox(
        "登録先カレンダーを選択" + ("（作業外予定）" if outside_mode else "（作業指示書）"),
        calendar_options,
        key=select_key,
        on_change=_save_calendar_selection,
        args=(user_id, outside_mode),
    )
    return st.session_state[select_key]


def _render_event_settings(user_id: str, outside_mode: bool) -> tuple:
    """イベント設定 Expander を描画し (all_day_override, private_event, description_columns) を返す"""
    with st.expander("📝 イベント設定", expanded=not outside_mode):
        all_day_override = st.checkbox(
            "終日イベントとして登録",
            value=False,
            key=f"all_day_override_{'outside' if outside_mode else 'work'}",
        )
        private_event = st.checkbox(
            "非公開イベントとして登録",
            value=True,
            key=f"private_event_{'outside' if outside_mode else 'work'}",
        )

        if outside_mode:
            return all_day_override, private_event, []

        description_columns_pool = st.session_state.get("description_columns_pool", [])
        saved_cols = get_user_setting(user_id, "description_columns_selected") or []
        default_selection = [c for c in saved_cols if c in description_columns_pool]

        desc_key = f"description_selector_register_{user_id}"
        desc_order_key = f"description_order_register_{user_id}"

        if desc_key not in st.session_state:
            st.session_state[desc_key] = list(default_selection)
        else:
            st.session_state[desc_key] = [c for c in st.session_state[desc_key] if c in description_columns_pool]

        st.multiselect(
            "説明欄に含める列（複数選択可）",
            description_columns_pool,
            key=desc_key,
            on_change=_save_description_settings,
            args=(user_id,),
        )
        selected_cols = st.session_state.get(desc_key, [])

        if not selected_cols:
            st.session_state.pop(desc_order_key, None)
            return all_day_override, private_event, []

        st.caption("↕️ ドラッグして説明欄の列の順番を変更できます")
        current_order = st.session_state.get(desc_order_key, [])
        current_order = [c for c in current_order if c in selected_cols]
        for c in selected_cols:
            if c not in current_order:
                current_order.append(c)

        edited_df = st.data_editor(
            pd.DataFrame({"列名（説明欄への出力順）": current_order}),
            num_rows="fixed",
            hide_index=False,
            use_container_width=True,
            column_config={
                "列名（説明欄への出力順）": st.column_config.SelectboxColumn(
                    "列名（説明欄への出力順）",
                    options=selected_cols,
                    required=True,
                )
            },
            key=f"{desc_order_key}_editor",
        )

        new_order = list(dict.fromkeys(
            c for c in edited_df["列名（説明欄への出力順）"].tolist() if c in selected_cols
        ))
        st.session_state[desc_order_key] = new_order
        return all_day_override, private_event, new_order


def _render_event_name_settings(user_id: str) -> tuple:
    """イベント名生成設定 Expander を描画し (add_task_type, fallback_col) を返す"""
    with st.expander("🧱 イベント名の生成設定", expanded=True):
        has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
        saved_col = get_user_setting(user_id, "event_name_col_selected")
        saved_flag = get_user_setting(user_id, "add_task_type_to_event_name")

        st.checkbox(
            "イベント名の先頭に作業タイプを追加する",
            value=bool(saved_flag),
            key=f"add_task_type_checkbox_{user_id}",
            on_change=_save_event_name_settings,
            args=(user_id,),
        )
        add_task_type = st.session_state[f"add_task_type_checkbox_{user_id}"]

        fallback_col = None
        if not (has_mng_data and has_name_data):
            available_cols = get_available_columns_for_event_name(st.session_state["merged_df_for_selector"])
            options = ["選択しない"] + available_cols
            idx = options.index(saved_col) if saved_col in options else 0

            st.selectbox(
                "イベント名として使用する代替列を選択してください:",
                options=options,
                index=idx,
                key=f"event_name_selector_register_{user_id}",
                on_change=_save_event_name_settings,
                args=(user_id,),
            )
            sel = st.session_state[f"event_name_selector_register_{user_id}"]
            if sel != "選択しない":
                fallback_col = sel
        else:
            st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")

        return add_task_type, fallback_col


# ============================================================
# 登録・更新ループ
# ============================================================

def _execute_registration(
    service,
    df: pd.DataFrame,
    calendar_id: str,
    outside_mode: bool,
):
    """Googleカレンダーへのイベント登録・更新を実行する"""
    progress = st.progress(0)
    added_count = updated_count = skipped_count = 0
    total = len(df)

    window = compute_fetch_window_from_df(df, buffer_days=30)
    time_min, time_max = window if window else default_fetch_window_years(2)

    with st.spinner("既存イベントを取得中..."):
        events = fetch_all_events(service, calendar_id, time_min, time_max) or []

    # ---- インデックス構築 ----
    worksheet_to_event: Dict[str, dict] = {}
    outside_key_to_event: Dict[str, dict] = {}

    for ev in events:
        if outside_mode:
            core = _strip_outside_suffix(ev.get("summary") or "")
            if not core:
                continue
            s_key, e_key = _normalize_event_times_to_key(ev.get("start") or {}, ev.get("end") or {})
            if s_key and e_key:
                outside_key_to_event[f"{core}|{s_key}|{e_key}"] = ev
        else:
            wid = extract_worksheet_id_from_description(ev.get("description") or "")
            if wid:
                worksheet_to_event[wid] = ev

    # ---- 行ごとの処理 ----
    for i, row in df.iterrows():
        desc_text = safe_get(row, "Description", "")
        subject = safe_get(row, "Subject", "")
        all_day_flag = safe_get(row, "All Day Event", _ALL_DAY_TRUE)
        private_flag = safe_get(row, "Private", _PRIVATE_TRUE)
        start_date_str = safe_get(row, "Start Date", "")
        end_date_str = safe_get(row, "End Date", "")
        start_time_str = safe_get(row, "Start Time", "")
        end_time_str = safe_get(row, "End Time", "")

        event_data = {
            "summary": subject,
            "location": safe_get(row, "Location", ""),
            "description": desc_text,
            "visibility": "private" if str(private_flag).strip() == _PRIVATE_TRUE else "default",
            "transparency": "opaque",
        }

        try:
            if all_day_flag == _ALL_DAY_TRUE:
                sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                ed = datetime.strptime(end_date_str or start_date_str, "%Y/%m/%d").date()
                event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                event_data["end"] = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
            else:
                sdt = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                edt = datetime.strptime(
                    f"{end_date_str or start_date_str} {end_time_str or start_time_str}", "%Y/%m/%d %H:%M"
                ).replace(tzinfo=JST)
                event_data["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
                event_data["end"] = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
        except Exception as e:
            st.error(f"行 {i} の日時パースに失敗しました: {e}")
            progress.progress((i + 1) / total)
            continue

        # 既存イベント検索
        if outside_mode:
            core = _strip_outside_suffix(subject)
            row_s, row_e = _normalize_row_times_to_key(
                {"Start Date": start_date_str, "End Date": end_date_str,
                 "Start Time": start_time_str, "End Time": end_time_str},
                all_day_flag,
            )
            existing = outside_key_to_event.get(f"{core}|{row_s}|{row_e}")
        else:
            worksheet_id = extract_worksheet_id_from_text(desc_text)
            existing = worksheet_to_event.get(worksheet_id) if worksheet_id else None

        try:
            if existing:
                if is_event_changed(existing, event_data):
                    update_event_if_needed(service, calendar_id, existing["id"], event_data)
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
                        wid = extract_worksheet_id_from_text(desc_text)
                        if wid:
                            worksheet_to_event[wid] = added_event
        except Exception as e:
            st.error(f"イベント '{event_data.get('summary', '(無題)')}' の登録/更新に失敗しました: {e}")

        progress.progress((i + 1) / total)

    st.success(
        f"✅ 登録: {added_count} 件 / 🔧 更新: {updated_count} 件 / ↪ スキップ: {skipped_count} 件 処理完了！"
    )


# ============================================================
# メインタブ描画
# ============================================================

def render_tab2_register(user_id: str, editable_calendar_options: dict, service):
    """タブ2: イベント登録・更新"""
    st.subheader("イベントを登録・更新")

    work_files = st.session_state.get("uploaded_files") or []
    has_work = (
        bool(work_files)
        and st.session_state.get("merged_df_for_selector") is not None
        and not st.session_state["merged_df_for_selector"].empty
    )
    outside_file = st.session_state.get("uploaded_outside_work_file")
    outside_mode = bool(outside_file) and not has_work

    if not has_work and not outside_mode:
        st.info("先に「1. ファイルのアップロード」タブでファイルをアップロードしてください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    base_calendar = (
        st.session_state.get("base_calendar_name")
        or st.session_state.get("selected_calendar_name")
        or get_user_setting(user_id, "selected_calendar_name")
        or calendar_options[0]
    )
    if base_calendar not in calendar_options:
        base_calendar = calendar_options[0]

    # --- UI ---
    selected_calendar_name = _render_calendar_selector(user_id, calendar_options, base_calendar, outside_mode)
    calendar_id = editable_calendar_options[selected_calendar_name]

    all_day_override, private_event, description_columns = _render_event_settings(user_id, outside_mode)

    if outside_mode:
        st.info("イベント名は『備考 + [作業外予定]』で登録します。")
        add_task_type = False
        fallback_col = None
    else:
        add_task_type, fallback_col = _render_event_name_settings(user_id)

    # --- 実行 ---
    st.subheader("➡️ イベント登録・更新実行")
    if not st.button("Googleカレンダーに登録・更新する"):
        return

    try:
        if outside_mode:
            raw_df = _read_outside_file_to_df(outside_file)
            df = _build_calendar_df_from_outside(raw_df, private_event=private_event, all_day_override=all_day_override)
        else:
            df = process_excel_data_for_calendar(
                st.session_state["uploaded_files"],
                description_columns,
                all_day_override,
                private_event,
                fallback_col,
                add_task_type,
            )
    except Exception as e:
        st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
        return

    if df.empty:
        st.warning("有効なイベントデータがありません。処理を中断しました。")
        return

    st.info(f"{len(df)} 件のイベントを処理します。")
    _execute_registration(service, df, calendar_id, outside_mode)
