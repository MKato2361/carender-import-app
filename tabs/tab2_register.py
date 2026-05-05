from ui.components import calendar_card
from core.utils.datetime_utils import default_fetch_window
from services.settings_service import get_setting as get_user_setting, set_setting as set_user_setting
import re
import streamlit as st
from streamlit_sortables import sort_items as _sort_items
import pandas as pd
from datetime import datetime, timedelta, date, time
from zoneinfo import ZoneInfo
from typing import Dict, Optional

from utils.helpers import safe_get
from core.parsers.description import extract_worksheet_id as extract_worksheet_id_from_text, is_event_changed
from excel_parser import (
    process_excel_data_for_calendar,
    get_available_columns_for_event_name,
    check_event_name_columns,
)
from services.calendar_service import (
    get_events as fetch_all_events,
    add_event_to_calendar,
    update_event_if_needed,
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


# ============================================================
# フェッチ期間計算
# ============================================================

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


def _is_blank(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.lower() in ("nan", "none")


def _count_missing_datetime_rows(df: pd.DataFrame, all_day_override: bool) -> int:
    if df is None or df.empty:
        return 0

    count = 0
    for _, row in df.iterrows():
        sd = safe_get(row, "Start Date", "")
        stime = safe_get(row, "Start Time", "")

        if all_day_override:
            if _is_blank(sd):
                count += 1
        else:
            if _is_blank(sd) or _is_blank(stime):
                count += 1

    return count


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
# UI 補助
# ============================================================


def _render_event_settings(user_id, outside_mode):
    """設定ウィジェットを描画する（値はセッション状態に保存済みのものを使う）"""
    st.markdown('<div class="section-heading"><span class="mi">tune</span>イベント基本設定</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.checkbox("すべて終日として扱う", key="reg_all_day")
    with col2:
        st.checkbox("すべて非公開で登録する", key="reg_private")

    if not outside_mode:
        pool = st.session_state.get("description_columns_pool") or []
        new_cols = st.multiselect("説明文に含める列", pool, key="reg_desc_cols")

        # 選択済みの列をドラッグで並び替え
        prev_order = st.session_state.get("reg_desc_cols_order", [])
        # 既存の順序を維持しつつ、削除された列を除き、新規列を末尾に追加
        synced = [c for c in prev_order if c in new_cols]
        synced += [c for c in new_cols if c not in synced]

        # 選択セットが変わったらソータブルを強制リマウント
        if set(prev_order) != set(new_cols):
            ver = st.session_state.get("_desc_sort_ver", 0) + 1
            st.session_state["_desc_sort_ver"] = ver
            st.session_state["reg_desc_cols_order"] = synced

        st.checkbox("説明文に列名を含める（例: 内容：〇〇）", key="reg_desc_include_header")
        saved_header = get_user_setting(user_id, "description_include_col_header") or False
        if st.session_state.get("reg_desc_include_header") != saved_header:
            set_user_setting(user_id, "description_include_col_header", st.session_state["reg_desc_include_header"])

        if new_cols:
            st.caption("順序（ドラッグで並び替え）")
            sorted_cols = _sort_items(
                synced,
                direction="horizontal",
                key=f"desc_sort_{st.session_state.get('_desc_sort_ver', 0)}",
                custom_style=".sortable-item { font-size: 13px; padding: 4px 12px; border-radius: 6px; cursor: grab; }",
            )
            if sorted_cols != st.session_state.get("reg_desc_cols_order"):
                st.session_state["reg_desc_cols_order"] = sorted_cols
        else:
            st.session_state["reg_desc_cols_order"] = []
            sorted_cols = []

        saved = get_user_setting(user_id, "description_columns_selected") or []
        current_order = st.session_state.get("reg_desc_cols_order", [])
        if current_order != saved:
            set_user_setting(user_id, "description_columns_selected", current_order)


def _render_bulk_datetime_settings(all_day_override: bool) -> None:
    """
    日時一括設定ウィジェットを描画する。
    return は行わず、値は session_state["bulk_datetime_enabled"] 等から取得する。
    （expander が折りたたまれていると内部コードが実行されず return 値が None になるバグを修正）
    """
    today = date.today()
    default_start_time = time(9, 0)
    st.session_state.setdefault("bulk_datetime_enabled", False)
    st.session_state.setdefault("bulk_start_date", today)
    st.session_state.setdefault("bulk_start_time", default_start_time)

    enabled = st.session_state.get("bulk_datetime_enabled", False)
    with st.expander("日時一括設定（日時が空の行に適用）", expanded=enabled):
        st.caption("日時が空の行だけに適用されます。1件ごとに1時間ずつずらして登録し、1日15件まで（16件目以降は翌日に繰り越し）。終了時刻は自動で開始の1時間後になります。")
        st.checkbox(
            "有効にする",
            key="bulk_datetime_enabled",
        )
        _enabled = st.session_state.get("bulk_datetime_enabled", False)
        c1, c2 = st.columns(2)
        with c1:
            st.date_input(
                "開始日",
                value=st.session_state.get("bulk_start_date", today),
                key="bulk_start_date",
                disabled=not _enabled,
            )
        with c2:
            st.time_input(
                "開始時刻",
                value=st.session_state.get("bulk_start_time", default_start_time),
                key="bulk_start_time",
                step=300,
                disabled=not _enabled,
            )


def _render_event_name_settings(user_id):
    """イベント名設定ウィジェットを描画する（値はセッション状態に保存済みのものを使う）"""
    pool = st.session_state.get("description_columns_pool") or []
    options = ["選択しない"] + pool
    # reg_fallback_col の値がプールに存在しない場合はリセット
    if st.session_state.get("reg_fallback_col") not in options:
        st.session_state["reg_fallback_col"] = "選択しない"

    is_customized = (
        st.session_state.get("reg_add_task_type", False)
        or st.session_state.get("reg_fallback_col", "選択しない") != "選択しない"
    )
    with st.expander("イベント名の構成（カスタマイズ）", expanded=is_customized):
        col1, col2 = st.columns(2)
        with col1:
            add_type = st.checkbox("先頭に作業種別を付与する", key="reg_add_task_type")
            set_user_setting(user_id, "add_task_type_to_event_name", add_type)
        with col2:
            fallback = st.selectbox("特定の列をイベント名にする（任意）", options, key="reg_fallback_col")
            set_user_setting(user_id, "event_name_col_selected", fallback)


def _execute_registration(
    service,
    df: pd.DataFrame,
    calendar_id: str,
    outside_mode: bool,
):
    """Googleカレンダーへのイベント登録・更新を実行する"""
    progress = st.progress(0)
    status_text = st.empty()
    added_count = 0
    updated_count = 0
    skipped_count = 0
    failed_count = 0
    failed_items = []
    total = len(df)

    window = compute_fetch_window_from_df(df, buffer_days=30)
    time_min, time_max = window if window else default_fetch_window(2)

    with st.spinner("既存イベントを取得中..."):
        events = fetch_all_events(service, calendar_id, time_min, time_max) or []

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
            wid = extract_worksheet_id_from_text(ev.get("description") or "")
            if wid:
                worksheet_to_event[wid] = ev

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
            failed_count += 1
            failed_items.append({
                "row_index": i,
                "subject": subject or "(無題)",
                "worksheet_id": extract_worksheet_id_from_text(desc_text) or "",
                "error": f"日時パース失敗: {e}",
            })
            progress.progress((i + 1) / total)
            continue

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
                    result = update_event_if_needed(service, calendar_id, existing["id"], event_data)
                    if result is None:
                        failed_count += 1
                        failed_items.append({
                            "row_index": i,
                            "subject": event_data.get("summary", "(無題)"),
                            "worksheet_id": extract_worksheet_id_from_text(desc_text) or "",
                            "error": "update_event_if_needed が None を返しました",
                        })
                    else:
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
                else:
                    failed_count += 1
                    failed_items.append({
                        "row_index": i,
                        "subject": event_data.get("summary", "(無題)"),
                        "worksheet_id": extract_worksheet_id_from_text(desc_text) or "",
                        "error": "add_event_to_calendar が None を返しました",
                    })
        except Exception as e:
            failed_count += 1
            failed_items.append({
                "row_index": i,
                "subject": event_data.get("summary", "(無題)"),
                "worksheet_id": extract_worksheet_id_from_text(desc_text) or "",
                "error": str(e),
            })

        done = i + 1
        progress.progress(done / total)
        status_text.caption(
            f"処理中 ({done}/{total})：{subject or '(無題)'} — 登録 {added_count} 更新 {updated_count} スキップ {skipped_count} 失敗 {failed_count}"
        )

    status_text.empty()

    st.success(
        f"登録 {added_count} 件 ／ 更新 {updated_count} 件 ／ スキップ {skipped_count} 件 ／ 失敗 {failed_count} 件"
    )

    accounted = added_count + updated_count + skipped_count + failed_count
    if accounted != total:
        st.warning(f"集計不一致: プレビュー {total} 件に対し、結果集計は {accounted} 件です。処理漏れの可能性があります。")
    else:
        st.caption(f"集計確認: プレビュー {total} 件 = 結果集計 {accounted} 件")

    if failed_items:
        with st.expander("登録失敗一覧", expanded=True):
            st.dataframe(pd.DataFrame(failed_items), use_container_width=True)

    st.caption("※ スキップ = カレンダー上の既存イベントと内容が同一のため更新不要だったもの")

    # 呼び出し元で結果表示 + 終了ボタンに切り替えるためセッションに保存
    _cal_name = st.session_state.get("selected_calendar_name_register", "")
    st.session_state["register_result"] = {
        "added":   added_count,
        "updated": updated_count,
        "skipped": skipped_count,
        "failed":  failed_count,
        "total":   total,
        "failed_items": failed_items,
        "calendar_name": _cal_name,
    }
    st.session_state["_last_registered_calendar"] = _cal_name


# ============================================================
# メインタブ描画 (AuthManager対応版)
# ============================================================

def render_tab2_register(user_id: str, manager):
    """タブ2: イベント登録・更新"""
    service = st.session_state.get("calendar_service")
    editable_calendar_options = st.session_state.get("editable_calendar_options", {})

    work_files = st.session_state.get("uploaded_files") or []
    has_work = (
        bool(work_files)
        and st.session_state.get("merged_df_for_selector") is not None
        and not st.session_state["merged_df_for_selector"].empty
    )
    outside_file = st.session_state.get("uploaded_outside_work_file")
    outside_mode = bool(outside_file) and not has_work

    if not has_work and not outside_mode:
        st.markdown("""
<div style="border:1.5px dashed var(--border-strong);border-radius:12px;padding:24px;text-align:center;color:var(--text-3);">
  <div style="margin-bottom:8px;"><span class="mi">folder_open</span></div>
  <div style="font-size:15px;font-weight:500;margin-bottom:4px;color:var(--text-1);">ファイルがアップロードされていません</div>
  <div style="font-size:13px;color:var(--text-2);">「1. ファイル取込」タブでExcel / CSVをアップロードしてから戻ってきてください。</div>
</div>
""", unsafe_allow_html=True)
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

    # ── セッション状態の初期化（ウィジェット描画前に1度だけ） ──
    pool = st.session_state.get("description_columns_pool") or []

    # プールが変わったら reg_desc_cols の存在しない列をフィルタして修正
    if "reg_desc_cols" in st.session_state:
        valid_cols = [c for c in st.session_state["reg_desc_cols"] if c in pool]
        if valid_cols != st.session_state["reg_desc_cols"]:
            st.session_state["reg_desc_cols"] = valid_cols

    if "reg_all_day" not in st.session_state:
        st.session_state["reg_all_day"] = get_user_setting(user_id, "default_allday_event") or False
    if "reg_private" not in st.session_state:
        v = get_user_setting(user_id, "default_private_event")
        st.session_state["reg_private"] = v if v is not None else True
    if "reg_desc_cols" not in st.session_state:
        saved = get_user_setting(user_id, "description_columns_selected") or ["内容", "詳細"]
        st.session_state["reg_desc_cols"] = [col for col in saved if col in pool]
    if "reg_desc_cols_order" not in st.session_state:
        st.session_state["reg_desc_cols_order"] = list(st.session_state["reg_desc_cols"])
    if "reg_desc_include_header" not in st.session_state:
        st.session_state["reg_desc_include_header"] = get_user_setting(user_id, "description_include_col_header") or False
    if "reg_add_task_type" not in st.session_state:
        st.session_state["reg_add_task_type"] = get_user_setting(user_id, "add_task_type_to_event_name") or False
    if "reg_fallback_col" not in st.session_state:
        st.session_state["reg_fallback_col"] = get_user_setting(user_id, "event_name_col_selected") or "選択しない"
    st.session_state.setdefault("bulk_datetime_enabled", False)
    st.session_state.setdefault("bulk_start_date", date.today())
    st.session_state.setdefault("bulk_start_time", time(9, 0))

    # ── セッション状態から設定値を読み取る ──
    all_day_override    = st.session_state["reg_all_day"]
    private_event       = st.session_state["reg_private"]
    description_columns   = st.session_state.get("reg_desc_cols_order") or st.session_state["reg_desc_cols"]
    include_col_header    = st.session_state.get("reg_desc_include_header", False)
    add_task_type       = st.session_state["reg_add_task_type"]
    saved_fb            = st.session_state["reg_fallback_col"]
    fallback_col        = None if saved_fb == "選択しない" else saved_fb
    bulk_enabled        = st.session_state["bulk_datetime_enabled"]
    bulk_start_date     = st.session_state["bulk_start_date"]
    bulk_start_time     = st.session_state["bulk_start_time"]

    if outside_mode:
        add_task_type = False
        fallback_col  = None
        bulk_enabled  = False

    # ── Step 1: 登録先カレンダー ──
    selected_calendar_name = calendar_card(
        calendar_names=calendar_options,
        session_key="selected_calendar_name_register",
        base_calendar=base_calendar,
        label="登録先カレンダー",
        share_on=st.session_state.get("share_calendar_selection_across_tabs", True),
        allow_change=not outside_mode or not st.session_state.get("share_calendar_selection_across_tabs", True),
    )
    calendar_id = editable_calendar_options[selected_calendar_name]

    # ── Step 2: df 計算 ──
    try:
        if outside_mode:
            raw_df = _read_outside_file_to_df(outside_file)
            df = _build_calendar_df_from_outside(
                raw_df, private_event=private_event, all_day_override=all_day_override
            )
        elif bulk_enabled:
            df = process_excel_data_for_calendar(
                st.session_state["uploaded_files"],
                description_columns, all_day_override, private_event,
                fallback_col, add_task_type,
                bulk_start_date=bulk_start_date, bulk_start_time=bulk_start_time,
                bulk_end_date=None, bulk_end_time=None,
                include_col_header=include_col_header,
            )
        else:
            df = process_excel_data_for_calendar(
                st.session_state["uploaded_files"],
                description_columns, all_day_override, private_event,
                fallback_col, add_task_type,
                include_col_header=include_col_header,
            )
    except Exception:
        st.error("ファイルの読み込み中にエラーが発生しました。ファイル形式と内容を確認してください。")
        return

    if df.empty:
        st.warning("有効なイベントデータがありません。処理を中断しました。")
        return

    if not outside_mode:
        remain_missing = _count_missing_datetime_rows(df, all_day_override)
        if remain_missing > 0:
            st.error(f"日時が未設定のイベントが {remain_missing} 件残っています。下の「日時一括設定」を有効にして設定してください。")
            with st.expander("未設定行の確認", expanded=True):
                st.dataframe(df, use_container_width=True)

    event_count = len(df)

    # ── Step 3: プレビュー + 確認カード + 登録ボタン ──
    st.divider()

    _preview_key = "register_preview_seen"
    _preview_expanded = not st.session_state.get(_preview_key, False)
    if _preview_expanded:
        st.session_state[_preview_key] = True
    with st.expander(f"登録内容プレビュー（{event_count} 件）", expanded=_preview_expanded):
        st.dataframe(df, use_container_width=True)

    st.markdown(f"""
<div style="border:1.5px solid var(--accent-border);border-radius:12px;padding:14px 18px;margin:8px 0;background:var(--surface);">
  <div style="font-size:12px;color:var(--text-3);margin-bottom:4px;">この内容でGoogleカレンダーに登録します</div>
  <div style="display:flex;align-items:baseline;gap:16px;flex-wrap:wrap;">
    <span style="font-size:14px;"><span class="mi">calendar_month</span> 登録先：<strong style="font-size:17px;color:var(--accent);">{selected_calendar_name}</strong></span>
    <span style="font-size:14px;">件数：<strong>{event_count} 件</strong></span>
  </div>
</div>
""", unsafe_allow_html=True)

    confirm_key = "register_confirm_pending"

    # カレンダーや設定が変わったら前回の登録結果をクリア
    prev_cal = st.session_state.get("_last_registered_calendar")
    if prev_cal and prev_cal != selected_calendar_name:
        st.session_state.pop("register_result", None)
        st.session_state.pop(confirm_key, None)
        st.session_state["_last_registered_calendar"] = selected_calendar_name

    result = st.session_state.get("register_result")

    # ── 登録完了後 → 結果表示 + 終了ボタン ──
    if result:
        r = result
        cal_name = r.get("calendar_name") or selected_calendar_name

        # 結果サマリーカード
        all_ok = r["failed"] == 0
        card_color = "var(--accent)" if all_ok else "var(--danger)"
        st.markdown(f"""
<div style="border:2px solid {card_color};border-radius:10px;padding:16px 20px;margin:8px 0;">
  <div style="font-size:13px;color:var(--text-2);margin-bottom:8px;">
    <span class="mi">check_circle</span> {cal_name} への登録が完了しました
  </div>
  <div style="display:flex;gap:24px;flex-wrap:wrap;font-size:15px;">
    <span><span class="mi">check_circle</span>登録 <strong>{r["added"]} 件</strong></span>
    <span><span class="mi">sync</span>更新 <strong>{r["updated"]} 件</strong></span>
    <span><span class="mi">skip_next</span>スキップ <strong>{r["skipped"]} 件</strong></span>
    {"" if all_ok else f'<span style="color:var(--danger);"><span class="mi">error</span>失敗 <strong>{r["failed"]} 件</strong></span>'}
  </div>
</div>
""", unsafe_allow_html=True)

        if r.get("failed_items"):
            with st.expander("登録失敗の一覧", expanded=True):
                st.dataframe(pd.DataFrame(r["failed_items"]), use_container_width=True)

        st.divider()

        col_done, col_retry = st.columns([3, 1])
        with col_done:
            if st.button(
                "アップロードファイルを削除して終了",
                type="primary",
                use_container_width=True,
            ):
                # ファイルとセッションをリセット
                from services.file_service import clear_files
                clear_files()
                st.session_state.pop("register_result", None)
                st.session_state.pop(confirm_key, None)
                st.session_state["upload_version"] = st.session_state.get("upload_version", 0) + 1
                st.session_state["gh_version"]     = st.session_state.get("gh_version", 0) + 1
                st.rerun()
        with col_retry:
            if st.button("続けて登録", use_container_width=True):
                st.session_state.pop("register_result", None)
                st.session_state.pop(confirm_key, None)
                st.rerun()
        return

    # ── 未実行 → 確認ボタン ──
    if not st.session_state.get(confirm_key):
        if st.button(
            f"「{selected_calendar_name}」に {event_count}件 登録する",
            type="primary",
            use_container_width=True,
        ):
            st.session_state[confirm_key] = True
            st.rerun()
    else:
        st.warning(f"「{selected_calendar_name}」に {event_count}件 を登録します。よろしいですか？")
        col_ok, col_cancel = st.columns([3, 1])
        with col_ok:
            if st.button("登録する", type="primary", use_container_width=True):
                st.session_state[confirm_key] = False
                _execute_registration(service, df, calendar_id, outside_mode)
                st.rerun()
        with col_cancel:
            if st.button("キャンセル", use_container_width=True):
                st.session_state[confirm_key] = False
                st.rerun()

    # ── Step 4: 設定（変更すると次回rerunでプレビューに反映） ──
    st.divider()
    _render_event_settings(user_id, outside_mode)

    if not outside_mode:
        _render_bulk_datetime_settings(all_day_override)
        _render_event_name_settings(user_id)
