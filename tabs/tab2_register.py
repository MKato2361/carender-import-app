import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

# ===== 他モジュール依存 =====
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


JST = timezone(timedelta(hours=9))


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
    from datetime import datetime, timezone, timedelta

    now_utc = datetime.now(timezone.utc)
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


# ===== 追加: 作業外予定ファイルを汎用的に読み取り、既存フロー互換のDataFrameへ整形 =====
def _read_outside_file_to_df(file_obj) -> pd.DataFrame:
    name = getattr(file_obj, "name", "")
    if name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj, dtype=str)
    else:
        # CSV: エンコーディングをいくつか試す
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                df = pd.read_csv(file_obj, dtype=str, encoding=enc, errors="ignore")
                break
            except Exception:
                df = None
        if df is None:
            raise ValueError("CSVの読み込みに失敗しました（対応エンコーディング不明）。")
    df = df.fillna("")
    return df


def _build_calendar_df_from_outside(df_raw: pd.DataFrame, private_event: bool, all_day_override: bool) -> pd.DataFrame:
    """
    作業外予定の生データから、既存処理と互換な列構成の DataFrame を生成する
    必要列:
      Subject, Description, All Day Event, Private, Start Date, End Date, Start Time, End Time, Location(任意)
    仕様:
      - イベント名: 備考 + " [作業外予定]"
      - Description: 「理由コード」列
      - 時刻が両方ない行は終日にフォールバック（Q1=Noに基づき“常に終日”ではない）
    """
    # 必須列チェック
    if "備考" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『備考』列が見つかりません。")
    if "理由コード" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『理由コード』列が見つかりません。")

    # 日付・時刻候補（柔軟に拾う）
    start_date_candidates = ["開始日", "日付", "開始日時", "Start Date", "Date"]
    end_date_candidates = ["終了日", "終了日時", "End Date"]
    start_time_candidates = ["開始時刻", "開始時間", "Start Time"]
    end_time_candidates = ["終了時刻", "終了時間", "End Time"]
    location_candidates = ["場所", "現場名", "所在地", "Location"]

    def pick(col_names):
        for c in col_names:
            if c in df_raw.columns:
                return c
        return None

    c_sd = pick(start_date_candidates)
    c_ed = pick(end_date_candidates)
    c_st = pick(start_time_candidates)
    c_et = pick(end_time_candidates)
    c_loc = pick(location_candidates)

    rows = []
    for _, r in df_raw.iterrows():
        subject = f"{str(r['備考']).strip()} [作業外予定]".strip()
        description = str(r["理由コード"]).strip()

        # 日付
        sd_raw = (str(r[c_sd]).strip() if c_sd else "")
        ed_raw = (str(r[c_ed]).strip() if c_ed else "")

        # 多くのフォーマットを想定してYYYY/MM/DDへ寄せる
        def norm_date(s: str) -> Optional[str]:
            s = s.replace("-", "/").replace(".", "/").strip()
            for fmt in ("%Y/%m/%d", "%Y/%m/%d %H:%M", "%m/%d/%Y", "%Y/%m/%d %H:%M:%S"):
                try:
                    return datetime.strptime(s, fmt).strftime("%Y/%m/%d")
                except Exception:
                    continue
            # 8桁数字(YYYYMMDD)も許容
            if s.isdigit() and len(s) == 8:
                return f"{s[0:4]}/{s[4:6]}/{s[6:8]}"
            return "" if not s else s  # そのまま返す（後工程で失敗時に終日化）

        sd = norm_date(sd_raw)
        ed = norm_date(ed_raw) if ed_raw else sd

        # 時刻
        st_raw = (str(r[c_st]).strip() if c_st else "")
        et_raw = (str(r[c_et]).strip() if c_et else "")

        def norm_time(t: str) -> Optional[str]:
            t = t.replace(".", ":").strip()
            for fmt in ("%H:%M", "%H:%M:%S"):
                try:
                    return datetime.strptime(t, fmt).strftime("%H:%M")
                except Exception:
                    continue
            # 数字3-4桁(HHMM)を許容
            if t.isdigit() and len(t) in (3, 4):
                t = t.zfill(4)
                return f"{t[:2]}:{t[2:]}"
            return ""

        stime = norm_time(st_raw)
        etime = norm_time(et_raw)

        # 時刻が無い/片方のみ → 後工程で安全に扱う
        location = (str(r[c_loc]).strip() if c_loc else "")

        rows.append(
            {
                "Subject": subject,
                "Description": description,
                "All Day Event": "True" if all_day_override else "False",  # 基本False（Q1=No）、ただしUIで上書き可
                "Private": "True" if private_event else "False",
                "Start Date": sd or "",
                "End Date": ed or (sd or ""),
                "Start Time": stime or "",
                "End Time": etime or "",
                "Location": location,
            }
        )

    df = pd.DataFrame(rows)

    # 行ごとに「時刻が両方空 or 日付欠落」は終日へフォールバック
    def apply_fallback(row):
        if row["All Day Event"] == "True":
            return row
        if not row["Start Date"]:
            row["All Day Event"] = "True"
            return row
        if (not row["Start Time"]) and (not row["End Time"]):
            row["All Day Event"] = "True"
            return row
        # 片側のみ時刻がある場合は1時間想定で補完
        if row["Start Time"] and not row["End Time"]:
            try:
                dt = datetime.strptime(row["Start Time"], "%H:%M")
                end_dt = (dt + timedelta(hours=1)).strftime("%H:%M")
                row["End Time"] = end_dt
            except Exception:
                row["All Day Event"] = "True"
        if row["End Time"] and not row["Start Time"]:
            try:
                dt = datetime.strptime(row["End Time"], "%H:%M")
                start_dt = (dt - timedelta(hours=1)).strftime("%H:%M")
                row["Start Time"] = start_dt
            except Exception:
                row["All Day Event"] = "True"
        return row

    df = df.apply(apply_fallback, axis=1)
    return df


def render_tab2_register(user_id: str, editable_calendar_options: dict, service, tasks_service=None, default_task_list_id=None):
    st.subheader("イベントを登録・更新")
