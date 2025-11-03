"""
register_handler.py
タブ2（イベント登録処理）用ロジック分離版
Part①: import & prepare_events()

役割：
- process_excel_data_for_calendar() で得た DataFrame から
  Googleカレンダー登録用のイベント候補データを作成
- エラー・警告を収集して返却
"""

from __future__ import annotations
from typing import Any, Dict, List

import pandas as pd


def prepare_events(
    df: pd.DataFrame,
    description_columns: List[str],
    fallback_event_name_column: str | None,
    add_task_type: bool,
) -> Dict[str, Any]:
    """
    DataFrame → Googleカレンダー登録用イベント候補データへ整形する

    Parameters
    ----------
    df : pd.DataFrame
        process_excel_data_for_calendar() の結果DataFrame（最終加工済み）
    description_columns : list[str]
        UIで選択された「説明欄に入れる列」
    fallback_event_name_column : str | None
        イベント名に使用する代替列（未設定なら None）
    add_task_type : bool
        イベント名の先頭に作業タイプを付けるかどうか

    Returns
    -------
    dict
        {
            "events": [ {...}, {...} ],
            "errors": [ "x行目: エラー内容", ... ],
            "warnings": [ "x行目: 警告内容", ... ]
        }
    """

    results: Dict[str, Any] = {"events": [], "errors": [], "warnings": []}

    if df.empty:
        results["errors"].append("Excelデータが空のため、イベントを生成できません。")
        return results

    required_cols = ["Start Date", "End Date", "Start Time", "End Time", "Subject", "Description"]

    # 必須列チェック（不足していれば致命的エラー扱い）
    for col in required_cols:
        if col not in df.columns:
            results["errors"].append(f"必須列 '{col}' が見つかりません。")
            return results

    for idx, row in df.iterrows():
        row_num = idx + 1  # UI表示用

        # --- イベント名生成 ---
        subject = str(row.get("Subject", "")).strip()

        # 代替タイトル列が指定されている場合
        if fallback_event_name_column and subject == "":
            alt_val = str(row.get(fallback_event_name_column, "")).strip()
            if alt_val:
                subject = alt_val
                results["warnings"].append(f"{row_num}行目: Subjectが空のため '{fallback_event_name_column}' をタイトルに使用しました")

        if add_task_type:
            task_type = str(row.get("作業タイプ", "")).strip()  # 存在しない場合は空
            if task_type:
                subject = f"{task_type} {subject}"

        if not subject:
            results["errors"].append(f"{row_num}行目: イベント名（Subject）が空のためスキップしました")
            continue

        # --- 説明欄作成（複数列結合） ---
        description_parts = []
        for col in description_columns:
            if col in df.columns:
                val = str(row.get(col, "")).strip()
                if val:
                    description_parts.append(f"{col}: {val}")
        description = "\n".join(description_parts) if description_parts else str(row.get("Description", "")).strip()

        # --- イベント用データオブジェクト化（時刻パースは後の build_event_payload に任せる） ---
        event_candidate = {
            "row_index": idx,
            "Subject": subject,
            "Description": description,
            "Start Date": str(row.get("Start Date", "")).strip(),
            "End Date": str(row.get("End Date", "")).strip(),
            "Start Time": str(row.get("Start Time", "")).strip(),
            "End Time": str(row.get("End Time", "")).strip(),
            "All Day Event": str(row.get("All Day Event", "True")).strip(),
            "Private": str(row.get("Private", "True")).strip(),
            "Location": str(row.get("Location", "")).strip(),
        }

        results["events"].append(event_candidate)

    return results
    
    from datetime import datetime, date, timedelta, timezone
from typing import Any, Dict, List, Tuple, Optional

JST = timezone(timedelta(hours=9))


def fetch_existing_events(service, calendar_id: str, time_min: str, time_max: str) -> Dict[str, dict]:
    """
    Googleカレンダーから既存イベントを取得し、
    作業指示書ID（worksheet_id）をキーにした辞書を返す。

    Returns
    -------
    dict
        { worksheet_id: event_obj }
    """
    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
        ).execute()

        items = events_result.get("items", [])
    except Exception:
        return {}

    mapping: Dict[str, dict] = {}
    for event in items:
        desc = (event.get("description") or "").strip()
        # [作業指示書: 123456] を抽出
        wid = _extract_worksheet_id(desc)
        if wid:
            mapping[wid] = event

    return mapping


def _extract_worksheet_id(desc: str) -> Optional[str]:
    """Description内の [作業指示書: 123456] からID抽出（Part①と同仕様）"""
    import re, unicodedata
    if not desc:
        return None
    m = re.search(r"\[作業指示書[：:]\s*([0-9０-９]+)\]", desc)
    if not m:
        return None
    return unicodedata.normalize("NFKC", m.group(1)).strip()


def build_event_payload(event_data: dict) -> dict:
    """
    prepare_events() が生成した event_candidate(dict) から
    Google Calendar API用 event payload を作成する。
    """
    subject = event_data["Subject"]
    description = event_data["Description"]
    location = event_data.get("Location", "")
    all_day_flag = event_data.get("All Day Event", "True")
    private_flag = event_data.get("Private", "True")

    start_date = event_data["Start Date"]
    end_date = event_data["End Date"]
    start_time = event_data["Start Time"]
    end_time = event_data["End Time"]

    payload: Dict[str, Any] = {
        "summary": subject,
        "description": description,
        "location": location,
        "transparency": "transparent" if private_flag == "True" else "opaque",
    }

    try:
        if all_day_flag == "True":
            sd = datetime.strptime(start_date, "%Y/%m/%d").date()
            ed = datetime.strptime(end_date, "%Y/%m/%d").date()
            payload["start"] = {"date": sd.strftime("%Y-%m-%d")}
            payload["end"] = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
        else:
            sdt = datetime.strptime(f"{start_date} {start_time}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
            edt = datetime.strptime(f"{end_date} {end_time}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
            payload["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
            payload["end"] = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
    except Exception:
        raise ValueError(f"日時パースに失敗しました: {subject}")

    return payload


def detect_changes(existing_event: dict, new_event: dict) -> bool:
    """
    既存イベントと新規イベントpayloadの差分を判定。
    True → 更新が必要
    """
    nz = lambda v: (v or "")

    if nz(existing_event.get("summary")) != nz(new_event.get("summary")):
        return True

    if nz(existing_event.get("description")) != nz(new_event.get("description")):
        return True

    if nz(existing_event.get("transparency")) != nz(new_event.get("transparency")):
        return True

    if (existing_event.get("start") or {}) != (new_event.get("start") or {}):
        return True

    if (existing_event.get("end") or {}) != (new_event.get("end") or {}):
        return True

    return False


def create_todo_for_event(tasks_service, task_list_id: str, title: str, event_id: str, due_date: date) -> bool:
    """
    対象イベント用のToDoを作成する。
    成功したら True、失敗なら False
    """
    try:
        tasks_service.tasks().insert(
            tasklist=task_list_id,
            body={
                "title": title,
                "notes": f"関連イベントID: {event_id}",
                "due": due_date.isoformat(),
            },
        ).execute()
        return True
    except Exception:
        return False


def register_or_update_events(
    service,
    calendar_id: str,
    event_candidates: List[dict],
    existing_event_map: Dict[str, dict],
) -> Dict[str, int]:
    """
    event_candidates: prepare_events() からの "events" リスト
    existing_event_map: fetch_existing_events() の戻り値
    UI側でループしprogress表示する前提で、加算のみ行う

    Returns
    -------
    dict { "added": int, "updated": int, "skipped": int }
    """
    results = {"added": 0, "updated": 0, "skipped": 0}

    for event_data in event_candidates:
        desc_text = event_data["Description"]
        worksheet_id = _extract_worksheet_id(desc_text)

        # payload生成（時間パース失敗時はUIでスキップ可）
        try:
            payload = build_event_payload(event_data)
        except Exception:
            results["skipped"] += 1
            continue

        existing_event = existing_event_map.get(worksheet_id) if worksheet_id else None

        try:
            if existing_event:
                if detect_changes(existing_event, payload):
                    service.events().update(
                        calendarId=calendar_id,
                        eventId=existing_event["id"],
                        body=payload,
                    ).execute()
                    results["updated"] += 1
                else:
                    results["skipped"] += 1
            else:
                added = service.events().insert(calendarId=calendar_id, body=payload).execute()
                results["added"] += 1
                if worksheet_id:
                    existing_event_map[worksheet_id] = added
        except Exception:
            results["skipped"] += 1

    return results