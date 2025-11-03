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