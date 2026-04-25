"""
utils/helpers.py — safe_get のみ残す（他は core/utils/datetime_utils.py に移行済み）
"""
import pandas as pd
from core.utils.datetime_utils import (  # 後方互換エイリアス
    to_utc_range,
    default_fetch_window as default_fetch_window_years,
)


def safe_get(row: pd.Series, key: str, default: str = "") -> str:
    """Pandas Series から安全に値を取得する（NaN/None → default）。"""
    try:
        val = row.get(key, default)
    except Exception:
        return default
    if pd.isna(val):
        return default
    return val
