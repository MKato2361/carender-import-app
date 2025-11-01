import pandas as pd
from datetime import datetime, timedelta, timezone

def safe_get(row: pd.Series, key: str, default: str = "") -> str:
    """
    Pandas Series から安全に値を取得する。
    - KeyError防止
    - NaN/None → default に変換
    """
    try:
        val = row.get(key, default)
    except Exception:
        return default

    if pd.isna(val):
        return default

    return val


def to_utc_range(start_date: datetime, end_date: datetime):
    """
    日付範囲をUTC形式の ISO8601 文字列に変換するヘルパー関数。
    Google Calendar API に渡す timeMin/timeMax 形式を生成。
    """
    if isinstance(start_date, datetime):
        start_utc = start_date.astimezone(timezone.utc).isoformat()
    else:
        start_utc = datetime.combine(start_date, datetime.min.time(), tzinfo=timezone.utc).isoformat()

    if isinstance(end_date, datetime):
        end_utc = end_date.astimezone(timezone.utc).isoformat()
    else:
        end_utc = datetime.combine(end_date, datetime.max.time(), tzinfo=timezone.utc).isoformat()

    return start_utc, end_utc


def default_fetch_window_years(years: int = 1):
    """
    年数を指定して、過去→未来の検索範囲を返す。
    例：years=1 の場合 → 過去1年〜未来1年
    """
    now = datetime.now(timezone.utc)
    start = now - timedelta(days=365 * years)
    end = now + timedelta(days=365 * years)
    return start, end

