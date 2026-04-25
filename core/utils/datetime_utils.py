from __future__ import annotations
"""
core/utils/datetime_utils.py
日時変換ユーティリティ（st.* 禁止）

to_utc_range が tab3/tab5/tab7/tab8 に重複して定義されていたものを統合。
"""
from datetime import date, datetime, timedelta, timezone

JST = timezone(timedelta(hours=9))


def to_utc_range(d1: date, d2: date) -> tuple[str, str]:
    """
    JST の日付範囲を UTC の ISO8601 文字列ペアに変換する。
    timeMin = d1 の JST 00:00:00, timeMax = d2 の JST 23:59:59.999999
    Google Calendar API の timeMin / timeMax に直接渡せる形式で返す。
    """
    start = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
    end   = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
    return (
        start.isoformat(timespec="microseconds").replace("+00:00", "Z"),
        end.isoformat(timespec="microseconds").replace("+00:00", "Z"),
    )


def default_fetch_window(years: int = 2) -> tuple[str, str]:
    """
    現在時刻を中心に ±years 年の UTC ISO8601 文字列ペアを返す。
    Google Calendar API の timeMin / timeMax に渡す用。
    """
    now = datetime.now(tz=timezone.utc)
    return (
        (now - timedelta(days=365 * years)).isoformat(),
        (now + timedelta(days=365 * years)).isoformat(),
    )


def to_jst_iso(s: str) -> str:
    """UTC/オフセット付き ISO 文字列を JST の ISO 文字列に変換する。"""
    try:
        if "T" in s and ("+" in s or s.endswith("Z")):
            dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
            return dt.isoformat(timespec="seconds")
    except (ValueError, AttributeError):
        pass
    return s
