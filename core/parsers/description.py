"""
core/parsers/description.py
Description 文字列パースユーティリティ（st.* 禁止）

utils/parsers.py と parsers/worksheet_parser.py の重複を統合。
"""
from __future__ import annotations
import re
import unicodedata
from typing import Optional, Dict, Any

# ── パターン ──
RE_WORKSHEET_ID = re.compile(
    r"(?:作業指示書[:：]?\s*|Worksheet[:：]?\s*)([0-9A-Za-z_\-]+)",
    flags=re.IGNORECASE,
)
RE_WONUM    = re.compile(r"\[作業指示書[：:]\s*(.*?)\]")
RE_ASSETNUM = re.compile(r"\[管理番号[：:]\s*(.*?)\]")
RE_WORKTYPE = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
RE_TITLE    = re.compile(r"\[タイトル[：:]\s*(.*?)\]")


def extract_worksheet_id(text: str) -> Optional[str]:
    """
    イベント Description から作業指示書 ID を抽出して返す。
    全角→半角正規化済み。見つからなければ None。
    """
    if not text:
        return None
    s = unicodedata.normalize("NFKC", text)
    m = RE_WORKSHEET_ID.search(s)
    return m.group(1).strip().upper() if m else None


def parse_description_fields(text: str) -> Dict[str, str]:
    """
    Description から各フィールドを抽出して辞書で返す。
    キー: worksheet_id, assetnum, worktype, title
    """
    if not text:
        return {"worksheet_id": "", "assetnum": "", "worktype": "", "title": ""}

    def _pick(pat: re.Pattern) -> str:
        m = pat.search(text)
        return (m.group(1).strip() if m else "") or ""

    return {
        "worksheet_id": _pick(RE_WONUM),
        "assetnum":     _pick(RE_ASSETNUM),
        "worktype":     _pick(RE_WORKTYPE),
        "title":        _pick(RE_TITLE),
    }


def is_event_changed(existing: Dict[str, Any], new: Dict[str, Any]) -> bool:
    """
    既存イベントと新しいイベントデータの差分を判定する。
    tab2_register.py と parsers/worksheet_parser.py の重複を統合。
    比較対象: summary / description / visibility / transparency / start / end
    ※ location は比較しない（意図的）
    """
    nz = lambda v: (v or "")
    for field in ("summary", "description", "visibility", "transparency"):
        if nz(existing.get(field)) != nz(new.get(field)):
            return True
    if (existing.get("start") or {}) != (new.get("start") or {}):
        return True
    if (existing.get("end") or {}) != (new.get("end") or {}):
        return True
    return False
