# -*- coding: utf-8 -*-
# worksheet_parser.py
#
# 役割:
# - Description からのID・各種フィールド抽出
# - ID正規化（全角→半角）
# - カレンダーイベントの差分判定（Locationは比較対象外）
#
# 依存: 標準ライブラリのみ（StreamlitやGoogle APIへは依存しません）

from __future__ import annotations
import re
import unicodedata
from typing import Optional, Dict, Any

# =========
# 正規表現
# =========
RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")
RE_WONUM        = re.compile(r"\[作業指示書[：:]\s*(.*?)\]")
RE_ASSETNUM     = re.compile(r"\[管理番号[：:]\s*(.*?)\]")
RE_WORKTYPE     = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
RE_TITLE        = re.compile(r"\[タイトル[：:]\s*(.*?)\]")

# ==================
# 正規化 & 抽出関数
# ==================
def normalize_worksheet_id(s: Optional[str]) -> Optional[str]:
    """作業指示書番号をNFKCで全角→半角に正規化し、前後の空白を除去。"""
    if not s:
        return s
    return unicodedata.normalize("NFKC", s).strip()

def extract_worksheet_id_from_description(desc: str) -> Optional[str]:
    """Description内の [作業指示書: 123456] からIDを抽出して正規化して返す。見つからなければ None。"""
    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return normalize_worksheet_id(m.group(1))

def parse_description_fields(desc: str) -> Dict[str, str]:
    """
    Description から各フィールドを抽出して返す（見つからなければ空文字）。
    返すキー: wonum, assetnum, worktype, title
    """
    if not desc:
        return {"wonum": "", "assetnum": "", "worktype": "", "title": ""}

    def _pick(pat: re.Pattern[str]) -> str:
        m = pat.search(desc)
        return (m.group(1).strip() if m else "") or ""

    return {
        "wonum": _pick(RE_WONUM),
        "assetnum": _pick(RE_ASSETNUM),
        "worktype": _pick(RE_WORKTYPE),
        "title": _pick(RE_TITLE),
    }

# ==================
# 差分判定
# ==================
def is_event_changed(existing_event: Dict[str, Any], new_event_data: Dict[str, Any]) -> bool:
    """
    既存eventと新しいevent_dataの差分を判定して True/False を返す。
    比較対象:
      1) summary（タイトル）
      2) description（説明）
      3) transparency（公開/非公開）
      4) start（終日/時間/TimeZone含むdict比較）
      5) end   （終日/時間/TimeZone含むdict比較）
    ※ Location（場所）は比較対象外。
    """
    nz = lambda v: (v or "")

    # 1) summary
    if nz(existing_event.get("summary")) != nz(new_event_data.get("summary")):
        return True

    # 2) description
    if nz(existing_event.get("description")) != nz(new_event_data.get("description")):
        return True

    # 3) transparency
    if nz(existing_event.get("transparency")) != nz(new_event_data.get("transparency")):
        return True

    # 4) start
    if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
        return True

    # 5) end
    if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
        return True

    return False
