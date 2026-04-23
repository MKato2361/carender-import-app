"""
utils/parsers.py — 後方互換ラッパー（実体は core/parsers/description.py）
"""
from core.parsers.description import (
    extract_worksheet_id as extract_worksheet_id_from_text,
    is_event_changed,
    parse_description_fields,
)
