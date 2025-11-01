import re
from typing import Optional

# 作業指示書IDパターン：
# [作業指示書: 1234567]、(作業指示書：1234567)、作業指示書 1234567、など幅広く対応
RE_WORKSHEET_ID = re.compile(
    r"(?:作業指示書[:：]?\s*|Worksheet[:：]?\s*)(\d+)"
)

def extract_worksheet_id_from_text(text: str) -> Optional[str]:
    """
    テキスト（主にイベントのDescriptionなど）から作業指示書IDを抽出する。
    マッチした場合はID（数字のみ）を文字列で返す。見つからない場合は None。
    """
    if not text:
        return None

    match = RE_WORKSHEET_ID.search(text)
    if not match:
        return None

    return match.group(1)
