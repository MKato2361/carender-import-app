import re
import unicodedata
from typing import Optional

# 作業指示書IDパターン：
# [作業指示書: 1234567]
# [作業指示書: WX-HK147-4307641]
# (作業指示書：WX-HK147-4307641)
# 作業指示書 WX-HK147-4307641
# Worksheet: WX-HK147-4307641
# など幅広く対応
RE_WORKSHEET_ID = re.compile(
    r"(?:作業指示書[:：]?\s*|Worksheet[:：]?\s*)([0-9A-Za-z_-]+)",
    flags=re.IGNORECASE,
)

def extract_worksheet_id_from_text(text: str) -> Optional[str]:
    """
    テキスト（主にイベントのDescriptionなど）から作業指示書IDを抽出する。
    数字のみでなく、英字・ハイフン・アンダースコアを含むIDにも対応する。
    マッチした場合はID文字列を返す。見つからない場合は None。
    """
    if not text:
        return None

    s = unicodedata.normalize("NFKC", text)
    match = RE_WORKSHEET_ID.search(s)
    if not match:
        return None

    return match.group(1).strip().upper()
