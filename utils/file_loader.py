"""
utils/file_loader.py — 後方互換ラッパー（実体は services/file_service.py）
"""
from services.file_service import (
    add_files      as update_uploaded_files,  # noqa: F401
    clear_files    as clear_uploaded_files,   # noqa: F401
    has_merged_data,                          # noqa: F401
)
import streamlit as st


def merge_uploaded_files() -> None:
    """
    services.file_service.merge_files を呼び出し、
    失敗ファイル名があれば st.warning で通知する。
    """
    from services.file_service import merge_files
    invalid_names = merge_files()
    if invalid_names:
        names = "、".join(invalid_names)
        st.warning(f"以下のファイルは読み込めなかったため除外しました：{names}")
