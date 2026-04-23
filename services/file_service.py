"""
services/file_service.py
ファイルアップロード管理（utils/file_loader.py を整理）

session_state のアクセスは許可。
st.warning / st.error は呼び出し元の UI 層が担うため、
このモジュールは結果を戻り値で返す。
"""
from __future__ import annotations
from typing import Any
import streamlit as st
from excel_parser import _load_and_merge_dataframes


def add_files(new_files: list[Any]) -> None:
    """重複なしで session_state['uploaded_files'] にファイルを追加する。"""
    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []
    existing = {getattr(f, "name", None) for f in st.session_state["uploaded_files"]}
    for f in new_files:
        if getattr(f, "name", None) not in existing:
            st.session_state["uploaded_files"].append(f)


def clear_files() -> None:
    """アップロード済みファイルと派生データをリセットする。"""
    st.session_state["uploaded_files"] = []
    st.session_state["merged_df_for_selector"] = None
    st.session_state["description_columns_pool"] = []


def merge_files() -> list[str]:
    """
    session_state['uploaded_files'] をマージして merged_df を更新する。
    読み込み失敗したファイルは除外し、ファイル名のリストを返す。
    """
    uploaded = st.session_state.get("uploaded_files", [])
    if not uploaded:
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []
        return []

    valid, invalid_names = [], []
    for f in uploaded:
        try:
            _load_and_merge_dataframes([f])
            valid.append(f)
        except Exception:
            invalid_names.append(getattr(f, "name", "不明なファイル"))

    if invalid_names:
        st.session_state["uploaded_files"] = [
            f for f in st.session_state["uploaded_files"] if f not in
            [x for x in uploaded if getattr(x, "name", "") in invalid_names]
        ]

    if valid:
        try:
            merged = _load_and_merge_dataframes(valid)
            st.session_state["merged_df_for_selector"] = merged
            st.session_state["description_columns_pool"] = merged.columns.tolist()
        except Exception:
            st.session_state["merged_df_for_selector"] = None
            st.session_state["description_columns_pool"] = []
    else:
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []

    return invalid_names  # 呼び出し元が警告表示するために返す


def has_merged_data() -> bool:
    """マージ済みデータが有効かどうかを返す。"""
    df = st.session_state.get("merged_df_for_selector")
    return df is not None and not df.empty
