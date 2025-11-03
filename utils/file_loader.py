"""
file_loader.py
タブ1（ファイルアップロード）専用の非UIロジックを管理するモジュール

役割：
- アップロード/ GitHub 選択ファイルの統合
- 重複ファイルの排除
- DataFrame統合処理（excel_parser を利用）
"""

from typing import List, Any
from io import BytesIO
import streamlit as st
from excel_parser import _load_and_merge_dataframes


def update_uploaded_files(new_files: List[Any]) -> None:
    """
    既存の session_state['uploaded_files'] に new_files を追加する。
    重複ファイル名は除外。
    """
    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []

    existing_names = {getattr(f, "name", None) for f in st.session_state["uploaded_files"]}

    for f in new_files:
        if getattr(f, "name", None) not in existing_names:
            st.session_state["uploaded_files"].append(f)


def clear_uploaded_files() -> None:
    """アップロード済みファイルとマージ済みデータをリセット"""
    st.session_state["uploaded_files"] = []
    st.session_state["merged_df_for_selector"] = None
    st.session_state["description_columns_pool"] = []


def merge_uploaded_files() -> None:
    """
    session_state['uploaded_files'] に保持されているファイル群をマージし、
    merged_df_for_selector と description_columns_pool を更新する。
    """
    if not st.session_state.get("uploaded_files"):
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []
        return

    try:
        merged_df = _load_and_merge_dataframes(st.session_state["uploaded_files"])
        st.session_state["merged_df_for_selector"] = merged_df
        st.session_state["description_columns_pool"] = merged_df.columns.tolist()
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました: {e}")
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []


def has_merged_data() -> bool:
    """マージ済みデータが有効か確認"""
    df = st.session_state.get("merged_df_for_selector")
    return df is not None and not df.empty