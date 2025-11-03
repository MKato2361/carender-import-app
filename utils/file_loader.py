"""
file_loader.py
タブ1（ファイルアップロード）専用の非UIロジックを管理するモジュール

追加機能:
- 壊れたファイルを自動除外
- 除外したファイル名を警告表示
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
    session_state['uploaded_files'] をマージ。
    壊れているファイルは除外し、警告を表示する。
    """
    uploaded = st.session_state.get("uploaded_files", [])
    if not uploaded:
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []
        return

    valid_files = []
    invalid_files = []

    # ✅ 1つずつ読み込みテストし、壊れたものを除外
    for f in uploaded:
        try:
            # 単体読み込みテスト（失敗するなら破損）
            _load_and_merge_dataframes([f])
            valid_files.append(f)
        except Exception:
            invalid_files.append(f)

    # 壊れたファイルを session_state から除去
    if invalid_files:
        for f in invalid_files:
            if f in st.session_state["uploaded_files"]:
                st.session_state["uploaded_files"].remove(f)

        # ⚠️ ユーザーへ通知（1行目に表示）
        names = "、".join(getattr(f, "name", "不明なファイル") for f in invalid_files)
        st.warning(f"⚠️ 以下のファイルは破損しているため除外しました：{names}\n別のファイルをアップロードしてください。")

    # 有効なファイルのみマージ
    if valid_files:
        try:
            merged_df = _load_and_merge_dataframes(valid_files)
            st.session_state["merged_df_for_selector"] = merged_df
            st.session_state["description_columns_pool"] = merged_df.columns.tolist()
        except Exception as e:
            st.error(f"ファイルの読み込み中にエラーが発生しました: {e}")
            st.session_state["merged_df_for_selector"] = None
            st.session_state["description_columns_pool"] = []
    else:
        st.session_state["merged_df_for_selector"] = None
        st.session_state["description_columns_pool"] = []


def has_merged_data() -> bool:
    """マージ済みデータが有効か確認"""
    df = st.session_state.get("merged_df_for_selector")
    return df is not None and not df.empty