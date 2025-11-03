# tabs/tab_upload.py
import streamlit as st
import pandas as pd
from typing import List
from io import BytesIO

def _read_excel(file) -> pd.DataFrame:
    try:
        return pd.read_excel(file)
    except Exception:
        file.seek(0)
        return pd.read_excel(file, engine="openpyxl")

def render_tab_upload():
    st.subheader("📂 ファイルのアップロード")
    files = st.file_uploader(
        "Excelファイルを選択（複数可）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uploader_tab1",
    )

    if files:
        dfs: List[pd.DataFrame] = []
        for f in files:
            try:
                df = _read_excel(f)
                df["__source_filename__"] = getattr(f, "name", "uploaded.xlsx")
                dfs.append(df)
            except Exception as e:
                st.error(f"{getattr(f, 'name', 'file')} の読み込みに失敗しました: {e}")

        if dfs:
            merged = pd.concat(dfs, ignore_index=True)
            st.session_state["uploaded_files"] = files
            st.session_state["merged_df_for_selector"] = merged

            # 説明欄候補（既知カラムを除外して候補に）
            known = {"Subject","Location","Description","All Day Event","Private",
                     "Start Date","End Date","Start Time","End Time"}
            pool = [c for c in merged.columns if c not in known]
            st.session_state["description_columns_pool"] = pool

            st.success(f"✅ {len(files)} ファイル、{len(merged)} 行を読み込みました。")
            with st.expander("先頭50行を表示", expanded=False):
                st.dataframe(merged.head(50), use_container_width=True)
        else:
            st.warning("有効なExcelデータがありません。")
    else:
        st.info("Excelファイル（.xlsx / .xls）をアップロードしてください。")
