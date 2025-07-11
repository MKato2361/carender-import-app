# excel_parser.py

import pandas as pd
import re
import datetime
import streamlit as st

def clean_mng_num(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).replace("HK", "")

def find_closest_column(columns, keywords):
    for kw in keywords:
        for col in columns:
            if kw.lower() in str(col).lower():
                return col
    return None

def format_description_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(round(val, 2))
    return str(val)

def format_worksheet_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(int(val))
    return str(val)

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event):
    dataframes = []

    if not uploaded_files:
        st.warning("Excelファイルをアップロードしてください。")
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                st.warning(f"ファイル '{uploaded_file.name}' に '管理番号' が見つかりません。スキップします。")
                continue
            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")
            return pd.DataFrame()

    if not dataframes:
        return pd.DataFrame()

    # --- 修正箇所：複数のファイルで同一名称の列がある場合の統合処理 ---
    # まず、すべてのDataFrameの管理番号列を文字列に変換
    for df in dataframes:
        df['管理番号'] = df['管理番号'].astype(str)

    # 最初のDataFrameを結合のベースとする
    merged_df = dataframes[0]
    
    # 2つ目以降のDataFrameを結合
    for i, df in enumerate(dataframes[1:], 2):
        # 共通の列を見つけ、管理番号以外は結合しない
        cols_to_merge = [col for col in df.columns if col == "管理番号" or col not in merged_df.columns]
        
        # 結合対象の列を抽出し、管理番号をキーに結合
        merged_df = pd.merge(merged_df, df[cols_to_merge], on="管理番号", how="outer")

    # 管理番号の重複を削除し、一意なエントリのみにする
    merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
    merged_df.drop_duplicates(subset="管理番号", inplace=True)
    # --- 修正ここまで ---

    name_col = find_closest_column(merged_df.columns, ["物件名"])
    start_col = find_closest_column(merged_df.columns, ["予定開始"])
    end_col = find_closest_column(merged_df.columns, ["予定終了"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])

    if not all([name_col, start_col, end_col]):
        st.error("必要な列（物件名・予定開始・予定終了）が見つかりません。")
        return pd.DataFrame()

    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        mng = clean_mng_num(row["管理番号"])
        name = row.get(name_col, "")
        subj = f"{mng}{name}"

        try:
            start = pd.to_datetime(row[start_col])
            end = pd.to_datetime(row[end_col])
        except Exception:
            continue

        location = row.get(addr_col, "")
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")

        description = " / ".join(
            [format_description_value(row.get(col)) for col in description_columns if col in row]
        )

        # 作業指示書を先頭に追加（整数化して表示）
        worksheet_value = row.get(worksheet_col, "") if worksheet_col else ""
        if pd.notna(worksheet_value) and str(worksheet_value).strip():
            formatted_ws = format_worksheet_value(worksheet_value)
            description = f"作業指示書：{formatted_ws}/ " + description

        output.append({
            "Subject": subj,
            "Start Date": start.strftime("%Y/%m/%d"),
            "Start Time": start.strftime("%H:%M"),
            "End Date": end.strftime("%Y/%m/%d"),
            "End Time": end.strftime("%H:%M"),
            "All Day Event": "True" if all_day_event else "False",
            "Description": description,
            "Location": location,
            "Private": "True" if private_event else "False"
        })

    return pd.DataFrame(output)
