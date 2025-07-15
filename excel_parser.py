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
                st.warning(f"ファイル '{uploaded_file.name}' に '管理番号' が見つかりません。このファイルでは代替の列を使用します。")
                df["管理番号"] = ""  # 管理番号がなくても続行できるように空列追加

            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")
            return pd.DataFrame()

    if not dataframes:
        return pd.DataFrame()

    for df in dataframes:
        df["管理番号"] = df["管理番号"].astype(str)

    merged_df = dataframes[0]

    for i, df in enumerate(dataframes[1:], 2):
        cols_to_merge = [col for col in df.columns if col == "管理番号" or col not in merged_df.columns]
        merged_df = pd.merge(merged_df, df[cols_to_merge], on="管理番号", how="outer")

    merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
    merged_df.drop_duplicates(subset="管理番号", inplace=True)

    name_col = find_closest_column(merged_df.columns, ["物件名"])
    # 🔽 ここを修正（代替候補を追加）
    start_col = find_closest_column(merged_df.columns, ["予定開始", "開始日時"])
    end_col = find_closest_column(merged_df.columns, ["予定終了", "終了日時"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])

    if not all([start_col, end_col]):
        st.error("必要な列（予定開始 / 開始日時・予定終了 / 終了日時）が見つかりません。")
        return pd.DataFrame()

    # 管理番号と物件名がどちらも空文字列の場合のみ、代替列を選択させる
    alt_subject_col = None
    mng_col_exists = (
        "管理番号" in merged_df.columns and
        merged_df["管理番号"].apply(lambda x: bool(str(x).strip())).any()
    )
    name_col_exists = (
        name_col is not None and
        merged_df[name_col].apply(lambda x: bool(str(x).strip())).any()
    )

    if not mng_col_exists and not name_col_exists:
        st.warning("管理番号と物件名の両方が見つかりません。代わりにイベント名として使用する列を選択してください。")
        alt_subject_col = st.selectbox("イベント名として使用する列を選んでください：", merged_df.columns)

    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        mng = clean_mng_num(row.get("管理番号", ""))
        name = row.get(name_col) if name_col else ""
        alt = row.get(alt_subject_col, "") if alt_subject_col else ""

        # イベント名（Subject）生成
        if mng or name:
            subj = f"{mng}{name}"
        elif alt:
            subj = str(alt)
        else:
            subj = "イベント"

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

