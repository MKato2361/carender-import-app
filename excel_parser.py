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
        return str(int(val)) if val.is_integer() else str(int(val))  # 小数はすべて整数として扱う
    return str(val)

def create_todo_checklist():
    """ToDoチェックリストを生成（クリック可能なHTMLチェックボックス）"""
    todo_items = [
        '<input type="checkbox" id="fax"> <label for="fax">点検通知（FAX）</label>',
        '<input type="checkbox" id="phone"> <label for="phone">点検通知（電話）</label>',
        '<input type="checkbox" id="paper"> <label for="paper">貼紙</label>'
    ]
    return "<br>".join(todo_items)

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event, include_todo=True):
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

    merged_df = dataframes[0]
    for df in dataframes[1:]:
        merged_df = pd.merge(merged_df, df, on="管理番号", how="outer")

    merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
    merged_df.drop_duplicates(subset="管理番号", inplace=True)

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

        # ToDoリストを追加
        if include_todo:
            todo_list = create_todo_checklist()
            if description.strip():
                description = description + "<br><br><strong>【作業ToDo】</strong><br>" + todo_list
            else:
                description = "<strong>【作業ToDo】</strong><br>" + todo_list

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
