import pandas as pd
import re
import datetime

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
        return str(int(val))
    return str(val)

def _load_and_merge_dataframes(uploaded_files):
    if not uploaded_files:
        raise ValueError("Excelファイルがアップロードされていません。")

    dfs = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                df["管理番号"] = ""
            dfs.append(df)
        except Exception as e:
            raise IOError(f"ファイル '{file.name}' の読み込みに失敗しました: {e}")

    if not dfs:
        raise ValueError("処理できる有効なデータがありません。")

    merged_df = pd.concat(dfs, ignore_index=True)
    merged_df.drop_duplicates(subset=["管理番号"], inplace=True)
    return merged_df

def get_available_columns_for_event_name(df):
    exclude_keywords = ["日時", "開始", "終了", "予定", "時間", "date", "time", "start", "end", "all day", "private", "subject", "description", "location", "作業タイプ"]
    return [col for col in df.columns if not any(kw in str(col).lower() for kw in exclude_keywords) and col != "管理番号"]

def check_event_name_columns(merged_df):
    mng_col = find_closest_column(merged_df.columns, ["管理番号"])
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    has_mng_data = mng_col is not None and not merged_df[mng_col].fillna("").astype(str).str.strip().eq("").all()
    has_name_data = name_col is not None and not merged_df[name_col].fillna("").astype(str).str.strip().eq("").all()
    return has_mng_data, has_name_data

def process_excel_data_for_calendar(
    uploaded_files,
    description_columns,
    all_day_event_override,
    private_event,
    fallback_event_name_column=None,
    add_task_type_to_event_name=False
):
    merged_df = _load_and_merge_dataframes(uploaded_files)
    
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    start_col = find_closest_column(merged_df.columns, ["予定開始", "開始日時", "開始時間", "開始"])
    end_col = find_closest_column(merged_df.columns, ["予定終了", "終了日時", "終了時間", "終了"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])
    task_type_col = find_closest_column(merged_df.columns, ["作業タイプ"])

    if not start_col:
        raise ValueError("必須の時刻列（'予定開始'、または'開始日時'など）が見つかりません。")

    def create_event_record(row):
        subj_parts = []
        if add_task_type_to_event_name and task_type_col and pd.notna(row.get(task_type_col)):
            subj_parts.append(f"【{row[task_type_col].strip()}】")
        
        mng = clean_mng_num(row.get("管理番号", ""))
        if mng:
            subj_parts.append(mng)

        name = ""
        if name_col and pd.notna(row.get(name_col)):
            name = str(row[name_col]).strip()
        elif fallback_event_name_column and pd.notna(row.get(fallback_event_name_column)):
            name = str(row[fallback_event_name_column]).strip()
        
        if name:
            subj_parts.append(name)

        subject = " ".join(subj_parts)
        if not subject:
            subject = "無題のイベント"

        description_text = ""
        for col in description_columns:
            if col in row and pd.notna(row[col]):
                description_text += f"{col}: {format_description_value(row[col])}\n"
        
        if worksheet_col and pd.notna(row.get(worksheet_col)):
            worksheet_id = format_worksheet_value(row[worksheet_col])
            description_text += f"\n作業指示書: {worksheet_id}"

        start_date = pd.to_datetime(row[start_col])
        end_date = pd.to_datetime(row[end_col]) if end_col and pd.notna(row.get(end_col)) else start_date + datetime.timedelta(hours=1)

        location = str(row[addr_col]) if addr_col and pd.notna(row.get(addr_col)) else ""

        return {
            "Subject": subject,
            "Start Date": start_date.strftime("%Y/%m/%d"),
            "Start Time": start_date.strftime("%H:%M:%S"),
            "End Date": end_date.strftime("%Y/%m/%d"),
            "End Time": end_date.strftime("%H:%M:%S"),
            "Description": description_text.strip(),
            "All Day Event": "True" if all_day_event_override or (end_date - start_date) >= datetime.timedelta(days=1) else "False",
            "Private": "True" if private_event else "False",
            "Location": location
        }

    output_df = merged_df.apply(create_event_record, axis=1, result_type='expand')
    return output_df
