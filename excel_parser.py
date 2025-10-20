import pandas as pd
import re
import datetime
import chardet  # ← 追加

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
    dataframes = []
    
    if not uploaded_files:
        raise ValueError("ExcelまたはCSVファイルがアップロードされていません。")

    for uploaded_file in uploaded_files:
        try:
            # 🔽 修正版：拡張子でCSVとExcelを判定し、CSVは自動エンコーディング検出
            if uploaded_file.name.lower().endswith(".csv"):
                raw_data = uploaded_file.read()
                detected = chardet.detect(raw_data)
                encoding = detected.get("encoding", "utf-8")
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=encoding, dtype=str)
            elif uploaded_file.name.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(uploaded_file, engine="openpyxl")
            else:
                raise ValueError(f"未対応のファイル形式です: {uploaded_file.name}")
            # 🔼 修正ここまで

            df.columns = [str(c).strip() for c in df.columns]
            
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                df["管理番号"] = ""
            dataframes.append(df)
        except Exception as e:
            raise IOError(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")

    if not dataframes:
        raise ValueError("処理できる有効なデータがありません。")

    merged_df = dataframes[0].copy()
    merged_df['管理番号'] = merged_df['管理番号'].astype(str)
    
    for df in dataframes[1:]:
        df_copy = df.copy()
        df_copy['管理番号'] = df_copy['管理番号'].astype(str)
        cols_to_merge = [col for col in df_copy.columns if col == "管理番号" or col not in merged_df.columns]
        merged_df = pd.merge(merged_df, df_copy[cols_to_merge], on="管理番号", how="outer")

    # 重複統合の条件を削除（以下の2行をコメントアウト）
    # if not merged_df["管理番号"].str.strip().eq("").all():
    #     merged_df.drop_duplicates(subset="管理番号", inplace=True)
        
    return merged_df

def get_available_columns_for_event_name(df):
    exclude_keywords = ["日時", "開始", "終了", "予定", "時間", "date", "time", "start", "end", "all day", "private", "subject", "description", "location", "作業タイプ"]
    available_columns = []
    
    for col in df.columns:
        col_lower = str(col).lower()
        if not any(keyword in col_lower for keyword in exclude_keywords) and col != "管理番号":
            available_columns.append(col)
            
    return available_columns

def check_event_name_columns(merged_df):
    mng_col = find_closest_column(merged_df.columns, ["管理番号"])
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    
    has_mng_data = (mng_col is not None and 
                    not merged_df[mng_col].fillna("").astype(str).str.strip().eq("").all())
    has_name_data = (name_col is not None and 
                     not merged_df[name_col].fillna("").astype(str).str.strip().eq("").all())
    
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

    output_records = []
    
    for index, row in merged_df.iterrows():
        subj_parts = []
        
        # 1. 「作業タイプ」を先頭に追加
        if add_task_type_to_event_name and task_type_col and pd.notna(row.get(task_type_col)):
            task_type = str(row.get(task_type_col)).strip()
            if task_type:
                subj_parts.append(f"【{task_type}】")
        
        # 2. 「管理番号」を追加
        mng = clean_mng_num(row["管理番号"])
        if mng and str(mng).strip():
            subj_parts.append(str(mng).strip())
        
        # 3. 「物件名」を追加
        name = row.get(name_col, "") if name_col else ""
        if name and str(name).strip():
            subj_parts.append(str(name).strip())

        # 4. パーツを結合してイベント名（`subj`）を生成
        subj = ""
        if subj_parts:
            if subj_parts[0].startswith("_
