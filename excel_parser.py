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
    dataframes = []

    if not uploaded_files:
        raise ValueError("ExcelまたはCSVファイルがアップロードされていません。")

    for uploaded_file in uploaded_files:
        try:
            # CSVかExcelかを判定
            if uploaded_file.name.lower().endswith(".csv"):
                success = False
                # 複数の文字コードを試す
                for enc in ["utf-8-sig", "cp932", "shift_jis", "utf-8"]:
                    uploaded_file.seek(0)
                    try:
                        df = pd.read_csv(
                            uploaded_file,
                            encoding=enc,
                            sep=None,
                            engine="python",
                            dtype=str
                        )
                        if not df.empty and len(df.columns) > 0:
                            success = True
                            break
                    except Exception:
                        continue
                if not success:
                    raise ValueError("CSVファイルの形式を自動判定できませんでした。")

            elif uploaded_file.name.lower().endswith((".xls", ".xlsx")):
                df = pd.read_excel(uploaded_file, engine="openpyxl")
            else:
                raise ValueError(f"未対応のファイル形式です: {uploaded_file.name}")

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
    merged_df["管理番号"] = merged_df["管理番号"].astype(str)

    for df in dataframes[1:]:
        df_copy = df.copy()
        df_copy["管理番号"] = df_copy["管理番号"].astype(str)
        cols_to_merge = [
            col for col in df_copy.columns
            if col == "管理番号" or col not in merged_df.columns
        ]
        merged_df = pd.merge(merged_df, df_copy[cols_to_merge], on="管理番号", how="outer")

    return merged_df

def get_available_columns_for_event_name(df):
    exclude_keywords = [
        "日時", "開始", "終了", "予定", "時間",
        "date", "time", "start", "end", "all day",
        "private", "subject", "description", "location"
    ]
    available_columns = []
    for col in df.columns:
        col_lower = str(col).lower()
        if not any(keyword in col_lower for keyword in exclude_keywords) and col != "管理番号":
            available_columns.append(col)
    return available_columns

def check_event_name_columns(merged_df):
    mng_col = find_closest_column(merged_df.columns, ["管理番号"])
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    has_mng_data = (
        mng_col is not None
        and not merged_df[mng_col].fillna("").astype(str).str.strip().eq("").all()
    )
    has_name_data = (
        name_col is not None
        and not merged_df[name_col].fillna("").astype(str).str.strip().eq("").all()
    )
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
        raise ValueError("必須の時刻列が見つかりません。")

    output_records = []

    for index, row in merged_df.iterrows():
        subj_parts = []

        if add_task_type_to_event_name and task_type_col and pd.notna(row.get(task_type_col)):
            task_type = str(row.get(task_type_col)).strip()
            if task_type:
                subj_parts.append(f"[{task_type}]")

        mng = clean_mng_num(row["管理番号"])
        if mng:
            subj_parts.append(mng)

        name = row.get(name_col, "") if name_col else ""
        if name and str(name).strip():
            subj_parts.append(str(name).strip())

        subj = " ".join(subj_parts) if subj_parts else "イベント"

        if not subj and fallback_event_name_column and fallback_event_name_column in row:
            subj = format_description_value(row.get(fallback_event_name_column, ""))

        try:
            start = pd.to_datetime(row[start_col])
            end = None
            if end_col and pd.notna(row.get(end_col)):
                try:
                    end = pd.to_datetime(row[end_col])
                except Exception:
                    pass
            if end is None:
                if start.time() == datetime.time(0, 0, 0):
                    end = start + datetime.timedelta(days=1)
                else:
                    end = start + datetime.timedelta(hours=1)
            if end < start:
                continue

            is_all_day = all_day_event_override or (
                start.time() == datetime.time(0, 0, 0)
                and end.time() == datetime.time(0, 0, 0)
            )

            if is_all_day:
                end_display = end - datetime.timedelta(days=1)
                start_time_display = ""
                end_time_display = ""
            else:
                end_display = end
                start_time_display = start.strftime("%H:%M")
                end_time_display = end.strftime("%H:%M")

        except Exception:
            continue

        location = row.get(addr_col, "") if addr_col else ""
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")

        # Description用の必須項目とオプション項目を整理
        required_items = []
        optional_items = []
        
        # 作業指示書（必須）
        worksheet_value = row.get(worksheet_col, "") if worksheet_col else ""
        if worksheet_col and pd.notna(worksheet_value):
            formatted_ws = format_worksheet_value(worksheet_value)
            if formatted_ws:
                required_items.append(f"[作業指示書: {formatted_ws}]")
        
        # 作業タイプ（必須）
        if task_type_col:
            task_type_value = row.get(task_type_col, "")
            if pd.notna(task_type_value):
                task_type = str(task_type_value).strip()
                if task_type:
                    required_items.append(f"[作業タイプ: {task_type}]")
        
        # タイトル（必須） - fallback_event_name_columnがある場合
        if fallback_event_name_column and fallback_event_name_column in row:
            title_value = format_description_value(row.get(fallback_event_name_column, ""))
            if title_value:
                required_items.append(f"[タイトル: {title_value}]")
        
        # 管理番号（必須）
        if mng:
            required_items.append(f"[管理番号: {mng}]")
        
        # 物件名（必須）
        if name_col:
            property_name_value = row.get(name_col, "")
            if pd.notna(property_name_value):
                property_name = str(property_name_value).strip()
                if property_name:
                    required_items.append(f"[物件名: {property_name}]")
        
        # 作業者（必須）
        worker_col = find_closest_column(merged_df.columns, ["作業者", "担当者"])
        if worker_col:
            worker_value = row.get(worker_col, "")
            if pd.notna(worker_value):
                worker = str(worker_value).strip()
                if worker:
                    required_items.append(f"[作業者: {worker}]")
        
        # ユーザーが選択したオプション項目
        for col in description_columns:
            if col in row:
                optional_items.append(format_description_value(row.get(col)))
        
        # Descriptionを組み立て: オプション項目 + 必須項目
        description_parts = optional_items.copy()
        if required_items:
            description_parts.extend(required_items)
        
        description = " / ".join(filter(None, description_parts))

        output_records.append({
            "Subject": subj,
            "Start Date": start.strftime("%Y/%m/%d"),
            "Start Time": start_time_display,
            "End Date": end_display.strftime("%Y/%m/%d"),
            "End Time": end_time_display,
            "All Day Event": "True" if is_all_day else "False",
            "Description": description,
            "Location": location,
            "Private": "True" if private_event else "False"
        })

    return pd.DataFrame(output_records) if output_records else pd.DataFrame()
