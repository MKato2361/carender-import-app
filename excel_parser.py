import pandas as pd
import re
import datetime

def clean_mng_num(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).replace("HK", "")

def restore_mng_format(cleaned_value):
    """clean_mng_num で整形された管理番号を HK 形式に復元する。

    仕様例:
      - 8906   -> HK8-906
      - 123    -> HK123
      - 12     -> HK012
      - 10123  -> HK10-123
      - 1234   -> HK1-234
      - R1234  -> HK-R1234
    """
    if cleaned_value is None:
        return ""

    s = str(cleaned_value).strip()
    if not s or s.lower() == "nan":
        return ""

    # 先頭が英字の場合（R1234 など）は HK-R1234 のようにする
    if s[0].isalpha():
        prefix = s[0].upper()
        rest = s[1:]
        return f"HK-{prefix}{rest}"

    # 想定は数字のみ
    if not s.isdigit():
        # 万一数字以外が混ざっていた場合は、とりあえず HK をつけて返す
        return "HK" + s

    n = len(s)

    if n <= 3:
        # 12   -> HK012
        # 123  -> HK123
        return "HK" + s.zfill(3)

    elif n == 4:
        # 8906 -> HK8-906
        # 1234 -> HK1-234
        return f"HK{s[0]}-{s[1:]}"

    elif n == 5:
        # 10123 -> HK10-123
        return f"HK{s[:2]}-{s[2:]}"

    else:
        # 6桁以上は「前半 - 下3桁」という汎用ルール
        # 例: 123456 -> HK123-456
        return f"HK{s[:-3]}-{s[-3:]}"

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
                df = pd.read_csv(uploaded_file, dtype=str)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
        except Exception as e:
            raise ValueError(f"ファイルの読み込みに失敗しました: {uploaded_file.name} ({e})")

        # 物件名を探す
        name_col = find_closest_column(df.columns, ["物件名", "物件名称", "ビル名"])
        addr_col = find_closest_column(df.columns, ["住所"])
        start_col = find_closest_column(df.columns, ["作業開始", "開始日時", "予定日", "予定開始"])
        end_col = find_closest_column(df.columns, ["作業終了", "終了日時", "予定終了"])
        task_type_col = find_closest_column(df.columns, ["作業タイプ", "作業種別", "種別"])
        worksheet_col = find_closest_column(df.columns, ["作業指示書", "作業指示書番号", "WONUM", "作業No"])

        # それぞれの列が見つからない場合のデフォルト
        if name_col is None:
            name_col = "物件名(自動補完)"
            df[name_col] = ""

        if addr_col is None:
            addr_col = "住所(自動補完)"
            df[addr_col] = ""

        if start_col is None:
            start_col = "開始日時(自動補完)"
            df[start_col] = ""

        if end_col is None:
            end_col = "終了日時(自動補完)"
            df[end_col] = ""

        if task_type_col is None:
            task_type_col = "作業タイプ(自動補完)"
            df[task_type_col] = ""

        if worksheet_col is None:
            worksheet_col = "作業指示書(自動補完)"
            df[worksheet_col] = ""

        # 管理番号を探してクリーニング
        mng_col = find_closest_column(df.columns, ["管理番号"])
        if mng_col:
            # 元の管理番号を保存してから整形
            df["元管理番号"] = df[mng_col].astype(str)
            df["管理番号"] = df[mng_col].apply(clean_mng_num)
        else:
            df["元管理番号"] = ""
            df["管理番号"] = ""

        # 統一された列名にリネーム
        df = df.rename(columns={
            name_col: "物件名",
            addr_col: "住所",
            start_col: "開始日時",
            end_col: "終了日時",
            task_type_col: "作業タイプ",
            worksheet_col: "作業指示書",
        })

        # 必須列が存在するかチェック
        required_cols = ["物件名", "住所", "開始日時", "終了日時", "作業タイプ", "作業指示書", "管理番号", "元管理番号"]
        for col in required_cols:
            if col not in df.columns:
                df[col] = ""

        dataframes.append(df)

    if not dataframes:
        raise ValueError("有効なデータが含まれていません。")

    merged_df = pd.concat(dataframes, ignore_index=True)

    # 重複排除（開始日時＋物件名＋作業指示書＋作業タイプ＋管理番号）
    merged_df["dup_key"] = (
        merged_df["開始日時"].astype(str)
        + "_" + merged_df["物件名"].astype(str)
        + "_" + merged_df["作業指示書"].astype(str)
        + "_" + merged_df["作業タイプ"].astype(str)
        + "_" + merged_df["管理番号"].astype(str)
    )
    merged_df = merged_df.drop_duplicates(subset=["dup_key"]).drop(columns=["dup_key"])

    return merged_df

def process_excel_data_for_calendar(
    uploaded_files,
    description_columns,
    all_day_event_override=False,
    private_event=True,
    fallback_event_name_column=None,
    add_task_type_to_event_name=False,
):
    # アップロードファイルからデータを読み込んで結合
    try:
        merged_df = _load_and_merge_dataframes(uploaded_files)
    except ValueError as e:
        raise e

    # 日付・時間の列を解析
    def parse_datetime_safe(val):
        if pd.isna(val) or val == "":
            return None
        try:
            return pd.to_datetime(val)
        except Exception:
            return None

    output_records = []

    # 物件名列
    name_col = "物件名" if "物件名" in merged_df.columns else None
    addr_col = "住所" if "住所" in merged_df.columns else None
    task_type_col = "作業タイプ" if "作業タイプ" in merged_df.columns else None
    worksheet_col = "作業指示書" if "作業指示書" in merged_df.columns else None

    for _, row in merged_df.iterrows():
        # 日時を取得
        start = parse_datetime_safe(row.get("開始日時"))
        end = parse_datetime_safe(row.get("終了日時"))

        if start is None:
            continue

        # 「元管理番号」（元の形式）を取得
        original_mng = ""
        if "元管理番号" in row:
            original_mng_value = row.get("元管理番号", "")
            if pd.notna(original_mng_value) and str(original_mng_value).strip() and str(original_mng_value).strip().lower() != 'nan':
                original_mng = str(original_mng_value).strip()

        # 件名
        if fallback_event_name_column and fallback_event_name_column in row:
            subj_base = str(row.get(fallback_event_name_column, "")).strip()
        else:
            subj_base = str(row.get("物件名", "")).strip() if name_col else ""

        # 作業タイプを件名にも含めるオプション
        if add_task_type_to_event_name and task_type_col:
            task_type_value = row.get(task_type_col, "")
            if pd.notna(task_type_value):
                task_type = str(task_type_value).strip()
                if task_type:
                    if subj_base:
                        subj = f"{subj_base}【{task_type}】"
                    else:
                        subj = f"{task_type}"
                else:
                    subj = subj_base
            else:
                subj = subj_base
        else:
            subj = subj_base

        # 時刻の表示形式を決定
        try:
            if all_day_event_override:
                is_all_day = True
                end_display = start
                start_time_display = ""
                end_time_display = ""
            else:
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
        
        # タイトル列（必須）
        title_col_name = find_closest_column(merged_df.columns, ["タイトル"])
        if title_col_name and title_col_name in row:
            title_value = format_description_value(row.get(title_col_name, ""))
            if title_value:
                required_items.append(f"[タイトル: {title_value}]")
        
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
        
        # 管理番号（必須）
        # 1. まず「元管理番号」（例: HK8-906）があればそれを使う
        # 2. 無ければ整形後の「管理番号」（例: 8906, 123, 10123, 12, R1234）から復元
        desc_mng = ""

        # 元管理番号（元データそのまま）を優先
        if original_mng:
            desc_mng = str(original_mng).strip()

        if not desc_mng:
            # 整形後の管理番号から復元を試みる
            cleaned_mng = row.get("管理番号", "")
            if pd.notna(cleaned_mng):
                desc_mng = restore_mng_format(cleaned_mng)

        # desc_mng が空でもタグ自体は必ず入れる
        required_items.append(f"[管理番号: {desc_mng}]")
        
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
        
        # ユーザーが選択したオプション項目（タイトル列との重複を避ける）
        for col in description_columns:
            if col in row and col != title_col_name:
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
