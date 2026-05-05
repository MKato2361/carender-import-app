import pandas as pd
import re
import datetime


def clean_mng_num(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).replace("HK", "")


def restore_mng_format(cleaned_value):
    """clean_mng_num で整形された管理番号を HK 形式に復元する。"""
    if cleaned_value is None:
        return ""

    s = str(cleaned_value).strip()
    if not s or s.lower() == "nan":
        return ""

    if s[0].isalpha():
        prefix = s[0].upper()
        rest = s[1:]
        return f"HK-{prefix}{rest}"

    if not s.isdigit():
        return "HK" + s

    n = len(s)

    if n <= 3:
        return "HK" + s.zfill(3)
    elif n == 4:
        return f"HK{s[0]}-{s[1:]}"
    elif n == 5:
        return f"HK{s[:2]}-{s[2:]}"
    else:
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
            if uploaded_file.name.lower().endswith(".csv"):
                success = False
                for enc in ["utf-8-sig", "cp932", "shift_jis", "utf-8"]:
                    uploaded_file.seek(0)
                    try:
                        df = pd.read_csv(
                            uploaded_file,
                            encoding=enc,
                            sep=None,
                            engine="python",
                            dtype=str,
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
                df["元管理番号"] = df[mng_col].astype(str)
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                df["元管理番号"] = ""
                df["管理番号"] = ""

            dataframes.append(df)

        except Exception as e:
            raise IOError(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")

    if not dataframes:
        raise ValueError("処理できる有効なデータがありません。")

    merged_df = dataframes[0].copy()
    merged_df["管理番号"] = merged_df["管理番号"].astype(str)
    merged_df["元管理番号"] = merged_df["元管理番号"].astype(str)

    for df in dataframes[1:]:
        df_copy = df.copy()
        df_copy["管理番号"] = df_copy["管理番号"].astype(str)
        df_copy["元管理番号"] = df_copy["元管理番号"].astype(str)

        cols_to_merge = []
        for col in df_copy.columns:
            if col == "管理番号":
                cols_to_merge.append(col)
            else:
                if col not in merged_df.columns:
                    cols_to_merge.append(col)

        if cols_to_merge == ["管理番号"]:
            continue

        merged_df = pd.merge(
            merged_df,
            df_copy[cols_to_merge],
            on="管理番号",
            how="outer",
        )

    return merged_df


def get_available_columns_for_event_name(df):
    exclude_keywords = [
        "日時",
        "開始",
        "終了",
        "予定",
        "時間",
        "date",
        "time",
        "start",
        "end",
        "all day",
        "private",
        "subject",
        "description",
        "location",
    ]
    available_columns = []
    for col in df.columns:
        col_lower = str(col).lower()
        if (
            not any(keyword in col_lower for keyword in exclude_keywords)
            and col != "管理番号"
            and col != "元管理番号"
        ):
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


def _is_blank(v):
    if v is None:
        return True
    if pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() in ("nan", "none", "nat")


def _safe_to_datetime(v):
    if _is_blank(v):
        return None
    try:
        return pd.to_datetime(v)
    except Exception:
        return None


def _to_date_str(dt):
    if dt is None:
        return ""
    return dt.strftime("%Y/%m/%d")


def _to_time_str(dt):
    if dt is None:
        return ""
    return dt.strftime("%H:%M")


def _has_valid_worksheet_value(row, worksheet_col):
    if worksheet_col is None:
        return False
    worksheet_value = row.get(worksheet_col, "")
    return (
        pd.notna(worksheet_value)
        and str(worksheet_value).strip() != ""
        and str(worksheet_value).strip().lower() != "nan"
    )


def _calc_shifted_bulk_start(base_dt, index_zero_based: int):
    """
    1件ごとに1時間ずらし、1日15件まで。
    16件目以降は翌日に繰り越す。
    """
    day_offset = index_zero_based // 15
    hour_offset = index_zero_based % 15
    return base_dt + datetime.timedelta(days=day_offset, hours=hour_offset)


def process_excel_data_for_calendar(
    uploaded_files,
    description_columns,
    all_day_event_override,
    private_event,
    fallback_event_name_column=None,
    add_task_type_to_event_name=False,
    bulk_start_date=None,
    bulk_start_time=None,
    bulk_end_date=None,
    bulk_end_time=None,
    apply_bulk_to_missing_only=True,
    include_col_header=False,
):
    merged_df = _load_and_merge_dataframes(uploaded_files)

    name_col = find_closest_column(merged_df.columns, ["物件名"])
    start_col = find_closest_column(
        merged_df.columns, ["予定開始", "開始日時", "開始時間", "開始"]
    )
    end_col = find_closest_column(
        merged_df.columns, ["予定終了", "終了日時", "終了時間", "終了"]
    )
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])
    task_type_col = find_closest_column(merged_df.columns, ["作業タイプ"])

    if not start_col and bulk_start_date is None:
        raise ValueError("必須の時刻列が見つかりません。一括設定の開始日を指定してください。")

    output_records = []

    # 一括日時で生成した「ユニーク作業指示書」の件数
    bulk_generated_count = 0

    # 作業指示書ごとの割り当て開始時刻を保持
    worksheet_time_map = {}

    for _, row in merged_df.iterrows():
        subj_parts = []

        if (
            add_task_type_to_event_name
            and task_type_col
            and pd.notna(row.get(task_type_col))
        ):
            task_type = str(row.get(task_type_col)).strip()
            if task_type:
                subj_parts.append(f"[{task_type}]")

        mng = clean_mng_num(row["管理番号"])
        if mng:
            subj_parts.append(mng)

        original_mng = ""
        if "元管理番号" in row:
            original_mng_value = row.get("元管理番号", "")
            if (
                pd.notna(original_mng_value)
                and str(original_mng_value).strip()
                and str(original_mng_value).strip().lower() != "nan"
            ):
                original_mng = str(original_mng_value).strip()

        name = row.get(name_col, "") if name_col else ""
        if name and str(name).strip():
            subj_parts.append(str(name).strip())

        subj = " ".join(subj_parts) if subj_parts else "イベント"

        if (not subj or subj == "イベント") and fallback_event_name_column and fallback_event_name_column in row:
            fallback_name = format_description_value(row.get(fallback_event_name_column, ""))
            if fallback_name:
                subj = fallback_name

        worksheet_value = row.get(worksheet_col, "") if worksheet_col else ""
        has_worksheet_value = _has_valid_worksheet_value(row, worksheet_col)
        formatted_ws = format_worksheet_value(worksheet_value) if has_worksheet_value else ""

        # -------------------------
        # 日時処理
        # -------------------------
        start = _safe_to_datetime(row.get(start_col)) if start_col else None
        end = _safe_to_datetime(row.get(end_col)) if end_col and pd.notna(row.get(end_col)) else None

        bulk_start_dt = None

        # 一括日時は「作業指示書がある行」だけ対象
        if has_worksheet_value and bulk_start_date is not None and bulk_start_time is not None:
            bulk_start_dt = pd.to_datetime(
                datetime.datetime.combine(bulk_start_date, bulk_start_time)
            )

        used_bulk_datetime = False

        # 開始が無い行だけ一括日時を適用
        if start is None and bulk_start_dt is not None:
            # 同じ作業指示書なら同じ開始時刻を使い回す
            if formatted_ws in worksheet_time_map:
                shifted_start = worksheet_time_map[formatted_ws]
            else:
                shifted_start = _calc_shifted_bulk_start(bulk_start_dt, bulk_generated_count)
                worksheet_time_map[formatted_ws] = shifted_start
                bulk_generated_count += 1

            start = shifted_start
            end = shifted_start + datetime.timedelta(hours=1)
            used_bulk_datetime = True

        # bulkを使っていない通常行の終了補完
        if start is not None and end is None:
            if start.time() == datetime.time(0, 0, 0):
                end = start + datetime.timedelta(days=1)
            else:
                end = start + datetime.timedelta(hours=1)

        # それでも開始が無いならイベント化しない
        if start is None:
            continue

        if end is not None and end < start:
            if used_bulk_datetime:
                end = start + datetime.timedelta(hours=1)
            else:
                continue

        # 終日イベントにはしない
        is_all_day = False

        end_display = end
        start_time_display = _to_time_str(start)
        end_time_display = _to_time_str(end)

        location = ""
        if addr_col and addr_col in row:
            location_value = row.get(addr_col, "")
            if pd.notna(location_value):
                location = str(location_value).strip()
                if "北海道札幌市" in location:
                    location = location.replace("北海道札幌市", "").strip()

        required_items = []
        optional_items = []

        title_col_name = find_closest_column(merged_df.columns, ["タイトル"])
        if title_col_name and title_col_name in row:
            title_value = format_description_value(row.get(title_col_name, ""))
            if title_value:
                required_items.append(f"[タイトル: {title_value}]")

        if worksheet_col and pd.notna(worksheet_value):
            if formatted_ws:
                required_items.append(f"[作業指示書: {formatted_ws}]")

        if task_type_col:
            task_type_value = row.get(task_type_col, "")
            if pd.notna(task_type_value):
                task_type = str(task_type_value).strip()
                if task_type:
                    required_items.append(f"[作業タイプ: {task_type}]")

        desc_mng = ""
        if original_mng:
            desc_mng = str(original_mng).strip()

        if not desc_mng:
            cleaned_mng = row.get("管理番号", "")
            if pd.notna(cleaned_mng):
                desc_mng = restore_mng_format(cleaned_mng)

        required_items.append(f"[管理番号: {desc_mng}]")

        if name_col:
            property_name_value = row.get(name_col, "")
            if pd.notna(property_name_value):
                property_name = str(property_name_value).strip()
                if property_name:
                    required_items.append(f"[物件名: {property_name}]")

        worker_col = find_closest_column(merged_df.columns, ["作業者", "担当者"])
        if worker_col:
            worker_value = row.get(worker_col, "")
            if pd.notna(worker_value):
                worker = str(worker_value).strip()
                if worker:
                    required_items.append(f"[作業者: {worker}]")

        for col in description_columns:
            if col in row and col != title_col_name:
                val = format_description_value(row.get(col))
                optional_items.append(f"{col}：{val}" if include_col_header else val)

        description_parts = optional_items.copy()
        if required_items:
            description_parts.extend(required_items)

        description = " / ".join(filter(None, description_parts))

        output_records.append(
            {
                "Subject": subj,
                "Start Date": _to_date_str(start),
                "Start Time": start_time_display,
                "End Date": _to_date_str(end_display),
                "End Time": end_time_display,
                "All Day Event": "False",
                "Description": description,
                "Location": location,
                "Private": "True" if private_event else "False",
            }
        )

    return pd.DataFrame(output_records) if output_records else pd.DataFrame()
