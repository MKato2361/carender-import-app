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

            # 列名の前後空白を削除
            df.columns = [str(c).strip() for c in df.columns]

            # 管理番号列を探してクリーニング
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                # 元の管理番号を保存してから整形
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

    # 複数ファイルを管理番号でマージ
    merged_df = dataframes[0].copy()
    merged_df["管理番号"] = merged_df["管理番号"].astype(str)
    merged_df["元管理番号"] = merged_df["元管理番号"].astype(str)

    # 複数ファイルを管理番号でマージ
    merged_df = dataframes[0].copy()
    merged_df["管理番号"] = merged_df["管理番号"].astype(str)
    merged_df["元管理番号"] = merged_df["元管理番号"].astype(str)

    for df in dataframes[1:]:
        df_copy = df.copy()
        df_copy["管理番号"] = df_copy["管理番号"].astype(str)
        df_copy["元管理番号"] = df_copy["元管理番号"].astype(str)

        # ① 結合キー「管理番号」は常に含める
        # ② それ以外は、merged_df にまだ存在しない列だけをマージ対象にする
        cols_to_merge = []
        for col in df_copy.columns:
            if col == "管理番号":
                cols_to_merge.append(col)
            else:
                # 既に merged_df に同名の列がある場合は追加しない（重複回避）
                if col not in merged_df.columns:
                    cols_to_merge.append(col)

        # もし cols_to_merge に管理番号しか入っていないなら、重複データは無視して continue する
        if cols_to_merge == ["管理番号"]:
            # このファイルにマージ追加する新しい列がない -> 次へ
            # ただし管理番号の存在確認のため outer 結合だけは行いたい場合は以下の1行を使う（必要なら）
            # merged_df = pd.merge(merged_df, df_copy[["管理番号"]], on="管理番号", how="outer")
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


def process_excel_data_for_calendar(
    uploaded_files,
    description_columns,
    all_day_event_override,
    private_event,
    fallback_event_name_column=None,
    add_task_type_to_event_name=False,
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

    if not start_col:
        raise ValueError("必須の時刻列が見つかりません。")

    output_records = []

    for index, row in merged_df.iterrows():
        subj_parts = []

        # 件名の先頭に [作業タイプ] を付けるオプション
        if (
            add_task_type_to_event_name
            and task_type_col
            and pd.notna(row.get(task_type_col))
        ):
            task_type = str(row.get(task_type_col)).strip()
            if task_type:
                subj_parts.append(f"[{task_type}]")

        # タイトル用の管理番号（整形済み）
        mng = clean_mng_num(row["管理番号"])
        if mng:
            subj_parts.append(mng)

        # Description用の管理番号（整形前の元データ）
        original_mng = ""
        if "元管理番号" in row:
            original_mng_value = row.get("元管理番号", "")
            if (
                pd.notna(original_mng_value)
                and str(original_mng_value).strip()
                and str(original_mng_value).strip().lower() != "nan"
            ):
                original_mng = str(original_mng_value).strip()

        # 件名に物件名を付ける
        name = row.get(name_col, "") if name_col else ""
        if name and str(name).strip():
            subj_parts.append(str(name).strip())

        subj = " ".join(subj_parts) if subj_parts else "イベント"

        if not subj and fallback_event_name_column and fallback_event_name_column in row:
            subj = format_description_value(row.get(fallback_event_name_column, ""))

        # 日時処理
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
            # 日時がどうしてもパースできない行は飛ばす
            continue

        # 場所
        location = ""
        if addr_col and addr_col in row:
        location_value = row.get(addr_col, "")
        if pd.notna(location_value):  # ← NaNチェックを追加
        location = str(location_value).strip()  # ← 必ず文字列に変換
        # "北海道札幌市" を削除
        if "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "").strip()
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

        output_records.append(
            {
                "Subject": subj,
                "Start Date": start.strftime("%Y/%m/%d"),
                "Start Time": start_time_display,
                "End Date": end_display.strftime("%Y/%m/%d"),
                "End Time": end_time_display,
                "All Day Event": "True" if is_all_day else "False",
                "Description": description,
                "Location": location,
                "Private": "True" if private_event else "False",
            }
        )

    return pd.DataFrame(output_records) if output_records else pd.DataFrame()


