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

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event, strict_work_order_match=False):
    dataframes = []

    if not uploaded_files:
        st.warning("Excelファイルをアップロードしてください。")
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            
            # 各ファイルの列情報を保持
            mng_col = find_closest_column(df.columns, ["管理番号", "作業指示書番号"])
            name_col = find_closest_column(df.columns, ["物件名", "イベント名", "概要"])
            start_col = find_closest_column(df.columns, ["予定開始", "開始日", "開始"])
            end_col = find_closest_column(df.columns, ["予定終了", "終了日", "終了"])
            start_time_col = find_closest_column(df.columns, ["開始時刻", "開始時間"])
            end_time_col = find_closest_column(df.columns, ["終了時刻", "終了時間"])
            addr_col = find_closest_column(df.columns, ["住所", "所在地", "場所"])

            # 必須列のチェック
            if not all([name_col, start_col, end_col]):
                st.warning(f"ファイル '{uploaded_file.name}' に必要な列（物件名・予定開始・予定終了）が見つかりません。このファイルはスキップされます。")
                continue

            # 作業指示書番号が必須の場合のフィルタリング
            if strict_work_order_match:
                if not mng_col:
                    st.warning(f"ファイル '{uploaded_file.name}' に '管理番号' または '作業指示書番号' 列が見つかりません。このファイルは更新/削除の処理対象外です。")
                    continue
                # 作業指示書番号の欠損値を持つ行を削除
                df = df.dropna(subset=[mng_col])
                if df.empty:
                    st.info(f"ファイル '{uploaded_file.name}' に有効な作業指示書番号を持つ行がありませんでした。")
                    continue
            
            # 日付列の欠損値を持つ行を削除
            df = df.dropna(subset=[start_col, end_col])

            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の処理中にエラーが発生しました: {e}")
            continue

    if not dataframes:
        return pd.DataFrame()

    merged_df = pd.concat(dataframes, ignore_index=True)

    output = []
    for _, row in merged_df.iterrows():
        # 作業指示書番号の処理
        wo_number_raw = row.get(mng_col, "") if mng_col else ""
        wo_number = clean_mng_num(wo_number_raw)

        # イベント名の処理
        name = row.get(name_col, "")
        subj = f"{wo_number} {name}".strip() if wo_number else name.strip()
        if not subj: # サブジェクトが空の場合はスキップ
            continue

        try:
            # 日付と時刻の結合
            start_date_val = row[start_col]
            end_date_val = row[end_col]

            # PandasのTimestampオブジェクトへの変換を試みる
            start = pd.to_datetime(start_date_val)
            end = pd.to_datetime(end_date_val)

            # 時間列が存在し、かつ終日イベントでない場合のみ時間情報を考慮
            start_time_val = row.get(start_time_col) if start_time_col else None
            end_time_val = row.get(end_time_col) if end_time_col else None

            if not all_day_event and pd.notna(start_time_val) and pd.notna(end_time_val):
                # 時刻は文字列またはtimedelta形式を想定
                # datetime.timeオブジェクトに変換して結合
                if isinstance(start_time_val, datetime.time):
                    start_time = start_time_val
                elif isinstance(start_time_val, datetime.timedelta):
                    total_seconds = int(start_time_val.total_seconds())
                    hours, remainder = divmod(total_seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    start_time = datetime.time(hours, minutes, seconds)
                else: # その他の形式 (例: '9:00')
                    start_time = pd.to_datetime(str(start_time_val)).time()

                if isinstance(end_time_val, datetime.time):
                    end_time = end_time_val
                elif isinstance(end_time_val, datetime.timedelta):
                    total_seconds = int(end_time_val.total_seconds())
                    hours, remainder = divmod(total_seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    end_time = datetime.time(hours, minutes, seconds)
                else: # その他の形式
                    end_time = pd.to_datetime(str(end_time_val)).time()
                
                # 日付と時刻を結合
                start_datetime_obj = datetime.datetime.combine(start.date(), start_time)
                end_datetime_obj = datetime.datetime.combine(end.date(), end_time)
            else:
                # 終日イベントの場合、時刻は無視し、日付のみを使用
                start_datetime_obj = datetime.datetime.combine(start.date(), datetime.time.min)
                end_datetime_obj = datetime.datetime.combine(end.date(), datetime.time.min)
                # 終日イベントの場合、Google Calendar APIは終了日を翌日に設定する必要がある
                if all_day_event:
                    end_datetime_obj += datetime.timedelta(days=1)

        except Exception as e:
            st.warning(f"日付または時刻の解析に失敗しました。この行はスキップされます: {subj} - エラー: {e}")
            continue

        # --- ここを修正 ---
        location_raw = row.get(addr_col) # None, NaN, または値
        if pd.isna(location_raw): # NaN かどうかをチェック
            location = ""
        else:
            location = str(location_raw).strip() # 文字列に変換してからstrip
            if "北海道札幌市" in location:
                location = location.replace("北海道札幌市", "")
        # --- 修正終わり ---

        # 説明欄の生成
        description_parts = []
        # 作業指示書番号を説明欄の先頭に追加
        if wo_number:
            description_parts.append(f"作業指示書:{wo_number}")

        # 選択された他の列を説明に追加
        for col in description_columns:
            if col in row and pd.notna(row[col]): # NaNではないことを確認
                description_parts.append(format_description_value(row[col]))
        
        description = " / ".join(description_parts)


        output_row = {
            "WorkOrderNumber": wo_number,
            "Subject": subj,
            "Location": location,
            "Description": description,
            "All Day Event": str(all_day_event), # Booleanを文字列で保存
            "Private": str(private_event) # Booleanを文字列で保存
        }
        
        if all_day_event:
            output_row["Start Date"] = start_datetime_obj.strftime("%Y/%m/%d")
            output_row["End Date"] = end_datetime_obj.strftime("%Y/%m/%d")
            output_row["Start Time"] = "" # 終日の場合は時間を空にする
            output_row["End Time"] = ""
        else:
            output_row["Start Date"] = start_datetime_obj.strftime("%Y/%m/%d")
            output_row["Start Time"] = start_datetime_obj.strftime("%H:%M")
            output_row["End Date"] = end_datetime_obj.strftime("%Y/%m/%d")
            output_row["End Time"] = end_datetime_obj.strftime("%H:%M")

        output.append(output_row)

    return pd.DataFrame(output)
