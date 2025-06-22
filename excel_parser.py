import pandas as pd
import re
import datetime
import streamlit as st

def clean_work_order_number(value):
    if pd.isna(value):
        return ""
    # 数字のみを抽出し、文字列として返す
    return re.sub(r"[^0-9]", "", str(value))

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
        # 整数値の場合は小数点以下を表示しない
        return str(int(val)) if val.is_integer() else str(round(val, 2))
    return str(val)

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event, strict_work_order_match=False):
    # strict_work_order_match: Trueの場合、WorkOrderNumberが取得できない行は出力から除外する
    # Falseの場合（登録タブ向け）、WorkOrderNumberがなくても全行を出力する
    dataframes = []

    if not uploaded_files:
        st.warning("Excelファイルをアップロードしてください。")
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            
            # 検索キーワードに「作業指示書」を追加
            work_order_col = find_closest_column(df.columns, ["作業指示書番号", "作業指示書", "wo_number", "workorder"])
            
            if work_order_col:
                # 'WorkOrderNumber' カラムを常に作成し、値があれば整形して格納
                df["WorkOrderNumber"] = df[work_order_col].apply(clean_work_order_number)
            else:
                # 作業指示書番号の列が見つからなければ、空文字列でカラムを作成
                df["WorkOrderNumber"] = "" 
                st.warning(f"ファイル '{uploaded_file.name}' に '作業指示書番号' または '作業指示書' を示す列が見つかりませんでした。このファイルのイベントは、作業指示書番号による更新対象にはなりません。")
            
            dataframes.append(df)
        except Exception as e:
            st.error(f"{uploaded_file.name} の読み込みに失敗しました: {e}")
            return pd.DataFrame() # エラー時は空のDataFrameを返す

    if not dataframes:
        return pd.DataFrame()

    merged_df = pd.concat(dataframes, ignore_index=True)

    name_col = find_closest_column(merged_df.columns, ["物件名", "イベント名", "概要"])
    start_col = find_closest_column(merged_df.columns, ["予定開始", "開始日時", "開始日"])
    end_col = find_closest_column(merged_df.columns, ["予定終了", "終了日時", "終了日"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地", "場所"])

    if not all([name_col, start_col, end_col]):
        st.error("必要な列（物件名・予定開始・予定終了）が見つかりません。")
        return pd.DataFrame()

    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        work_order_number = str(row.get("WorkOrderNumber", "")).strip()

        # strict_work_order_match が True の場合、WorkOrderNumber がない行はスキップ
        if strict_work_order_match and not work_order_number:
            continue # この行は更新対象外なのでスキップ

        name = row.get(name_col, "")
        
        # Subject の整形: 作業指示書番号がある場合のみ先頭に追加
        subj = f"{work_order_number} {name}" if work_order_number else name
        subj = subj.strip() # 前後の空白を削除

        try:
            start = pd.to_datetime(row[start_col])
            end = pd.to_datetime(row[end_col])
        except Exception:
            continue

        location = row.get(addr_col, "")
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")
        location = location.strip() # 前後の空白を削除

        # Description の先頭に「作業指示書:12345678 / 」を挿入（WorkOrderNumberがある場合のみ）
        base_description_parts = []
        if work_order_number: # ここでWorkOrderNumberがあるか確認
            base_description_parts.append(f"作業指示書:{work_order_number}")

        # 選択された項目をDescriptionに追加
        selected_description_parts = [
            format_description_value(row.get(col)) for col in description_columns if col in row
        ]
        
        # 空でない部分だけを結合し、Descriptionを生成
        # 作業指示書番号と選択項目が両方ない場合は空にする
        if not base_description_parts and not selected_description_parts:
            description = ""
        else:
            description = " / ".join(base_description_parts + [p for p in selected_description_parts if p.strip()])
        description = description.strip()


        # 終日イベントの場合の調整
        if all_day_event:
            start_date_str = start.strftime("%Y/%m/%d")
            # Google Calendar APIの終日イベントは終了日-1日なので、表示上は元の終了日
            end_date_str = end.strftime("%Y/%m/%d")
            start_time_str = ""
            end_time_str = ""
            is_all_day = "True"
        else:
            start_date_str = start.strftime("%Y/%m/%d")
            end_date_str = end.strftime("%Y/%m/%d")
            start_time_str = start.strftime("%H:%M")
            end_time_str = end.strftime("%H:%M")
            is_all_day = "False"

        output.append({
            "WorkOrderNumber": work_order_number, # 作業指示書番号
            "Subject": subj,
            "Start Date": start_date_str,
            "Start Time": start_time_str,
            "End Date": end_date_str,
            "End Time": end_time_str,
            "Location": location,
            "Description": description,
            "All Day Event": is_all_day,
            "Private": "True" if private_event else "False"
        })

    return pd.DataFrame(output)
