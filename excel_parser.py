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

def generate_event_title(row, mng_col, name_col, available_columns):
    """
    イベントタイトルを生成する関数
    管理番号+物件名が基本だが、ない場合は他の項目を使用
    """
    title_parts = []
    
    # 管理番号がある場合は追加
    if mng_col and mng_col in row:
        mng_value = clean_mng_num(row[mng_col])
        if mng_value:
            title_parts.append(mng_value)
    
    # 物件名がある場合は追加
    if name_col and name_col in row:
        name_value = row.get(name_col, "")
        if pd.notna(name_value) and str(name_value).strip():
            title_parts.append(str(name_value))
    
    # 管理番号と物件名の両方がある場合はそれを使用
    if len(title_parts) >= 2:
        return "".join(title_parts)
    elif len(title_parts) == 1:
        return title_parts[0]
    
    # 管理番号と物件名がない場合、他の項目を使用
    # 使用可能な項目を優先度順に定義
    fallback_keywords = [
        ["件名", "タイトル", "項目", "名称", "title", "subject"],
        ["内容", "詳細", "description", "content"],
        ["作業", "work", "task"],
        ["種別", "種類", "type", "category"],
        ["顧客", "customer", "client"],
        ["会社", "company", "corp"]
    ]
    
    for keywords in fallback_keywords:
        fallback_col = find_closest_column(available_columns, keywords)
        if fallback_col and fallback_col in row:
            fallback_value = row.get(fallback_col, "")
            if pd.notna(fallback_value) and str(fallback_value).strip():
                return str(fallback_value)
    
    # 最後の手段として、最初の非空の列を使用
    for col in available_columns:
        if col in row:
            value = row.get(col, "")
            if pd.notna(value) and str(value).strip():
                return str(value)
    
    # すべて失敗した場合は無名イベントとして扱う
    return "無名イベント"

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event):
    dataframes = []

    if not uploaded_files:
        st.warning("Excelファイルをアップロードしてください。")
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            
            # 管理番号列の処理（あれば処理、なければスキップ）
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            
            dataframes.append(df)
            
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")
            return pd.DataFrame()

    if not dataframes:
        return pd.DataFrame()

    # 複数のファイルで同一名称の列がある場合の統合処理
    # 管理番号がある場合のみ管理番号で結合、ない場合は単純に縦に結合
    has_mng_num = any("管理番号" in df.columns for df in dataframes)
    
    if has_mng_num:
        # 管理番号がある場合は従来の結合方法を使用
        for df in dataframes:
            if "管理番号" in df.columns:
                df['管理番号'] = df['管理番号'].astype(str)
        
        merged_df = dataframes[0]
        for i, df in enumerate(dataframes[1:], 2):
            if "管理番号" in df.columns:
                cols_to_merge = [col for col in df.columns if col == "管理番号" or col not in merged_df.columns]
                merged_df = pd.merge(merged_df, df[cols_to_merge], on="管理番号", how="outer")
            else:
                # 管理番号がないファイルは単純に縦結合
                merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)
        
        # 管理番号の重複を削除
        if "管理番号" in merged_df.columns:
            merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
            merged_df.drop_duplicates(subset="管理番号", inplace=True)
    else:
        # 管理番号がない場合は全て縦結合
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)

    # 必要な列を検索（存在しない場合もある）
    name_col = find_closest_column(merged_df.columns, ["物件名", "件名", "タイトル", "項目"])
    start_col = find_closest_column(merged_df.columns, ["予定開始", "開始", "開始日時", "start"])
    end_col = find_closest_column(merged_df.columns, ["予定終了", "終了", "終了日時", "end"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地", "場所", "location"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書", "指示書", "番号"])

    # 開始・終了時間は必須
    if not start_col or not end_col:
        st.error("必要な列（開始時間・終了時間）が見つかりません。")
        st.info("以下のような列名を使用してください：")
        st.info("- 開始時間: 予定開始, 開始, 開始日時, start")
        st.info("- 終了時間: 予定終了, 終了, 終了日時, end")
        return pd.DataFrame()

    # 開始・終了時間が空でない行のみを残す
    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        # 管理番号列の取得（ない場合はNone）
        mng_col = "管理番号" if "管理番号" in merged_df.columns else None
        
        # イベントタイトルを生成
        subj = generate_event_title(row, mng_col, name_col, merged_df.columns)

        try:
            start = pd.to_datetime(row[start_col])
            end = pd.to_datetime(row[end_col])
        except Exception as e:
            st.warning(f"日時の解析に失敗しました（行をスキップ）: {e}")
            continue

        # 場所の処理
        location = row.get(addr_col, "") if addr_col else ""
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")

        # 説明文の作成
        description = " / ".join(
            [format_description_value(row.get(col)) for col in description_columns if col in row and pd.notna(row.get(col))]
        )

        # 作業指示書を先頭に追加（整数化して表示）
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
