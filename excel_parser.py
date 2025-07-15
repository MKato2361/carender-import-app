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

def get_available_columns_for_event_name(merged_df):
    """イベント名に使用可能な列を取得"""
    # 日時系の列を除外
    exclude_keywords = ["日時", "開始", "終了", "予定", "時間", "date", "time", "start", "end"]
    available_columns = []
    
    for col in merged_df.columns:
        col_lower = str(col).lower()
        if not any(keyword in col_lower for keyword in exclude_keywords):
            available_columns.append(col)
    
    return available_columns

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event, fallback_event_name_column=None, show_selector=True):
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
                st.info(f"ファイル '{uploaded_file.name}' に '管理番号' が見つかりません。代替のイベント名を使用します。")
                df["管理番号"] = ""  # 空の管理番号列を作成
            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")
            continue  # エラーが発生したファイルはスキップして続行

    if not dataframes:
        return pd.DataFrame()

    # --- 修正箇所：複数のファイルで同一名称の列がある場合の統合処理 ---
    # まず、すべてのDataFrameの管理番号列を文字列に変換
    for df in dataframes:
        df['管理番号'] = df['管理番号'].astype(str)

    # 最初のDataFrameを結合のベースとする
    merged_df = dataframes[0]
    
    # 2つ目以降のDataFrameを結合
    for i, df in enumerate(dataframes[1:], 2):
        # 共通の列を見つけ、管理番号以外は結合しない
        cols_to_merge = [col for col in df.columns if col == "管理番号" or col not in merged_df.columns]
        
        # 結合対象の列を抽出し、管理番号をキーに結合
        merged_df = pd.merge(merged_df, df[cols_to_merge], on="管理番号", how="outer")

    # 管理番号の重複を削除し、一意なエントリのみにする
    merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
    
    # 管理番号が空の場合は重複削除をスキップ
    if not merged_df["管理番号"].str.strip().eq("").all():
        merged_df.drop_duplicates(subset="管理番号", inplace=True)
    # --- 修正ここまで ---

    # 列の検索（既存の処理）
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    
    # 日時列の検索（代替処理を追加）
    start_col = find_closest_column(merged_df.columns, ["予定開始"])
    if not start_col:
        start_col = find_closest_column(merged_df.columns, ["開始日時", "開始時間", "開始"])
    
    end_col = find_closest_column(merged_df.columns, ["予定終了"])
    if not end_col:
        end_col = find_closest_column(merged_df.columns, ["終了日時", "終了時間", "終了"])
    
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])

    # 必要な列（日時）が見つからない場合はエラー
    if not all([start_col, end_col]):
        st.error("必要な列（日時関連）が見つかりません。予定開始・予定終了、または開始日時・終了日時が必要です。")
        return pd.DataFrame()

    # イベント名選択が必要かチェック
    if show_selector:
        has_mng, has_name = check_missing_columns(merged_df)
        if not (has_mng and has_name):
            # イベント名選択が必要な場合は、処理を中断してUIを表示
            return "SHOW_SELECTOR"

    merged_df = merged_df.dropna(subset=[start_col, end_col])

    output = []
    for _, row in merged_df.iterrows():
        mng = clean_mng_num(row["管理番号"])
        name = row.get(name_col, "")
        
        # イベント名の決定処理
        if mng and name:
            # 管理番号と物件名がある場合（既存の処理）
            subj = f"{mng}{name}"
        elif name:
            # 物件名のみがある場合
            subj = name
        elif mng:
            # 管理番号のみがある場合
            subj = mng
        elif fallback_event_name_column and fallback_event_name_column in row:
            # 代替列が指定されている場合
            fallback_value = row.get(fallback_event_name_column, "")
            subj = format_description_value(fallback_value)
        else:
            # どれもない場合はデフォルト値
            subj = "イベント"

        try:
            start = pd.to_datetime(row[start_col])
            end = pd.to_datetime(row[end_col])
            
            # 時間が00:00:00の場合は全日イベントとして処理
            if start.time() == datetime.time(0, 0, 0) and end.time() == datetime.time(0, 0, 0):
                # 終了日を1日前に調整（Outlookの全日イベント形式）
                end = end - datetime.timedelta(days=1)
                
        except Exception as e:
            st.warning(f"日時の変換に失敗しました: {row[start_col]} - {row[end_col]}")
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

        output.append({
            "Subject": subj,
            "Start Date": start.strftime("%Y/%m/%d"),
            "Start Time": start.strftime("%H:%M") if start.time() != datetime.time(0, 0, 0) else "",
            "End Date": end.strftime("%Y/%m/%d"),
            "End Time": end.strftime("%H:%M") if end.time() != datetime.time(0, 0, 0) else "",
            "All Day Event": "True" if (start.time() == datetime.time(0, 0, 0) and end.time() == datetime.time(0, 0, 0)) or all_day_event else "False",
            "Description": description,
            "Location": location,
            "Private": "True" if private_event else "False"
        })

    return pd.DataFrame(output)

def get_available_columns_for_event_name(merged_df):
    """イベント名に使用可能な列を取得"""
    # 日時系の列を除外
    exclude_keywords = ["日時", "開始", "終了", "予定", "時間", "date", "time", "start", "end", "管理番号"]
    available_columns = []
    
    for col in merged_df.columns:
        col_lower = str(col).lower()
        if not any(keyword in col_lower for keyword in exclude_keywords) and col != "管理番号":
            available_columns.append(col)
    
    return available_columns

def check_missing_columns(merged_df):
    """管理番号と物件名の有無をチェック"""
    mng_col = find_closest_column(merged_df.columns, ["管理番号"])
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    
    has_mng = mng_col is not None and not merged_df[mng_col].fillna("").astype(str).str.strip().eq("").all()
    has_name = name_col is not None and not merged_df[name_col].fillna("").astype(str).str.strip().eq("").all()
    
    return has_mng, has_name

def create_event_name_selector(merged_df):
    """イベント名選択用のUI要素を作成"""
    if merged_df.empty:
        return None
    
    # 管理番号と物件名の存在チェック
    has_mng, has_name = check_missing_columns(merged_df)
    
    # 管理番号と物件名の両方がある場合はドロップダウンを表示しない
    if has_mng and has_name:
        return None
        
    # データフレームから利用可能な列を取得
    available_columns = get_available_columns_for_event_name(merged_df)
    
    if not available_columns:
        st.warning("イベント名に使用可能な列がありません。")
        return None
    
    st.write("### イベント名の設定")
    
    if not has_mng and not has_name:
        st.write("管理番号と物件名が見つからないため、以下の列をイベント名として使用できます：")
    elif not has_mng:
        st.write("管理番号が見つからないため、以下の列を追加のイベント名として使用できます：")
    elif not has_name:
        st.write("物件名が見つからないため、以下の列を追加のイベント名として使用できます：")
    
    selected_column = st.selectbox(
        "イベント名に使用する列を選択してください：",
        options=["使用しない"] + available_columns,
        index=0,
        key="event_name_selector"
    )
    
    return selected_column if selected_column != "使用しない" else None

def should_show_event_name_selector(uploaded_files):
    """イベント名選択UIを表示すべきかどうかを判定"""
    if not uploaded_files:
        return False
    
    merged_df = get_merged_dataframe(uploaded_files)
    if merged_df.empty:
        return False
    
    has_mng, has_name = check_missing_columns(merged_df)
    return not (has_mng and has_name)

def process_with_event_name_selection(uploaded_files, description_columns, all_day_event, private_event):
    """イベント名選択を含む完全な処理フロー"""
    
    # 1. イベント名選択が必要かチェック
    if should_show_event_name_selector(uploaded_files):
        merged_df = get_merged_dataframe(uploaded_files)
        
        # 2. イベント名選択UIを表示
        selected_column = create_event_name_selector(merged_df)
        
        # 3. 選択された列で処理実行
        if st.button("イベント登録を実行", key="execute_import"):
            result_df = process_excel_files(
                uploaded_files, 
                description_columns, 
                all_day_event, 
                private_event, 
                selected_column,
                show_selector=False
            )
            return result_df
        else:
            # まだボタンが押されていない場合は空のDataFrameを返す
            return pd.DataFrame()
    else:
        # 4. イベント名選択が不要な場合は直接処理
        return process_excel_files(
            uploaded_files, 
            description_columns, 
            all_day_event, 
            private_event, 
            show_selector=False
        )

def get_merged_dataframe(uploaded_files):
    """アップロードされたファイルを統合したDataFrameを返す（列選択用）"""
    dataframes = []

    if not uploaded_files:
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                df["管理番号"] = ""
            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")
            continue

    if not dataframes:
        return pd.DataFrame()

    # DataFrameの統合
    for df in dataframes:
        df['管理番号'] = df['管理番号'].astype(str)

    merged_df = dataframes[0]
    
    for i, df in enumerate(dataframes[1:], 2):
        cols_to_merge = [col for col in df.columns if col == "管理番号" or col not in merged_df.columns]
        merged_df = pd.merge(merged_df, df[cols_to_merge], on="管理番号", how="outer")

    merged_df["管理番号"] = merged_df["管理番号"].apply(clean_mng_num)
    
    if not merged_df["管理番号"].str.strip().eq("").all():
        merged_df.drop_duplicates(subset="管理番号", inplace=True)
    
    return merged_df
