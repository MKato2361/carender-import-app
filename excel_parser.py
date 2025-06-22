import pandas as pd
import re
import datetime
import streamlit as st

def clean_mng_num(value):
    if pd.isna(value):
        return ""
    # 数字と英字のみを抽出し、"HK" を削除
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).replace("HK", "")

def find_closest_column(columns, keywords):
    # 列名の前後の空白を除去し、小文字にして比較
    cleaned_columns = {str(col).strip().lower(): col for col in columns}
    
    for kw in keywords:
        # キーワードも小文字にして比較
        kw_lower = kw.lower()
        for cleaned_col_name, original_col_name in cleaned_columns.items():
            # 完全に一致するか、部分一致でより適切なものを見つけるロジック
            # まず完全一致を優先
            if kw_lower == cleaned_col_name:
                return original_col_name
            # 部分一致でキーワードが含まれるものを探す
            if kw_lower in cleaned_col_name:
                return original_col_name # 元の列名を返す
    return None

def format_description_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        # 整数であれば整数として、そうでなければ小数点以下2桁で表示
        return str(int(val)) if val.is_integer() else str(round(val, 2))
    return str(val)

def parse_date_robustly(date_val):
    """様々な形式の日付文字列をdatetime.dateオブジェクトにパースする"""
    if pd.isna(date_val):
        return None
    
    # Pandasのto_datetimeで一般的な形式を試す
    try:
        dt_obj = pd.to_datetime(date_val)
        return dt_obj.date()
    except (ValueError, TypeError):
        pass # 次の形式を試す

    # よくある日付形式を明示的に試す
    date_formats = [
        "%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y",
        "%Y年%m月%d日", "%m月%d日" # 日本語形式
    ]
    for fmt in date_formats:
        try:
            # datetime.strptimeは文字列にしか使えないため、str()でキャスト
            return datetime.datetime.strptime(str(date_val), fmt).date()
        except (ValueError, TypeError):
            pass

    return None # どれも解析できない場合

def parse_time_robustly(time_val):
    """様々な形式の時刻をdatetime.timeオブジェクトにパースする"""
    if pd.isna(time_val):
        return datetime.time.min # 時刻がない場合は00:00:00を返す

    if isinstance(time_val, datetime.time):
        return time_val
    if isinstance(time_val, datetime.timedelta):
        total_seconds = int(time_val.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return datetime.time(hours, minutes, seconds)
    
    # 文字列または数値の場合
    s_time_val = str(time_val).strip()

    # '9:00', '09:00', '9時00分' など
    time_formats = ["%H:%M", "%H時%M分", "%H時"] # %H時00分のようなケースも対応
    for fmt in time_formats:
        try:
            return datetime.datetime.strptime(s_time_val, fmt).time()
        except ValueError:
            pass
    
    # '900', '0900' (HHMM形式)
    try:
        # 数値の場合はゼロ埋めして4桁にする
        if s_time_val.isdigit() and len(s_time_val) <= 4:
            s_time_val = s_time_val.zfill(4) # 例: '900' -> '0900'
            return datetime.datetime.strptime(s_time_val, "%H%M").time()
    except ValueError:
        pass
        
    return datetime.time.min # どれも解析できない場合はデフォルト値を返す

def process_excel_files(uploaded_files, description_columns, all_day_event, private_event, strict_work_order_match=False):
    dataframes = []

    if not uploaded_files:
        st.warning("Excelファイルをアップロードしてください。")
        return pd.DataFrame()

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            # 列名を読み込んだ時点で前後の空白を除去
            df.columns = [str(c).strip() for c in df.columns]
            
            # 各ファイルの列情報を取得
            # キーワードをさらに網羅的にする
            mng_col = find_closest_column(df.columns, ["管理番号", "作業指示書番号", "作業指示書No", "案件No", "工事番号"])
            name_col = find_closest_column(df.columns, ["物件名", "イベント名", "概要", "件名", "工事名"])
            # 日付関連のキーワードをさらに増やす
            start_col = find_closest_column(df.columns, ["予定開始", "開始日", "開始", "工事開始日", "着工日", "開始日付", "日付"])
            end_col = find_closest_column(df.columns, ["予定終了", "終了日", "終了", "工事終了日", "完工日", "終了日付", "完了日"])
            
            start_time_col = find_closest_column(df.columns, ["開始時刻", "開始時間", "開始時", "開始時間From"])
            end_time_col = find_closest_column(df.columns, ["終了時刻", "終了時間", "終了時", "終了時間To"])
            addr_col = find_closest_column(df.columns, ["住所", "所在地", "場所", "現場住所", "現場"])

            # 必須列のチェック
            missing_cols = []
            if not name_col: missing_cols.append("物件名/イベント名")
            if not start_col: missing_cols.append("予定開始/開始日")
            if not end_col: missing_cols.append("予定終了/終了日")

            if missing_cols:
                st.warning(f"ファイル '{uploaded_file.name}' に必要な列（{', '.join(missing_cols)}）が見つかりません。このファイルはスキップされます。")
                continue

            # strict_work_order_matchがTrueの場合、作業指示書番号が必須
            if strict_work_order_match:
                if not mng_col:
                    st.warning(f"ファイル '{uploaded_file.name}' に '管理番号' または '作業指示書番号' 列が見つかりません。このファイルは更新/削除の処理対象外です。")
                    continue
                # 作業指示書番号の欠損値を持つ行を削除
                df = df.dropna(subset=[mng_col]).copy()
                if df.empty:
                    st.info(f"ファイル '{uploaded_file.name}' に有効な作業指示書番号を持つ行がありませんでした。")
                    continue
            
            # 日付列の欠損値を持つ行を削除
            df = df.dropna(subset=[start_col, end_col]).copy()
            if df.empty: # 日付列が欠損した結果、空になった場合
                st.info(f"ファイル '{uploaded_file.name}' に有効な開始日/終了日を持つ行がありませんでした。")
                continue

            dataframes.append(df)
        except Exception as e:
            st.error(f"ファイル '{uploaded_file.name}' の処理中にエラーが発生しました: {e}")
            continue

    if not dataframes:
        return pd.DataFrame()

    merged_df = pd.concat(dataframes, ignore_index=True)

    output = []
    for _, row in merged_df.iterrows():
        # 各行の処理時に再度列名を探すのではなく、上で見つけた列名を使用する
        # 作業指示書番号の処理
        wo_number_raw = row.get(mng_col, "") if mng_col else ""
        wo_number = clean_mng_num(wo_number_raw)

        # イベント名の処理
        name = row.get(name_col, "")
        subj = f"{wo_number} {name}".strip() if wo_number else name.strip()
        if not subj: # サブジェクトが空の場合はスキップ
            st.warning(f"空のイベント名を持つ行がスキップされました: {row.to_dict()}")
            continue

        try:
            # 日付と時刻の結合
            start_date_obj = parse_date_robustly(row[start_col])
            end_date_obj = parse_date_robustly(row[end_col])

            if start_date_obj is None or end_date_obj is None:
                st.warning(f"日付の解析に失敗しました。この行はスキップされます: {subj} - 開始日:'{row.get(start_col)}', 終了日:'{row.get(end_col)}'")
                continue

            # 時間列が存在し、かつ終日イベントでない場合のみ時間情報を考慮
            start_time_val = row.get(start_time_col) if start_time_col else None
            end_time_val = row.get(end_time_col) if end_time_col else None

            if not all_day_event and pd.notna(start_time_val) and pd.notna(end_time_val):
                start_time_obj = parse_time_robustly(start_time_val)
                end_time_obj = parse_time_robustly(end_time_val)
                
                start_datetime_obj = datetime.datetime.combine(start_date_obj, start_time_obj)
                end_datetime_obj = datetime.datetime.combine(end_date_obj, end_time_obj)
            else:
                # 終日イベント、または時刻情報がない場合、時刻は無視し、日付のみを使用
                start_datetime_obj = datetime.datetime.combine(start_date_obj, datetime.time.min)
                end_datetime_obj = datetime.datetime.combine(end_date_obj, datetime.time.min)
                # 終日イベントの場合、Google Calendar APIは終了日を翌日に設定する必要がある
                if all_day_event:
                    end_datetime_obj += datetime.timedelta(days=1)

        except Exception as e:
            st.warning(f"日付または時刻の解析中に予期せぬエラーが発生しました。この行はスキップされます: {subj} - エラー: {e}")
            continue

        # 場所の処理
        location_raw = row.get(addr_col) # None, NaN, または値
        if pd.isna(location_raw): # NaN かどうかをチェック
            location = ""
        else:
            location = str(location_raw).strip() # 文字列に変換してからstrip
            if "北海道札幌市" in location:
                location = location.replace("北海道札幌市", "")

        # 説明欄の生成
        description_parts = []
        # 作業指示書番号を説明欄の先頭に追加
        if wo_number:
            description_parts.append(f"作業指示書:{wo_number}")

        # 選択された他の列を説明に追加
        for col in description_columns:
            # find_closest_columnの結果を直接使用するため、`col`がrowに存在するかチェック
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
