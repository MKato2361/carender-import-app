import pandas as pd
import re
import datetime # datetimeモジュールからtimedeltaとdatetimeを直接インポートするために残す

# UI要素（streamlit）はここから削除。純粋なデータ処理ロジックのみに特化。

def clean_mng_num(value):
    """
    管理番号から特定の文字を除去し、クリーンな文字列を返します。
    例: 'ABC-123HK' -> 'ABC123'
    """
    if pd.isna(value):
        return ""
    # 数字とアルファベット以外の文字を除去し、"HK"を削除
    return re.sub(r"[^0-9A-Za-z]", "", str(value)).replace("HK", "")

def find_closest_column(columns, keywords):
    """
    指定されたキーワードに最も近い列名を検索します（大文字小文字を区別しない）。
    見つからない場合はNoneを返します。
    """
    for kw in keywords:
        for col in columns:
            if kw.lower() in str(col).lower():
                return col
    return None

def format_description_value(val):
    """
    説明文用の値をフォーマットします。
    NaNは空文字列、浮動小数点数は整数または小数点以下2桁の文字列に変換します。
    """
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        # 整数であれば整数として、そうでなければ小数点以下2桁に丸めて文字列化
        return str(int(val)) if val.is_integer() else str(round(val, 2))
    return str(val)

def format_worksheet_value(val):
    """
    作業指示書用の値をフォーマットします。
    NaNは空文字列、浮動小数点数は整数として文字列に変換します。
    """
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        # 浮動小数点数であっても整数として表示（例: 123.0 -> "123"）
        return str(int(val))
    return str(val)

def _load_and_merge_dataframes(uploaded_files):
    """
    アップロードされたExcelファイルを読み込み、'管理番号'をキーに統合します。
    """
    dataframes = []
    
    if not uploaded_files:
        raise ValueError("Excelファイルがアップロードされていません。")

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            # 列名をクリーニング（前後の空白除去）
            df.columns = [str(c).strip() for c in df.columns]
            
            mng_col = find_closest_column(df.columns, ["管理番号"])
            if mng_col:
                df["管理番号"] = df[mng_col].apply(clean_mng_num)
            else:
                # 管理番号が見つからない場合でも処理を続行できるよう空の列を作成
                df["管理番号"] = ""
            dataframes.append(df)
        except Exception as e:
            raise IOError(f"ファイル '{uploaded_file.name}' の読み込みに失敗しました: {e}")

    if not dataframes:
        raise ValueError("処理できる有効なデータがありません。")

    # 最初のDataFrameを結合のベースとする
    merged_df = dataframes[0].copy() # copy()でSettingWithCopyWarningを避ける
    merged_df['管理番号'] = merged_df['管理番号'].astype(str)
    
    # 2つ目以降のDataFrameを結合
    for df in dataframes[1:]:
        df_copy = df.copy() # copy()でSettingWithCopyWarningを避ける
        df_copy['管理番号'] = df_copy['管理番号'].astype(str)
        # 共通の列を見つけ、管理番号以外は結合しない
        # 注意: pd.mergeはデフォルトで共通の列を自動で結合しますが、
        # ここでは'管理番号'をキーに、それ以外の固有の列を追加する動作を意図しています。
        # 厳密には、同名で内容が異なる列の扱いには注意が必要ですが、このコードでは'outer'結合で対応。
        cols_to_merge = [col for col in df_copy.columns if col == "管理番号" or col not in merged_df.columns]
        merged_df = pd.merge(merged_df, df_copy[cols_to_merge], on="管理番号", how="outer")

    # 管理番号が空でない場合に重複を削除
    if not merged_df["管理番号"].str.strip().eq("").all():
        merged_df.drop_duplicates(subset="管理番号", inplace=True)
        
    return merged_df

def get_available_columns_for_event_name(df):
    """
    イベント名に使用可能な列名を取得します。
    日時関連や特定の除外キーワードを含む列、および'管理番号'を除外します。
    """
    exclude_keywords = ["日時", "開始", "終了", "予定", "時間", "date", "time", "start", "end", "all day", "private", "subject", "description", "location", "作業タイプ"]
    available_columns = []
    
    for col in df.columns:
        col_lower = str(col).lower()
        # 除外キーワードを含まず、かつ「管理番号」列でないもの
        if not any(keyword in col_lower for keyword in exclude_keywords) and col != "管理番号":
            available_columns.append(col)
            
    return available_columns

def check_event_name_columns(merged_df):
    """
    統合されたDataFrameに'管理番号'と'物件名'の列が存在し、かつデータが空でないかをチェックします。
    """
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
    all_day_event_override, # all_day_eventから名称変更し、上書き設定であることを明確に
    private_event, 
    fallback_event_name_column=None,
    add_task_type_to_event_name=False # 新しい引数を追加
):
    """
    アップロードされたExcelファイルを処理し、Outlookカレンダーインポート用のDataFrameを生成します。

    Args:
        uploaded_files (list): アップロードされたファイルのリスト。
        description_columns (list): 説明文として結合する列名のリスト。
        all_day_event_override (bool): Trueの場合、すべてのイベントを終日イベントとして扱う。
                                       Falseの場合、開始時刻と終了時刻が00:00:00の場合に終日と判断。
        private_event (bool): Trueの場合、すべてのイベントをプライベートとしてマーク。
        fallback_event_name_column (str, optional): '管理番号'や'物件名'がない場合の
                                                    代替イベント名に使用する列名。Defaults to None.
        add_task_type_to_event_name (bool): Trueの場合、イベント名の先頭に「作業タイプ」列の値を追加する。Defaults to False.

    Returns:
        pandas.DataFrame: カレンダーインポート用の整形されたデータ。
    
    Raises:
        ValueError: 必須の列が見つからない場合や、ファイルの読み込みに失敗した場合。
    """
    
    merged_df = _load_and_merge_dataframes(uploaded_files)

    # 必要な列の検索
    name_col = find_closest_column(merged_df.columns, ["物件名"])
    start_col = find_closest_column(merged_df.columns, ["予定開始", "開始日時", "開始時間", "開始"])
    end_col = find_closest_column(merged_df.columns, ["予定終了", "終了日時", "終了時間", "終了"])
    addr_col = find_closest_column(merged_df.columns, ["住所", "所在地"])
    worksheet_col = find_closest_column(merged_df.columns, ["作業指示書"])
    
    # 新しく作業タイプ列を検索
    task_type_col = find_closest_column(merged_df.columns, ["作業タイプ"])


    if not start_col: # 'end_col' がなくても1時間イベントとして処理するので、'start_col'のみを必須とする
        raise ValueError("必須の時刻列（'予定開始'、または'開始日時'など）が見つかりません。")

    output_records = []
    # 既存のdropnaを削除し、すべての行を処理する
    for index, row in merged_df.iterrows(): 
        mng = clean_mng_num(row["管理番号"])
        name = row.get(name_col, "") if name_col else ""
        
        # イベント名の決定ロジック
        subj = ""
        
        # 新しいロジック: 作業タイプを先頭に追加するかどうか
        if add_task_type_to_event_name and task_type_col and pd.notna(row.get(task_type_col)):
            task_type = str(row.get(task_type_col)).strip()
            if task_type:
                subj += f"【{task_type}】"
        
        if mng and str(mng).strip(): # 管理番号にデータがある場合
            subj += str(mng).strip()
        if name and str(name).strip(): # 物件名にデータがある場合
            if subj and not subj.endswith("【】"): # 管理番号または作業タイプがある場合は間にスペース
                subj += " " 
            subj += str(name).strip()
        
        # 管理番号も物件名も空で、代替列が指定されている場合
        if not subj and fallback_event_name_column and fallback_event_name_column in row:
            fallback_value = row.get(fallback_event_name_column, "")
            subj = format_description_value(fallback_value)
        
        # 最終的にイベント名が空の場合はデフォルト値
        if not subj:
            subj = "イベント"

        try:
            # 予定開始時刻の解析は必須
            start = pd.to_datetime(row[start_col])
            
            # 予定終了時刻の解析を試みる
            end = None
            if end_col and pd.notna(row.get(end_col)):
                try:
                    end = pd.to_datetime(row[end_col])
                except Exception:
                    # 予定終了の形式が不正な場合は、1時間イベントとして処理
                    pass # end は None のままにする
            
            # 予定終了がNoneの場合（データがない、または解析失敗）、開始時刻から1時間後を設定
            if end is None:
                if start.time() == datetime.time(0, 0, 0): # 時刻情報がないと判断できる場合 (00:00:00)
                     # 終日イベントとしたい場合
                    end = start + datetime.timedelta(days=1) # 終日イベントは翌日0時まで
                    print(f"Info: 行 {index+2} の「予定終了」が空または無効なため、終日イベントとして設定されました。")
                else: # 時刻情報がある場合
                    end = start + datetime.timedelta(hours=1)
                    print(f"Info: 行 {index+2} の「予定終了」が空または無効なため、1時間のイベントとして設定されました。")


            # 終了日時が開始日時より前の場合は調整 (重要: ここは調整せず、エラーとしてスキップが安全)
            if end < start:
                print(f"Warning: 開始日時({start})が終了日時({end})より後です。この行はスキップされます。")
                continue 
            
            # 全日イベントの判定
            # all_day_event_override が True の場合は常に True
            # そうでない場合は、開始と終了の時間が00:00:00で、かつ終了日が開始日の翌日である場合にTrue
            # または、終了日と開始日が同じで、時間が00:00:00の場合（終日イベントとして扱う）
            is_all_day = all_day_event_override or (
                start.time() == datetime.time(0, 0, 0) and 
                end.time() == datetime.time(0, 0, 0) and 
                (end.date() == start.date() + datetime.timedelta(days=1) or end.date() == start.date()) # 翌日0時まで、または当日0時まで
            )

            if is_all_day:
                # Outlookの全日イベントは終了日を1日前に設定する必要がある
                end_display = end - datetime.timedelta(days=1)
                start_time_display = ""
                end_time_display = ""
            else:
                end_display = end
                start_time_display = start.strftime("%H:%M")
                end_time_display = end.strftime("%H:%M")

        except Exception as e:
            # 日時変換エラーが発生した場合、その行をスキップ
            print(f"Warning: 日時の変換に失敗しました（行 {index+2} のデータ: {row.to_dict()}）: {e}")
            continue

        location = row.get(addr_col, "") if addr_col else ""
        if isinstance(location, str) and "北海道札幌市" in location:
            location = location.replace("北海道札幌市", "")

        # 説明文の生成
        description_parts = []
        for col in description_columns:
            if col in row:
                description_parts.append(format_description_value(row.get(col)))
        description = " / ".join(filter(None, description_parts)) # 空文字列を除外して結合

        # 作業指示書を先頭に追加（存在する場合のみ）
        worksheet_value = row.get(worksheet_col, "") if worksheet_col else ""
        if pd.notna(worksheet_value) and str(worksheet_value).strip():
            formatted_ws = format_worksheet_value(worksheet_value)
            description = f"作業指示書：{formatted_ws}/ " + description if description else f"作業指示書：{formatted_ws}"


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
