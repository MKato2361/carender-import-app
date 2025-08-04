import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# ヘルパー関数: ExcelからDataFrameを読み込む
def _read_excel_to_dataframe(file_content, filename):
    try:
        if isinstance(file_content, BytesIO):
            return pd.read_excel(file_content, engine='openpyxl')
        else:
            return pd.read_excel(filename, engine='openpyxl')
    except Exception as e:
        st.error(f"ファイル '{filename}' の読み込み中にエラーが発生しました: {e}")
        return None

# ヘルパー関数: 複数のExcelファイルを読み込み、DataFrameをマージする
def _load_and_merge_dataframes(files):
    all_dfs = []
    for f in files:
        df = _read_excel_to_dataframe(f, f.name)
        if df is not None:
            df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            all_dfs.append(df)

    if not all_dfs:
        return pd.DataFrame()

    merged_df = pd.concat(all_dfs, ignore_index=True)
    return merged_df

# ヘルパー関数: イベント名の候補となる列を抽出する
def get_available_columns_for_event_name(df):
    candidates = df.columns.tolist()
    # 既に使われがちな列は除外
    exclude_cols = ['開始日', '終了日', '開始時刻', '終了時刻', '日付', '曜日', '時刻', '場所', '住所', '作業タイプ']
    return [col for col in candidates if col not in exclude_cols]

# ヘルパー関数: イベント名生成に必須の列があるか確認
def check_event_name_columns(df):
    has_mng_data = '管理番号' in df.columns
    has_name_data = '物件名' in df.columns
    return has_mng_data, has_name_data

# ヘルパー関数: ワークシートIDを正規表現で整形する
def format_worksheet_value(value):
    if pd.notna(value):
        value_str = str(value)
        match = re.search(r'\d+', value_str)
        if match:
            return match.group(0)
    return None

def process_excel_data_for_calendar(files, description_columns, all_day_event_override, private_event, fallback_event_name_column, prepend_event_type):
    """
    Excelファイルを処理し、Googleカレンダーイベント用のDataFrameを返す。
    
    Args:
        files (list): アップロードされたExcelファイルのリスト。
        description_columns (list): 説明欄に含める列名のリスト。
        all_day_event_override (bool): Trueの場合、すべてのイベントを終日イベントとして扱う。
        private_event (bool): Trueの場合、すべてのイベントを非公開として扱う。
        fallback_event_name_column (str): イベント名に使用する代替列名。
        prepend_event_type (bool): イベント名の先頭に作業タイプを追加するかどうか。
        
    Returns:
        pd.DataFrame: Googleカレンダーイベント用のデータフレーム。
    """
    merged_df = _load_and_merge_dataframes(files)
    if merged_df.empty:
        return pd.DataFrame()

    # 日付・時間列の標準化
    date_cols = ['開始日', '日付']
    time_cols = ['開始時刻', '終了時刻', '時刻']
    
    # 終日イベントと通常イベントを識別するための列
    merged_df['All Day Event'] = 'False'
    
    # 開始日・終了日を特定
    if '開始日' in merged_df.columns:
        merged_df['Start Date'] = merged_df['開始日']
    elif '日付' in merged_df.columns:
        merged_df['Start Date'] = merged_df['日付']
    else:
        return pd.DataFrame() # 日付列がない場合は処理を中断

    # 終了日がない場合は開始日と同じにする
    if '終了日' in merged_df.columns:
        merged_df['End Date'] = merged_df['終了日']
    else:
        merged_df['End Date'] = merged_df['Start Date']

    # 終日イベントかどうかの判定
    if all_day_event_override:
        merged_df['All Day Event'] = 'True'
    else:
        has_time_data = '開始時刻' in merged_df.columns and '終了時刻' in merged_df.columns
        if not has_time_data:
            merged_df['All Day Event'] = 'True'

    # 時間列の特定とデフォルト値の設定
    if '開始時刻' in merged_df.columns:
        merged_df['Start Time'] = merged_df['開始時刻']
    else:
        merged_df['Start Time'] = '09:00'

    if '終了時刻' in merged_df.columns:
        merged_df['End Time'] = merged_df['終了時刻']
    else:
        merged_df['End Time'] = '17:00'

    # フィルタリング
    df_filtered = merged_df.dropna(subset=['Start Date']).copy()
    
    if df_filtered.empty:
        return pd.DataFrame()

    # 終日イベントのデータ型を文字列に変換
    if 'Start Date' in df_filtered.columns:
        df_filtered['Start Date'] = pd.to_datetime(df_filtered['Start Date'], errors='coerce').dt.strftime('%Y/%m/%d')
    if 'End Date' in df_filtered.columns:
        df_filtered['End Date'] = pd.to_datetime(df_filtered['End Date'], errors='coerce').dt.strftime('%Y/%m/%d')

    # 出力用DataFrameの準備
    result_df = pd.DataFrame(columns=['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Private', 'Location', 'Description'])
    
    for index, row in df_filtered.iterrows():
        # イベント名の生成
        event_name_parts = []
        has_mng_data, has_name_data = check_event_name_columns(df_filtered)

        if has_mng_data and pd.notna(row['管理番号']):
            event_name_parts.append(str(row['管理番号']))
        if has_name_data and pd.notna(row['物件名']):
            event_name_parts.append(str(row['物件名']))
        
        # prepend_event_typeがTrueで、かつ作業タイプ列が存在し、値が空でない場合、先頭に追加
        if prepend_event_type and '作業タイプ' in row and pd.notna(row['作業タイプ']):
            event_name_parts.insert(0, f"【{str(row['作業タイプ'])}】")

        if event_name_parts:
            event_name = "".join(event_name_parts)
        else:
            # 代替列が指定されていればそれを使用
            if fallback_event_name_column and fallback_event_name_column in row and pd.notna(row[fallback_event_name_column]):
                event_name = str(row[fallback_event_name_column])
            else:
                event_name = "タイトルなし"
        
        # 説明欄の生成
        description_text = ""
        worksheet_id = format_worksheet_value(row.get('管理番号'))
        if worksheet_id:
            description_text += f"作業指示書: {worksheet_id}\n"
        
        for col in description_columns:
            if col in row and pd.notna(row[col]):
                description_text += f"{col}: {row[col]}\n"

        # 場所の特定
        location = ""
        if '住所' in row and pd.notna(row['住所']):
            location = str(row['住所'])

        # 結果DataFrameに追加
        result_df.loc[len(result_df)] = {
            'Subject': event_name,
            'Start Date': row['Start Date'],
            'Start Time': row['Start Time'],
            'End Date': row['End Date'],
            'End Time': row['End Time'],
            'All Day Event': 'True' if row['All Day Event'] == 'True' else 'False',
            'Private': 'True' if private_event else 'False',
            'Location': location,
            'Description': description_text
        }

    return result_df
