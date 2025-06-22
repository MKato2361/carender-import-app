import pickle
import os
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import pandas as pd
import re

SCOPES = ["https://www.googleapis.com/auth/calendar"]
# TOKEN_FILE = "token.pickle" # 複数ユーザー対応のため、このファイルは使用しない

def authenticate_google():
    creds = None
    
    # 1. まず現在のセッションの認証情報がst.session_stateにあるか確認します
    if 'credentials' in st.session_state and st.session_state['credentials'] and st.session_state['credentials'].valid:
        creds = st.session_state['credentials']
        return creds

    # 2. 認証情報が有効でない、または期限切れの場合、更新または再認証を行います
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # トークンが期限切れでリフレッシュトークンがある場合、トークンをリフレッシュします
            try:
                creds.refresh(Request())
                # リフレッシュされた認証情報をsession_stateに保存します
                st.session_state['credentials'] = creds
                st.info("認証トークンを更新しました。")
                st.rerun() # トークン更新後、アプリを再実行して変更を反映
            except Exception as e:
                st.error(f"トークンのリフレッシュに失敗しました。再認証してください: {e}")
                st.session_state['credentials'] = None
                creds = None
        else: # 有効な認証情報がない場合、新しい認証フローを開始します
            try:
                # Streamlit Secretsからクライアント情報を取得
                client_config = {
                    "installed": {
                        "client_id": st.secrets["google"]["client_id"],
                        "client_secret": st.secrets["google"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"] # コンソール認証用
                    }
                }
                flow = Flow.from_client_config(client_config, SCOPES)
                flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
                auth_url, _ = flow.authorization_url(prompt='consent')

                st.info("以下のURLをブラウザで開いて、表示されたコードをここに貼り付けてください：")
                st.write(auth_url)
                code = st.text_input("認証コードを貼り付けてください:")

                if code:
                    flow.fetch_token(code=code)
                    creds = flow.credentials
                    # 新しい認証情報をsession_stateに保存します
                    st.session_state['credentials'] = creds
                    st.success("Google認証が完了しました！")
                    st.rerun() # 認証成功後、アプリを再読み込み
            except Exception as e:
                st.error(f"Google認証に失敗しました: {e}")
                st.session_state['credentials'] = None
                return None

    return creds

def add_event_to_calendar(service, calendar_id, event_data):
    """
    Googleカレンダーにイベントを追加します。
    """
    try:
        event = service.events().insert(calendarId=calendar_id, body=event_data).execute()
        # st.success(f"イベント '{event.get('summary')}' を追加しました。")
        return event # 登録されたイベントオブジェクトを返す
    except Exception as e:
        st.error(f"イベントの追加に失敗しました: {e}")
        return None

def delete_events_from_calendar(service, calendar_id, start_datetime, end_datetime):
    # Google Calendar APIはUTC時間を要求するため、JSTをUTCに変換
    jst = timezone(timedelta(hours=9))
    start_utc = start_datetime.astimezone(timezone.utc).isoformat() + 'Z'
    end_utc = end_datetime.astimezone(timezone.utc).isoformat() + 'Z'

    st.info(f"{start_datetime.strftime('%Y/%m/%d')}から{end_datetime.strftime('%Y/%m/%d')}までの削除対象イベントを検索中...")
    
    all_events_to_delete = []
    page_token = None

    with st.spinner("イベントを検索中..."):
        while True:
            try:
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=start_utc,
                    timeMax=end_utc,
                    singleEvents=True,
                    orderBy='startTime',
                    pageToken=page_token
                ).execute()
                events = events_result.get('items', [])
                all_events_to_delete.extend(events) # 取得したイベントをリストに追加

                page_token = events_result.get('nextPageToken')
                if not page_token:
                    break 
            except Exception as e:
                st.error(f"イベントの検索中にエラーが発生しました: {e}")
                return 0 # エラーが発生したら処理を中断

    total_events = len(all_events_to_delete)
    
    # 削除対象イベントがない場合、ここでリターン
    if total_events == 0:
        return 0

    # Step 2: 取得したイベントを削除（プログレスバー表示）
    progress_bar = st.progress(0)
    
    for i, event in enumerate(all_events_to_delete):
        event_summary = event.get('summary', '不明なイベント')
        try:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            progress_bar.progress((i + 1) / total_events, text=f"削除中: {event_summary} ({i+1}/{total_events})")
        except Exception as e:
            st.error(f"イベント '{event_summary}' の削除に失敗しました: {e}")
    
    progress_bar.empty() # プログレスバーをクリア
    return total_events # 削除したイベントの総数を返す

# 新規追加: 指定範囲のカレンダーイベントを取得する関数
def get_existing_calendar_events(service, calendar_id, start_datetime, end_datetime):
    jst = timezone(timedelta(hours=9))
    start_utc = start_datetime.astimezone(timezone.utc).isoformat() + 'Z'
    end_utc = end_datetime.astimezone(timezone.utc).isoformat() + 'Z'
    
    events = []
    page_token = None
    try:
        while True:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=start_utc,
                timeMax=end_utc,
                singleEvents=True, # 繰り返しイベントを展開
                orderBy='startTime',
                pageToken=page_token
            ).execute()
            events.extend(events_result.get('items', []))
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        return events
    except Exception as e:
        st.error(f"既存のGoogleカレンダーイベントの取得に失敗しました: {e}")
        return []

# 新規追加: イベントを更新する関数
def update_event_in_calendar(service, calendar_id, event_id, new_event_data):
    try:
        updated_event = service.events().update(
            calendarId=calendar_id, eventId=event_id, body=new_event_data
        ).execute()
        # st.success(f"イベント '{updated_event.get('summary')}' を更新しました。")
        return updated_event
    except Exception as e:
        st.error(f"イベントの更新に失敗しました (ID: {event_id}): {e}")
        return None

# 新規追加: ExcelデータとGoogleカレンダーデータを比較し、アクションを決定する関数
def reconcile_events(excel_df: pd.DataFrame, existing_gcal_events: list):
    # 'WorkOrderNumber'が空のExcel行は既にexcel_parserでフィルタリングされている前提
    # ここでは、WorkOrderNumberを持つExcelイベントとGCalイベントの比較を行う
    
    events_to_add_to_gcal = []    # ExcelにWO番号があり、GCalに該当WO番号がないもの
    events_to_update_in_gcal = [] # ExcelにWO番号があり、GCalに該当WO番号があり、内容に変更があるもの
    events_to_skip_due_to_no_change = [] # 新たに追加: 変更がないためスキップされたイベント

    # 既存のGoogleカレンダーイベントをWorkOrderNumberでマッピング（Descriptionから抽出）
    gcal_events_by_wo_number = {}
    for event in existing_gcal_events:
        description = event.get('description', '')
        # Descriptionの先頭から「作業指示書:XXXX / 」の形式を抽出し、数字のみを取得
        match = re.match(r"^作業指示書:(\d+)\s*/?\s*", description) # \d+ で数字のみに限定
        if match:
            wo_number = match.group(1).strip()
            # 同じ作業指示書番号のイベントが複数ある可能性も考慮し、リストで保持
            if wo_number not in gcal_events_by_wo_number:
                gcal_events_by_wo_number[wo_number] = []
            gcal_events_by_wo_number[wo_number].append(event)


    for index, excel_row in excel_df.iterrows():
        excel_wo_number = excel_row['WorkOrderNumber']
        
        # このreconcile_events関数は、excel_parser側でWorkOrderNumberがない行を既に除外しているため
        # ここではexcel_wo_numberが空であることは基本的にないはずだが、念のためガード
        if not excel_wo_number:
            # st.warning(f"Excel行に作業指示書番号がありません。この行はスキップされます。: {excel_row.get('Subject', '無題')}")
            continue # WorkOrderNumberがない場合は更新・新規登録の対象外とする

        # Excel行からGoogleカレンダーAPIのbodyを生成
        # (このevent_data_from_excelは、GCalへの登録/更新に使われる最終データ)
        event_data_from_excel = {
            'summary': excel_row['Subject'],
            'location': excel_row['Location'] if pd.notna(excel_row['Location']) else '',
            'description': excel_row['Description'] if pd.notna(excel_row['Description']) else '',
            # 'private'ではなく'transparency'を使用
            'transparency': 'transparent' if excel_row['Private'] == "True" else 'opaque'
        }

        # 日時情報の整形
        start_dt_excel_obj = None
        end_dt_excel_obj = None
        if excel_row['All Day Event'] == "True":
            start_date_obj = datetime.strptime(excel_row['Start Date'], "%Y/%m/%d").date()
            end_date_obj = datetime.strptime(excel_row['End Date'], "%Y/%m/%d").date() + timedelta(days=1) # Google Calendar APIの仕様
            event_data_from_excel['start'] = {'date': start_date_obj.strftime("%Y-%m-%d")}
            event_data_from_excel['end'] = {'date': end_date_obj.strftime("%Y-%m-%d")}
            start_dt_excel_obj = start_date_obj # 比較用
            end_dt_excel_obj = end_date_obj # 比較用、API仕様と合わせる
        else:
            start_dt_str = f"{excel_row['Start Date']} {excel_row['Start Time']}"
            end_dt_str = f"{excel_row['End Date']} {excel_row['End Time']}"
            start_dt_excel_obj = datetime.strptime(start_dt_str, "%Y/%m/%d %H:%M")
            end_dt_excel_obj = datetime.strptime(end_dt_str, "%Y/%m/%d %H:%M")
            event_data_from_excel['start'] = {'dateTime': start_dt_excel_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}
            event_data_from_excel['end'] = {'dateTime': end_dt_excel_obj.isoformat(), 'timeZone': 'Asia/Tokyo'}

        # Excelの作業指示書番号が既存のイベントにあるかチェック
        matched_gcal_events = gcal_events_by_wo_number.get(excel_wo_number, [])

        if not matched_gcal_events:
            # ExcelにあるがGoogleカレンダーにない場合は新規登録リストに追加
            # WorkOrderNumberもイベントデータに含めて、後で表示できるようにする
            event_data_from_excel['WorkOrderNumber'] = excel_wo_number
            events_to_add_to_gcal.append(event_data_from_excel)
        else:
            # マッチした既存イベントがある場合、更新が必要かチェック
            # 同じWO_NUMBERで複数のGCalイベントがある場合、内容が変更されている最初のイベントを更新対象とする
            found_event_to_update = None
            found_no_change_event = None # 変更がなかったイベントを追跡

            for gcal_event in matched_gcal_events:
                # 既存GCalイベントのデータを取得
                gcal_summary = gcal_event.get('summary', '')
                gcal_location = gcal_event.get('location', '')
                gcal_description = gcal_event.get('description', '')
                gcal_transparency = gcal_event.get('transparency', 'opaque') # デフォルトは'opaque'

                # GoogleカレンダーイベントのDescriptionからも作業指示書部分を削除して比較
                # gcal_description_for_comp = re.sub(r"^作業指示書:\d+\s*/?\s*", "", gcal_description).strip()
                # summaryからWO番号部分を削除して比較
                # gcal_summary_for_comp = re.sub(r"^(\d+)\s+", "", gcal_summary).strip() # 不要になったためコメントアウト
                
                # ExcelのDescriptionから作業指示書部分を削除して比較
                excel_description_for_comp = event_data_from_excel['description']
                # excel_description_for_comp = re.sub(rf"^作業指示書:{re.escape(excel_wo_number)}\s*/?\s*", "", excel_description_for_comp).strip() # 不要になったためコメントアウト


                # 開始/終了日時をdatetimeオブジェクトに変換して比較
                gcal_start_dt_obj_comp = None
                gcal_end_dt_obj_comp = None
                
                if 'date' in gcal_event['start']: # 終日イベント
                    gcal_start_dt_obj_comp = datetime.strptime(gcal_event['start']['date'], '%Y-%m-%d').date()
                    gcal_end_dt_obj_comp = datetime.strptime(gcal_event['end']['date'], '%Y-%m-%d').date() - timedelta(days=1)
                elif 'dateTime' in gcal_event['start']: # 時間指定イベント
                    # Googleカレンダーから取得した日時文字列はUTCであることが多いので、タイムゾーンを考慮してJSTに変換
                    gcal_start_dt_obj_comp = datetime.fromisoformat(gcal_event['start']['dateTime'].replace('Z', '+00:00')).astimezone(timezone(timedelta(hours=9))).replace(tzinfo=None) # タイムゾーン情報を除去して比較
                    gcal_end_dt_obj_comp = datetime.fromisoformat(gcal_event['end']['dateTime'].replace('Z', '+00:00')).astimezone(timezone(timedelta(hours=9))).replace(tzinfo=None) # タイムゾーン情報を除去して比較


                has_changed = False

                # 各フィールドの比較
                if gcal_summary != event_data_from_excel['summary']: # Subject全体を比較
                    has_changed = True
                if gcal_location != event_data_from_excel['location']:
                    has_changed = True
                if gcal_description != event_data_from_excel['description']: # Description全体を比較
                    has_changed = True
                if gcal_transparency != event_data_from_excel['transparency']: # 非公開設定の比較
                    has_changed = True

                # 日時の比較
                if excel_row['All Day Event'] == "True":
                    # Excel側が終日
                    if not ('date' in gcal_event['start'] and \
                            gcal_start_dt_obj_comp == start_dt_excel_obj and \
                            gcal_end_dt_obj_comp == end_dt_excel_obj):
                        has_changed = True
                else:
                    # Excel側が時間指定
                    if not ('dateTime' in gcal_event['start'] and \
                            gcal_start_dt_obj_comp == start_dt_excel_obj and \
                            gcal_end_dt_obj_comp == end_dt_excel_obj):
                        has_changed = True
                
                # ここで has_changed が true なら、このイベントを更新対象とする
                if has_changed:
                    found_event_to_update = gcal_event
                    break # 最初に見つかった変更済みイベントを更新対象とする
                else:
                    # 変更がなかったイベントも記録
                    found_no_change_event = excel_row.to_dict() # DataFrameの行を辞書に変換して保存

            if found_event_to_update:
                # 更新対象イベントの情報を追加
                events_to_update_in_gcal.append({
                    'id': found_event_to_update['id'],
                    'old_summary': found_event_to_update.get('summary', '不明'),
                    'new_data': event_data_from_excel
                })
            elif found_no_change_event:
                # 変更がなかった場合はスキップリストに追加
                events_to_skip_due_to_no_change.append(found_no_change_event)

    return events_to_add_to_gcal, events_to_update_in_gcal, events_to_skip_due_to_no_change
