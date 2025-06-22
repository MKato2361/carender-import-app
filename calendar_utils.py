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
    # 修正: token.pickleからの読み込みを削除し、st.session_stateのみに依存
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
    elif not creds: # 有効な認証情報がない場合、新しい認証フローを開始します
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
    # タイムゾーン情報を付与
    jst = timezone(timedelta(hours=9))
    
    # datetimeオブジェクトにタイムゾーン情報を付与し、UTCに変換してからフォーマット
    # timeMin は開始日の00:00:00 JST から
    time_min_with_tz = jst.localize(start_datetime) if start_datetime.tzinfo is None else start_datetime
    time_min_utc = time_min_with_tz.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

    # timeMax は終了日の23:59:59 JST まで
    # main.pyから渡される end_datetime は既に 23:59:59 を含んだ datetime オブジェクトのはず
    time_max_with_tz = jst.localize(end_datetime) if end_datetime.tzinfo is None else end_datetime
    time_max_utc = time_max_with_tz.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


    st.info(f"{start_datetime.strftime('%Y/%m/%d')}から{end_datetime.strftime('%Y/%m/%d')}までの削除対象イベントを検索中...")
    
    all_events_to_delete = []
    page_token = None

    with st.spinner("イベントを検索中..."):
        while True:
            try:
                events_result = service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min_utc,
                    timeMax=time_max_utc,
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

# 修正: 指定範囲のカレンダーイベントを取得する関数
def get_existing_calendar_events(service, calendar_id, start_datetime, end_datetime):
    jst = timezone(timedelta(hours=9))
    
    # datetimeオブジェクトにタイムゾーン情報を付与し、UTCに変換してからフォーマット
    # timeMin は開始日の00:00:00 JST から
    time_min_with_tz = jst.localize(start_datetime) if start_datetime.tzinfo is None else start_datetime
    time_min_utc = time_min_with_tz.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    # timeMax は終了日の23:59:59 JST まで
    # main.pyから渡される end_datetime は既に 23:59:59 を含んだ datetime オブジェクトのはず
    time_max_with_tz = jst.localize(end_datetime) if end_datetime.tzinfo is None else end_datetime
    time_max_utc = time_max_with_tz.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    events = []
    page_token = None
    try:
        while True:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min_utc,
                timeMax=time_max_utc,
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

# イベントを更新する関数
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

# ExcelデータとGoogleカレンダーデータを比較し、アクションを決定する関数
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
        excel_wo_number = str(excel_row['WorkOrderNumber']).strip()
        
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
        
        # Excelからの日付と時刻は、既にexcel_parserでdatetimeオブジェクトに変換されている前提
        # ここでは、そのオブジェクトをGoogle Calendar APIが要求する形式に変換する
        start_date_str = excel_row['Start Date'] # "%Y/%m/%d"形式の文字列
        end_date_str = excel_row['End Date']   # "%Y/%m/%d"形式の文字列
        start_time_str = excel_row['Start Time'] # "%H:%M"形式の文字列
        end_time_str = excel_row['End Time']   # "%H:%M"形式の文字列


        if excel_row['All Day Event'] == "True": # Excel側が終日イベント
            # excel_parserで既に翌日になっているので、そのまま使用
            start_date_obj_for_api = datetime.strptime(start_date_str, "%Y/%m/%d").strftime("%Y-%m-%d")
            end_date_obj_for_api = datetime.strptime(end_date_str, "%Y/%m/%d").strftime("%Y-%m-%d")
            event_data_from_excel['start'] = {'date': start_date_obj_for_api}
            event_data_from_excel['end'] = {'date': end_date_obj_for_api}
            
            # 比較用に日付オブジェクトを保持
            start_dt_excel_obj = datetime.strptime(start_date_str, "%Y/%m/%d").date()
            end_dt_excel_obj = datetime.strptime(end_date_str, "%Y/%m/%d").date() # Google API仕様に合わせた後の日付
        else: # Excel側が時間指定イベント
            # datetime.datetimeオブジェクトに変換してからISO形式にする
            start_dt_excel_full = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M")
            end_dt_excel_full = datetime.strptime(f"{end_date_str} {end_time_str}", "%Y/%m/%d %H:%M")

            event_data_from_excel['start'] = {'dateTime': start_dt_excel_full.isoformat(), 'timeZone': 'Asia/Tokyo'}
            event_data_from_excel['end'] = {'dateTime': end_dt_excel_full.isoformat(), 'timeZone': 'Asia/Tokyo'}

            start_dt_excel_obj = start_dt_excel_full # 比較用
            end_dt_excel_obj = end_dt_excel_full   # 比較用

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
                gcal_start_dt_obj_comp = None
                gcal_end_dt_obj_comp = None
                
                # Googleカレンダーのイベントが終日か時間指定か
                is_gcal_all_day = 'date' in gcal_event['start']

                if excel_row['All Day Event'] == "True": # Excel側が終日イベント
                    if is_gcal_all_day: # GCal側も終日イベント
                        gcal_start_dt_obj_comp = datetime.strptime(gcal_event['start']['date'], '%Y-%m-%d').date()
                        gcal_end_dt_obj_comp = datetime.strptime(gcal_event['end']['date'], '%Y-%m-%d').date() # GCalの終日イベントの終了日は翌日が含まれていない

                        # Excelのend_dt_excel_objは既にexcel_parserで翌日になっているので、
                        # Google Calendar APIの終日イベントの終了日と比較するために、
                        # GCalの終了日もExcelと同じロジックで翌日を期待する形に変換してから比較
                        # (GCal APIは終日イベントの場合、開始日+1日をend.dateに指定するのが慣習)
                        # なので、gcal_end_dt_obj_compも+1日してExcel側と比較
                        gcal_end_dt_obj_comp_plus_one = gcal_end_dt_obj_comp + timedelta(days=1)

                        if not (start_dt_excel_obj == gcal_start_dt_obj_comp and 
                                end_dt_excel_obj == gcal_end_dt_obj_comp_plus_one):
                            has_changed = True
                    else: # GCal側が時間指定イベント -> 終日イベントへの変更は変更とみなす
                        has_changed = True
                else: # Excel側が時間指定イベント
                    if is_gcal_all_day: # GCal側が終日イベント -> 時間指定イベントへの変更は変更とみなす
                        has_changed = True
                    else: # GCal側も時間指定イベント
                        # Googleカレンダーから取得した日時文字列をdatetimeオブジェクトに変換
                        try:
                            gcal_start_dt_obj_comp = datetime.fromisoformat(gcal_event['start']['dateTime'].replace('Z', '+00:00'))
                            gcal_end_dt_obj_comp = datetime.fromisoformat(gcal_event['end']['dateTime'].replace('Z', '+00:00'))
                            
                            # タイムゾーンを考慮せずに純粋な日時を比較 (日本時間基準)
                            # GCalのdateTimeはUTCなので、JSTに変換してから比較
                            jst_tz = timezone(timedelta(hours=9))
                            gcal_start_dt_obj_comp_jst = gcal_start_dt_obj_comp.astimezone(jst_tz)
                            gcal_end_dt_obj_comp_jst = gcal_end_dt_obj_comp.astimezone(jst_tz)

                            if not (start_dt_excel_obj == gcal_start_dt_obj_comp_jst and 
                                    end_dt_excel_obj == gcal_end_dt_obj_comp_jst):
                                has_changed = True
                        except ValueError as e:
                            # 日時形式のパースエラーが発生した場合も変更とみなす（データ異常）
                            st.warning(f"既存Googleカレンダーイベントの日時解析エラー({gcal_event.get('summary')}): {e}。このイベントは変更として扱われます。")
                            has_changed = True
                
                if has_changed:
                    found_event_to_update = {
                        'id': gcal_event['id'],
                        'old_summary': gcal_summary, # 変更前のサマリーを保持
                        'new_data': event_data_from_excel
                    }
                    break # 変更があったイベントが見つかったらループを抜ける
                else:
                    # 変更がなかった場合、スキップリストに追加するために情報を保持
                    found_no_change_event = excel_row.to_dict() # 元のExcel行データを保持
                    # 最初の変更なしイベントが見つかったら、それを候補とする
                    # ただし、後続で変更ありイベントが見つかる可能性があるので、breakはしない

            if found_event_to_update:
                events_to_update_in_gcal.append(found_event_to_update)
            elif found_no_change_event:
                events_to_skip_due_to_no_change.append(found_no_change_event)

    return events_to_add_to_gcal, events_to_update_in_gcal, events_to_skip_due_to_no_change
