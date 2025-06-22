# calendar_utils.py
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
TOKEN_FILE = "token.pickle"

# authenticate_google (変更なし)
# add_event_to_calendar (変更なし)
# delete_events_from_calendar (変更なし)
# list_events_in_range (変更なし)
# get_existing_calendar_events (変更なし)
# update_event_in_calendar (変更なし)


def reconcile_events(excel_df: pd.DataFrame, existing_gcal_events: list):
    # 'WorkOrderNumber'が空のExcel行は既にexcel_parserでフィルタリングされている前提
    # ここでは、WorkOrderNumberを持つExcelイベントとGCalイベントの比較を行う
    
    events_to_add_to_gcal = []    # ExcelにWO番号があり、GCalに該当WO番号がないもの
    events_to_update_in_gcal = [] # ExcelにWO番号があり、GCalに該当WO番号があり、内容に変更があるもの

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
            end_dt_excel_obj = end_date_obj - timedelta(days=1) # 比較用、API仕様と合わせる
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
            events_to_add_to_gcal.append(event_data_from_excel)
        else:
            # マッチした既存イベントがある場合、更新が必要かチェック
            # 同じWO_NUMBERで複数のGCalイベントがある場合、内容が変更されている最初のイベントを更新対象とする
            found_event_to_update = None
            
            for gcal_event in matched_gcal_events:
                # 既存GCalイベントのデータを取得
                gcal_summary = gcal_event.get('summary', '')
                gcal_location = gcal_event.get('location', '')
                gcal_description = gcal_event.get('description', '')
                gcal_transparency = gcal_event.get('transparency', 'opaque') # デフォルトは'opaque'

                # GoogleカレンダーイベントのDescriptionからも作業指示書部分を削除して比較
                gcal_description_for_comp = re.sub(r"^作業指示書:\d+\s*/?\s*", "", gcal_description).strip()
                
                # ExcelのDescriptionから作業指示書部分を削除して比較
                excel_description_for_comp = event_data_from_excel['description']
                excel_description_for_comp = re.sub(rf"^作業指示書:{re.escape(excel_wo_number)}\s*/?\s*", "", excel_description_for_comp).strip()


                # 開始/終了日時をdatetimeオブジェクトに変換して比較
                gcal_start_dt_obj_comp = None
                gcal_end_dt_obj_comp = None
                
                if 'date' in gcal_event['start']: # 終日イベント
                    gcal_start_dt_obj_comp = datetime.strptime(gcal_event['start']['date'], '%Y-%m-%d').date()
                    gcal_end_dt_obj_comp = datetime.strptime(gcal_event['end']['date'], '%Y-%m-%d').date() - timedelta(days=1)
                elif 'dateTime' in gcal_event['start']: # 時間指定イベント
                    # Googleカレンダーから取得した日時文字列はUTCであることが多いので、タイムゾーンを考慮してJSTに変換
                    gcal_start_dt_obj_comp = datetime.fromisoformat(gcal_event['start']['dateTime'].replace('Z', '+00:00')).astimezone(timezone(timedelta(hours=9)))
                    gcal_end_dt_obj_comp = datetime.fromisoformat(gcal_event['end']['dateTime'].replace('Z', '+00:00')).astimezone(timezone(timedelta(hours=9)))

                has_changed = False

                # 各フィールドの比較
                if gcal_summary != event_data_from_excel['summary']:
                    has_changed = True
                if gcal_location != event_data_from_excel['location']:
                    has_changed = True
                if gcal_description_for_comp != excel_description_for_comp:
                    has_changed = True
                if gcal_transparency != event_data_from_excel['transparency']: # 非公開設定の比較
                    has_changed = True

                # 日時の比較
                # 比較対象のオブジェクトタイプが異なる可能性があるため、日付部分のみ、または日時全体で慎重に比較
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

            if found_event_to_update:
                # 更新対象イベントの情報を追加
                events_to_update_in_gcal.append({
                    'id': found_event_to_update['id'],
                    'old_summary': found_event_to_update.get('summary', '不明'),
                    'new_data': event_data_from_excel
                })
            # else: 同じWO_NUMBERだが変更がない場合は、events_to_add_to_gcalにもevents_to_update_in_gcalにも追加しない（スキップ）

    return events_to_add_to_gcal, events_to_update_in_gcal