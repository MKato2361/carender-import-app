import streamlit as st
import pandas as pd
import datetime
import calendar_utils
import excel_parser

from firebase_auth import login_page, get_firebase_user_id

st.set_page_config(
    page_title="イベントインポートアプリ",
    page_icon="📅",
    layout="wide"
)

def main():
    """
    Streamlitアプリケーションのメイン関数
    """
    user = login_page()
    if user:
        st.title("ExcelからGoogleカレンダーへイベントをインポート")
        
        # ファイルアップローダー
        uploaded_file = st.file_uploader(
            "Excelファイルをアップロードしてください",
            type=["xlsx", "xls"],
            help="ファイルには'開始日時', '終了日時', 'イベント名'の3つの列が必要です。"
        )
        
        if uploaded_file is not None:
            # Excelファイルの解析
            df = excel_parser.parse_excel(uploaded_file)
            if df is not None:
                st.write("### アップロードされたデータプレビュー")
                st.dataframe(df)

                # Googleカレンダー認証
                creds = calendar_utils.authenticate_google()

                if creds:
                    service = calendar_utils.get_google_service(creds)
                    
                    if service:
                        calendars = calendar_utils.get_all_calendars(service)
                        
                        if calendars:
                            calendar_summaries = [cal['summary'] for cal in calendars]
                            selected_calendar_summary = st.selectbox(
                                "イベントを追加するカレンダーを選択してください",
                                calendar_summaries
                            )
                            
                            # 既存イベントの削除オプション
                            delete_existing = st.checkbox(
                                "登録前に同じ名前の既存イベントを削除しますか？"
                            )

                            if st.button("カレンダーに登録"):
                                # 選択されたカレンダーIDを取得
                                selected_calendar_id = calendar_utils.get_calendar_id_by_summary(calendars, selected_calendar_summary)

                                if selected_calendar_id:
                                    success_count = 0
                                    fail_count = 0
                                    
                                    with st.spinner("カレンダーにイベントを登録中..."):
                                        for index, row in df.iterrows():
                                            event_summary = row["イベント名"]
                                            start_time_str = row["開始日時"]
                                            end_time_str = row["終了日時"]
                                            
                                            # 既存イベントの削除
                                            if delete_existing:
                                                calendar_utils.delete_events_by_summary(
                                                    service, 
                                                    selected_calendar_id, 
                                                    event_summary
                                                )

                                            # イベントの作成
                                            created_event = calendar_utils.create_event(
                                                service, 
                                                selected_calendar_id, 
                                                event_summary, 
                                                start_time_str, 
                                                end_time_str,
                                                'Asia/Tokyo'
                                            )
                                            if created_event:
                                                success_count += 1
                                            else:
                                                fail_count += 1

                                    st.success(f"{success_count}個のイベントがカレンダーに登録されました！")
                                    if fail_count > 0:
                                        st.warning(f"{fail_count}個のイベントは登録に失敗しました。")
                                else:
                                    st.error("選択されたカレンダーが見つかりません。")

# アプリケーションのエントリポイント
if __name__ == "__main__":
    main()
