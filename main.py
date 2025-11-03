# main.py
import streamlit as st

from tabs.tab_upload import render_tab_upload
from tabs.tab_register import render_tab_register
from tabs.tab_delete import render_tab_delete
from tabs.tab_duplicates import render_tab_duplicates
from tabs.tab_export import render_tab_export

from state.calendar_state import get_calendar

# ---- Google API service / カレンダー一覧の取得はあなたの既存処理を利用してください ----
# ここではダミー/例示として editable_calendar_options を仮置きします。
# 実際はあなたの main (14).py の認証・取得ロジックをそのまま前段で実行し、下記に渡してください。
def _get_service_and_calendars():
    service = st.session_state.get("google_calendar_service")  # 既存のserviceをセッションから受け取る想定
    editable_calendar_options = st.session_state.get("editable_calendar_options", {})
    return service, editable_calendar_options

def main():
    st.set_page_config(page_title="Calendar Import App", layout="wide")

    service, editable_calendar_options = _get_service_and_calendars()

    tabs = st.tabs([
        "1. ファイルのアップロード",
        "2. イベントの登録",
        "3. イベントの削除",
        "4. 重複イベントの検出・削除",
        "5. イベントのExcel出力",
    ])

    # 現在のカレンダー名（共通）
    current_calendar_name = get_calendar(
        user_id=st.session_state.get("user_id", "default_user"),
        editable_calendar_options=editable_calendar_options,
        default_name=list(editable_calendar_options.keys())[0] if editable_calendar_options else None
    )

    with tabs[0]:
        render_tab_upload()

    with tabs[1]:
        if not service or not editable_calendar_options:
            st.warning("Googleカレンダーの認証/カレンダー取得が未設定です。既存の初期化処理を main() 上部に追加してください。")
        else:
            render_tab_register(
                service=service,
                editable_calendar_options=editable_calendar_options,
                user_id=st.session_state.get("user_id", "default_user"),
                current_calendar_name=current_calendar_name,
            )

    with tabs[2]:
        if not service or not editable_calendar_options:
            st.warning("Googleカレンダーの認証/カレンダー取得が未設定です。")
        else:
            render_tab_delete(
                service=service,
                editable_calendar_options=editable_calendar_options,
                user_id=st.session_state.get("user_id", "default_user"),
                current_calendar_name=current_calendar_name,
            )

    with tabs[3]:
        if not service or not editable_calendar_options:
            st.warning("Googleカレンダーの認証/カレンダー取得が未設定です。")
        else:
            render_tab_duplicates(
                service=service,
                editable_calendar_options=editable_calendar_options,
                user_id=st.session_state.get("user_id", "default_user"),
                current_calendar_name=current_calendar_name,
            )

    with tabs[4]:
        render_tab_export(
            editable_calendar_options=editable_calendar_options,
            user_id=st.session_state.get("user_id", "default_user"),
            current_calendar_name=current_calendar_name,
        )

if __name__ == "__main__":
    main()
