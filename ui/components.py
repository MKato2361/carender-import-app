"""
ui/components.py
全タブで使う共通 UI コンポーネント

各タブに散在していたカレンダーカード・確認ボタンパターンをここに集約。
st.* の使用は許可。
"""
from __future__ import annotations
import streamlit as st


def calendar_card(
    calendar_names: list[str],
    session_key: str,
    base_calendar: str,
    label: str = "カレンダー",
    share_on: bool = True,
    allow_change: bool = True,
) -> str:
    """
    カレンダー選択カード（青枠）を描画し、選択中のカレンダー名を返す。

    - share_on=True のとき session_key をサイドバーの基準カレンダーに同期する。
    - allow_change=False のとき変更 expander を表示しない。
    """
    if share_on:
        st.session_state[session_key] = base_calendar
    elif (session_key not in st.session_state) or (
        st.session_state.get(session_key) not in calendar_names
    ):
        st.session_state[session_key] = base_calendar

    current = st.session_state.get(session_key, base_calendar)

    st.markdown(
        f"""
<div style="border:2px solid #1E88E5;border-radius:10px;padding:14px 18px;
            margin-bottom:8px;background:var(--color-background-info);">
  <div style="font-size:12px;font-weight:600;color:var(--color-text-info);
              margin-bottom:4px;">📅 {label}（必ず確認）</div>
  <div style="font-size:20px;font-weight:700;color:var(--color-text-info);">{current}</div>
</div>
""",
        unsafe_allow_html=True,
    )

    if share_on:
        st.caption("サイドバーの「基準カレンダー」と連動しています。")
    elif allow_change:
        with st.expander("カレンダーを変更する"):
            st.selectbox(
                "カレンダーを選択",
                calendar_names,
                key=session_key,
                label_visibility="collapsed",
            )

    return st.session_state.get(session_key, base_calendar)


def confirm_action_button(
    button_label: str,
    confirm_label: str,
    session_key: str,
    on_confirm,
    on_cancel=None,
    button_type: str = "primary",
) -> None:
    """
    2ステップ確認ボタン。

    1 回目クリック → confirm_label の警告を表示。
    2 回目クリック（「実行」）→ on_confirm() を呼ぶ。
    「キャンセル」 → on_cancel() または session_key フラグをリセット。
    """
    if not st.session_state.get(session_key):
        if st.button(button_label, type=button_type, use_container_width=True):
            st.session_state[session_key] = True
            st.rerun()
    else:
        st.warning(confirm_label)
        col_ok, col_cancel = st.columns([3, 1])
        with col_ok:
            if st.button("✅ 実行する", type="primary", use_container_width=True):
                st.session_state[session_key] = False
                on_confirm()
        with col_cancel:
            if st.button("キャンセル", use_container_width=True):
                st.session_state[session_key] = False
                if on_cancel:
                    on_cancel()
                st.rerun()


def file_summary_bar(has_work: bool, has_outside: bool, on_confirm, on_clear) -> None:
    """
    ファイル取込済み時のコンパクト 1 行サマリー + 確定 / クリアボタン。
    tab1 専用。
    """
    files   = st.session_state.get("uploaded_files", [])
    outside = st.session_state.get("uploaded_outside_work_file")
    df      = st.session_state.get("merged_df_for_selector")

    if has_work:
        names     = [getattr(f, "name", "Unknown") for f in files]
        row_count = len(df) if df is not None else "—"
        badge     = f"{len(names)} ファイル / {row_count} 行"
        summary   = "、".join(names)
        kind      = "作業指示書"
    else:
        summary = getattr(outside, "name", "")
        badge   = "1 ファイル"
        kind    = "作業外予定"

    col_info, col_btn, col_clear = st.columns([5, 3, 1])

    with col_info:
        st.markdown(
            f"**{kind}** &nbsp;"
            f"<span style='background:var(--color-background-success);"
            f"color:var(--color-text-success);font-size:12px;font-weight:600;"
            f"padding:2px 8px;border-radius:4px;'>{badge}</span>  \n"
            f"<span style='font-size:12px;color:var(--color-text-secondary);'>{summary}</span>",
            unsafe_allow_html=True,
        )
    with col_btn:
        if st.button("✅ 確定してカレンダー登録へ →", type="primary", use_container_width=True):
            on_confirm()
    with col_clear:
        if st.button("🗑️", help="アップロードをクリア", use_container_width=True):
            on_clear()

def handle_http_error(e, action: str = "操作") -> None:
    """Google API HttpError をユーザー向けメッセージに変換して表示する。"""
    import streamlit as st
    try:
        status = e.resp.status
    except Exception:
        status = None
    messages = {
        401: f"{action}に失敗しました。Googleセッションが切れています。ページを再読み込みして再連携してください。",
        403: f"{action}に失敗しました。このカレンダーへの書き込み権限がありません。",
        404: f"{action}に失敗しました。対象のイベントが見つかりません（すでに削除済みの可能性があります）。",
        429: f"{action}に失敗しました。APIのリクエスト上限に達しました。しばらく待ってから再試行してください。",
    }
    st.error(messages.get(status, f"{action}に失敗しました（エラーコード: {status}）。しばらく待ってから再試行してください。"))

