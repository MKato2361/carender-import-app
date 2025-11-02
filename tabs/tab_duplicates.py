# tabs/tab_duplicates.py
import streamlit as st
import pandas as pd
from typing import List, Optional
from datetime import datetime, date, timedelta, timezone

from state.calendar_state import get_calendar, set_calendar

# ---- 依存関数のインポート（存在すれば利用、無ければフォールバック定義） -----------------
# 時間帯（JST）
try:
    from utils.timezone import JST  # プロジェクトにある場合
except Exception:
    JST = timezone(timedelta(hours=9))

# カレンダーイベント取得・期間ユーティリティ
try:
    from utils.event_utils import fetch_all_events, default_fetch_window_years
except Exception:
    # 最低限のフォールバック：必要に応じて本体側の utils を使ってください
    def fetch_all_events(service, calendar_id, time_min, time_max):
        events, page_token = [], None
        while True:
            resp = (
                service.events()
                .list(calendarId=calendar_id, timeMin=time_min, timeMax=time_max, pageToken=page_token, singleEvents=True)
                .execute()
            )
            events.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return events

    def default_fetch_window_years(years: int = 2):
        now_utc = datetime.now(timezone.utc)
        return (now_utc - timedelta(days=365 * years)).isoformat(), (now_utc + timedelta(days=365 * years)).isoformat()

# worksheet_id の抽出に使う正規表現と正規化関数
# （プロジェクト側に存在すればそれを利用。無ければフォールバック定義）
RE_WORKSHEET_ID = None
normalize_worksheet_id = None
try:
    # 例：utils.worksheet_utils に入っているケース
    from utils.worksheet_utils import RE_WORKSHEET_ID as _RE_WS, normalize_worksheet_id as _norm_ws
    RE_WORKSHEET_ID, normalize_worksheet_id = _RE_WS, _norm_ws
except Exception:
    pass

if RE_WORKSHEET_ID is None:
    import re
    # 説明文中に "worksheet_id: XXX" もしくは "作業指示書: XXX" のような形式が含まれる想定のフォールバック
    RE_WORKSHEET_ID = re.compile(r"(?:worksheet[_\- ]?id|作業指示書(?:番号)?)\s*[:：]\s*([A-Za-z0-9_\-]+)", re.IGNORECASE)

if normalize_worksheet_id is None:
    def normalize_worksheet_id(s: str) -> str:
        return (s or "").strip()


# ====================================================================================
# タブ4：重複イベントの検出・削除
#  - 仕様：現行踏襲（検出 → 手動削除 / 自動削除（古い/新しい））
#  - カレンダー選択はサイドバーと同期（全タブ共通）
#  - ロジックは元コードと同一。外部依存のみ安全化。
# ====================================================================================
def render_tab_duplicates(service, editable_calendar_options, user_id, current_calendar_name: str):
    st.subheader("🔍 重複イベントの検出・削除")

    # ---------- カレンダー選択（タブ上部 × サイドバー同期） ----------
    if not editable_calendar_options:
        st.error("対象カレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    try:
        idx = calendar_options.index(current_calendar_name)
    except Exception:
        idx = 0

    selected_tab_calendar = st.selectbox(
        "対象カレンダーを選択",
        calendar_options,
        index=idx,
        key=f"dup_calendar_select_tab_{user_id}",
    )

    if selected_tab_calendar != current_calendar_name:
        set_calendar(user_id, selected_tab_calendar)
        st.session_state["selected_calendar_name"] = selected_tab_calendar
        st.rerun()

    selected_calendar = selected_tab_calendar
    calendar_id = editable_calendar_options[selected_calendar]

    # ---------- 以降、元コードの挙動を踏襲 ----------
    # メッセージ復元
    if "last_dup_message" in st.session_state and st.session_state["last_dup_message"]:
        msg_type, msg_text = st.session_state["last_dup_message"]
        if msg_type in {"success", "error", "info", "warning"}:
            getattr(st, msg_type)(msg_text)
        else:
            st.info(msg_text)
        st.session_state["last_dup_message"] = None

    delete_mode = st.radio(
        "削除モードを選択",
        ["手動で選択して削除", "古い方を自動削除", "新しい方を自動削除"],
        horizontal=True,
        key=f"dup_delete_mode_{user_id}",
    )

    # セッションキー初期化
    if "dup_df" not in st.session_state:
        st.session_state["dup_df"] = pd.DataFrame()
    if "auto_delete_ids" not in st.session_state:
        st.session_state["auto_delete_ids"] = []
    if "last_dup_message" not in st.session_state:
        st.session_state["last_dup_message"] = None

    def parse_created(dt_str: Optional[str]) -> datetime:
        try:
            if dt_str:
                return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        except Exception:
            pass
        return datetime.min.replace(tzinfo=timezone.utc)

    # 重複チェック実行
    if st.button("重複イベントをチェック", key=f"run_dup_check_{user_id}"):
        with st.spinner("カレンダー内のイベントを取得中..."):
            time_min, time_max = default_fetch_window_years(2)
            events = fetch_all_events(service, calendar_id, time_min, time_max)

        if not events:
            st.session_state["last_dup_message"] = ("info", "イベントが見つかりませんでした。")
            st.session_state["dup_df"] = pd.DataFrame()
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode
            st.rerun()

        st.success(f"{len(events)} 件のイベントを取得しました。")

        rows = []
        for e in events:
            desc = (e.get("description") or "").strip()
            m = RE_WORKSHEET_ID.search(desc) if desc else None
            worksheet_id = normalize_worksheet_id(m.group(1)) if m else None
            start_time = e["start"].get("dateTime", e["start"].get("date"))
            end_time = e["end"].get("dateTime", e["end"].get("date"))
            rows.append(
                {
                    "id": e.get("id"),
                    "summary": e.get("summary", ""),
                    "worksheet_id": worksheet_id,
                    "created": e.get("created"),
                    "start": start_time,
                    "end": end_time,
                }
            )

        df = pd.DataFrame(rows)
        df_valid = df[df["worksheet_id"].notna()].copy()
        dup_mask = df_valid.duplicated(subset=["worksheet_id"], keep=False)
        dup_df = df_valid[dup_mask].sort_values(["worksheet_id", "created"])

        st.session_state["dup_df"] = dup_df
        if dup_df.empty:
            st.session_state["last_dup_message"] = ("info", "重複している作業指示書番号は見つかりませんでした。")
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode
            st.rerun()

        # 自動削除モードなら対象IDを計算して保持
        if delete_mode != "手動で選択して削除":
            auto_delete_ids: List[str] = []
            for _, group in dup_df.groupby("worksheet_id"):
                group_sorted = group.sort_values(
                    ["created", "id"],
                    key=lambda s: s.map(parse_created) if s.name == "created" else s,
                    ascending=True,
                )
                if len(group_sorted) <= 1:
                    continue
                if delete_mode == "古い方を自動削除":
                    delete_targets = group_sorted.iloc[:-1]
                elif delete_mode == "新しい方を自動削除":
                    delete_targets = group_sorted.iloc[1:]
                else:
                    continue
                auto_delete_ids.extend(delete_targets["id"].tolist())

            st.session_state["auto_delete_ids"] = auto_delete_ids
            st.session_state["current_delete_mode"] = delete_mode
        else:
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode

        st.rerun()

    # 結果表示と削除操作
    if not st.session_state["dup_df"].empty:
        dup_df = st.session_state["dup_df"]
        current_mode = st.session_state.get("current_delete_mode", "手動で選択して削除")

        st.warning(
            f"⚠️ {dup_df['worksheet_id'].nunique()} 種類の重複作業指示書が見つかりました。（合計 {len(dup_df)} イベント）"
        )
        st.dataframe(
            dup_df[["worksheet_id", "summary", "created", "start", "end", "id"]],
            use_container_width=True,
        )

        # 手動削除
        if current_mode == "手動で選択して削除":
            delete_ids = st.multiselect(
                "削除するイベントを選択してください（イベントIDで指定）",
                dup_df["id"].tolist(),
                key=f"manual_delete_ids_{user_id}",
            )
            confirm = st.checkbox(
                "削除操作を確認しました",
                value=False,
                key=f"manual_del_confirm_{user_id}",
            )

            if st.button("🗑️ 選択したイベントを削除", type="primary", disabled=not confirm, key=f"run_manual_delete_{user_id}"):
                deleted_count = 0
                errors: List[str] = []
                for eid in delete_ids:
                    try:
                        service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                        deleted_count += 1
                    except Exception as e:
                        errors.append(f"イベントID {eid} の削除に失敗: {e}")

                if deleted_count > 0:
                    st.session_state["last_dup_message"] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")

                if errors:
                    st.error("以下のイベントの削除に失敗しました:\n" + "\n".join(errors))
                    if deleted_count == 0:
                        st.session_state["last_dup_message"] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                st.session_state["dup_df"] = pd.DataFrame()
                st.rerun()

        # 自動削除
        else:
            auto_delete_ids = st.session_state["auto_delete_ids"]
            if not auto_delete_ids:
                st.info("削除対象のイベントが見つかりませんでした。")
            else:
                st.warning(f"以下のモードで {len(auto_delete_ids)} 件のイベントを自動削除します: **{current_mode}**")
                st.write(auto_delete_ids)

                confirm = st.checkbox(
                    "削除操作を確認しました",
                    value=False,
                    key=f"auto_del_confirm_final_{user_id}",
                )
                if st.button("🗑️ 自動削除を実行", type="primary", disabled=not confirm, key=f"run_auto_delete_{user_id}"):
                    deleted_count = 0
                    errors: List[str] = []
                    for eid in auto_delete_ids:
                        try:
                            service.events().delete(calendarId=calendar_id, eventId=eid).execute()
                            deleted_count += 1
                        except Exception as e:
                            errors.append(f"イベントID {eid} の削除に失敗: {e}")

                    if deleted_count > 0:
                        st.session_state["last_dup_message"] = ("success", f"✅ {deleted_count} 件のイベントを削除しました。")

                    if errors:
                        st.error("以下のイベントの削除に失敗しました:\n" + "\n".join(errors))
                        if deleted_count == 0:
                            st.session_state["last_dup_message"] = ("error", "⚠️ 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                    st.session_state["dup_df"] = pd.DataFrame()
                    st.rerun()
