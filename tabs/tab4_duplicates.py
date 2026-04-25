from services.settings_service import get_setting as get_user_setting, set_setting as set_user_setting
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from typing import List, Optional
import re
import unicodedata

def _get_current_user_key(fallback: str = "") -> str:
    """設定保存用のユーザーキーを取得（優先: uid -> email）。"""
    return (
        st.session_state.get("user_id")
        or st.session_state.get("firebase_uid")
        or st.session_state.get("localId")
        or st.session_state.get("uid")
        or st.session_state.get("user_email")
        or fallback
        or ""
    )

# --- 正規表現（main.pyと同じものをコピー） ---
RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")

def normalize_worksheet_id(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    return unicodedata.normalize("NFKC", s).strip()

def parse_created(dt_str: Optional[str]) -> datetime:
    try:
        if dt_str:
            return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        pass
    return datetime.min.replace(tzinfo=timezone.utc)


def render_tab4_duplicates(service, editable_calendar_options, fetch_all_events):
    st.subheader("重複イベントの検出・削除")

    # メッセージ復元
    if "last_dup_message" in st.session_state and st.session_state["last_dup_message"]:
        msg_type, msg_text = st.session_state["last_dup_message"]
        if msg_type in {"success", "error", "info", "warning"}:
            getattr(st, msg_type)(msg_text)
        else:
            st.info(msg_text)
        st.session_state["last_dup_message"] = None

    # カレンダー選択（サイドバーの基準カレンダーを初期値に。タブ側は永続化しない）
    calendar_options = list(editable_calendar_options.keys())
    if not calendar_options:
        st.error("利用可能なカレンダーがありません。")
        return

    base_calendar = (
        st.session_state.get("base_calendar_name")
        or st.session_state.get("selected_calendar_name")
        or calendar_options[0]
    )
    if base_calendar not in calendar_options:
        base_calendar = calendar_options[0]

    select_key = "dup_calendar_select"
    if (select_key not in st.session_state) or (st.session_state.get(select_key) not in calendar_options):
        st.session_state[select_key] = base_calendar

    selected_calendar = st.selectbox(
        "対象カレンダーを選択",
        calendar_options,
        key=select_key,
    )

    calendar_id = editable_calendar_options[selected_calendar]


    # 削除モード
    delete_mode = st.radio(
        "削除モードを選択",
        ["手動で選択して削除", "古い方を自動削除", "新しい方を自動削除"],
        horizontal=True,
        key="dup_delete_mode"
    )

    # Session 初期化
    if "dup_df" not in st.session_state:
        st.session_state["dup_df"] = pd.DataFrame()
    if "auto_delete_ids" not in st.session_state:
        st.session_state["auto_delete_ids"] = []
    if "last_dup_message" not in st.session_state:
        st.session_state["last_dup_message"] = None

    # ===== 重複チェック =====
    if st.button("重複イベントをチェック", key="run_dup_check"):

        with st.spinner("カレンダー内のイベントを取得中..."):
            # 2年分の検索範囲（default_fetch_window_years の代替）
            now_utc = datetime.now(timezone.utc)
            time_min = (now_utc - timedelta(days=365*2)).isoformat()
            time_max = (now_utc + timedelta(days=365*2)).isoformat()
            events = fetch_all_events(service, calendar_id, time_min, time_max)

        if not events:
            st.session_state["last_dup_message"] = ("info", "イベントが見つかりませんでした。")
            st.session_state["dup_df"] = pd.DataFrame()
            st.session_state["auto_delete_ids"] = []
            st.session_state["current_delete_mode"] = delete_mode
            st.rerun()

        st.success(f"{len(events)} 件のイベントを取得しました。")

        # worksheet_id を抽出
        rows = []
        for e in events:
            desc = (e.get("description") or "").strip()
            m = RE_WORKSHEET_ID.search(desc)
            worksheet_id = normalize_worksheet_id(m.group(1)) if m else None
            start_time = e["start"].get("dateTime", e["start"].get("date"))
            end_time   = e["end"].get("dateTime", e["end"].get("date"))
            rows.append({
                "id": e["id"],
                "summary": e.get("summary", ""),
                "worksheet_id": worksheet_id,
                "created": e.get("created"),
                "start": start_time,
                "end": end_time,
            })

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

        # 自動削除モード
        if delete_mode != "手動で選択して削除":
            auto_delete_ids: List[str] = []
            for _, group in dup_df.groupby("worksheet_id"):
                group_sorted = group.sort_values(
                    ["created", "id"],
                    key=lambda s: s.map(parse_created) if s.name == "created" else s,
                    ascending=True
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

    # ===== テーブル & 削除UI =====
    if not st.session_state["dup_df"].empty:
        dup_df = st.session_state["dup_df"]
        current_mode = st.session_state.get("current_delete_mode", "手動で選択して削除")

        st.warning(f" {dup_df['worksheet_id'].nunique()} 種類の重複作業指示書が見つかりました。（合計 {len(dup_df)} イベント）")
        st.dataframe(dup_df[["worksheet_id", "summary", "created", "start", "end", "id"]], use_container_width=True)

        # ===== 手動削除 =====
        if current_mode == "手動で選択して削除":
            delete_ids = st.multiselect(
                "削除するイベントを選択してください（イベントIDで指定）",
                dup_df["id"].tolist(),
                key="manual_delete_ids"
            )
            confirm = st.checkbox("削除操作を確認しました", value=False, key="manual_del_confirm")

            if st.button("選択したイベントを削除", type="primary", disabled=not confirm, key="run_manual_delete"):
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
                        st.session_state["last_dup_message"] = ("error", " 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                st.session_state["dup_df"] = pd.DataFrame()
                st.rerun()

        # ===== 自動削除 =====
        else:
            auto_delete_ids = st.session_state["auto_delete_ids"]

            if not auto_delete_ids:
                st.info("削除対象のイベントが見つかりませんでした。")
            else:
                st.warning(f"以下のモードで {len(auto_delete_ids)} 件のイベントを自動削除します: **{current_mode}**")
                st.dataframe({"削除対象イベントID": auto_delete_ids}, use_container_width=True)

                confirm = st.checkbox("削除操作を確認しました", value=False, key="auto_del_confirm_final")

                if st.button("自動削除を実行", type="primary", disabled=not confirm, key="run_auto_delete"):
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
                            st.session_state["last_dup_message"] = ("error", " 削除処理中にエラーが発生しました。詳細はログを確認してください。")

                    st.session_state["dup_df"] = pd.DataFrame()
                    st.rerun()
