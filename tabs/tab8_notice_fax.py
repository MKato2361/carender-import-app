from __future__ import annotations

from datetime import datetime, date, time, timedelta, timezone
from typing import Any, Dict, List, Optional
import io
import os
import re
import unicodedata
import zipfile

import pandas as pd
import streamlit as st
from firebase_admin import firestore

from tabs.tab6_property_master import (
    MASTER_COLUMNS,
    BASIC_COLUMNS,
    load_sheet_as_df,
    _normalize_df,
)
from utils.helpers import safe_get
from session_utils import get_user_setting, set_user_setting

# ─────────────────────────────────────────────────────────
# 定数
# ─────────────────────────────────────────────────────────
JST = timezone(timedelta(hours=9))

ASSETNUM_PATTERN = re.compile(
    r"[［\[]?\s*管理番号[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?"
)
WORKTYPE_PATTERN = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")

DISPLAY_COLS = [
    "作成",
    "管理番号",
    "物件名",
    "予定日",
    "予定時間",
    "作業タイプ",
    "テンプレファイル",
    "イベントタイトル",
    "備考",
]

STEPS = ["① 設定", "② 取得", "③ 確認", "④ 生成"]

# ─────────────────────────────────────────────────────────
# harigami モジュール（任意依存）
# ─────────────────────────────────────────────────────────
try:
    from utils.harigami_generator import (
        DEFAULT_TEMPLATE_MAP,
        extract_tags_from_description,
        build_replacements_from_event,
        generate_docx_from_template_like,
    )
    HARIGAMI_AVAILABLE = True
    HARIGAMI_IMPORT_ERROR: Optional[Exception] = None
except Exception as e:
    HARIGAMI_AVAILABLE = False
    HARIGAMI_IMPORT_ERROR = e


# ─────────────────────────────────────────────────────────
# ユーティリティ関数
# ─────────────────────────────────────────────────────────

def extract_assetnum(text: str) -> str:
    if not text:
        return ""
    m = ASSETNUM_PATTERN.search(unicodedata.normalize("NFKC", str(text)))
    return m.group(1).strip() if m else ""


def extract_worktype(text: str) -> str:
    if not text:
        return ""
    m = WORKTYPE_PATTERN.search(unicodedata.normalize("NFKC", str(text)))
    return (m.group(1) or "").strip() if m else ""


def to_utc_range_from_dates(d1: date, d2: date) -> tuple[str, str]:
    start = datetime.combine(d1, time.min, tzinfo=JST).astimezone(timezone.utc)
    end = datetime.combine(d2, time.max, tzinfo=JST).astimezone(timezone.utc)
    return start.isoformat(), end.isoformat()


def get_event_start_datetime(event: Dict[str, Any]) -> Optional[datetime]:
    start = event.get("start", {})
    if "dateTime" in start:
        try:
            dt = pd.to_datetime(start["dateTime"])
            if dt.tzinfo is None:
                dt = dt.tz_localize(timezone.utc)
            return dt.astimezone(JST).to_pydatetime()
        except Exception:
            return None
    if "date" in start:
        try:
            return datetime.combine(date.fromisoformat(start["date"]), time.min, tzinfo=JST)
        except Exception:
            return None
    return None


def fetch_events_in_range(service: Any, calendar_id: str, start_date: date, end_date: date) -> List[Dict[str, Any]]:
    if not service:
        return []
    time_min, time_max = to_utc_range_from_dates(start_date, end_date)
    events: List[Dict[str, Any]] = []
    page_token: Optional[str] = None
    while True:
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min,
            timeMax=time_max,
            maxResults=2500,
            singleEvents=True,
            orderBy="startTime",
            pageToken=page_token,
        ).execute()
        events.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return events


def get_property_master_spreadsheet_id(current_user_email: Optional[str]) -> str:
    if not current_user_email:
        return ""
    try:
        db = firestore.client()
        doc = db.collection("user_settings").document(current_user_email).get()
        return (doc.to_dict() or {}).get("property_master_spreadsheet_id") or ""
    except Exception:
        return ""


def load_property_master_view(sheets_service: Any, spreadsheet_id: str) -> pd.DataFrame:
    if not sheets_service or not spreadsheet_id:
        return pd.DataFrame()
    try:
        basic_df = load_sheet_as_df(sheets_service, spreadsheet_id, "物件基本情報", BASIC_COLUMNS)
        master_df = load_sheet_as_df(sheets_service, spreadsheet_id, "物件マスタ", MASTER_COLUMNS)
        basic_df = _normalize_df(basic_df, BASIC_COLUMNS)
        master_df = _normalize_df(master_df, MASTER_COLUMNS)

        if master_df.empty:
            merged = basic_df.copy()
            for col in MASTER_COLUMNS:
                if col not in merged.columns:
                    merged[col] = ""
            return merged

        return master_df.merge(
            basic_df[["管理番号", "物件名", "住所", "窓口会社"]],
            on="管理番号",
            how="left"
        )
    except Exception as e:
        raise RuntimeError(f"物件マスタの読み込みに失敗しました: {e}") from e


def _pm_index(pm_view_df: pd.DataFrame) -> Optional[pd.DataFrame]:
    if isinstance(pm_view_df, pd.DataFrame) and not pm_view_df.empty and "管理番号" in pm_view_df.columns:
        return pm_view_df.set_index("管理番号")
    return None


def _clear_candidates() -> None:
    for k in (
        "notice_fax_candidates_df",
        "notice_fax_events_by_id",
        "notice_fax_pm_view_df",
        "fax_zip_ready",
    ):
        st.session_state.pop(k, None)


# ─────────────────────────────────────────────────────────
# UI コンポーネント
# ─────────────────────────────────────────────────────────

def _render_step_indicator(current: int) -> None:
    cols = st.columns(len(STEPS))
    for i, (col, label) in enumerate(zip(cols, STEPS)):
        is_done = i < current
        is_current = i == current
        color = "#4CAF50" if is_done else "#1976D2" if is_current else "#9E9E9E"
        weight = "700" if is_current else "600" if is_done else "400"
        bg = "#E3F2FD" if is_current else "transparent"
        icon = "✅" if is_done else "▶" if is_current else "○"

        col.markdown(
            f"<div style='text-align:center; color:{color}; font-size:0.9rem; font-weight:{weight}; "
            f"background:{bg}; padding:8px 4px; border-radius:8px; border: 1px solid {color if is_current else 'transparent'}'>"
            f"{icon} {label}</div>",
            unsafe_allow_html=True,
        )
    st.write("")


# ─────────────────────────────────────────────────────────
# メイン関数
# ─────────────────────────────────────────────────────────

def render_tab8_notice_fax(manager, current_user_email: str) -> None:
    if not HARIGAMI_AVAILABLE:
        st.error(f"貼り紙生成モジュールが読み込めませんでした: {HARIGAMI_IMPORT_ERROR}")
        return

    if not manager.calendar_service:
        st.warning("⚠️ カレンダーサービスが初期化されていません。")
        return

    col_title, col_reset = st.columns([5, 1])
    with col_title:
        st.markdown("### 📄 貼り紙自動作成 (FAX・メール通知)")
    with col_reset:
        if st.button("🔄 リセット", use_container_width=True, help="データをクリアして最初からやり直します"):
            _clear_candidates()
            st.rerun()

    st.divider()

    # STEP 1
    _render_step_indicator(0)

    with st.container(border=True):
        st.markdown("**① 取得条件の設定**")
        c1, c2, c3 = st.columns([2, 2, 2])

        calendar_options = manager.editable_calendar_options
        calendar_name = c1.selectbox("対象カレンダー", list(calendar_options.keys()), key="fax_cal_select")
        calendar_id = calendar_options[calendar_name]

        start_date = c2.date_input("開始日", value=date.today(), key="fax_start_date")
        if "fax_end_date" not in st.session_state:
            st.session_state["fax_end_date"] = start_date + timedelta(days=7)
        end_date = c3.date_input("終了日", key="fax_end_date")

        st.markdown("---")
        c_mode, c_btn = st.columns([3, 1])
        use_master_filter = c_mode.toggle("物件マスタの『貼り紙テンプレ種別=自社』のみを対象にする", value=True)
        fetch_clicked = c_btn.button("🔍 イベントを取得", type="primary", use_container_width=True)

    # STEP 2
    if fetch_clicked:
        st.session_state.pop("fax_zip_ready", None)

        with st.status("イベントを取得中...", expanded=True) as status:
            try:
                spreadsheet_id = get_property_master_spreadsheet_id(current_user_email)
                pm_view_df = load_property_master_view(manager.sheets_service, spreadsheet_id)
                events = fetch_events_in_range(manager.calendar_service, calendar_id, start_date, end_date)

                if not events:
                    status.update(label="イベントが見つかりませんでした", state="error")
                    st.warning("指定された期間にイベントがありません。")
                else:
                    candidates = []
                    events_by_id = {}
                    pm_idx = _pm_index(pm_view_df)

                    for ev in events:
                        desc = safe_get(ev, "description") or ""
                        mgmt = extract_assetnum(desc)
                        if not mgmt:
                            continue

                        if use_master_filter and pm_idx is not None:
                            if mgmt not in pm_idx.index:
                                continue
                            if str(pm_idx.loc[mgmt].get("貼り紙テンプレ種別", "")).strip() != "自社":
                                continue

                        start_dt = get_event_start_datetime(ev)
                        candidates.append({
                            "作成": True,
                            "event_id": ev["id"],
                            "管理番号": mgmt,
                            "物件名": pm_idx.loc[mgmt].get("物件名", "") if pm_idx is not None and mgmt in pm_idx.index else ev.get("summary", ""),
                            "予定日": start_dt.strftime("%m/%d") if start_dt else "-",
                            "予定時間": start_dt.strftime("%H:%M") if start_dt else "-",
                            "作業タイプ": extract_worktype(desc),
                            "イベントタイトル": ev.get("summary", ""),
                            "備考": desc[:50] + "..." if len(desc) > 50 else desc,
                        })
                        events_by_id[ev["id"]] = ev

                    if not candidates:
                        status.update(label="条件に合うイベントがありませんでした", state="error")
                    else:
                        st.session_state["notice_fax_candidates_df"] = pd.DataFrame(candidates)
                        st.session_state["notice_fax_events_by_id"] = events_by_id
                        st.session_state["notice_fax_pm_view_df"] = pm_view_df
                        status.update(label=f"成功: {len(candidates)} 件の候補を取得しました", state="complete")
                        st.rerun()
            except Exception as e:
                status.update(label=f"エラーが発生しました: {e}", state="error")

    # STEP 3
    cand_df = st.session_state.get("notice_fax_candidates_df")
    if cand_df is not None and not cand_df.empty:
        st.divider()
        _render_step_indicator(2)

        with st.container(border=True):
            st.markdown(f"**② 作成対象の確認 ({len(cand_df)} 件)**")

            # 保存ボタンはここにだけ出す
            if "fax_zip_ready" in st.session_state:
                st.caption("保存ファイル")
                st.download_button(
                    label="📦 ZIPファイルを保存",
                    data=st.session_state["fax_zip_ready"]["data"],
                    file_name=st.session_state["fax_zip_ready"]["name"],
                    mime="application/zip",
                    use_container_width=True,
                    type="primary",
                    key="fax_zip_download_top_only",
                )
                st.markdown("---")

            col_op1, col_op2, _ = st.columns([1, 1, 3])
            if col_op1.button("✅ 全選択", key="fax_all_on"):
                cand_df["作成"] = True
                st.session_state["notice_fax_candidates_df"] = cand_df
                st.rerun()

            if col_op2.button("☐ 全解除", key="fax_all_off"):
                cand_df["作成"] = False
                st.session_state["notice_fax_candidates_df"] = cand_df
                st.rerun()

            edited_df = st.data_editor(
                cand_df,
                column_config={
                    "作成": st.column_config.CheckboxColumn(width="small"),
                    "event_id": None,
                    "管理番号": st.column_config.TextColumn(disabled=True),
                    "物件名": st.column_config.TextColumn(disabled=True),
                    "予定日": st.column_config.TextColumn(disabled=True),
                    "作業タイプ": st.column_config.TextColumn(disabled=True),
                },
                hide_index=True,
                use_container_width=True,
                key="fax_editor",
            )
            st.session_state["notice_fax_candidates_df"] = edited_df

        # STEP 4
        st.divider()
        _render_step_indicator(3)

        target_count = int(edited_df["作成"].sum())
        if target_count > 0:
            if st.button(
                f"🖨️ {target_count} 件の貼り紙を生成",
                type="primary",
                use_container_width=True,
                key="fax_generate_button_only",
            ):
                with st.status("貼り紙を生成中...") as status:
                    pm_idx = _pm_index(st.session_state["notice_fax_pm_view_df"])
                    events_by_id = st.session_state["notice_fax_events_by_id"]
                    outputs = []
                    errors = []

                    for _, row in edited_df[edited_df["作成"]].iterrows():
                        ev = events_by_id.get(row["event_id"])
                        if not ev:
                            continue
                        try:
                            out_name, content = _generate_single_docx(ev, row, pm_idx)
                            outputs.append((out_name, content))
                        except Exception as e:
                            errors.append(f"{row['管理番号']}: {e}")

                    if outputs:
                        zip_data = _pack_zip(outputs)
                        st.session_state["fax_zip_ready"] = {
                            "data": zip_data,
                            "name": f"harigami_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
                        }
                        status.update(label=f"生成完了: {len(outputs)} 件", state="complete")
                        st.rerun()
                    else:
                        status.update(label="生成に失敗しました", state="error")

                if errors:
                    with st.expander("エラー詳細"):
                        for e in errors:
                            st.error(e)
    else:
        st.info("条件を設定してイベントを取得してください。")


# ─────────────────────────────────────────────────────────
# 内部補助関数
# ─────────────────────────────────────────────────────────

def _generate_single_docx(ev, row, pm_idx):
    mgmt = row["管理番号"]
    desc = safe_get(ev, "description") or ""
    summary = safe_get(ev, "summary") or ""
    tags = extract_tags_from_description(desc)
    if "ASSETNUM" not in tags and mgmt:
        tags["ASSETNUM"] = mgmt

    work_type = extract_worktype(desc) or "default"
    template_file = DEFAULT_TEMPLATE_MAP.get(work_type, DEFAULT_TEMPLATE_MAP.get("default"))
    template_path = os.path.join("templates", template_file)
    start_dt = get_event_start_datetime(ev)
    when_str = start_dt.date().strftime("%Y-%m-%d") if start_dt else ""

    name_for_doc = (
        str(pm_idx.loc[mgmt].get("物件名") or summary or mgmt).strip()
        if pm_idx is not None and mgmt in pm_idx.index
        else str(summary or mgmt).strip()
    )

    replacements = build_replacements_from_event(ev, name_for_doc, tags)
    safe_title = f"{when_str}_{mgmt}_{name_for_doc}"
    return generate_docx_from_template_like(template_path, replacements, safe_title)


def _pack_zip(outputs):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used = set()
        for fname, content in outputs:
            name = base = fname or "harigami.docx"
            i = 1
            while name in used:
                root, ext = os.path.splitext(base)
                name = f"{root}_{i}{ext}"
                i += 1
            used.add(name)
            zf.writestr(name, content)
    return buf.getvalue()
