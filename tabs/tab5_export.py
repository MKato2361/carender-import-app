import re
import logging
import unicodedata
import calendar as cal_mod
from datetime import datetime, date, timedelta, timezone
from typing import List, Callable
from io import BytesIO

import pandas as pd
import streamlit as st

# 認証・カレンダー関連のユーティリティ
from calendar_utils import fetch_all_events

# ==============================
# 正規表現（全角/半角/表記ゆれ対応）
# ==============================
WONUM_PATTERN = re.compile(
    r"[［\[]?\s*作業指示書(?:番号)?[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?",
    flags=re.IGNORECASE
)

ASSETNUM_PATTERN = re.compile(
    r"[［\[]?\s*管理番号[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?",
    flags=re.IGNORECASE
)

WORKTYPE_PATTERN = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
TITLE_PATTERN = re.compile(r"\[タイトル[：:]\s*(.*?)\]")

JST = timezone(timedelta(hours=9))

DEFAULT_SITE_ID = "JES"


# ==============================
# 抽出 & クリーニング関数
# ==============================
def extract_wonum(description_text: str) -> str:
    """Descriptionから作業指示書番号を抽出（全角→半角、表記ゆれ吸収）"""
    if not description_text:
        return ""
    s = unicodedata.normalize("NFKC", description_text)
    m = WONUM_PATTERN.search(s)
    return (m.group(1).strip() if m else "")


def extract_assetnum(description_text: str) -> str:
    """Descriptionから管理番号を抽出（全角→半角、表記ゆれ吸収）"""
    if not description_text:
        return ""
    s = unicodedata.normalize("NFKC", description_text)
    m = ASSETNUM_PATTERN.search(s)
    return (m.group(1).strip() if m else "")


def _clean(val) -> str:
    """"実質空"を厳密判定するためのクリーナー（WONUM/ASSETNUM共通）"""
    if val is None:
        return ""
    s = str(val)
    s = unicodedata.normalize("NFKC", s)
    if s.lower() in ("nan", "none"):
        return ""
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Cf")
    s = s.replace("\ufeff", "").replace("\u00A0", " ").replace("\u3000", " ")
    return s.strip()


# ==============================
# 日付処理
# ==============================
def to_utc_range(d1: date, d2: date):
    start_dt_utc = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
    end_dt_utc = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
    return (
        start_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
        end_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
    )


def to_jst_iso(s: str) -> str:
    """UTC/オフセット付きISO文字列をJSTのISO文字列に変換する。"""
    try:
        if "T" in s and ("+" in s or s.endswith("Z")):
            dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
            return dt.isoformat(timespec="seconds")
    except ValueError as e:
        logging.warning(f"日時パース失敗: {s!r} → {e}")
    return s


def safe_filename(name: str) -> str:
    """ファイル名に使用できない文字を除去・変換する。"""
    name = unicodedata.normalize("NFKC", name)
    name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', "", name)
    name = re.sub(r'[@.]', "_", name)
    name = name.strip("_ ")
    return name or "output"


# ==============================
# ロジック分離
# ==============================
def _fetch_and_extract(
    service,
    calendar_id: str,
    start_date: date,
    end_date: date,
) -> tuple[pd.DataFrame, int]:
    """イベント取得・抽出・除外を担当"""
    time_min_utc, time_max_utc = to_utc_range(start_date, end_date)
    events = fetch_all_events(service, calendar_id, time_min_utc, time_max_utc)

    extracted_data: List[dict] = []
    excluded_count = 0

    for event in events:
        description_text = event.get("description", "") or ""
        normalized_desc = unicodedata.normalize("NFKC", description_text)

        wonum = _clean(extract_wonum(description_text))
        assetnum = _clean(extract_assetnum(description_text))

        if not wonum or not assetnum:
            excluded_count += 1
            continue

        worktype_match = WORKTYPE_PATTERN.search(normalized_desc)
        title_match = TITLE_PATTERN.search(normalized_desc)
        worktype = (worktype_match.group(1).strip() if worktype_match else "") or ""
        description_val = title_match.group(1).strip() if title_match else ""

        start_time = event["start"].get("dateTime") or event["start"].get("date") or ""
        end_time = event["end"].get("dateTime") or event["end"].get("date") or ""

        extracted_data.append({
            "WONUM": wonum,
            "ASSETNUM": assetnum,
            "DESCRIPTION": description_val,
            "WORKTYPE": worktype,
            "SCHEDSTART": to_jst_iso(start_time),
            "SCHEDFINISH": to_jst_iso(end_time),
            "LEAD": "",
            "JESSCHEDFIXED": "",
            "SITEID": DEFAULT_SITE_ID,
        })

    return pd.DataFrame(extracted_data), excluded_count


def _build_download_section(df: pd.DataFrame, file_base_name: str, export_format: str) -> None:
    """ダウンロードボタン描画"""
    if export_format == "CSV":
        csv_buffer = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="✅ CSVファイルとしてダウンロード",
            data=csv_buffer,
            file_name=f"{file_base_name}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    else:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="カレンダーイベント")
        buffer.seek(0)
        st.download_button(
            label="✅ Excelファイルとしてダウンロード",
            data=buffer,
            file_name=f"{file_base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ==============================
# メインタブ描画 (AuthManager対応版)
# ==============================
def render_tab5_export(manager) -> None:
    """タブ5: カレンダーイベントをExcel/CSVへ出力"""
    # manager から必要なサービスとオプションを取得
    service = manager.calendar_service
    editable_calendar_options = manager.editable_calendar_options

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。Google認証を確認してください。")
        return

    calendar_options = list(editable_calendar_options.keys())
    base_calendar = (
        st.session_state.get("base_calendar_name")
        or st.session_state.get("selected_calendar_name")
        or calendar_options[0]
    )
    if base_calendar not in calendar_options:
        base_calendar = calendar_options[0]

    select_key = "export_calendar_select"
    share_calendar = st.session_state.get("share_calendar_selection_across_tabs", True)

    if share_calendar:
        st.session_state[select_key] = base_calendar
    elif (select_key not in st.session_state) or (st.session_state.get(select_key) not in calendar_options):
        st.session_state[select_key] = base_calendar

    current_cal = st.session_state.get(select_key, base_calendar)
    st.markdown(f"""
<div style="border:2px solid #1E88E5;border-radius:10px;padding:14px 18px;margin-bottom:8px;background:var(--color-background-info);">
  <div style="font-size:12px;font-weight:600;color:var(--color-text-info);margin-bottom:4px;">📅 出力対象カレンダー（必ず確認）</div>
  <div style="font-size:20px;font-weight:700;color:var(--color-text-info);">{current_cal}</div>
</div>
""", unsafe_allow_html=True)

    if share_calendar:
        st.caption("サイドバーの「基準カレンダー」と連動しています。")
    else:
        with st.expander("カレンダーを変更する"):
            st.selectbox("カレンダーを選択", calendar_options, key=select_key, label_visibility="collapsed")

    selected_calendar_name_export = st.session_state.get(select_key, base_calendar)
    calendar_id_export = editable_calendar_options[selected_calendar_name_export]

    export_format = st.radio("出力形式", ("CSV", "Excel"), index=0, horizontal=True)

    st.divider()
    st.markdown("##### 出力期間")
    today_date = date.today()

    if "export_start_date" not in st.session_state:
        st.session_state["export_start_date"] = today_date - timedelta(days=30)
    if "export_end_date" not in st.session_state:
        st.session_state["export_end_date"] = today_date

    def _on_start_date_change():
        new_start = st.session_state["export_start_date"]
        month = new_start.month + 1
        year = new_start.year + (1 if month > 12 else 0)
        month = month if month <= 12 else 1
        last_day = cal_mod.monthrange(year, month)[1]
        auto_end = new_start.replace(year=year, month=month, day=min(new_start.day, last_day))
        st.session_state["export_end_date"] = auto_end

    col1, col2 = st.columns(2)
    with col1:
        st.date_input(
            "開始日",
            key="export_start_date",
            on_change=_on_start_date_change,
            help="開始日を変えると終了日が自動で1ヶ月後にセットされます。"
        )
    with col2:
        st.date_input(
            "終了日",
            key="export_end_date",
            min_value=st.session_state["export_start_date"],
        )

    export_start_date: date = st.session_state["export_start_date"]
    export_end_date: date = st.session_state["export_end_date"]

    if export_start_date > export_end_date:
        st.error("⚠️ 終了日は開始日以降に設定してください。")
        return

    st.divider()

    # 実行ボタン
    if st.button(f"🚀 {export_format} データを生成する", type="primary", use_container_width=True):
        progress = st.progress(0, text="📡 カレンダーからデータを取得中...")
        try:
            df_filtered, excluded_count = _fetch_and_extract(
                service,
                calendar_id_export,
                export_start_date,
                export_end_date
            )

            if df_filtered.empty:
                progress.empty()
                st.info("条件に一致するイベントが見つかりませんでした。")
                return

            progress.progress(80, text="📄 ダウンロード準備中...")
            
            start_str = export_start_date.strftime("%Y%m%d")
            end_str = export_end_date.strftime("%m%d")
            file_base_name = f"{safe_filename(selected_calendar_name_export)}_{start_str}_{end_str}"

            st.success(f"✅ {len(df_filtered)} 件のデータを抽出しました。")
            if excluded_count > 0:
                st.caption(f"※ 作業指示書番号がないイベント {excluded_count} 件を除外しました。")
            
            _build_download_section(df_filtered, file_base_name, export_format)
            
            with st.expander("🔍 抽出データプレビュー", expanded=True):
                st.dataframe(df_filtered, use_container_width=True)
            
            progress.progress(100, text="✅ 完了")

        except Exception as e:
            progress.empty()
            st.error("エクスポート中にエラーが発生しました。しばらく待ってから再試行してください。")
