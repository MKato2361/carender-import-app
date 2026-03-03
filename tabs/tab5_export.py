import re
import logging
import unicodedata
from datetime import datetime, date, timedelta, timezone
from typing import List, Callable
from io import BytesIO

import pandas as pd
import streamlit as st

# 🔴 修正1: 未使用インポートを削除（get_user_setting, set_user_setting, _get_current_user_key）
# from session_utils import get_user_setting, set_user_setting

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

# 🟢 定数化（レビュー軽微項目だが修正に含める）
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

    s = s.replace("\ufeff", "")
    s = s.replace("\u00A0", " ")
    s = s.replace("\u3000", " ")

    s = s.strip()

    return s


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


# 🔴 修正1: ループ内ネスト定義 → モジュールレベルに移動
def to_jst_iso(s: str) -> str:
    """UTC/オフセット付きISO文字列をJSTのISO文字列に変換する。パース失敗時は元の文字列を返す。"""
    try:
        if "T" in s and ("+" in s or s.endswith("Z")):
            dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
            return dt.isoformat(timespec="seconds")
    except ValueError as e:
        # 🟡 修正3: 例外を握りつぶさずログ出力
        logging.warning(f"日時パース失敗: {s!r} → {e}")
    return s


# 🟡 修正4: safe_filename をモジュールレベルに移動（テスト可能・再利用可能）
def safe_filename(name: str) -> str:
    """ファイル名に使用できない文字を除去・変換する。"""
    name = unicodedata.normalize("NFKC", name)
    name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', "", name)
    name = re.sub(r'[@.]', "_", name)
    name = name.strip("_ ")
    return name or "output"


# ==============================
# 🟡 修正4: render_tab5_export の責務を分離
# ==============================
def _fetch_and_extract(
    service,
    calendar_id: str,
    start_date: date,
    end_date: date,
    fetch_all_events: Callable,
) -> tuple[pd.DataFrame, int]:
    """
    イベント取得・抽出・除外のみを担当する。
    Returns:
        df: 抽出済みDataFrame
        excluded_count: 除外件数
    """
    time_min_utc, time_max_utc = to_utc_range(start_date, end_date)
    events = fetch_all_events(service, calendar_id, time_min_utc, time_max_utc)

    extracted_data: List[dict] = []
    excluded_count = 0

    for event in events:
        description_text = event.get("description", "") or ""

        # 🟡 修正5: WORKTYPE/TITLE も NFKC 正規化後の文字列に対してマッチ（不整合を解消）
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
    """ダウンロードボタン描画のみを担当する。"""
    if export_format == "CSV":
        csv_buffer = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="✅ CSVファイルとしてダウンロード",
            data=csv_buffer,
            file_name=f"{file_base_name}.csv",
            mime="text/csv",
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
        )


# ==============================
# タブ5本体
# ==============================
def render_tab5_export(
    editable_calendar_options: dict,
    service,
    fetch_all_events: Callable,  # 🟢 型ヒントを追加
) -> None:
    """タブ5: カレンダーイベントをExcel/CSVへ出力"""

    st.subheader("カレンダーイベントをExcelに出力")

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。")
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
    if (select_key not in st.session_state) or (st.session_state.get(select_key) not in calendar_options):
        st.session_state[select_key] = base_calendar

    selected_calendar_name_export = st.selectbox(
        "出力対象カレンダーを選択",
        calendar_options,
        key=select_key,
    )

    calendar_id_export = editable_calendar_options[selected_calendar_name_export]

    st.subheader("🗓️ 出力期間の選択")
    today_date_export = date.today()

    # session_state の初期化（初回のみ）
    if "export_start_date" not in st.session_state:
        st.session_state["export_start_date"] = today_date_export - timedelta(days=30)
    if "export_end_date" not in st.session_state:
        st.session_state["export_end_date"] = today_date_export

    def _on_start_date_change():
        """開始日が変わったら終了日を自動で1ヶ月後にセット"""
        import calendar as cal_mod
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
            "📅 開始日",
            key="export_start_date",
            min_value=today_date_export - timedelta(days=365 * 3),
            on_change=_on_start_date_change,
        )
    with col2:
        st.date_input(
            "📅 終了日（開始日選択で自動セット・調整可）",
            key="export_end_date",
            min_value=st.session_state["export_start_date"],
            max_value=today_date_export + timedelta(days=365),
        )

    export_start_date: date = st.session_state["export_start_date"]
    export_end_date: date = st.session_state["export_end_date"]

    if export_start_date > export_end_date:
        st.error("⚠️ 終了日は開始日以降に設定してください。")
        return

    period_days = (export_end_date - export_start_date).days + 1
    st.caption(f"📆 選択期間: {export_start_date} 〜 {export_end_date}（{period_days} 日間）")

    # UI修正2: 出力形式を読み込みボタンの前に配置し、1アクションで完結させる
    export_format = st.radio("出力形式を選択", ("CSV", "Excel"), index=0, horizontal=True)

    # UI修正2: ボタンラベルを「読み込み → ダウンロード」の1アクションであることを明示
    button_label = f"📥 イベントを読み込んで {export_format} でダウンロード"

    if st.button(button_label, type="primary"):

        # UI修正3: 進捗バーで処理の進捗を可視化
        progress = st.progress(0, text="📡 カレンダーからイベントを取得中...")

        try:
            time_min_utc, time_max_utc = to_utc_range(export_start_date, export_end_date)
            events_raw = fetch_all_events(service, calendar_id_export, time_min_utc, time_max_utc)

            if not events_raw:
                progress.empty()
                st.info("指定期間内にイベントは見つかりませんでした。")
                return

            progress.progress(40, text=f"🔍 {len(events_raw)} 件のイベントを解析中...")

            df_filtered, excluded_count = _fetch_and_extract(
                service,
                calendar_id_export,
                export_start_date,
                export_end_date,
                fetch_all_events,
            )

            progress.progress(80, text="📄 ファイルを生成中...")

            start_str = export_start_date.strftime("%Y%m%d")
            end_str = export_end_date.strftime("%m%d")
            file_base_name = f"{safe_filename(selected_calendar_name_export)}_{start_str}_{end_str}"

            _build_download_section(df_filtered, file_base_name, export_format)

            progress.progress(100, text="✅ 完了")

            if df_filtered.empty:
                st.info("番号が抽出できるイベントが見つかりませんでした。")
                return

            # UI修正（提案6）: 成功メッセージ → 警告 → テーブルの順に整理
            st.success(f"✅ {len(df_filtered)} 件を出力しました。（期間: {export_start_date} 〜 {export_end_date}）")

            if excluded_count > 0:
                st.warning(f"⚠️ 作業指示書番号/管理番号なし（抽出不可） {excluded_count} 件を除外しました。")

            st.dataframe(df_filtered)

        except Exception as e:
            progress.empty()
            st.error(f"イベントの読み込み中にエラーが発生しました: {e}")
