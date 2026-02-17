import re
import unicodedata
from datetime import datetime, date, timedelta, timezone
from typing import List
from io import BytesIO

import pandas as pd
import streamlit as st
from session_utils import get_user_setting, set_user_setting

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
    """“実質空”を厳密判定するためのクリーナー（WONUM/ASSETNUM共通）"""
    if val is None:
        return ""
    s = str(val)

    # 正規化（全角→半角など）
    s = unicodedata.normalize("NFKC", s)

    # NaN/None の文字列化を空扱い
    if s.lower() in ("nan", "none"):
        return ""

    # 不可視文字（Cfカテゴリ＝ゼロ幅スペース等）除去
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Cf")

    # BOM/特殊空白除去
    s = s.replace("\ufeff", "")     # BOM
    s = s.replace("\u00A0", " ")    # NBSP
    s = s.replace("\u3000", " ")    # 全角スペース

    # 空白・改行・タブ除去
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


# ==============================
# タブ5本体
# ==============================
def render_tab5_export(editable_calendar_options, service, fetch_all_events):
    """タブ5: カレンダーイベントをExcel/CSVへ出力（WONUM & ASSETNUM 抽出不可完全除外版）"""

    st.subheader("カレンダーイベントをExcelに出力")

    def safe_filename(name: str) -> str:
        name = unicodedata.normalize("NFKC", name)
        name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', "", name)
        name = re.sub(r'[@.]', "_", name)   # @ と . を _ に変換
        name = name.strip("_ ")             # 先頭・末尾の _ や空白を除去
        return name or "output"

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。")
        return

    # 出力対象カレンダー（サイドバーの基準カレンダーを初期値に。タブ側は永続化しない）
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
    export_start_date = st.date_input("出力開始日", value=today_date_export - timedelta(days=30))
    export_end_date = st.date_input("出力終了日", value=today_date_export)
    export_format = st.radio("出力形式を選択", ("CSV", "Excel"), index=0)

    if export_start_date > export_end_date:
        st.error("出力開始日は終了日より前に設定してください。")
        return

    if st.button("指定期間のイベントを読み込む"):
        with st.spinner("イベントを読み込み中..."):
            try:
                time_min_utc, time_max_utc = to_utc_range(export_start_date, export_end_date)
                events_to_export = fetch_all_events(service, calendar_id_export, time_min_utc, time_max_utc)

                if not events_to_export:
                    st.info("指定期間内にイベントは見つかりませんでした。")
                    return

                extracted_data: List[dict] = []
                excluded_count = 0

                for event in events_to_export:
                    description_text = event.get("description", "") or ""

                    # 抽出処理（+クリーニング）
                    wonum_raw = extract_wonum(description_text)
                    assetnum_raw = extract_assetnum(description_text)
                    wonum = _clean(wonum_raw)
                    assetnum = _clean(assetnum_raw)

                    # ★ 抽出不可（空扱い）は即除外（WONUM or ASSETNUMが空なら除外）
                    if not wonum or not assetnum:
                        excluded_count += 1
                        continue

                    # 追加情報（任意）
                    worktype_match = WORKTYPE_PATTERN.search(description_text or "")
                    title_match = TITLE_PATTERN.search(description_text or "")
                    worktype = (worktype_match.group(1).strip() if worktype_match else "") or ""
                    description_val = title_match.group(1).strip() if title_match else ""

                    # 日付
                    start_time = event["start"].get("dateTime") or event["start"].get("date") or ""
                    end_time = event["end"].get("dateTime") or event["end"].get("date") or ""

                    def to_jst_iso(s: str) -> str:
                        try:
                            if "T" in s and ("+" in s or s.endswith("Z")):
                                dt = datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)
                                return dt.isoformat(timespec="seconds")
                        except Exception:
                            pass
                        return s

                    schedstart = to_jst_iso(start_time)
                    schedfinish = to_jst_iso(end_time)

                    extracted_data.append({
                        "WONUM": wonum,
                        "ASSETNUM": assetnum,
                        "DESCRIPTION": description_val,
                        "WORKTYPE": worktype,
                        "SCHEDSTART": schedstart,
                        "SCHEDFINISH": schedfinish,
                        "LEAD": "",
                        "JESSCHEDFIXED": "",
                        "SITEID": "JES",
                    })

                df_filtered = pd.DataFrame(extracted_data)

                # 表示（番号ありのみ）
                st.dataframe(df_filtered)

                if excluded_count > 0:
                    st.warning(f"⚠️ 作業指示書番号/管理番号なし（抽出不可） {excluded_count} 件を除外しました。")

                output_df = df_filtered.copy()

                # ファイル名
                start_str = export_start_date.strftime("%Y%m%d")
                end_str = export_end_date.strftime("%m%d")
                safe_cal_name = safe_filename(selected_calendar_name_export)
                file_base_name = f"{safe_cal_name}_{start_str}_{end_str}"

                # ダウンロード
                if export_format == "CSV":
                    csv_buffer = output_df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        label="✅ CSVファイルとしてダウンロード",
                        data=csv_buffer,
                        file_name=f"{file_base_name}.csv",
                        mime="text/csv",
                    )
                else:
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        output_df.to_excel(writer, index=False, sheet_name="カレンダーイベント")
                    buffer.seek(0)
                    st.download_button(
                        label="✅ Excelファイルとしてダウンロード",
                        data=buffer,
                        file_name=f"{file_base_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                st.success(f"{len(output_df)} 件のイベントを読み込みました。（※番号なし抽出不可は除外済）")

            except Exception as e:
                st.error(f"イベントの読み込み中にエラーが発生しました: {e}")
