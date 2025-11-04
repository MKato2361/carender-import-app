import re
import unicodedata
from datetime import datetime, date, timedelta, timezone
from typing import List
from io import BytesIO

import pandas as pd
import streamlit as st

# 旧のRE_WONUM等は残してOKですが、wonum抽出は下の関数で統一します
# 汎用：全角/半角カッコ、番号表記ゆれ、改行混在に対応
WONUM_PATTERN = re.compile(
    r"[［\[]?\s*作業指示書(?:番号)?[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?",
    flags=re.IGNORECASE
)

JST = timezone(timedelta(hours=9))

def to_utc_range(d1: date, d2: date):
    start_dt_utc = datetime.combine(d1, datetime.min.time(), tzinfo=JST).astimezone(timezone.utc)
    end_dt_utc = datetime.combine(d2, datetime.max.time(), tzinfo=JST).astimezone(timezone.utc)
    return (
        start_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
        end_dt_utc.isoformat(timespec="microseconds").replace("+00:00", "Z"),
    )

# ★ 新規：頑強な作業指示書番号抽出
def extract_wonum(description_text: str) -> str:
    if not description_text:
        return ""
    s = unicodedata.normalize("NFKC", description_text)  # 全角→半角・表記統一
    m = WONUM_PATTERN.search(s)
    return (m.group(1).strip() if m else "")

# 参考：他タグは従来ロジックのままでもOK
RE_ASSETNUM = re.compile(r"\[管理番号[：:]\s*(.*?)\]")
RE_WORKTYPE = re.compile(r"\[作業タイプ[：:]\s*(.*?)\]")
RE_TITLE = re.compile(r"\[タイトル[：:]\s*(.*?)\]")

def render_tab5_export(editable_calendar_options, service, fetch_all_events):
    """タブ5: カレンダーイベントをExcel/CSVへ出力（番号なし除外＋除外件数表示：堅牢抽出版）"""

    st.subheader("カレンダーイベントをExcelに出力")

    def safe_filename(name: str) -> str:
        name = unicodedata.normalize("NFKC", name)
        name = re.sub(r'[\/\\\:\*\?\"\<\>\|]', "", name)
        name = name.strip(" .")
        return name or "output"

    if not editable_calendar_options:
        st.error("利用可能なカレンダーが見つかりません。")
        return

    selected_calendar_name_export = st.selectbox(
        "出力対象カレンダーを選択",
        list(editable_calendar_options.keys()),
        key="export_calendar_select",
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
                total_count = 0
                excluded_count = 0

                for event in events_to_export:
                    total_count += 1
                    description_text = event.get("description", "") or ""

                    # ★ ここで堅牢にWONUMを抽出
                    wonum = extract_wonum(description_text)

                    # 他フィールド（従来通りでOK）
                    assetnum_match = RE_ASSETNUM.search(description_text or "")
                    worktype_match = RE_WORKTYPE.search(description_text or "")
                    title_match = RE_TITLE.search(description_text or "")

                    assetnum = (assetnum_match.group(1).strip() if assetnum_match else "") or ""
                    worktype = (worktype_match.group(1).strip() if worktype_match else "") or ""
                    description_val = title_match.group(1).strip() if title_match else ""

                    # ★ 作業指示書番号が無いイベントは除外（カウント）
                    if not wonum:
                        excluded_count += 1
                        continue

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
                        "DESCRIPTION": description_val,
                        "ASSETNUM": assetnum,
                        "WORKTYPE": worktype,
                        "SCHEDSTART": schedstart,
                        "SCHEDFINISH": schedfinish,
                        "LEAD": "",
                        "JESSCHEDFIXED": "",
                        "SITEID": "JES",
                    })

                output_df = pd.DataFrame(extracted_data)
                st.dataframe(output_df)

                # 除外件数表示
                kept_count = len(output_df)
                if excluded_count > 0:
                    st.warning(f"⚠️ 作業指示書番号なしのイベント {excluded_count} 件を除外しました。")

                start_str = export_start_date.strftime("%Y%m%d")
                end_str = export_end_date.strftime("%m%d")
                safe_cal_name = safe_filename(selected_calendar_name_export)
                file_base_name = f"{safe_cal_name}_{start_str}_{end_str}"

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

                st.success(f"{kept_count} 件のイベントを読み込みました。（※番号なし除外済）")

            except Exception as e:
                st.error(f"イベントの読み込み中にエラーが発生しました: {e}")