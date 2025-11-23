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

# 物件マスタ関連（tab6 と同じロジックを再利用）
from tabs.tab6_property_master import (
    MASTER_COLUMNS,
    BASIC_COLUMNS,
    load_sheet_as_df,
    _normalize_df,
)
from utils.helpers import safe_get

# 貼り紙テンプレ生成ユーティリティ
try:
    from utils.harigami_generator import (
        DEFAULT_TEMPLATE_MAP,
        extract_tags_from_description,
        build_replacements_from_event,
        generate_docx_from_template_like,
    )

    HARIGAMI_AVAILABLE = True
    HARIGAMI_IMPORT_ERROR: Optional[Exception] = None
except Exception as e:  # ImportError や python-docx 未インストールもまとめて拾う
    HARIGAMI_AVAILABLE = False
    HARIGAMI_IMPORT_ERROR = e

JST = timezone(timedelta(hours=9))

# [管理番号: HK5-123] 対応
ASSETNUM_PATTERN = re.compile(
    r"[［\[]?\s*管理番号[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?"
)

# [作業タイプ: 点検] など
WORKTYPE_PATTERN = re.compile(
    r"\[作業タイプ[：:]\s*(.*?)\]"
)


def extract_assetnum(text: str) -> str:
    """説明 or タイトル から管理番号を抽出"""
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    m = ASSETNUM_PATTERN.search(s)
    if not m:
        return ""
    return m.group(1).strip()


def extract_worktype(text: str) -> str:
    """説明から作業タイプを抽出（なければ空文字）"""
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    m = WORKTYPE_PATTERN.search(s)
    if not m:
        return ""
    return (m.group(1) or "").strip()


def to_utc_range_from_dates(d1: date, d2: date) -> tuple[str, str]:
    """JST の日付範囲 → Calendar API 用 UTC ISO 文字列"""
    start_dt_utc = datetime.combine(d1, time.min, tzinfo=JST).astimezone(timezone.utc)
    end_dt_utc = datetime.combine(d2, time.max, tzinfo=JST).astimezone(timezone.utc)
    return start_dt_utc.isoformat(), end_dt_utc.isoformat()


def get_event_start_datetime(event: Dict[str, Any]) -> Optional[datetime]:
    """Google カレンダーイベントから開始日時（JST）を取得"""
    start = event.get("start", {})
    # 時間付き予定
    if "dateTime" in start:
        try:
            dt = pd.to_datetime(start["dateTime"])
            if dt.tzinfo is None:
                dt = dt.tz_localize(timezone.utc)
            dt = dt.astimezone(JST)
            return dt.to_pydatetime()
        except Exception:
            return None
    # 終日予定
    if "date" in start:
        try:
            d = date.fromisoformat(start["date"])
            return datetime.combine(d, time.min, tzinfo=JST)
        except Exception:
            return None
    return None


def fetch_events_in_range(
    service: Any,
    calendar_id: str,
    start_date: date,
    end_date: date,
) -> List[Dict[str, Any]]:
    """指定期間のイベントをカレンダーから取得"""
    if not service:
        return []

    time_min, time_max = to_utc_range_from_dates(start_date, end_date)

    events: List[Dict[str, Any]] = []
    page_token: Optional[str] = None

    while True:
        resp = (
            service.events()
            .list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                maxResults=2500,
                singleEvents=True,
                orderBy="startTime",
                pageToken=page_token,
            )
            .execute()
        )
        items = resp.get("items", [])
        events.extend(items)
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return events


def get_property_master_spreadsheet_id(current_user_email: Optional[str]) -> str:
    """Firestore user_settings から物件マスタ用 Spreadsheet ID を取得"""
    if not current_user_email:
        return ""
    try:
        db = firestore.client()
        doc = db.collection("user_settings").document(current_user_email).get()
        if not doc.exists:
            return ""
        data = doc.to_dict() or {}
        return data.get("property_master_spreadsheet_id") or ""
    except Exception as e:
        st.warning(f"物件マスタ用スプレッドシートIDの取得に失敗しました: {e}")
        return ""


def load_property_master_view(
    sheets_service: Any,
    spreadsheet_id: str,
    basic_sheet_title: str = "物件基本情報",
    master_sheet_title: str = "物件マスタ",
) -> pd.DataFrame:
    """
    物件基本情報 ＋ 物件マスタ を管理番号でマージしたビュー DF を返す
    """
    if not sheets_service or not spreadsheet_id:
        return pd.DataFrame()

    try:
        basic_df = load_sheet_as_df(
            sheets_service,
            spreadsheet_id,
            basic_sheet_title,
            BASIC_COLUMNS,
        )
        master_df = load_sheet_as_df(
            sheets_service,
            spreadsheet_id,
            master_sheet_title,
            MASTER_COLUMNS,
        )
    except Exception as e:
        st.error(f"物件マスタの読み込みに失敗しました: {e}")
        return pd.DataFrame()

    basic_df = _normalize_df(basic_df, BASIC_COLUMNS)
    master_df = _normalize_df(master_df, MASTER_COLUMNS)

    if master_df.empty:
        merged = basic_df.copy()
        for col in MASTER_COLUMNS:
            if col not in merged.columns:
                merged[col] = ""
        return merged

    merged = master_df.merge(
        basic_df[["管理番号", "物件名", "住所", "窓口会社"]],
        on="管理番号",
        how="left",
    )

    return merged


def _display_value(val: Any) -> str:
    """nan / None / 空文字 を '-' に揃える"""
    if val is None:
        return "-"
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return "-"
    return s


def _build_safe_title(when: str, mgmt: str, name: str) -> str:
    """ファイル名用のタイトル生成"""
    parts = []
    if when:
        parts.append(when)
    if mgmt:
        parts.append(mgmt)
    if name:
        parts.append(name)
    if not parts:
        return "harigami"
    return "_".join(parts)


# ============================================================
# メイン：貼り紙自動作成タブ
# ============================================================

def render_tab8_notice_fax(
    service: Any,
    editable_calendar_options: Dict[str, str],
    sheets_service: Any = None,
    current_user_email: Optional[str] = None,
    **kwargs,
) -> None:
    """
    貼り紙自動生成タブ

    - 指定期間のカレンダーイベントを取得
    - 物件マスタと管理番号で突合
    - 「貼り紙テンプレ種別」が「自社」のものだけを対象
    - 作業タイプ → DEFAULT_TEMPLATE_MAP でテンプレファイルを選択
    - Word 貼り紙を一括生成し、ZIP でダウンロード
    """
    st.subheader("📄 貼り紙自動作成")

    # harigami モジュールが使えない場合
    if not HARIGAMI_AVAILABLE:
        st.error(
            "貼り紙生成モジュール（utils.harigami_generator）が読み込めませんでした。\n\n"
            "・requirements.txt に `python-docx` が追加されているか\n"
            "・`utils/harigami_generator.py` が正しく配置されているか\n"
            "を確認してください。\n\n"
            f"詳細エラー: {HARIGAMI_IMPORT_ERROR}"
        )
        return

    if not service or not editable_calendar_options:
        st.warning("カレンダーサービスが初期化されていません。タブ1〜2で認証を完了してください。")
        return

    # 物件マスタ Spreadsheet ID
    spreadsheet_id = get_property_master_spreadsheet_id(current_user_email)
    if not spreadsheet_id:
        st.error(
            "物件マスタ用スプレッドシートIDが設定されていません。\n\n"
            "タブ『物件マスタ管理』で物件マスタを作成・保存すると、この機能を利用できます。"
        )
        return

    st.markdown(
        f"物件マスタスプレッドシート: "
        f"[こちらを開く](https://docs.google.com/spreadsheets/d/{spreadsheet_id})"
    )

    # 物件マスタビュー
    pm_view_df = load_property_master_view(
        sheets_service,
        spreadsheet_id,
        basic_sheet_title="物件基本情報",
        master_sheet_title="物件マスタ",
    )
    if pm_view_df is None or pm_view_df.empty:
        st.error("物件マスタ（＋基本情報）が空です。先にタブ『物件マスタ管理』で登録してください。")
        return

    # 「自社」フラグ件数の表示（単なる参考）
    col_flag = "貼り紙テンプレ種別"
    if col_flag in pm_view_df.columns:
        flag_series = pm_view_df[col_flag].fillna("").astype(str).str.strip()
        cnt_jisha = (flag_series == "自社").sum()
        st.caption(f"物件マスタ登録件数: {len(pm_view_df)} 件 / 貼り紙対象（自社）: {cnt_jisha} 件")
    else:
        st.caption(f"物件マスタ登録件数: {len(pm_view_df)} 件")

    # カレンダー選択
    cal_names = list(editable_calendar_options.keys())
    default_cal = cal_names[0] if cal_names else None
    calendar_name = st.selectbox(
        "対象カレンダー",
        cal_names,
        index=(cal_names.index(default_cal) if default_cal in cal_names else 0),
        key="notice_fax_calendar",
    )
    calendar_id = editable_calendar_options.get(calendar_name)

    # 期間指定
    today = date.today()
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input("対象イベントの開始日", value=today, key="notice_fax_start_date")
    with col_d2:
        end_date = st.date_input("対象イベントの終了日", value=today + timedelta(days=60), key="notice_fax_end_date")

    if start_date > end_date:
        st.error("開始日は終了日以前の日付を指定してください。")
        return

    # イベント取得ボタン
    fetch_btn = st.button("カレンダーから対象イベントを取得し、貼り紙候補を作成する", type="primary")

    if fetch_btn:
        with st.spinner("カレンダーイベントを取得中..."):
            events = fetch_events_in_range(service, calendar_id, start_date, end_date)

        st.write(f"取得したイベント件数: {len(events)} 件")

        # 物件マスタを管理番号で引けるように
        pm_idx = pm_view_df.set_index("管理番号")

        candidates: List[Dict[str, Any]] = []
        events_by_id: Dict[str, Dict[str, Any]] = {}

        for ev in events:
            desc = safe_get(ev, "description") or ""
            summary = safe_get(ev, "summary") or ""

            mgmt = extract_assetnum(desc) or extract_assetnum(summary)
            if not mgmt:
                continue
            mgmt_norm = mgmt.strip()

            if mgmt_norm not in pm_idx.index:
                continue

            pm_row = pm_idx.loc[mgmt_norm]

            # 物件マスタ側のフラグが「自社」のものだけを対象とする
            flag_val = str(pm_row.get("貼り紙テンプレ種別", "")).strip()
            if flag_val != "自社":
                continue

            start_dt = get_event_start_datetime(ev)
            start_date_val = start_dt.date() if start_dt else None
            date_str = start_date_val.strftime("%Y-%m-%d") if start_date_val else ""
            time_str = start_dt.strftime("%H:%M") if start_dt and "dateTime" in (ev.get("start") or {}) else ""

            work_type = extract_worktype(desc)
            if not work_type:
                work_type = "default"
            template_file = DEFAULT_TEMPLATE_MAP.get(work_type, DEFAULT_TEMPLATE_MAP.get("default"))

            candidates.append(
                {
                    "作成": True,
                    "event_id": ev.get("id") or "",
                    "管理番号": mgmt_norm,
                    "物件名": _display_value(pm_row.get("物件名", "")),
                    "予定日": date_str,
                    "予定時間": time_str,
                    "作業タイプ": work_type,
                    "テンプレファイル": template_file,
                    "貼り紙フラグ": flag_val,
                    "イベントタイトル": _display_value(summary),
                    "備考": _display_value(pm_row.get("備考", "")),
                }
            )

            ev_id = ev.get("id")
            if ev_id:
                events_by_id[ev_id] = ev

        if not candidates:
            st.info("貼り紙対象（物件マスタの『貼り紙テンプレ種別』が『自社』）となるイベントは見つかりませんでした。")
            st.session_state.pop("notice_fax_candidates_df", None)
            st.session_state.pop("notice_fax_events_by_id", None)
        else:
            cand_df = pd.DataFrame(candidates).fillna("")
            cand_df["作成"] = cand_df["作成"].astype(bool)

            st.session_state["notice_fax_candidates_df"] = cand_df
            st.session_state["notice_fax_events_by_id"] = events_by_id
            st.session_state["notice_fax_pm_view_df"] = pm_view_df

            st.success(f"貼り紙候補 {len(candidates)} 件を作成しました。この下で内容を確認し、作成対象を選択できます。")

    # ここからは既に候補がある場合の表示
    cand_df: Optional[pd.DataFrame] = st.session_state.get("notice_fax_candidates_df")
    events_by_id: Dict[str, Dict[str, Any]] = st.session_state.get("notice_fax_events_by_id", {})
    pm_view_df = st.session_state.get("notice_fax_pm_view_df", pm_view_df)

    if cand_df is None or cand_df.empty:
        return

    st.markdown("### 貼り紙作成候補一覧")
    st.caption("※『作成』チェックが ON の行だけが貼り紙として出力されます。必要に応じて OFF にしてください。")

    display_cols = [
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

    for col in display_cols:
        if col not in cand_df.columns:
            cand_df[col] = True if col == "作成" else ""

    disp_df = cand_df[display_cols].copy()
    for col in disp_df.columns:
        if col == "作成":
            disp_df[col] = disp_df[col].astype(bool)
        else:
            disp_df[col] = (
                disp_df[col]
                .fillna("")
                .astype(str)
                .replace({"nan": "-", "NaN": "-", "None": "-"})
            )

    edit_df = st.data_editor(
        disp_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "作成": st.column_config.CheckboxColumn("作成"),
            "管理番号": st.column_config.TextColumn("管理番号", disabled=True),
            "物件名": st.column_config.TextColumn("物件名", disabled=True),
            "予定日": st.column_config.TextColumn("予定日", disabled=True),
            "予定時間": st.column_config.TextColumn("予定時間", disabled=True),
            "作業タイプ": st.column_config.TextColumn("作業タイプ", disabled=True),
            "テンプレファイル": st.column_config.TextColumn("テンプレファイル", disabled=True),
            "イベントタイトル": st.column_config.TextColumn("イベントタイトル", disabled=True),
            "備考": st.column_config.TextColumn("備考", disabled=True),
        },
        key="notice_fax_editor",
    )

    cand_df["作成"] = edit_df["作成"].values
    st.session_state["notice_fax_candidates_df"] = cand_df

    generate_btn = st.button("✅ 選択された行の貼り紙を一括生成し、ZIPでダウンロードする", type="primary")

    if not generate_btn:
        return

    target_df = cand_df[cand_df["作成"] == True].copy()
    if target_df.empty:
        st.warning("『作成』にチェックが入っている行がありません。")
        return

    if not events_by_id:
        st.error("内部イベント情報が見つかりません。再度『貼り紙候補を作成』ボタンからやり直してください。")
        return

    pm_idx = pm_view_df.set_index("管理番号")

    outputs: List[tuple[str, bytes]] = []
    errors: List[str] = []

    with st.spinner("貼り紙（Word）を生成中..."):
        for _, row in target_df.iterrows():
            event_id = row.get("event_id")
            mgmt = row.get("管理番号")
            if not event_id or not mgmt:
                continue

            ev = events_by_id.get(event_id)
            if ev is None:
                continue

            if mgmt not in pm_idx.index:
                continue
            pm_row = pm_idx.loc[mgmt]

            desc = safe_get(ev, "description") or ""
            summary = safe_get(ev, "summary") or ""

            tags = extract_tags_from_description(desc or "")
            # description に管理番号タグが入っていない場合もあるので補完
            if "ASSETNUM" not in tags and mgmt:
                tags["ASSETNUM"] = mgmt

            work_type = extract_worktype(desc)
            if not work_type:
                work_type = "default"
            template_file = DEFAULT_TEMPLATE_MAP.get(work_type, DEFAULT_TEMPLATE_MAP.get("default"))
            template_path = os.path.join("templates", template_file)

            start_dt = get_event_start_datetime(ev)
            when_str = ""
            if start_dt:
                d = start_dt.date()
                when_str = d.strftime("%Y-%m-%d")

            # NAME は物件マスタの物件名を優先
            name_for_doc = str(pm_row.get("物件名") or summary or mgmt or "").strip()
            replacements = build_replacements_from_event(ev, name_for_doc, tags)

            safe_title = _build_safe_title(when_str, mgmt, name_for_doc)

            try:
                out_name, content = generate_docx_from_template_like(
                    template_path,
                    replacements,
                    safe_title,
                )
                outputs.append((out_name, content))
            except Exception as e:
                errors.append(f"{mgmt}: {e}")

    if not outputs:
        st.error("貼り紙の生成に失敗しました。テンプレートファイルの配置や権限を確認してください。")
        if errors:
            st.warning(f"エラー例: {errors[0]}")
        return

    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = set()
        for fname, content in outputs:
            base_name = fname or "harigami.docx"
            name = base_name
            idx = 1
            while name in used_names:
                root, ext = os.path.splitext(base_name)
                name = f"{root}_{idx}{ext}"
                idx += 1
            used_names.add(name)
            zf.writestr(name, content)

    mem_zip.seek(0)
    zip_filename = f"harigami_{datetime.now().strftime('%Y%m%d')}.zip"

    st.download_button(
        "📦 貼り紙ZIPをダウンロード",
        data=mem_zip.getvalue(),
        file_name=zip_filename,
        mime="application/zip",
    )

    if errors:
        st.warning(f"一部の貼り紙生成でエラーが発生しました（{len(errors)} 件）。先頭のエラー: {errors[0]}")
