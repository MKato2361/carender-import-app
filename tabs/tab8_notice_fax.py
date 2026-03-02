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

# ウィザードステップ定義
STEPS = ["① 設定", "② イベント取得", "③ 候補確認", "④ 生成・DL"]

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

def _get_current_user_key(fallback: str = "") -> str:
    for key in ("user_id", "firebase_uid", "localId", "uid", "user_email"):
        val = st.session_state.get(key)
        if val:
            return val
    return fallback or ""


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


def fetch_events_in_range(
    service: Any,
    calendar_id: str,
    start_date: date,
    end_date: date,
) -> List[Dict[str, Any]]:
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
        if not doc.exists:
            return ""
        return (doc.to_dict() or {}).get("property_master_spreadsheet_id") or ""
    except Exception as e:
        st.warning(f"物件マスタ用スプレッドシートIDの取得に失敗しました: {e}")
        return ""


def load_property_master_view(
    sheets_service: Any,
    spreadsheet_id: str,
    basic_sheet_title: str = "物件基本情報",
    master_sheet_title: str = "物件マスタ",
) -> pd.DataFrame:
    if not sheets_service or not spreadsheet_id:
        return pd.DataFrame()
    try:
        basic_df = load_sheet_as_df(sheets_service, spreadsheet_id, basic_sheet_title, BASIC_COLUMNS)
        master_df = load_sheet_as_df(sheets_service, spreadsheet_id, master_sheet_title, MASTER_COLUMNS)
    except Exception as e:
        raise RuntimeError(f"物件マスタの読み込みに失敗しました: {e}") from e

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
        how="left",
    )


def _display_value(val: Any) -> str:
    if val is None:
        return "-"
    s = str(val).strip()
    return "-" if not s or s.lower() in ("nan", "none") else s


def _build_safe_title(when: str, mgmt: str, name: str) -> str:
    parts = [p for p in (when, mgmt, name) if p]
    return "_".join(parts) if parts else "harigami"



def _generate_single_docx(
    ev: Dict[str, Any],
    row: pd.Series,
    pm_idx: Optional[pd.DataFrame],
) -> tuple[str, bytes]:
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

    if pm_idx is not None and mgmt in pm_idx.index:
        name_for_doc = str(pm_idx.loc[mgmt].get("物件名") or summary or mgmt).strip()
    else:
        name_for_doc = str(summary or mgmt).strip()

    replacements = build_replacements_from_event(ev, name_for_doc, tags)
    safe_title = _build_safe_title(when_str, mgmt, name_for_doc)
    return generate_docx_from_template_like(template_path, replacements, safe_title)


def _pack_zip(outputs: List[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used: set[str] = set()
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


def _pm_index(pm_view_df: pd.DataFrame) -> Optional[pd.DataFrame]:
    if (
        isinstance(pm_view_df, pd.DataFrame)
        and not pm_view_df.empty
        and "管理番号" in pm_view_df.columns
    ):
        return pm_view_df.set_index("管理番号")
    return None


def _clear_candidates() -> None:
    for k in ("notice_fax_candidates_df", "notice_fax_events_by_id", "notice_fax_pm_view_df"):
        st.session_state.pop(k, None)


# ─────────────────────────────────────────────────────────
# ウィザード UI
# ─────────────────────────────────────────────────────────

def _render_step_indicator(current: int) -> None:
    """上部に進行状況バーを表示"""
    cols = st.columns(len(STEPS))
    for i, (col, label) in enumerate(zip(cols, STEPS)):
        if i < current:
            col.markdown(
                f"<div style='text-align:center;color:#4CAF50;font-size:0.85rem;font-weight:600'>"
                f"✅ {label}</div>",
                unsafe_allow_html=True,
            )
        elif i == current:
            col.markdown(
                f"<div style='text-align:center;color:#1976D2;font-size:0.85rem;font-weight:700;"
                f"background:#E3F2FD;padding:4px 2px;border-radius:6px'>"
                f"▶ {label}</div>",
                unsafe_allow_html=True,
            )
        else:
            col.markdown(
                f"<div style='text-align:center;color:#9E9E9E;font-size:0.85rem'>"
                f"{label}</div>",
                unsafe_allow_html=True,
            )
    st.write("")


def _step1_settings(
    editable_calendar_options: Dict[str, str],
    sheets_service: Any,
    current_user_email: Optional[str],
) -> tuple[bool, str, str, date, date, bool, pd.DataFrame]:
    """
    STEP1: 設定
    戻り値: (valid, calendar_name, calendar_id, start_date, end_date, use_master_filter, pm_view_df)
    """
    _render_step_indicator(0)
    st.markdown("#### カレンダー・期間・対象の設定")

    # ── 物件マスタ ──
    spreadsheet_id = get_property_master_spreadsheet_id(current_user_email)
    pm_view_df: pd.DataFrame = pd.DataFrame()
    has_master = False

    with st.expander("📋 物件マスタの状態を確認する", expanded=False):
        if sheets_service is not None and spreadsheet_id:
            st.markdown(
                f"スプレッドシート: "
                f"[こちらを開く](https://docs.google.com/spreadsheets/d/{spreadsheet_id})"
            )
            try:
                pm_view_df = load_property_master_view(
                    sheets_service, spreadsheet_id,
                    basic_sheet_title="物件基本情報",
                    master_sheet_title="物件マスタ",
                )
                if pm_view_df is not None and not pm_view_df.empty:
                    has_master = True
                    col_flag = "貼り紙テンプレ種別"
                    total = len(pm_view_df)
                    if col_flag in pm_view_df.columns:
                        cnt = (
                            pm_view_df[col_flag].fillna("").astype(str).str.strip() == "自社"
                        ).sum()
                        st.success(f"読み込み完了: {total} 件登録 ／ 貼り紙対象（自社）: {cnt} 件")
                    else:
                        st.success(f"読み込み完了: {total} 件登録（貼り紙テンプレ種別列なし）")
                else:
                    st.warning(
                        "物件マスタが空です。タブ『物件マスタ管理』で登録してください。"
                    )
            except Exception as e:
                st.error(str(e))
        else:
            st.info(
                "物件マスタ用スプレッドシートIDが未設定、またはスプレッドシートサービスが利用できません。\n"
                "管理番号付きイベントをすべて一覧表示するモードで動作します。"
            )

    # ── カレンダー ──
    cal_names = list(editable_calendar_options.keys())
    base = (
        st.session_state.get("base_calendar_name")
        or st.session_state.get("selected_calendar_name")
        or cal_names[0]
    )
    if base not in cal_names:
        base = cal_names[0]
    key = "notice_fax_calendar"
    if key not in st.session_state or st.session_state[key] not in cal_names:
        st.session_state[key] = base

    calendar_name = st.selectbox("📅 対象カレンダー", cal_names, key=key)
    calendar_id = editable_calendar_options[calendar_name]

    # ── 期間 ──
    today = date.today()
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("開始日", value=today, key="notice_fax_start_date")
    with col2:
        end_date = st.date_input("終了日", value=today + timedelta(days=60), key="notice_fax_end_date")

    if start_date > end_date:
        st.error("⚠️ 開始日は終了日以前の日付を指定してください。")
        return False, calendar_name, calendar_id, start_date, end_date, False, pm_view_df

    # ── 対象モード ──
    if has_master:
        opts = [
            "物件マスタの『貼り紙テンプレ種別=自社』のみ対象にする",
            "管理番号付きイベントをすべて一覧表示して手動選択する",
        ]
        choice = st.radio(
            "🔍 貼り紙対象の選び方",
            opts,
            index=0,
            key="notice_fax_mode",
            help="「自社」モードは物件マスタに登録済みかつテンプレ種別が『自社』の物件のみが対象になります。",
        )
        use_master_filter = (choice == opts[0])
    else:
        st.info("物件マスタが利用できないため、管理番号付きイベントをすべて対象とします。")
        use_master_filter = False

    return True, calendar_name, calendar_id, start_date, end_date, use_master_filter, pm_view_df


def _step2_fetch(
    service: Any,
    calendar_id: str,
    start_date: date,
    end_date: date,
    use_master_filter: bool,
    pm_view_df: pd.DataFrame,
) -> bool:
    """
    STEP2: イベント取得ボタン
    取得・整形して session_state に保存。成功したら True を返す。
    """
    _render_step_indicator(1)
    st.markdown("#### カレンダーイベントの取得")

    col1, col2, _ = st.columns([2, 2, 3])
    with col1:
        st.metric("対象期間", f"{start_date.strftime('%m/%d')} 〜 {end_date.strftime('%m/%d')}")
    with col2:
        days = (end_date - start_date).days + 1
        st.metric("日数", f"{days} 日間")

    fetch_btn = st.button(
        "📅 イベントを取得して候補を作成する",
        type="primary",
        use_container_width=True,
    )
    if not fetch_btn:
        # 既に取得済みなら次へ進める
        return bool(st.session_state.get("notice_fax_candidates_df") is not None)

    with st.spinner("カレンダーイベントを取得中..."):
        events = fetch_events_in_range(service, calendar_id, start_date, end_date)

    st.caption(f"取得イベント: {len(events)} 件")

    pm_idx = _pm_index(pm_view_df)
    candidates: List[Dict[str, Any]] = []
    events_by_id: Dict[str, Dict[str, Any]] = {}

    for ev in events:
        desc = safe_get(ev, "description") or ""
        summary = safe_get(ev, "summary") or ""

        # 管理番号抽出（説明 or タイトル）
        mgmt = extract_assetnum(desc) or extract_assetnum(summary)
        if not mgmt:
            continue
        mgmt_norm = mgmt.strip()

        pm_row = None
        flag_val = ""
        pm_name = ""
        pm_remark = ""

        if pm_idx is not None and mgmt_norm in pm_idx.index:
            pm_row = pm_idx.loc[mgmt_norm]
            flag_val = str(pm_row.get("貼り紙テンプレ種別", "")).strip()
            pm_name = _display_value(pm_row.get("物件名", ""))
            pm_remark = _display_value(pm_row.get("備考", ""))

        # 物件マスタの「貼り紙テンプレ種別=自社」で絞り込むモード
        if use_master_filter:
            if pm_row is None:
                continue
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

        name_for_list = _display_value(pm_name or summary or mgmt_norm or "")
        remark_for_list = _display_value(pm_remark or "")

        candidates.append(
            {
                "作成": True,
                "event_id": ev.get("id") or "",
                "管理番号": mgmt_norm,
                "物件名": name_for_list,
                "予定日": date_str,
                "予定時間": time_str,
                "作業タイプ": work_type,
                "テンプレファイル": template_file,
                "貼り紙フラグ": flag_val,
                "イベントタイトル": _display_value(summary),
                "備考": remark_for_list,
            }
        )

        ev_id = ev.get("id")
        if ev_id:
            events_by_id[ev_id] = ev

    if not candidates:
        _clear_candidates()
        if use_master_filter:
            st.warning(
                "物件マスタの『貼り紙テンプレ種別=自社』に該当するイベントは見つかりませんでした。\n"
                "モードを変更するか、物件マスタを確認してください。"
            )
        else:
            st.warning("管理番号付きイベントは見つかりませんでした。対象期間やカレンダーを確認してください。")
        return False

    cand_df = pd.DataFrame(candidates).fillna("")
    cand_df["作成"] = cand_df["作成"].astype(bool)
    st.session_state["notice_fax_candidates_df"] = cand_df
    st.session_state["notice_fax_events_by_id"] = events_by_id
    st.session_state["notice_fax_pm_view_df"] = pm_view_df

    st.success(f"✅ 貼り紙候補 {len(candidates)} 件を作成しました。次のステップで内容を確認してください。")
    return True


def _step3_review() -> bool:
    """
    STEP3: 候補確認・チェックボックス操作
    確定したら True を返す。
    """
    cand_df: Optional[pd.DataFrame] = st.session_state.get("notice_fax_candidates_df")
    if cand_df is None or cand_df.empty:
        st.info("候補データがありません。STEP2 でイベントを取得してください。")
        return False

    _render_step_indicator(2)
    st.markdown("#### 貼り紙作成候補の確認")

    # 統計表示
    total = len(cand_df)
    checked = int(cand_df["作成"].sum())
    col1, col2, col3 = st.columns(3)
    col1.metric("候補件数", f"{total} 件")
    col2.metric("作成チェック済み", f"{checked} 件")
    col3.metric("スキップ", f"{total - checked} 件")

    # 一括操作
    bc1, bc2, _ = st.columns([1, 1, 4])
    if bc1.button("✅ すべてON", key="check_all"):
        cand_df["作成"] = True
        st.session_state["notice_fax_candidates_df"] = cand_df
    if bc2.button("☐ すべてOFF", key="uncheck_all"):
        cand_df["作成"] = False
        st.session_state["notice_fax_candidates_df"] = cand_df

    st.caption("『作成』がチェックされている行のみ Word ファイルを生成します。")

    # DISPLAY_COLS の補完
    for col in DISPLAY_COLS:
        if col not in cand_df.columns:
            cand_df[col] = True if col == "作成" else ""

    disp_df = cand_df[DISPLAY_COLS].copy()
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

    col_cfg = {
        "作成": st.column_config.CheckboxColumn("作成"),
        **{
            c: st.column_config.TextColumn(c, disabled=True)
            for c in DISPLAY_COLS
            if c != "作成"
        },
    }

    edit_df = st.data_editor(
        disp_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config=col_cfg,
        key="notice_fax_editor",
    )

    cand_df["作成"] = edit_df["作成"].values
    st.session_state["notice_fax_candidates_df"] = cand_df

    checked_now = int(cand_df["作成"].sum())
    if checked_now == 0:
        st.warning("⚠️ 『作成』にチェックが入っている行がありません。1件以上チェックしてください。")
        return False

    return True


def _step4_generate() -> None:
    """STEP4: 貼り紙生成 & ZIP ダウンロード"""
    _render_step_indicator(3)
    st.markdown("#### 貼り紙の一括生成")

    cand_df: Optional[pd.DataFrame] = st.session_state.get("notice_fax_candidates_df")
    events_by_id: Dict[str, Dict[str, Any]] = st.session_state.get("notice_fax_events_by_id", {})
    pm_view_df = st.session_state.get("notice_fax_pm_view_df", pd.DataFrame())

    if cand_df is None or cand_df.empty:
        st.error("候補データがありません。STEP2 からやり直してください。")
        return
    if not events_by_id:
        st.error("内部イベント情報が見つかりません。STEP2 からやり直してください。")
        return

    target_df = cand_df[cand_df["作成"] == True].copy()
    if target_df.empty:
        st.warning("作成チェックが付いている行がありません。STEP3 で確認してください。")
        return

    total = len(target_df)
    col1, col2 = st.columns(2)
    col1.metric("生成対象", f"{total} 件")

    generate_btn = st.button(
        "🖨️ 貼り紙を一括生成して ZIP をダウンロードする",
        type="primary",
        use_container_width=True,
    )
    if not generate_btn:
        return

    pm_idx = _pm_index(pm_view_df)
    outputs: List[tuple[str, bytes]] = []
    errors: List[str] = []

    progress = st.progress(0, text="生成を開始します...")
    status = st.empty()

    for i, (_, row) in enumerate(target_df.iterrows(), start=1):
        mgmt = row.get("管理番号", "")
        progress.progress(i / total, text=f"生成中... {i}/{total} 件目 ({mgmt})")
        status.caption(f"処理中: 管理番号 {mgmt}")

        event_id = row.get("event_id")
        if not event_id or not mgmt:
            errors.append(f"行 {i}: event_id または管理番号が空のためスキップ")
            continue

        ev = events_by_id.get(event_id)
        if ev is None:
            errors.append(f"{mgmt}: イベント情報が見つかりません（event_id={event_id}）")
            continue

        try:
            out_name, content = _generate_single_docx(ev, row, pm_idx)
            outputs.append((out_name, content))
        except FileNotFoundError as e:
            errors.append(f"{mgmt}: テンプレートファイルが見つかりません – {e}")
        except Exception as e:
            errors.append(f"{mgmt}: {e}")

    progress.empty()
    status.empty()

    # ── 結果表示 ──
    if not outputs:
        st.error("すべての生成に失敗しました。テンプレートの配置・権限を確認してください。")
        _show_errors(errors)
        return

    success_count = len(outputs)
    error_count = len(errors)

    res_col1, res_col2 = st.columns(2)
    res_col1.metric("✅ 生成成功", f"{success_count} 件")
    if error_count:
        res_col2.metric("⚠️ エラー", f"{error_count} 件")
        _show_errors(errors)
    else:
        res_col2.metric("エラー", "0 件")

    zip_data = _pack_zip(outputs)
    zip_filename = f"harigami_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"

    st.success(f"ZIP ファイルを準備しました（{success_count} ファイル）")
    st.download_button(
        label="📦 貼り紙 ZIP をダウンロード",
        data=zip_data,
        file_name=zip_filename,
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )


def _show_errors(errors: List[str]) -> None:
    if not errors:
        return
    with st.expander(f"⚠️ エラー詳細（{len(errors)} 件）を確認する"):
        for err in errors:
            st.error(err, icon="🚨")


# ─────────────────────────────────────────────────────────
# メイン関数
# ─────────────────────────────────────────────────────────

def render_tab8_notice_fax(
    service: Any,
    editable_calendar_options: Dict[str, str],
    sheets_service: Any = None,
    current_user_email: Optional[str] = None,
    **kwargs,
) -> None:
    """貼り紙自動生成タブ（ウィザード形式）"""
    st.subheader("📄 貼り紙自動作成")

    # ── 前提チェック ──────────────────────────────────────
    if not HARIGAMI_AVAILABLE:
        st.error(
            "貼り紙生成モジュール（utils.harigami_generator）が読み込めませんでした。\n\n"
            "**確認事項:**\n"
            "- `requirements.txt` に `python-docx` が追加されているか\n"
            "- `utils/harigami_generator.py` が正しく配置されているか\n\n"
            f"詳細エラー: `{HARIGAMI_IMPORT_ERROR}`"
        )
        return

    if not service or not editable_calendar_options:
        st.warning(
            "⚠️ カレンダーサービスが初期化されていません。\n"
            "タブ1〜2で認証を完了してから再度お試しください。"
        )
        return

    # ── リセットボタン（右端に配置） ─────────────────────
    _, reset_col = st.columns([6, 1])
    with reset_col:
        if st.button("🔄 リセット", help="取得済みの候補データをクリアして最初からやり直します"):
            _clear_candidates()
            st.rerun()

    st.divider()

    # ── STEP1: 設定 ───────────────────────────────────────
    valid, calendar_name, calendar_id, start_date, end_date, use_master_filter, pm_view_df = (
        _step1_settings(editable_calendar_options, sheets_service, current_user_email)
    )
    if not valid:
        return

    st.divider()

    # ── STEP2: イベント取得 ───────────────────────────────
    fetched = _step2_fetch(
        service, calendar_id, start_date, end_date, use_master_filter, pm_view_df
    )

    cand_exists = st.session_state.get("notice_fax_candidates_df") is not None
    if not cand_exists:
        return

    st.divider()

    # ── STEP3: 候補確認 ───────────────────────────────────
    ready = _step3_review()
    if not ready:
        return

    st.divider()

    # ── STEP4: 生成・DL ───────────────────────────────────
    _step4_generate()
