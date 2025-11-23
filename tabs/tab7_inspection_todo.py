# tabs/tab7_inspection_todo.py
from __future__ import annotations

from datetime import datetime, date, time, timedelta, timezone
from typing import Any, Dict, Optional, List
import re
import unicodedata

import pandas as pd
import streamlit as st
from firebase_admin import firestore

# 物件マスタの列定義などを流用するために import
from tabs.tab6_property_master import (
    MASTER_COLUMNS,
    BASIC_COLUMNS,
    load_sheet_as_df,
    _normalize_df,
)
from utils.helpers import safe_get  # あれば便利なので


# JST（日時計算用）
JST = timezone(timedelta(hours=9))


# ==========================
# 管理番号抽出ヘルパー
# ==========================

# [管理番号: HK5-123] みたいな表記対応
ASSETNUM_PATTERN = re.compile(
    r"[［\[]?\s*管理番号[：:]\s*([0-9A-Za-z\-]+)\s*[］\]]?"
)


def extract_assetnum(text: str) -> str:
    """Description 等から管理番号を抽出"""
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    m = ASSETNUM_PATTERN.search(s)
    if not m:
        return ""
    return m.group(1).strip()


# ==========================
# イベント関連ヘルパー
# ==========================

def to_utc_range_from_dates(d1: date, d2: date) -> tuple[str, str]:
    """JSTの date 範囲 → Calendar API 用の UTC ISO文字列範囲"""
    start_dt_utc = datetime.combine(d1, time.min, tzinfo=JST).astimezone(timezone.utc)
    end_dt_utc = datetime.combine(d2, time.max, tzinfo=JST).astimezone(timezone.utc)
    return start_dt_utc.isoformat(), end_dt_utc.isoformat()


def get_event_start_date(event: Dict[str, Any]) -> Optional[date]:
    """Googleカレンダーイベントから date 型の開始日を取得"""
    start = event.get("start", {})
    # 終日予定
    if "date" in start:
        try:
            return date.fromisoformat(start["date"])
        except Exception:
            return None
    # 時間付き予定
    if "dateTime" in start:
        try:
            dt = pd.to_datetime(start["dateTime"])
            return dt.date()
        except Exception:
            return None
    return None


def fetch_events_in_range(
    service: Any,
    calendar_id: str,
    start_date: date,
    end_date: date,
) -> List[Dict[str, Any]]:
    """指定カレンダーのイベントを期間指定で全取得"""
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


# ==========================
# 物件マスタ読み込み
# ==========================

def get_property_master_spreadsheet_id(current_user_email: Optional[str]) -> str:
    """Firestore の user_settings からユーザーごとの物件マスタ用 Spreadsheet ID を取得"""
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
    物件基本情報（BASIC_COLUMNS）＋物件マスタ（MASTER_COLUMNS）を読み込み、
    管理番号でマージした DataFrame を返す。
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
        # マスタが空なら管理番号だけでも返しておく
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


# ==========================
# TODO 生成ロジック
# ==========================

def build_methods_str(row: pd.Series) -> str:
    methods = []
    if str(row.get("連絡方法_電話1", "")).strip():
        methods.append("電話1")
    if str(row.get("連絡方法_電話2", "")).strip():
        methods.append("電話2")
    if str(row.get("連絡方法_FAX1", "")).strip():
        methods.append("FAX1")
    if str(row.get("連絡方法_FAX2", "")).strip():
        methods.append("FAX2")
    if str(row.get("連絡方法_メール1", "")).strip():
        methods.append("メール1")
    if str(row.get("連絡方法_メール2", "")).strip():
        methods.append("メール2")
    return " / ".join(methods)


def build_contacts_str(row: pd.Series, kind: str) -> str:
    """
    kind: "電話", "FAX", "メール"
    """
    parts = []

    if kind == "電話":
        tel1 = str(row.get("電話番号1", "")).strip()
        tel2 = str(row.get("電話番号2", "")).strip()
        if tel1:
            parts.append(f"① {tel1}")
        if tel2:
            parts.append(f"② {tel2}")
    elif kind == "FAX":
        fax1 = str(row.get("FAX番号1", "")).strip()
        fax2 = str(row.get("FAX番号2", "")).strip()
        if fax1:
            parts.append(f"① {fax1}")
        if fax2:
            parts.append(f"② {fax2}")
    elif kind == "メール":
        m1 = str(row.get("メールアドレス1", "")).strip()
        m2 = str(row.get("メールアドレス2", "")).strip()
        if m1:
            parts.append(f"① {m1}")
        if m2:
            parts.append(f"② {m2}")
    return " / ".join(parts)


def build_task_title(row: pd.Series, event: Dict[str, Any]) -> str:
    mgmt = str(row.get("管理番号", "")).strip()
    name = str(row.get("物件名", "")).strip()
    if not name:
        name = safe_get(event, "summary") or ""
    base = f"点検連絡: {name}" if name else "点検連絡"
    if mgmt:
        return f"{base}（{mgmt}）"
    return base


def build_task_notes(
    row: pd.Series,
    event: Dict[str, Any],
    start_date: Optional[date],
    due_days_str: str,
) -> str:
    mgmt = str(row.get("管理番号", "")).strip()
    name = str(row.get("物件名", "")).strip()
    addr = str(row.get("住所", "")).strip()
    window = str(row.get("窓口会社", "")).strip()
    note_deadline = str(row.get("備考", "")).strip()

    methods = build_methods_str(row)
    tel = build_contacts_str(row, "電話")
    fax = build_contacts_str(row, "FAX")
    mail = build_contacts_str(row, "メール")

    event_title = safe_get(event, "summary") or ""
    event_desc = safe_get(event, "description") or ""

    lines = []
    lines.append(f"イベントタイトル: {event_title}")
    if start_date:
        lines.append(f"点検予定日: {start_date.strftime('%Y-%m-%d')}")
    if due_days_str:
        lines.append(f"連絡期限_日前: {due_days_str}")
    if note_deadline:
        lines.append(f"備考: {note_deadline}")

    if mgmt:
        lines.append(f"管理番号: {mgmt}")
    if name:
        lines.append(f"物件名: {name}")
    if addr:
        lines.append(f"住所: {addr}")
    if window:
        lines.append(f"窓口: {window}")

    if methods:
        lines.append(f"連絡方法: {methods}")
    if tel:
        lines.append(f"電話: {tel}")
    if fax:
        lines.append(f"FAX: {fax}")
    if mail:
        lines.append(f"メール: {mail}")

    if event_desc:
        lines.append("")
        lines.append("------ イベント説明 ------")
        lines.append(event_desc)

    return "\n".join(lines)


def build_due_iso(due_date: Optional[date]) -> Optional[str]:
    """Tasks API 用の due（RFC3339）を日付ベースで生成"""
    if not due_date:
        return None
    dt_utc = datetime.combine(due_date, time.min, tzinfo=JST).astimezone(timezone.utc)
    return dt_utc.isoformat().replace("+00:00", "Z")


def build_task_candidates(
    events: List[Dict[str, Any]],
    pm_view_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    カレンダーイベントと物件マスタビューを突合して、
    ToDo作成候補の DataFrame を作成。
    """
    if pm_view_df is None or pm_view_df.empty:
        return pd.DataFrame()

    # 管理番号をキーにして検索しやすく
    pm_view_idx = pm_view_df.set_index("管理番号")

    rows = []

    for ev in events:
        desc = safe_get(ev, "description") or ""
        summary = safe_get(ev, "summary") or ""

        mgmt = extract_assetnum(desc) or extract_assetnum(summary)
        if not mgmt:
            continue

        mgmt_norm = mgmt.strip()
        if mgmt_norm not in pm_view_idx.index:
            # マスタに存在しない管理番号はスキップ
            continue

        pm_row = pm_view_idx.loc[mgmt_norm]

        start_date = get_event_start_date(ev)
        due_days_str = str(pm_row.get("連絡期限_日前", "")).strip()
        due_date: Optional[date] = None
        if due_days_str.isdigit() and start_date:
            due_date = start_date - timedelta(days=int(due_days_str))

        methods = build_methods_str(pm_row)
        if not methods:
            # 連絡方法が1つも指定されていない場合は一旦スキップしてもよいが、
            # 今回は候補としては出しておく（ユーザー側で調整可能）
            pass

        tel = build_contacts_str(pm_row, "電話")
        fax = build_contacts_str(pm_row, "FAX")
        mail = build_contacts_str(pm_row, "メール")

        title = build_task_title(pm_row, ev)
        notes = build_task_notes(pm_row, ev, start_date, due_days_str)
        due_iso = build_due_iso(due_date)

        row = {
            "作成": True,
            "event_id": ev.get("id"),
            "管理番号": mgmt_norm,
            "物件名": str(pm_row.get("物件名", "")).strip(),
            "予定日": start_date.strftime("%Y-%m-%d") if start_date else "",
            "連絡期限_日前": due_days_str,
            "ToDo期限日": due_date.strftime("%Y-%m-%d") if due_date else "",
            "連絡方法": methods,
            "電話": tel,
            "FAX": fax,
            "メール": mail,
            "貼り紙テンプレ種別": str(pm_row.get("貼り紙テンプレ種別", "")).strip(),
            "備考": str(pm_row.get("備考", "")).strip(),
            # 実際に Tasks API に投げる情報
            "_todo_title": title,
            "_todo_notes": notes,
            "_todo_due_iso": due_iso,
        }
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    # 作成フラグは bool に
    df["作成"] = df["作成"].astype(bool)
    return df


# ==========================
# メイン UI
# ==========================

def render_tab7_inspection_todo(
    service: Any,
    editable_calendar_options: Dict[str, str],
    tasks_service: Any,
    default_task_list_id: Optional[str],
    sheets_service: Any,
    current_user_email: Optional[str] = None,
):
    """
    点検イベント → 物件マスタ突合 → Google ToDo 自動生成タブ
    """
    st.subheader("点検連絡用 ToDo 自動作成")

    if not service or not editable_calendar_options:
        st.warning("カレンダーサービスが初期化されていません。タブ1〜2で認証を完了してください。")
        return

    if not tasks_service or not default_task_list_id:
        st.warning("Google ToDo（Tasks）サービスが利用できません。サイドバーの認証状態を確認してください。")
        return

    # 物件マスタ用スプレッドシートIDを取得
    spreadsheet_id = get_property_master_spreadsheet_id(current_user_email)
    if not spreadsheet_id:
        st.error("物件マスタ用スプレッドシートIDが設定されていません。タブ『物件マスタ管理』側で一度設定＆保存してください。")
        return

    st.markdown(
        f"物件マスタスプレッドシート: "
        f"[リンクを開く](https://docs.google.com/spreadsheets/d/{spreadsheet_id})"
    )

    # カレンダー選択
    cal_names = list(editable_calendar_options.keys())
    default_cal = cal_names[0] if cal_names else None
    calendar_name = st.selectbox(
        "対象カレンダー",
        cal_names,
        index=(cal_names.index(default_cal) if default_cal in cal_names else 0),
        key="ins_todo_calendar",
    )
    calendar_id = editable_calendar_options.get(calendar_name)

    # 期間指定
    today = date.today()
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input("点検予定の検索開始日", value=today, key="ins_todo_start_date")
    with col_d2:
        end_date = st.date_input("点検予定の検索終了日", value=today + timedelta(days=60), key="ins_todo_end_date")

    if start_date > end_date:
        st.error("開始日は終了日以前の日付を指定してください。")
        return

    # 物件マスタビュー読み込み
    pm_view_df = load_property_master_view(
        sheets_service,
        spreadsheet_id,
        basic_sheet_title="物件基本情報",
        master_sheet_title="物件マスタ",
    )
    if pm_view_df is None or pm_view_df.empty:
        st.error("物件マスタ（＋基本情報）が空です。先にタブ『物件マスタ管理』で登録してください。")
        return

    st.caption(f"物件マスタ登録件数: {len(pm_view_df)} 件（管理番号単位）")

    # イベント取得ボタン
    fetch_btn = st.button("カレンダーから点検イベントを取得し、物件マスタと照合する", type="primary")

    if fetch_btn:
        with st.spinner("カレンダーイベントを取得中..."):
            events = fetch_events_in_range(service, calendar_id, start_date, end_date)

        st.write(f"取得したイベント件数: {len(events)} 件")

        with st.spinner("物件マスタとの照合＆ToDo候補の作成中..."):
            candidates_df = build_task_candidates(events, pm_view_df)

        if candidates_df.empty:
            st.info("管理番号が付与され、かつ物件マスタに登録されている点検イベントは見つかりませんでした。")
            return

        st.session_state["ins_todo_candidates_df"] = candidates_df
        st.success(f"ToDo候補 {len(candidates_df)} 件を作成しました。この下で内容を確認して作成対象を選択できます。")

    # 既に候補が作られている場合はそれを表示
    candidates_df: Optional[pd.DataFrame] = st.session_state.get("ins_todo_candidates_df")
    if candidates_df is None or candidates_df.empty:
        return

    st.markdown("### ToDo 作成候補一覧")
    st.caption("※『作成』チェックが ON の行だけが Google ToDo に登録されます。必要に応じてOFFにしてください。")

    # 表示用にコピー（内部用カラムは隠す）
    display_cols = [
        "作成",
        "管理番号",
        "物件名",
        "予定日",
        "連絡期限_日前",
        "ToDo期限日",
        "連絡方法",
        "電話",
        "FAX",
        "メール",
        "貼り紙テンプレ種別",
        "備考",
    ]
    hidden_cols = [c for c in candidates_df.columns if c not in display_cols]

    edit_df = st.data_editor(
        candidates_df[display_cols],
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key="ins_todo_editor",
    )

    # edit_df には内部用カラムがないので、元DFの該当カラムを更新して戻す
    updated_candidates = candidates_df.copy()
    for col in display_cols:
        updated_candidates[col] = edit_df[col].values
    st.session_state["ins_todo_candidates_df"] = updated_candidates

    # ToDo作成ボタン
    create_btn = st.button("選択された行の ToDo を Google ToDo に一括作成する", type="primary")

    if create_btn:
        df = st.session_state.get("ins_todo_candidates_df")
        if df is None or df.empty:
            st.error("ToDo候補がありません。先にカレンダーから取得してください。")
            return

        target_df = df[df["作成"] == True].copy()
        if target_df.empty:
            st.warning("『作成』にチェックが入っている行がありません。")
            return

        created = 0
        errors: List[str] = []

        with st.spinner("Google ToDo に登録中..."):
            for _, row in target_df.iterrows():
                title = row.get("_todo_title")
                notes = row.get("_todo_notes")
                due_iso = row.get("_todo_due_iso")

                if not title:
                    continue

                body = {
                    "title": title,
                }
                if notes:
                    body["notes"] = notes
                if due_iso:
                    body["due"] = due_iso

                try:
                    tasks_service.tasks().insert(
                        tasklist=default_task_list_id,
                        body=body,
                    ).execute()
                    created += 1
                except Exception as e:
                    errors.append(str(e))

        if created > 0:
            st.success(f"{created} 件の ToDo を作成しました。")
        if errors:
            st.warning(f"一部のToDo作成でエラーが発生しました（{len(errors)} 件）。詳細: {errors[0]}")
