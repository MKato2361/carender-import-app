import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

# ===== 他モジュール依存 =====
from utils.helpers import safe_get
from utils.parsers import extract_worksheet_id_from_text
from excel_parser import (
    process_excel_data_for_calendar,
    get_available_columns_for_event_name,
    check_event_name_columns,
)
from calendar_utils import (
    fetch_all_events,
    add_event_to_calendar,
    update_event_if_needed,
)
from session_utils import (
    get_user_setting,
    set_user_setting,
)
from firebase_admin import firestore


JST = timezone(timedelta(hours=9))


def is_event_changed(existing_event: dict, new_event_data: dict) -> bool:
    nz = lambda v: (v or "")

    if nz(existing_event.get("summary")) != nz(new_event_data.get("summary")):
        return True

    if nz(existing_event.get("description")) != nz(new_event_data.get("description")):
        return True

    if nz(existing_event.get("transparency")) != nz(new_event_data.get("transparency")):
        return True

    if (existing_event.get("start") or {}) != (new_event_data.get("start") or {}):
        return True

    if (existing_event.get("end") or {}) != (new_event_data.get("end") or {}):
        return True

    return False


def default_fetch_window_years(years: int = 2):
    from datetime import datetime, timezone, timedelta

    now_utc = datetime.now(timezone.utc)
    return (
        (now_utc - timedelta(days=365 * years)).isoformat(),
        (now_utc + timedelta(days=365 * years)).isoformat(),
    )


def extract_worksheet_id_from_description(desc: str) -> str | None:
    import re
    import unicodedata

    RE_WORKSHEET_ID = re.compile(r"\[作業指示書[：:]\s*([0-9０-９]+)\]")

    if not desc:
        return None
    m = RE_WORKSHEET_ID.search(desc)
    if not m:
        return None
    return unicodedata.normalize("NFKC", m.group(1)).strip()


# ===== 追加: 作業外予定ファイルを汎用的に読み取り、既存フロー互換のDataFrameへ整形 =====
def _read_outside_file_to_df(file_obj) -> pd.DataFrame:
    name = getattr(file_obj, "name", "")
    if name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj, dtype=str)
    else:
        # CSV: エンコーディングをいくつか試す
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                df = pd.read_csv(file_obj, dtype=str, encoding=enc, errors="ignore")
                break
            except Exception:
                df = None
        if df is None:
            raise ValueError("CSVの読み込みに失敗しました（対応エンコーディング不明）。")
    df = df.fillna("")
    return df


def _build_calendar_df_from_outside(df_raw: pd.DataFrame, private_event: bool, all_day_override: bool) -> pd.DataFrame:
    """
    作業外予定の生データから、既存処理と互換な列構成の DataFrame を生成する
    必要列:
      Subject, Description, All Day Event, Private, Start Date, End Date, Start Time, End Time, Location(任意)
    仕様:
      - イベント名: 備考 + " [作業外予定]"
      - Description: 「理由コード」列
      - 時刻が両方ない行は終日にフォールバック（Q1=Noに基づき“常に終日”ではない）
    """
    # 必須列チェック
    if "備考" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『備考』列が見つかりません。")
    if "理由コード" not in df_raw.columns:
        raise ValueError("作業外予定ファイルに『理由コード』列が見つかりません。")

    # 日付・時刻候補（柔軟に拾う）
    start_date_candidates = ["開始日", "日付", "開始日時", "Start Date", "Date"]
    end_date_candidates = ["終了日", "終了日時", "End Date"]
    start_time_candidates = ["開始時刻", "開始時間", "Start Time"]
    end_time_candidates = ["終了時刻", "終了時間", "End Time"]
    location_candidates = ["場所", "現場名", "所在地", "Location"]

    def pick(col_names):
        for c in col_names:
            if c in df_raw.columns:
                return c
        return None

    c_sd = pick(start_date_candidates)
    c_ed = pick(end_date_candidates)
    c_st = pick(start_time_candidates)
    c_et = pick(end_time_candidates)
    c_loc = pick(location_candidates)

    rows = []
    for _, r in df_raw.iterrows():
        subject = f"{str(r['備考']).strip()} [作業外予定]".strip()
        description = str(r["理由コード"]).strip()

        # 日付
        sd_raw = (str(r[c_sd]).strip() if c_sd else "")
        ed_raw = (str(r[c_ed]).strip() if c_ed else "")

        # 多くのフォーマットを想定してYYYY/MM/DDへ寄せる
        def norm_date(s: str) -> Optional[str]:
            s = s.replace("-", "/").replace(".", "/").strip()
            for fmt in ("%Y/%m/%d", "%Y/%m/%d %H:%M", "%m/%d/%Y", "%Y/%m/%d %H:%M:%S"):
                try:
                    return datetime.strptime(s, fmt).strftime("%Y/%m/%d")
                except Exception:
                    continue
            # 8桁数字(YYYYMMDD)も許容
            if s.isdigit() and len(s) == 8:
                return f"{s[0:4]}/{s[4:6]}/{s[6:8]}"
            return "" if not s else s  # そのまま返す（後工程で失敗時に終日化）

        sd = norm_date(sd_raw)
        ed = norm_date(ed_raw) if ed_raw else sd

        # 時刻
        st_raw = (str(r[c_st]).strip() if c_st else "")
        et_raw = (str(r[c_et]).strip() if c_et else "")

        def norm_time(t: str) -> Optional[str]:
            t = t.replace(".", ":").strip()
            for fmt in ("%H:%M", "%H:%M:%S"):
                try:
                    return datetime.strptime(t, fmt).strftime("%H:%M")
                except Exception:
                    continue
            # 数字3-4桁(HHMM)を許容
            if t.isdigit() and len(t) in (3, 4):
                t = t.zfill(4)
                return f"{t[:2]}:{t[2:]}"
            return ""

        stime = norm_time(st_raw)
        etime = norm_time(et_raw)

        # 時刻が無い/片方のみ → 後工程で安全に扱う
        location = (str(r[c_loc]).strip() if c_loc else "")

        rows.append(
            {
                "Subject": subject,
                "Description": description,
                "All Day Event": "True" if all_day_override else "False",  # 基本False（Q1=No）、ただしUIで上書き可
                "Private": "True" if private_event else "False",
                "Start Date": sd or "",
                "End Date": ed or (sd or ""),
                "Start Time": stime or "",
                "End Time": etime or "",
                "Location": location,
            }
        )

    df = pd.DataFrame(rows)

    # 行ごとに「時刻が両方空 or 日付欠落」は終日へフォールバック
    def apply_fallback(row):
        if row["All Day Event"] == "True":
            return row
        if not row["Start Date"]:
            row["All Day Event"] = "True"
            return row
        if (not row["Start Time"]) and (not row["End Time"]):
            row["All Day Event"] = "True"
            return row
        # 片側のみ時刻がある場合は1時間想定で補完
        if row["Start Time"] and not row["End Time"]:
            try:
                dt = datetime.strptime(row["Start Time"], "%H:%M")
                end_dt = (dt + timedelta(hours=1)).strftime("%H:%M")
                row["End Time"] = end_dt
            except Exception:
                row["All Day Event"] = "True"
        if row["End Time"] and not row["Start Time"]:
            try:
                dt = datetime.strptime(row["End Time"], "%H:%M")
                start_dt = (dt - timedelta(hours=1)).strftime("%H:%M")
                row["Start Time"] = start_dt
            except Exception:
                row["All Day Event"] = "True"
        return row

    df = df.apply(apply_fallback, axis=1)
    return df


def render_tab2_register(user_id: str, editable_calendar_options: dict, service, tasks_service=None, default_task_list_id=None):
    st.subheader("イベントを登録・更新")

    # ===== モード判定 =====
    work_files = st.session_state.get("uploaded_files") or []
    has_work = bool(work_files) and st.session_state.get("merged_df_for_selector") is not None and not st.session_state["merged_df_for_selector"].empty
    outside_file = st.session_state.get("uploaded_outside_work_file")
    outside_mode = bool(outside_file) and not has_work  # tab1の仕様上どちらかのみ

    if not has_work and not outside_mode:
        st.info("先に「1. ファイルのアップロード」タブでファイルをアップロードしてください。")
        return

    if not editable_calendar_options:
        st.error("登録可能なカレンダーが見つかりませんでした。Googleカレンダーの設定を確認してください。")
        return

    # ===== カレンダー選択（モード別に保存キーを変える）=====
    calendar_options = list(editable_calendar_options.keys())
    if outside_mode:
        saved_calendar_name = get_user_setting(user_id, "selected_calendar_name_outside")
    else:
        saved_calendar_name = get_user_setting(user_id, "selected_calendar_name")

    try:
        default_index = calendar_options.index(saved_calendar_name)
    except Exception:
        default_index = 0

    selected_calendar_name = st.selectbox(
        "登録先カレンダーを選択" + ("（作業外予定）" if outside_mode else "（作業指示書）"),
        calendar_options,
        index=default_index,
        key="reg_calendar_select_outside" if outside_mode else "reg_calendar_select",
    )
    calendar_id = editable_calendar_options[selected_calendar_name]

    if outside_mode:
        set_user_setting(user_id, "selected_calendar_name_outside", selected_calendar_name)
    else:
        set_user_setting(user_id, "selected_calendar_name", selected_calendar_name)

    # ===== イベント設定（共通UIを使用、デフォルトはQ1=Noに合わせ終日=False）=====
    with st.expander("📝 イベント設定", expanded=not outside_mode):
        all_day_event_override = st.checkbox("終日イベントとして登録", value=False, key=f"all_day_override_{'outside' if outside_mode else 'work'}")
        private_event = st.checkbox("非公開イベントとして登録", value=True, key=f"private_event_{'outside' if outside_mode else 'work'}")

        if outside_mode:
            st.caption("※ 作業外予定では説明列の選択は不要です（Description は『理由コード』列が使用されます）")
            description_columns = []  # 使わない
        else:
            description_columns_pool = st.session_state.get("description_columns_pool", [])
            saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
            default_selection = [col for col in saved_description_cols if col in description_columns_pool]
            description_columns = st.multiselect(
                "説明欄に含める列（複数選択可）",
                description_columns_pool,
                default=default_selection,
                key=f"description_selector_register_{user_id}",
            )
            set_user_setting(user_id, "description_columns_selected", description_columns)

    # ===== イベント名設定（作業外予定は固定仕様 / 作業指示書は従来のまま）=====
    if outside_mode:
        st.info("イベント名は『備考 + [作業外予定]』で登録します。")
        add_task_type_to_event_name = False
        fallback_event_name_column = None
    else:
        with st.expander("🧱 イベント名の生成設定", expanded=True):
            has_mng_data, has_name_data = check_event_name_columns(st.session_state["merged_df_for_selector"])
            saved_event_name_col = get_user_setting(user_id, "event_name_col_selected")
            saved_task_type_flag = get_user_setting(user_id, "add_task_type_to_event_name")
            add_task_type_to_event_name = st.checkbox(
                "イベント名の先頭に作業タイプを追加する",
                value=bool(saved_task_type_flag),
                key=f"add_task_type_checkbox_{user_id}",
            )
            fallback_event_name_column = None
            if not (has_mng_data and has_name_data):
                available_event_name_cols = get_available_columns_for_event_name(
                    st.session_state["merged_df_for_selector"]
                )
                event_name_options = ["選択しない"] + available_event_name_cols
                try:
                    name_index = event_name_options.index(saved_event_name_col) if saved_event_name_col else 0
                except Exception:
                    name_index = 0
                selected_event_name_col = st.selectbox(
                    "イベント名として使用する代替列を選択してください:",
                    options=event_name_options,
                    index=name_index,
                    key=f"event_name_selector_register_{user_id}",
                )
                if selected_event_name_col != "選択しない":
                    fallback_event_name_column = selected_event_name_col
                set_user_setting(user_id, "event_name_col_selected", selected_event_name_col)
            else:
                st.info("「管理番号」と「物件名」のデータが両方存在するため、それらがイベント名として使用されます。")
            set_user_setting(user_id, "add_task_type_to_event_name", add_task_type_to_event_name)

    # ===== ToDo設定（作業外予定は表示しない／作成しない）=====
    if not outside_mode:
        st.subheader("✅ ToDoリスト連携設定 (オプション)")
        with st.expander("ToDoリスト作成オプション", expanded=False):
            create_todo = st.checkbox(
                "このイベントに対応するToDoリストを作成する",
                value=bool(get_user_setting(user_id, "create_todo_checkbox_state")),
                key="create_todo_checkbox",
            )
            set_user_setting(user_id, "create_todo_checkbox_state", create_todo)
            if create_todo:
                st.markdown("以下のToDoが**常にすべて**作成されます: `点検通知`")
            else:
                st.markdown("ToDoリストの作成は無効です。")
            # 期限設定UIは既存ロジックを流用（省略可）

    # ===== 実行 =====
    st.subheader("➡️ イベント登録・更新実行")
    if st.button("Googleカレンダーに登録・更新する"):
        try:
            if outside_mode:
                # 外予定: 独自読み込み → 互換DFへ整形
                raw_df = _read_outside_file_to_df(outside_file)
                df = _build_calendar_df_from_outside(
                    raw_df,
                    private_event=private_event,
                    all_day_override=all_day_event_override,
                )
            else:
                # 従来フロー
                df = process_excel_data_for_calendar(
                    st.session_state["uploaded_files"],
                    # description_columns は上のUIで決定済み
                    description_columns,
                    all_day_event_override,
                    private_event,
                    # フォールバック列（従来仕様）
                    fallback_event_name_column,
                    add_task_type_to_event_name,
                )
        except Exception as e:
            st.error(f"Excelデータ処理中にエラーが発生しました: {e}")
            return

        if df.empty:
            st.warning("有効なイベントデータがありません。処理を中断しました。")
            return

        st.info(f"{len(df)} 件のイベントを処理します。")
        progress = st.progress(0)

        added_count = 0
        updated_count = 0
        skipped_count = 0

        time_min, time_max = default_fetch_window_years(2)
        with st.spinner("既存イベントを取得中..."):
            events = fetch_all_events(service, calendar_id, time_min, time_max)

        worksheet_to_event: Dict[str, dict] = {}
        for event in events or []:
            wid = extract_worksheet_id_from_description(event.get("description") or "")
            if wid:
                worksheet_to_event[wid] = event

        total = len(df)

        for i, row in df.iterrows():
            desc_text = safe_get(row, "Description", "")

            # 外予定は作業指示書IDがない想定 → 既存照合は機能しない（常に新規扱い）
            worksheet_id = extract_worksheet_id_from_text(desc_text) if not outside_mode else None

            all_day_flag = safe_get(row, "All Day Event", "True" if outside_mode else "True")
            private_flag = safe_get(row, "Private", "True")
            start_date_str = safe_get(row, "Start Date", "")
            end_date_str = safe_get(row, "End Date", "")
            start_time_str = safe_get(row, "Start Time", "")
            end_time_str = safe_get(row, "End Time", "")

            event_data = {
                "summary": safe_get(row, "Subject", ""),
                "location": safe_get(row, "Location", ""),
                "description": desc_text,
                "transparency": "transparent" if private_flag == "True" else "opaque",
            }

            try:
                if all_day_flag == "True":
                    sd = datetime.strptime(start_date_str, "%Y/%m/%d").date()
                    ed = datetime.strptime(end_date_str or start_date_str, "%Y/%m/%d").date()
                    event_data["start"] = {"date": sd.strftime("%Y-%m-%d")}
                    event_data["end"] = {"date": (ed + timedelta(days=1)).strftime("%Y-%m-%d")}
                else:
                    sdt = datetime.strptime(f"{start_date_str} {start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                    edt = datetime.strptime(f"{end_date_str or start_date_str} {end_time_str or start_time_str}", "%Y/%m/%d %H:%M").replace(tzinfo=JST)
                    event_data["start"] = {"dateTime": sdt.isoformat(), "timeZone": "Asia/Tokyo"}
                    event_data["end"] = {"dateTime": edt.isoformat(), "timeZone": "Asia/Tokyo"}
            except Exception as e:
                st.error(f"行 {i} の日時パースに失敗しました: {e}")
                progress.progress((i + 1) / total)
                continue

            existing_event = worksheet_to_event.get(worksheet_id) if worksheet_id else None

            try:
                if existing_event:
                    if is_event_changed(existing_event, event_data):
                        _ = update_event_if_needed(service, calendar_id, existing_event["id"], event_data)
                        updated_count += 1
                    else:
                        skipped_count += 1
                else:
                    added_event = add_event_to_calendar(service, calendar_id, event_data)
                    if added_event:
                        added_count += 1
                        if worksheet_id:
                            worksheet_to_event[worksheet_id] = added_event
            except Exception as e:
                st.error(f"イベント '{event_data.get('summary','(無題)')}' の登録/更新に失敗しました: {e}")

            progress.progress((i + 1) / total)

        st.success(f"✅ 登録: {added_count} / 🔧 更新: {updated_count} / ↪ スキップ: {skipped_count}")
