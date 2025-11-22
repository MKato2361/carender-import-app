# tabs/tab6_property_master.py
from __future__ import annotations

from datetime import datetime
from typing import Optional, Any
from io import BytesIO
import re
import unicodedata

import pandas as pd
import streamlit as st
from firebase_admin import firestore  # ユーザーごとのID保存用


# ==========================
# 列定義
# ==========================

# 物件マスタ（点検条件・連絡方法など）
MASTER_COLUMNS = [
    "管理番号",
    "点検実施月",
    "連絡期限_日前",
    "連絡方法_電話1",
    "連絡方法_電話2",
    "連絡方法_FAX1",
    "連絡方法_FAX2",
    "連絡方法_メール1",
    "連絡方法_メール2",
    "電話番号1",
    "電話番号2",
    "FAX番号1",
    "FAX番号2",
    "メールアドレス1",
    "メールアドレス2",
    "連絡宛名1",
    "連絡宛名2",
    "OK曜日",
    "NG曜日",
    "OK時間帯_開始",
    "OK時間帯_終了",
    "NG時間帯_開始",
    "NG時間帯_終了",
    "貼り紙テンプレ種別",
    "貼り紙テンプレ_ドライブID",
    "FAXテンプレ種別",
    "FAXテンプレ_ドライブID",
    "メールテンプレ_ドライブID",
    "備考",
    "更新日時",
    "最終更新者",
]

# 物件基本情報（Excel/CSV から取り込む）
BASIC_COLUMNS = [
    "管理番号",
    "物件名",
    "住所",
    "窓口会社",
    "担当部署",
    "担当者名",
    "契約種別",
]


# ==========================
# 共通ヘルパー
# ==========================

def _normalize_df(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """指定列だけに揃えて、文字列 + strip に統一"""
    df = df.copy() if df is not None else pd.DataFrame()
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns].copy()
    if not df.empty:
        df = df.astype(str).apply(lambda col: col.str.strip())
    return df


def parse_notice_deadline_to_days(text: str) -> tuple[str, str]:
    """
    「点検通知先１通知期限」の文字列 → 日数（文字列）と、解析できなかった場合用のメモ
      - 例: "1週間前" → ("7", "")
      - 例: "10日前" → ("10", "")
      - それ以外 → ("", 元の文字列)
    """
    s = str(text or "").strip()
    if not s:
        return "", ""

    s_norm = unicodedata.normalize("NFKC", s)  # 全角→半角など
    # 〇週間
    m = re.search(r"(\d+)\s*週", s_norm)
    if m:
        days = int(m.group(1)) * 7
        return str(days), ""
    # 〇日前 / 〇日
    m = re.search(r"(\d+)\s*日", s_norm)
    if m:
        days = int(m.group(1))
        return str(days), ""

    # 解析できないものは備考側へ回す
    return "", s


# ==========================
# Sheets ヘルパー
# ==========================

def ensure_sheet_and_headers(
    sheets_service: Any,
    spreadsheet_id: str,
    sheet_title: str,
    headers: list[str],
) -> None:
    """
    指定スプレッドシート内にシートを作成し、
    1行目にヘッダーをセットする（なければ）。
    """
    if not sheets_service or not spreadsheet_id:
        return

    # シート一覧取得
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = meta.get("sheets", [])
    existing_titles = {s["properties"]["title"] for s in sheets}

    # シートがなければ追加
    if sheet_title not in existing_titles:
        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_title,
                        }
                    }
                }
            ]
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body,
        ).execute()

    # ヘッダー行の確認
    range_header = f"{sheet_title}!1:1"
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_header,
    ).execute()
    values = result.get("values", [])

    need_update_header = False
    if not values:
        need_update_header = True
    else:
        current_header = values[0]
        if current_header != headers:
            need_update_header = True

    if need_update_header:
        body = {"values": [headers]}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_title}!A1",
            valueInputOption="RAW",
            body=body,
        ).execute()


def create_property_master_spreadsheet(
    sheets_service: Any,
    user_email: Optional[str] = None,
) -> str:
    """
    物件基本情報 / 物件マスタ の2シートを持つスプレッドシートを新規作成し、
    ヘッダーを設定して Spreadsheet ID を返す。
    """
    if not sheets_service:
        raise RuntimeError("Sheets service is not initialized")

    title_suffix = user_email or "property_master"
    body = {
        "properties": {
            "title": f"物件マスタ_{title_suffix}",
        },
        "sheets": [
            {"properties": {"title": "物件基本情報"}},
            {"properties": {"title": "物件マスタ"}},
        ],
    }
    resp = sheets_service.spreadsheets().create(body=body).execute()
    spreadsheet_id = resp["spreadsheetId"]

    # ヘッダー書き込み
    ensure_sheet_and_headers(sheets_service, spreadsheet_id, "物件基本情報", BASIC_COLUMNS)
    ensure_sheet_and_headers(sheets_service, spreadsheet_id, "物件マスタ", MASTER_COLUMNS)

    return spreadsheet_id


def load_sheet_as_df(
    sheets_service: Any,
    spreadsheet_id: str,
    sheet_title: str,
    columns: list[str],
) -> pd.DataFrame:
    """
    A1 からの内容を DataFrame として取得し、指定列に揃えて返す。
    ※ 行によって列数がバラバラでも、ヘッダー数に合わせてパディングする。
    """
    if not sheets_service or not spreadsheet_id:
        return pd.DataFrame(columns=columns)

    range_name = f"{sheet_title}!A1:ZZ"
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name,
        ).execute()
    except Exception as e:
        st.error(f"{sheet_title} シートの読み込みに失敗しました: {e}")
        return pd.DataFrame(columns=columns)

    values = result.get("values", [])
    if not values:
        return pd.DataFrame(columns=columns)

    header = values[0]
    rows = values[1:] if len(values) > 1 else []

    # 行ごとの差をパディング／切り詰め
    padded_rows = []
    for row in rows:
        if len(row) < len(header):
            padded_rows.append(row + [""] * (len(header) - len(row)))
        elif len(row) > len(header):
            padded_rows.append(row[:len(header)])
        else:
            padded_rows.append(row)

    df = pd.DataFrame(padded_rows, columns=header)
    df = df.astype(str).apply(lambda col: col.str.strip())

    # 足りない列補完
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[columns].copy()


def save_df_to_sheet(
    sheets_service: Any,
    spreadsheet_id: str,
    sheet_title: str,
    df: pd.DataFrame,
    columns: list[str],
) -> None:
    """指定 DataFrame をヘッダー込みでシートにまるごと書き戻す。"""
    if not sheets_service or not spreadsheet_id:
        return

    df_to_save = _normalize_df(df, columns)
    values = [columns] + df_to_save.values.tolist()

    try:
        # シート全体クリア
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=sheet_title,
        ).execute()

        body = {"values": values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_title}!A1",
            valueInputOption="RAW",
            body=body,
        ).execute()
    except Exception as e:
        st.error(f"{sheet_title} シートへの保存に失敗しました: {e}")
        raise


# ==========================
# 物件基本情報：Excel/CSV 読み込み & 差分
# ==========================

def load_raw_from_uploaded(uploaded_file) -> pd.DataFrame:
    """
    アップロードされた Excel/CSV を「そのまま」DataFrame にする（全列保持）。
    文字列化＆stripだけ実施。
    """
    if uploaded_file is None:
        return pd.DataFrame()

    name = uploaded_file.name.lower()

    # Excel
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, dtype=str)
        if not df.empty:
            df = df.astype(str).apply(lambda col: col.str.strip())
        return df

    # CSV
    raw_bytes = uploaded_file.read()
    encodings_to_try = ["utf-8", "utf-8-sig", "cp932"]
    last_err: Optional[Exception] = None

    for enc in encodings_to_try:
        try:
            df = pd.read_csv(BytesIO(raw_bytes), dtype=str, encoding=enc)
            if not df.empty:
                df = df.astype(str).apply(lambda col: col.str.strip())
            return df
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue

    st.error(f"CSVファイルの読み込みに失敗しました。エンコーディングを確認してください。（最後のエラー: {last_err}）")
    return pd.DataFrame()


def _map_basic_from_raw_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    元の DataFrame（どんなヘッダー名でもOK）から BASIC_COLUMNS を構成する。
    例:
      - 管理番号  ← 物件の管理番号 / 物件管理番号 / 管理番号
      - 住所     ← 物件情報-住所1 / 住所 / 所在地
      - 契約種別 ← 契約種類 / 契約種別
      - 窓口会社 ← 窓口名優先、なければ契約先名、なければ窓口会社
    """
    df = df.copy()
    if not df.empty:
        df = df.astype(str).apply(lambda col: col.str.strip())
    n = len(df)

    def pick(*names: str) -> pd.Series:
        for name in names:
            if name in df.columns:
                return df[name]
        return pd.Series([""] * n)

    mapped = pd.DataFrame()
    mapped["管理番号"] = pick("管理番号", "物件の管理番号", "物件管理番号", "物件番号")
    mapped["物件名"] = pick("物件名", "施設名")
    mapped["住所"] = pick("住所", "物件情報-住所1", "住所1", "所在地")
    # 窓口会社：窓口名 → 契約先名 → 窓口会社 の順で優先
    mapped["窓口会社"] = pick("窓口名", "契約先名", "窓口会社")
    mapped["担当部署"] = pick("担当部署", "部署名")
    mapped["担当者名"] = pick("担当者名", "担当者")
    mapped["契約種別"] = pick("契約種別", "契約種類")

    # 管理番号が空の行は除外
    mapped = mapped[mapped["管理番号"].astype(str).str.strip() != ""].reset_index(drop=True)

    return _normalize_df(mapped, BASIC_COLUMNS)


def diff_basic_info(current_df: pd.DataFrame, new_df: pd.DataFrame):
    """
    current_df: 現在シートに入っている基本情報
    new_df    : 新しくアップロードされた基本情報（BASIC_COLUMNS）

    戻り値:
      - new_rows     : 新規追加行
      - updated_rows : 更新行（新しい値。旧値は *_旧 列で持つ）
      - deleted_rows : 削除候補行
    """
    cur = _normalize_df(current_df, BASIC_COLUMNS)
    new = _normalize_df(new_df, BASIC_COLUMNS)

    cur_ids = set(cur["管理番号"])
    new_ids = set(new["管理番号"])

    new_only_ids = new_ids - cur_ids
    deleted_ids = cur_ids - new_ids
    common_ids = cur_ids & new_ids

    new_rows = new[new["管理番号"].isin(new_only_ids)].copy()
    deleted_rows = cur[cur["管理番号"].isin(deleted_ids)].copy()

    cur_common = cur[cur["管理番号"].isin(common_ids)].set_index("管理番号")
    new_common = new[new["管理番号"].isin(common_ids)].set_index("管理番号")

    changed_ids = []
    for mid in common_ids:
        if not cur_common.loc[mid].equals(new_common.loc[mid]):
            changed_ids.append(mid)

    updated_cur = cur_common.loc[changed_ids].reset_index()
    updated_new = new_common.loc[changed_ids].reset_index()

    updated_rows = updated_new.copy()
    for col in BASIC_COLUMNS:
        if col == "管理番号":
            continue
        updated_rows[f"{col}_旧"] = updated_cur[col].values

    return new_rows, updated_rows, deleted_rows


# ==========================
# 物件マスタへの自動マッピング
# ==========================

def _map_master_from_raw_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    元の DataFrame から物件マスタ用の MASTER_COLUMNS を構成する。
    （管理番号が空の行は除外）
    """
    df = df.copy()
    if not df.empty:
        df = df.astype(str).apply(lambda col: col.str.strip())
    if df.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    n_all = len(df)

    def pick0(*names: str) -> pd.Series:
        for name in names:
            if name in df.columns:
                return df[name]
        return pd.Series([""] * n_all)

    # まず管理番号を見て、空行を除外
    mgmt_all = pick0("管理番号", "物件の管理番号", "物件管理番号", "物件番号")
    mask = mgmt_all.astype(str).str.strip() != ""
    df2 = df[mask].reset_index(drop=True)
    if df2.empty:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    n = len(df2)

    def pick(*names: str) -> pd.Series:
        for name in names:
            if name in df2.columns:
                return df2[name]
        return pd.Series([""] * n)

    # 出力用 DataFrame（全列空で初期化）
    out = pd.DataFrame({col: [""] * n for col in MASTER_COLUMNS})

    # 管理番号
    out["管理番号"] = pick("管理番号", "物件の管理番号", "物件管理番号", "物件番号")

    # 点検実施月 ← 点検月そのまま
    out["点検実施月"] = pick("点検月", "点検実施月")

    # 連絡期限_日前 ＋ 通知期限の原文を備考用に保持
    deadline_series = pick("点検通知先１通知期限", "点検通知先1通知期限")
    days_list: list[str] = []
    notes_from_deadline: list[str] = []
    for v in deadline_series:
        days, note = parse_notice_deadline_to_days(v)
        days_list.append(days)
        notes_from_deadline.append(note)
    out["連絡期限_日前"] = pd.Series(days_list)

    # 通知方法
    method1 = pick("点検通知先１通知方法", "点検通知先1通知方法")
    method2 = pick("点検通知先２通知方法", "点検通知先2通知方法")

    tel1_series = pick("点検通知先１TEL", "点検通知先1TEL")
    tel2_series = pick("点検通知先２TEL", "点検通知先2TEL")
    tel_fallback = pick("TEL")

    fax1_series = pick("点検通知先１FAX", "点検通知先1FAX")
    fax2_series = pick("点検通知先２FAX", "点検通知先2FAX")
    fax_fallback = pick("FAX")

    mail1_series = pick("点検通知先１Email/URL", "点検通知先1Email/URL", "点検通知先１Email", "点検通知先1Email")
    mail2_series = pick("点検通知先２Email/URL", "点検通知先2Email/URL", "点検通知先２Email", "点検通知先2Email")

    window_name = pick("窓口名")
    contract_name = pick("契約先名")
    contact2_name = pick("点検通知先２点検通知先", "点検通知先2点検通知先")

    sticker_type = pick("貼紙貼付書式", "貼紙貼付様式")
    sticker_count = pick("貼紙枚数")

    notes_combined: list[str] = []

    for i in range(n):
        # --- 連絡方法1 ---
        m1_raw = str(method1.iloc[i]) if i < len(method1) else ""
        m1_norm = unicodedata.normalize("NFKC", m1_raw).upper()
        if ("TEL" in m1_norm) or ("電話" in m1_raw):
            out.at[i, "連絡方法_電話1"] = "1"
        if ("FAX" in m1_norm) or ("ＦＡＸ" in m1_raw):
            out.at[i, "連絡方法_FAX1"] = "1"
        if ("MAIL" in m1_norm) or ("ﾒｰﾙ" in m1_raw) or ("メール" in m1_raw):
            out.at[i, "連絡方法_メール1"] = "1"

        # --- 連絡方法2 ---
        m2_raw = str(method2.iloc[i]) if i < len(method2) else ""
        m2_norm = unicodedata.normalize("NFKC", m2_raw).upper()
        if ("TEL" in m2_norm) or ("電話" in m2_raw):
            out.at[i, "連絡方法_電話2"] = "2"
        if ("FAX" in m2_norm) or ("ＦＡＸ" in m2_raw):
            out.at[i, "連絡方法_FAX2"] = "2"
        if ("MAIL" in m2_norm) or ("ﾒｰﾙ" in m2_raw) or ("メール" in m2_raw):
            out.at[i, "連絡方法_メール2"] = "2"

        # --- 電話番号 ---
        tel1 = str(tel1_series.iloc[i]) if i < len(tel1_series) else ""
        tel_fb = str(tel_fallback.iloc[i]) if i < len(tel_fallback) else ""
        tel2 = str(tel2_series.iloc[i]) if i < len(tel2_series) else ""
        out.at[i, "電話番号1"] = tel1 or tel_fb
        out.at[i, "電話番号2"] = tel2

        # --- FAX番号 ---
        fax1 = str(fax1_series.iloc[i]) if i < len(fax1_series) else ""
        fax_fb = str(fax_fallback.iloc[i]) if i < len(fax_fallback) else ""
        fax2 = str(fax2_series.iloc[i]) if i < len(fax2_series) else ""
        out.at[i, "FAX番号1"] = fax1 or fax_fb
        out.at[i, "FAX番号2"] = fax2

        # --- メールアドレス ---
        mail1 = str(mail1_series.iloc[i]) if i < len(mail1_series) else ""
        mail2 = str(mail2_series.iloc[i]) if i < len(mail2_series) else ""
        out.at[i, "メールアドレス1"] = mail1
        out.at[i, "メールアドレス2"] = mail2

        # --- 連絡宛名 ---
        win = str(window_name.iloc[i]) if i < len(window_name) else ""
        con = str(contract_name.iloc[i]) if i < len(contract_name) else ""
        out.at[i, "連絡宛名1"] = win or con

        cn2 = str(contact2_name.iloc[i]) if i < len(contact2_name) else ""
        out.at[i, "連絡宛名2"] = cn2

        # --- 貼り紙テンプレ種別 ---
        stype = str(sticker_type.iloc[i]) if i < len(sticker_type) else ""
        out.at[i, "貼り紙テンプレ種別"] = stype

        # --- 備考 ---
        note_parts = []
        if notes_from_deadline[i]:
            note_parts.append(f"通知期限: {notes_from_deadline[i]}")
        sc = str(sticker_count.iloc[i]) if i < len(sticker_count) else ""
        if sc:
            note_parts.append(f"貼紙枚数: {sc}")
        notes_combined.append(" / ".join([p for p in note_parts if p]))

    if "備考" in out.columns:
        out["備考"] = notes_combined

    return _normalize_df(out, MASTER_COLUMNS)


# ==========================
# マージ処理
# ==========================

def merge_master_and_basic(master_df: pd.DataFrame, basic_df: pd.DataFrame) -> pd.DataFrame:
    """管理番号で物件マスタと基本情報をマージして表示用 DataFrame にする。"""
    master_df = _normalize_df(master_df, MASTER_COLUMNS)
    basic_df = _normalize_df(basic_df, BASIC_COLUMNS)

    if master_df.empty:
        merged = basic_df.copy()
        for col in MASTER_COLUMNS:
            if col not in merged.columns:
                merged[col] = ""
        return merged

    merged = master_df.merge(
        basic_df,
        on="管理番号",
        how="left",
        suffixes=("", "_基本"),
    )

    display_cols = (
        ["管理番号", "物件名", "住所", "窓口会社", "担当部署", "担当者名", "契約種別"]
        + [col for col in MASTER_COLUMNS if col != "管理番号"]
    )
    display_cols = [c for c in display_cols if c in merged.columns]
    return merged[display_cols]


# ==========================
# メイン UI
# ==========================

def render_tab6_property_master(
    sheets_service: Any,
    default_spreadsheet_id: str = "",
    basic_sheet_title: str = "物件基本情報",
    master_sheet_title: str = "物件マスタ",
    current_user_email: Optional[str] = None,
):
    """
    物件マスタ管理タブ
    - 物件基本情報 / 物件マスタ を同一スプレッドシートの別シートとして管理
    - Excel/CSV から基本情報を取り込み、差分プレビュー → シート反映
    - 物件マスタは「インポート時に新規管理番号だけ自動初期化」
    - 物件マスタ用スプレッドシートIDはユーザーごとに Firestore に保存
    """
    st.subheader("物件マスタ管理")

    # ------------------------------
    # Firestore からユーザーごとのIDを読み込み
    # ------------------------------
    db = None
    stored_sheet_id: Optional[str] = None
    if current_user_email:
        try:
            db = firestore.client()
            doc = db.collection("user_settings").document(current_user_email).get()
            if doc.exists:
                stored_sheet_id = (doc.to_dict() or {}).get("property_master_spreadsheet_id") or None
        except Exception as e:
            st.warning(f"物件マスタ用スプレッドシートIDの読み込みに失敗しました: {e}")

    # 初回ロード時：session_state にまだ入ってなければ Firestore or default からセット
    if "pm_spreadsheet_id" not in st.session_state or not st.session_state["pm_spreadsheet_id"]:
        initial_id = stored_sheet_id or default_spreadsheet_id
        if initial_id:
            st.session_state["pm_spreadsheet_id"] = initial_id

    # ------------------------------
    # スプレッドシート設定 & 新規作成
    # ------------------------------
    with st.expander("スプレッドシート設定", expanded=True):
        col1, col2 = st.columns([3, 2])

        # 1) 先に「新規作成ボタン」を処理し、必要なら session_state に ID をセット
        with col2:
            st.write("　")
            if st.button("🆕 新規スプレッドシート作成", use_container_width=True):
                if not sheets_service:
                    st.error("Sheets API のサービスが初期化されていません。")
                else:
                    try:
                        new_id = create_property_master_spreadsheet(
                            sheets_service,
                            user_email=current_user_email,
                        )
                        st.session_state["pm_spreadsheet_id"] = new_id
                        st.success(f"新しいスプレッドシートを作成しました。\nID: {new_id}")

                        # Firestore に保存（ユーザーごと）
                        if db and current_user_email:
                            try:
                                db.collection("user_settings").document(current_user_email).set(
                                    {"property_master_spreadsheet_id": new_id},
                                    merge=True,
                                )
                                st.info("このスプレッドシートIDをユーザー設定に保存しました。次回から自動で読み込まれます。")
                            except Exception as ee:
                                st.warning(f"スプレッドシートIDの保存に失敗しました: {ee}")

                    except Exception as e:
                        st.error(f"スプレッドシートの新規作成に失敗しました: {e}")

        # 2) session_state に入っている値 or default から text_input を表示
        default_id = st.session_state.get("pm_spreadsheet_id", default_spreadsheet_id)
        with col1:
            spreadsheet_id = st.text_input(
                "物件マスタ用スプレッドシートID",
                value=default_id,
                key="pm_spreadsheet_id",
                help="物件基本情報 / 物件マスタ を保存する Google スプレッドシートの ID を入力してください。",
            )

        col3, col4 = st.columns(2)
        with col3:
            basic_title = st.text_input(
                "物件基本情報シート名",
                value=st.session_state.get("pm_basic_sheet_title", basic_sheet_title),
                key="pm_basic_sheet_title",
            )
        with col4:
            master_title = st.text_input(
                "物件マスタシート名",
                value=st.session_state.get("pm_master_sheet_title", master_sheet_title),
                key="pm_master_sheet_title",
            )

        load_btn = st.button("物件マスタ ＋ 基本情報を読み込む", type="primary")

        # 手入力でIDを変更した場合も Firestore に保存
        if db and current_user_email and spreadsheet_id:
            try:
                if stored_sheet_id != spreadsheet_id:
                    db.collection("user_settings").document(current_user_email).set(
                        {"property_master_spreadsheet_id": spreadsheet_id},
                        merge=True,
                    )
            except Exception as e:
                st.warning(f"スプレッドシートIDの保存に失敗しました: {e}")

    # ------------------------------
    # 物件基本情報：Excel/CSV → シート
    # ------------------------------
    with st.expander("物件基本情報（Excel/CSV インポート）", expanded=False):
        st.caption("※ 原本となる Excel/CSV から『物件基本情報』シートを更新します。通常は最初に1回行い、変更があったときのみ再実行します。")

        uploaded_basic = st.file_uploader(
            "物件基本情報ファイル（Excel or CSV）",
            type=["xlsx", "xls", "csv"],
            key="pm_basic_file_upload",
        )

        col_u1, col_u2 = st.columns(2)
        with col_u1:
            preview_diff_btn = st.button("差分をプレビュー", key="pm_preview_diff")
        with col_u2:
            apply_diff_btn = st.button("差分をシートに反映", key="pm_apply_diff")

        # 差分プレビュー
        if preview_diff_btn:
            if not spreadsheet_id:
                st.error("スプレッドシートIDを先に設定してください。")
            elif not sheets_service:
                st.error("Sheets API のサービスが初期化されていません。")
            elif uploaded_basic is None:
                st.error("Excel/CSV ファイルをアップロードしてください。")
            else:
                try:
                    # シートとヘッダーを事前に準備
                    ensure_sheet_and_headers(
                        sheets_service,
                        spreadsheet_id,
                        basic_title,
                        BASIC_COLUMNS,
                    )
                    current_df = load_sheet_as_df(
                        sheets_service,
                        spreadsheet_id,
                        basic_title,
                        BASIC_COLUMNS,
                    )

                    # アップロードファイル → 生のDF → 基本情報DFにマッピング
                    raw_df = load_raw_from_uploaded(uploaded_basic)
                    new_df = _map_basic_from_raw_df(raw_df)

                    new_rows, updated_rows, deleted_rows = diff_basic_info(current_df, new_df)

                    st.session_state["pm_basic_uploaded_raw_df"] = raw_df
                    st.session_state["pm_basic_uploaded_df"] = new_df
                    st.session_state["pm_basic_new_rows"] = new_rows
                    st.session_state["pm_basic_updated_rows"] = updated_rows
                    st.session_state["pm_basic_deleted_rows"] = deleted_rows

                    st.success("差分を計算しました。")
                except Exception as e:
                    st.error(f"差分計算中にエラーが発生しました: {e}")

        # 差分表示
        new_rows = st.session_state.get("pm_basic_new_rows")
        updated_rows = st.session_state.get("pm_basic_updated_rows")
        deleted_rows = st.session_state.get("pm_basic_deleted_rows")

        if isinstance(new_rows, pd.DataFrame):
            st.write(f"✅ 新規追加候補: {len(new_rows)} 件")
            if len(new_rows) > 0:
                st.dataframe(new_rows, use_container_width=True, height=200)

        if isinstance(updated_rows, pd.DataFrame):
            st.write(f"✅ 更新候補: {len(updated_rows)} 件")
            if len(updated_rows) > 0:
                st.dataframe(updated_rows, use_container_width=True, height=200)

        if isinstance(deleted_rows, pd.DataFrame):
            st.write(f"⚠️ 削除候補: {len(deleted_rows)} 件（※反映時は新しいファイルの内容でシート全体を置き換えます）")
            if len(deleted_rows) > 0:
                st.dataframe(deleted_rows, use_container_width=True, height=200)

        # 差分反映（基本情報＋物件マスタ自動初期登録）
        if apply_diff_btn:
            new_df = st.session_state.get("pm_basic_uploaded_df")
            raw_df = st.session_state.get("pm_basic_uploaded_raw_df")

            if not spreadsheet_id:
                st.error("スプレッドシートIDを先に設定してください。")
            elif not sheets_service:
                st.error("Sheets API のサービスが初期化されていません。")
            elif new_df is None or raw_df is None:
                st.error("差分が計算されていません。先に『差分をプレビュー』を実行してください。")
            else:
                try:
                    # --- 物件基本情報シートを新しい内容で全置換 ---
                    ensure_sheet_and_headers(
                        sheets_service,
                        spreadsheet_id,
                        basic_title,
                        BASIC_COLUMNS,
                    )
                    save_df_to_sheet(
                        sheets_service,
                        spreadsheet_id,
                        basic_title,
                        new_df,
                        BASIC_COLUMNS,
                    )
                    st.success("物件基本情報シートを更新しました。（新しいファイルの内容で全行を置き換えています）")

                    # セッション上の基本情報も更新
                    basic_df_norm = _normalize_df(new_df, BASIC_COLUMNS)
                    st.session_state["pm_basic_df"] = basic_df_norm

                    # --- 物件マスタ：新規管理番号だけ自動初期登録 ---
                    ensure_sheet_and_headers(
                        sheets_service,
                        spreadsheet_id,
                        master_title,
                        MASTER_COLUMNS,
                    )
                    current_master_df = load_sheet_as_df(
                        sheets_service,
                        spreadsheet_id,
                        master_title,
                        MASTER_COLUMNS,
                    )

                    candidate_master_df = _map_master_from_raw_df(raw_df)

                    if not candidate_master_df.empty:
                        existing_ids = set(current_master_df["管理番号"].astype(str).str.strip())
                        cand_ids = candidate_master_df["管理番号"].astype(str).str.strip()
                        mask_new = ~cand_ids.isin(existing_ids)
                        new_master_rows = candidate_master_df[mask_new].copy()

                        if not new_master_rows.empty:
                            updated_master_df = pd.concat(
                                [current_master_df, new_master_rows],
                                ignore_index=True,
                            )
                            save_df_to_sheet(
                                sheets_service,
                                spreadsheet_id,
                                master_title,
                                updated_master_df,
                                MASTER_COLUMNS,
                            )
                            st.session_state["pm_master_df"] = updated_master_df
                            st.success(f"物件マスタシートに新規 {len(new_master_rows)} 件を自動登録しました。")
                        else:
                            updated_master_df = current_master_df
                            st.session_state["pm_master_df"] = updated_master_df
                            st.info("物件マスタシートに新規登録する管理番号はありませんでした。")
                    else:
                        updated_master_df = current_master_df
                        st.session_state["pm_master_df"] = updated_master_df
                        st.info("物件マスタ用にマッピングできるデータがありませんでした。")

                    # --- マージ結果も更新しておく ---
                    merged_df_latest = merge_master_and_basic(updated_master_df, basic_df_norm)
                    st.session_state["pm_merged_df"] = merged_df_latest

                except Exception as e:
                    st.error(f"物件基本情報 / 物件マスタ更新中にエラーが発生しました: {e}")

    # ------------------------------
    # 物件マスタ＋基本情報 読み込み（手動）
    # ------------------------------
    if load_btn:
        if not spreadsheet_id:
            st.error("スプレッドシートIDを入力してください。")
        elif not sheets_service:
            st.error("Sheets API のサービスが初期化されていません。")
        else:
            try:
                ensure_sheet_and_headers(
                    sheets_service,
                    spreadsheet_id,
                    basic_title,
                    BASIC_COLUMNS,
                )
                ensure_sheet_and_headers(
                    sheets_service,
                    spreadsheet_id,
                    master_title,
                    MASTER_COLUMNS,
                )

                basic_df = load_sheet_as_df(
                    sheets_service,
                    spreadsheet_id,
                    basic_title,
                    BASIC_COLUMNS,
                )
                master_df = load_sheet_as_df(
                    sheets_service,
                    spreadsheet_id,
                    master_title,
                    MASTER_COLUMNS,
                )

                merged_df = merge_master_and_basic(master_df, basic_df)

                st.session_state["pm_basic_df"] = basic_df
                st.session_state["pm_master_df"] = master_df
                st.session_state["pm_merged_df"] = merged_df
                st.success("物件マスタ ＋ 基本情報を読み込みました。")
            except Exception as e:
                st.error(f"シート読み込み中にエラーが発生しました: {e}")

    merged_df: Optional[pd.DataFrame] = st.session_state.get("pm_merged_df")

    if merged_df is None or merged_df.empty:
        st.info("上部の『物件マスタ ＋ 基本情報を読み込む』ボタン、またはインポート反映を実行してください。")
        return

    # ------------------------------
    # フィルター
    # ------------------------------
    with st.expander("フィルター", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            keyword = st.text_input("キーワード検索（管理番号 / 物件名 / 住所など）", key="pm_keyword")
        with col2:
            only_has_master = st.checkbox(
                "物件マスタに登録がある管理番号のみ表示",
                value=False,
                key="pm_only_has_master",
            )

    df_view = merged_df.copy()

    if keyword:
        kw = keyword.strip()
        mask = pd.Series(False, index=df_view.index)
        for col in ["管理番号", "物件名", "住所", "窓口会社", "担当部署", "担当者名"]:
            if col in df_view.columns:
                mask |= df_view[col].astype(str).str.contains(kw, case=False, na=False)
        df_view = df_view[mask]

    if only_has_master:
        master_cols_for_check = [
            "点検実施月",
            "連絡期限_日前",
            "連絡方法_電話1",
            "連絡方法_電話2",
            "連絡方法_FAX1",
            "連絡方法_FAX2",
            "連絡方法_メール1",
            "連絡方法_メール2",
        ]
        has_any = pd.Series(False, index=df_view.index)
        for col in master_cols_for_check:
            if col in df_view.columns:
                has_any |= df_view[col].astype(str).str.strip() != ""
        df_view = df_view[has_any]

    # 削除用の「選択」列追加
    if "選択" not in df_view.columns:
        df_view.insert(0, "選択", False)

    st.caption("※ 物件基本情報は『物件基本情報』シート、物件マスタは『物件マスタ』シートに保存されます。基本情報を編集したい場合は、Excel/CSV を更新して再インポートしてください。")

    edited_df = st.data_editor(
        df_view,
        num_rows="dynamic",
        key="pm_editor",
        use_container_width=True,
        hide_index=True,
    )

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        if st.button("選択行を削除"):
            if "選択" in edited_df.columns:
                edited_df = edited_df[~edited_df["選択"]].copy()
                st.session_state["pm_merged_df"] = edited_df.drop(columns=["選択"])
                st.success("選択された行を削除しました。（保存ボタンを押すと『物件マスタ』シートに反映されます）")
            else:
                st.warning("選択列が見つかりませんでした。")

    with col_b:
        if st.button("新規行を追加"):
            new_row = {col: "" for col in edited_df.columns}
            new_row["選択"] = False
            edited_df = pd.concat([edited_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["pm_merged_df"] = edited_df.drop(columns=["選択"])
            st.success("空の行を追加しました。（保存ボタンを押すと『物件マスタ』シートに反映されます）")

    with col_c:
        save_btn = st.button("『物件マスタ』シートに保存", type="primary")

    # ------------------------------
    # 物件マスタシートへの保存
    # ------------------------------
    if save_btn:
        if not spreadsheet_id:
            st.error("スプレッドシートIDが未入力です。")
            return
        if not sheets_service:
            st.error("Sheets API のサービスが初期化されていません。")
            return

        save_df = edited_df.drop(columns=["選択"]) if "選択" in edited_df.columns else edited_df.copy()

        # 物件マスタ用の列だけ抽出
        master_only = _normalize_df(save_df, MASTER_COLUMNS)

        # 更新日時・最終更新者
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "更新日時" in master_only.columns:
            master_only["更新日時"] = now_str
        if "最終更新者" in master_only.columns and current_user_email:
            master_only["最終更新者"] = current_user_email

        try:
            ensure_sheet_and_headers(
                sheets_service,
                spreadsheet_id,
                master_title,
                MASTER_COLUMNS,
            )
            save_df_to_sheet(
                sheets_service,
                spreadsheet_id,
                master_title,
                master_only,
                MASTER_COLUMNS,
            )
            st.session_state["pm_master_df"] = master_only

            # 最新の基本情報と再マージ
            basic_df = st.session_state.get("pm_basic_df") or load_sheet_as_df(
                sheets_service,
                spreadsheet_id,
                basic_title,
                BASIC_COLUMNS,
            )
            merged_df_latest = merge_master_and_basic(master_only, basic_df)
            st.session_state["pm_merged_df"] = merged_df_latest

            st.success("『物件マスタ』シートに保存しました。")
        except Exception:
            # エラーは save_df_to_sheet / ensure 内で表示済み
            pass
