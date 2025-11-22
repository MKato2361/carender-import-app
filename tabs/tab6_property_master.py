# tabs/tab6_property_master.py
from __future__ import annotations

from datetime import datetime
from typing import Optional, Any
from io import BytesIO

import pandas as pd
import streamlit as st


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

    df = pd.DataFrame(rows, columns=header)
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

def load_basic_info_from_uploaded(uploaded_file) -> pd.DataFrame:
    """
    アップロードされた Excel/CSV から物件基本情報 DataFrame を作成。
    - Excel: そのまま read_excel
    - CSV : まず UTF-8 / UTF-8-SIG を試し、ダメなら CP932(Shift_JIS) で再トライ
    """
    if uploaded_file is None:
        return pd.DataFrame(columns=BASIC_COLUMNS)

    name = uploaded_file.name.lower()

    # --- Excel の場合 ---
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, dtype=str)
        df = df.astype(str).apply(lambda col: col.str.strip())
        return _normalize_df(df, BASIC_COLUMNS)

    # --- CSV の場合 ---
    # 一度バイト列として読み込み、複数エンコーディングでトライする
    raw_bytes = uploaded_file.read()

    # 以降、この関数の中だけで raw_bytes を使い切る前提
    encodings_to_try = ["utf-8", "utf-8-sig", "cp932"]

    last_err: Optional[Exception] = None
    for enc in encodings_to_try:
        try:
            df = pd.read_csv(BytesIO(raw_bytes), dtype=str, encoding=enc)
            df = df.astype(str).apply(lambda col: col.str.strip())
            return _normalize_df(df, BASIC_COLUMNS)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            continue

    st.error(f"CSVファイルの読み込みに失敗しました。エンコーディングを確認してください。（最後のエラー: {last_err}）")
    return pd.DataFrame(columns=BASIC_COLUMNS)


def diff_basic_info(current_df: pd.DataFrame, new_df: pd.DataFrame):
    """
    current_df: 現在シートに入っている基本情報
    new_df    : 新しくアップロードされた Excel/CSV を読み込んだ基本情報

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
    - 物件マスタは Data Editor で編集 → シート保存
    """
    st.subheader("物件マスタ管理")

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
                        st.info("必要であれば、このIDを secrets.toml の PROPERTY_MASTER_SHEET_ID に保存してください。")
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
                    new_df = load_basic_info_from_uploaded(uploaded_basic)

                    new_rows, updated_rows, deleted_rows = diff_basic_info(current_df, new_df)

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

        # 差分反映（実際には「新しいファイルの内容でシート全体を置き換え」）
        if apply_diff_btn:
            new_df = st.session_state.get("pm_basic_uploaded_df")
            if not spreadsheet_id:
                st.error("スプレッドシートIDを先に設定してください。")
            elif not sheets_service:
                st.error("Sheets API のサービスが初期化されていません。")
            elif new_df is None:
                st.error("差分が計算されていません。先に『差分をプレビュー』を実行してください。")
            else:
                try:
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
                    st.session_state["pm_basic_df"] = _normalize_df(new_df, BASIC_COLUMNS)
                except Exception as e:
                    st.error(f"物件基本情報シートの更新中にエラーが発生しました: {e}")

    # ------------------------------
    # 物件マスタ＋基本情報 読み込み
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
        st.info("上部の『物件マスタ ＋ 基本情報を読み込む』ボタンからデータを読み込んでください。")
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
