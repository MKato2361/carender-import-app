# tabs/tab6_property_master.py
from __future__ import annotations

from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
from googleapiclient.discovery import Resource
from firebase_admin import firestore


# ==========================
# 定数（列定義）
# ==========================

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
# ヘルパー（基本情報：アップロードファイル → DataFrame）
# ==========================

def load_basic_info_from_uploaded(uploaded_file) -> pd.DataFrame:
    """
    アップロードされた Excel/CSV から物件基本情報 DataFrame を作成。
    初回インポート / 差分更新用。
    """
    if uploaded_file is None:
        return pd.DataFrame(columns=BASIC_COLUMNS)

    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, dtype=str)
    else:
        # CSV想定。必要に応じて encoding 変更（例: encoding="cp932"）
        df = pd.read_csv(uploaded_file, dtype=str)

    df = df.astype(str).apply(lambda col: col.str.strip())

    for col in BASIC_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[BASIC_COLUMNS].copy()
    return df


# ==========================
# ヘルパー（基本情報：Firestore 連携）
# ==========================

def load_basic_info_from_firestore(user_id: str) -> pd.DataFrame:
    """
    Firestore から、指定ユーザーの物件基本情報を DataFrame で取得。
    コレクション: property_basic_info
    ドキュメントID例: {user_id}_{管理番号}
    """
    if not user_id:
        return pd.DataFrame(columns=BASIC_COLUMNS)

    db = firestore.client()

    docs = (
        db.collection("property_basic_info")
        .where("user_id", "==", user_id)
        .stream()
    )

    rows = []
    for doc in docs:
        data = doc.to_dict() or {}
        row = {col: data.get(col, "") for col in BASIC_COLUMNS}
        rows.append(row)

    if not rows:
        return pd.DataFrame(columns=BASIC_COLUMNS)

    df = pd.DataFrame(rows)
    df = df.astype(str).apply(lambda col: col.str.strip())
    return df[BASIC_COLUMNS].copy()


def apply_basic_info_diff_to_firestore(
    user_id: str,
    new_rows: pd.DataFrame,
    updated_rows: pd.DataFrame,
    deleted_rows: pd.DataFrame,
    do_delete: bool = False,
):
    """
    物件基本情報の差分を Firestore に反映する。
    - new_rows     : 新規追加行（BASIC_COLUMNS 構成）
    - updated_rows : 更新行（BASIC_COLUMNS + *_旧 の列を持つが、保存時は BASIC_COLUMNS を使う）
    - deleted_rows : 削除候補行
    """
    if not user_id:
        st.error("ユーザーIDが未設定のため、Firestore への反映ができません。")
        return

    db = firestore.client()
    col_ref = db.collection("property_basic_info")
    batch = db.batch()

    # 新規追加
    for _, row in new_rows.iterrows():
        mid = str(row.get("管理番号", "")).strip()
        if not mid:
            continue
        doc_id = f"{user_id}_{mid}"
        doc_ref = col_ref.document(doc_id)
        data = {col: str(row.get(col, "") or "").strip() for col in BASIC_COLUMNS}
        data["user_id"] = user_id
        data["updated_at"] = firestore.SERVER_TIMESTAMP
        batch.set(doc_ref, data, merge=True)

    # 更新
    for _, row in updated_rows.iterrows():
        mid = str(row.get("管理番号", "")).strip()
        if not mid:
            continue
        doc_id = f"{user_id}_{mid}"
        doc_ref = col_ref.document(doc_id)
        data = {col: str(row.get(col, "") or "").strip() for col in BASIC_COLUMNS}
        data["user_id"] = user_id
        data["updated_at"] = firestore.SERVER_TIMESTAMP
        batch.set(doc_ref, data, merge=True)

    # 削除（ユーザーが許可した場合のみ）
    if do_delete:
        for _, row in deleted_rows.iterrows():
            mid = str(row.get("管理番号", "")).strip()
            if not mid:
                continue
            doc_id = f"{user_id}_{mid}"
            doc_ref = col_ref.document(doc_id)
            batch.delete(doc_ref)

    batch.commit()


def diff_basic_info(current_df: pd.DataFrame, new_df: pd.DataFrame):
    """
    current_df: Firestore から取得した現状の基本情報
    new_df    : 新しくアップロードされた Excel/CSV を読み込んだ DataFrame

    戻り値:
      - new_rows     : 新規追加行
      - updated_rows : 更新行（新しい値が入った DataFrame。旧値は *_旧 列に持つ）
      - deleted_rows : 削除候補行
    """
    def _normalize(df):
        df = df.copy()
        for col in BASIC_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[BASIC_COLUMNS].copy()
        df = df.astype(str).apply(lambda col: col.str.strip())
        return df

    cur = _normalize(current_df)
    new = _normalize(new_df)

    cur_ids = set(cur["管理番号"])
    new_ids = set(new["管理番号"])

    # 新規追加
    new_only_ids = new_ids - cur_ids
    new_rows = new[new["管理番号"].isin(new_only_ids)].copy()

    # 削除候補
    deleted_ids = cur_ids - new_ids
    deleted_rows = cur[cur["管理番号"].isin(deleted_ids)].copy()

    # 更新候補（IDは共通だが、内容が違うもの）
    common_ids = cur_ids & new_ids
    cur_common = cur[cur["管理番号"].isin(common_ids)].set_index("管理番号")
    new_common = new[new["管理番号"].isin(common_ids)].set_index("管理番号")

    changed_ids = []
    for mid in common_ids:
        if not cur_common.loc[mid].equals(new_common.loc[mid]):
            changed_ids.append(mid)

    updated_cur = cur_common.loc[changed_ids].reset_index()
    updated_new = new_common.loc[changed_ids].reset_index()

    updated_rows = updated_new.copy()
    # 旧値を *_旧 列として持たせる
    for col in BASIC_COLUMNS:
        if col == "管理番号":
            continue
        updated_rows[f"{col}_旧"] = updated_cur[col].values

    return new_rows, updated_rows, deleted_rows


# ==========================
# ヘルパー（物件マスタ：Sheets 連携）
# ==========================

def load_master_from_sheets(
    sheets_service: Resource,
    spreadsheet_id: str,
    sheet_name: str = "物件マスタ",
) -> pd.DataFrame:
    """
    Google スプレッドシートから物件マスタを読み込んで DataFrame にする。
    """
    range_name = f"{sheet_name}!A1:AE"  # 列数に応じて十分広めに

    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name,
        ).execute()
    except Exception as e:
        st.error(f"物件マスタの読み込みに失敗しました: {e}")
        return pd.DataFrame(columns=MASTER_COLUMNS)

    values = result.get("values", [])
    if not values:
        return pd.DataFrame(columns=MASTER_COLUMNS)

    header = values[0]
    rows = values[1:]

    df = pd.DataFrame(rows, columns=header)
    df = df.astype(str).apply(lambda col: col.str.strip())

    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[MASTER_COLUMNS].copy()


def save_master_to_sheets(
    sheets_service: Resource,
    spreadsheet_id: str,
    sheet_name: str,
    df: pd.DataFrame,
) -> None:
    """
    DataFrame の内容をヘッダー込みでシート全体に書き戻す。
    既存の内容は一度クリアしてから上書き。
    """
    df_to_save = df.copy()
    df_to_save = df_to_save.fillna("").astype(str)

    values = [list(df_to_save.columns)] + df_to_save.values.tolist()

    range_all = sheet_name
    range_start = f"{sheet_name}!A1"

    try:
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=range_all,
        ).execute()

        body = {"values": values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_start,
            valueInputOption="RAW",
            body=body,
        ).execute()
    except Exception as e:
        st.error(f"物件マスタの保存に失敗しました: {e}")
        raise


def merge_master_and_basic(
    master_df: pd.DataFrame,
    basic_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    管理番号をキーに、物件マスタ（通知・条件）と物件基本情報DBを左結合する。
    物件マスタ側がベース。
    """
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
    sheets_service: Resource,
    default_spreadsheet_id: str = "",
    default_sheet_name: str = "物件マスタ",
    current_user_email: Optional[str] = None,
    current_user_id: Optional[str] = None,
):
    """
    物件マスタ管理タブの描画。

    - 物件マスタ（Google Sheets）はそのまま
    - 物件基本情報は Firestore に保存し、普段は Firestore から読み取り
    - Excel/CSV アップロードで差分を計算し、Firestore を更新可能
    """
    st.subheader("物件マスタ管理")

    # ---- 基本設定エリア ----
    with st.expander("物件マスタ（スプレッドシート）設定", expanded=True):
        col1, col2 = st.columns([2, 1])
        with col1:
            spreadsheet_id = st.text_input(
                "物件マスタ用スプレッドシートID",
                value=st.session_state.get("pm_spreadsheet_id", default_spreadsheet_id),
                key="pm_spreadsheet_id",
                help="物件マスタを保存する Google スプレッドシートの ID を入力してください。",
            )
        with col2:
            sheet_name = st.text_input(
                "シート名",
                value=st.session_state.get("pm_sheet_name", default_sheet_name),
                key="pm_sheet_name",
            )

        load_btn = st.button("物件マスタ＋基本情報を読み込む", type="primary")

    # ---- 基本情報DBの更新（Excel/CSV → Firestore 差分更新） ----
    with st.expander("物件基本情報DBの更新（Excel/CSV → Firestore）", expanded=False):
        st.caption("※ 初回はここから Excel/CSV をアップロードして Firestore に登録します。以降は更新があるときだけでOKです。")

        uploaded_basic = st.file_uploader(
            "物件基本情報ファイル（Excel or CSV）をアップロード",
            type=["xlsx", "xls", "csv"],
            key="pm_basic_file_upload",
            help="管理番号・物件名・住所などを持つ基本情報の原本ファイルです。",
        )

        col_u1, col_u2 = st.columns(2)
        with col_u1:
            preview_diff_btn = st.button("差分をプレビュー", key="pm_preview_diff")
        with col_u2:
            do_delete = st.checkbox(
                "新ファイルに存在しない管理番号は Firestore から削除する",
                value=False,
                key="pm_do_delete",
                help="チェックした場合、旧データのみ存在する管理番号は削除されます。",
            )

        if preview_diff_btn:
            if not current_user_id:
                st.error("ユーザーIDが未設定のため、差分計算ができません。")
            elif uploaded_basic is None:
                st.error("Excel/CSV ファイルをアップロードしてください。")
            else:
                with st.spinner("Firestore から現在の基本情報を読み込み中..."):
                    current_df = load_basic_info_from_firestore(current_user_id)
                with st.spinner("アップロードファイルを読み込み中..."):
                    new_df = load_basic_info_from_uploaded(uploaded_basic)

                new_rows, updated_rows, deleted_rows = diff_basic_info(current_df, new_df)

                st.session_state["pm_basic_new_rows"] = new_rows
                st.session_state["pm_basic_updated_rows"] = updated_rows
                st.session_state["pm_basic_deleted_rows"] = deleted_rows

                st.success("差分を計算しました。")

        # 差分結果の表示
        new_rows = st.session_state.get("pm_basic_new_rows")
        updated_rows = st.session_state.get("pm_basic_updated_rows")
        deleted_rows = st.session_state.get("pm_basic_deleted_rows")

        if isinstance(new_rows, pd.DataFrame):
            st.write(f"✅ 新規追加候補: {len(new_rows)} 件")
            if len(new_rows) > 0:
                st.dataframe(new_rows, use_container_width=True)

        if isinstance(updated_rows, pd.DataFrame):
            st.write(f"✅ 更新候補: {len(updated_rows)} 件")
            if len(updated_rows) > 0:
                st.dataframe(updated_rows, use_container_width=True)

        if isinstance(deleted_rows, pd.DataFrame):
            st.write(f"⚠️ 削除候補: {len(deleted_rows)} 件")
            if len(deleted_rows) > 0:
                st.dataframe(deleted_rows, use_container_width=True)

        apply_diff_btn = st.button("差分を Firestore に反映", key="pm_apply_diff", type="secondary")

        if apply_diff_btn:
            if not current_user_id:
                st.error("ユーザーIDが未設定のため、Firestore への反映ができません。")
            elif new_rows is None or updated_rows is None or deleted_rows is None:
                st.error("差分が計算されていません。先に『差分をプレビュー』を実行してください。")
            else:
                try:
                    with st.spinner("Firestore に差分を反映中..."):
                        apply_basic_info_diff_to_firestore(
                            user_id=current_user_id,
                            new_rows=new_rows,
                            updated_rows=updated_rows,
                            deleted_rows=deleted_rows,
                            do_delete=do_delete,
                        )
                    st.success("Firestore に差分を反映しました。")

                    # 反映後は Firestore から再読み込みしてセッションを更新しておく
                    refreshed_basic = load_basic_info_from_firestore(current_user_id)
                    st.session_state["pm_basic_df"] = refreshed_basic

                    # 差分はクリア
                    st.session_state["pm_basic_new_rows"] = None
                    st.session_state["pm_basic_updated_rows"] = None
                    st.session_state["pm_basic_deleted_rows"] = None
                except Exception as e:
                    st.error(f"Firestore への反映中にエラーが発生しました: {e}")

    # ---- 物件マスタ＋基本情報の読み込み ----
    if load_btn:
        if not spreadsheet_id:
            st.error("スプレッドシートIDを入力してください。")
        elif sheets_service is None:
            st.error("Sheets API の service が初期化されていません。")
        elif not current_user_id:
            st.error("ユーザーIDが未設定のため、物件基本情報が読み込めません。")
        else:
            with st.spinner("物件マスタを読み込み中..."):
                master_df = load_master_from_sheets(sheets_service, spreadsheet_id, sheet_name)
            with st.spinner("物件基本情報を Firestore から読み込み中..."):
                basic_df = load_basic_info_from_firestore(current_user_id)

            merged_df = merge_master_and_basic(master_df, basic_df)

            st.session_state["pm_master_df"] = master_df
            st.session_state["pm_basic_df"] = basic_df
            st.session_state["pm_merged_df"] = merged_df
            st.success("物件マスタ＋基本情報を読み込みました。")

    merged_df: Optional[pd.DataFrame] = st.session_state.get("pm_merged_df")

    if merged_df is None or merged_df.empty:
        st.info("上部の『物件マスタ＋基本情報を読み込む』ボタンからデータを読み込んでください。")
        return

    # ---- フィルターエリア ----
    with st.expander("フィルター", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            keyword = st.text_input("キーワード検索（管理番号 / 物件名 / 住所など）", key="pm_keyword")
        with col2:
            only_has_master = st.checkbox(
                "物件マスタに登録がある管理番号のみ表示",
                value=False,
                key="pm_only_has_master",
                help="点検実施月や連絡方法が設定されているものだけを表示します。",
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

    # 選択列を追加（削除用）
    if "選択" not in df_view.columns:
        df_view.insert(0, "選択", False)

    st.caption("※ 物件基本情報（物件名・住所など）は Firestore に保存されています。編集したい場合は、原本ファイルを更新して『物件基本情報DBの更新』から差分反映してください。")

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
                st.success("選択された行を削除しました。（保存ボタンを押すとスプレッドシートに反映されます）")
            else:
                st.warning("選択列が見つかりませんでした。")

    with col_b:
        if st.button("新規行を追加"):
            new_row = {col: "" for col in edited_df.columns}
            new_row["選択"] = False
            edited_df = pd.concat([edited_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["pm_merged_df"] = edited_df.drop(columns=["選択"])
            st.success("空の行を追加しました。（保存ボタンを押すとスプレッドシートに反映されます）")

    with col_c:
        save_btn = st.button("スプレッドシートに保存", type="primary")

    if save_btn:
        if not spreadsheet_id:
            st.error("スプレッドシートIDが未入力です。")
            return

        save_df = edited_df.drop(columns=["選択"]) if "選択" in edited_df.columns else edited_df.copy()

        for col in MASTER_COLUMNS:
            if col not in save_df.columns:
                save_df[col] = ""

        save_df = save_df[MASTER_COLUMNS].copy()

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "更新日時" in save_df.columns:
            save_df["更新日時"] = now_str
        if "最終更新者" in save_df.columns and current_user_email:
            save_df["最終更新者"] = current_user_email

        try:
            with st.spinner("物件マスタをスプレッドシートに保存中..."):
                save_master_to_sheets(sheets_service, spreadsheet_id, sheet_name, save_df)

            st.session_state["pm_master_df"] = save_df

            # 最新の基本情報（Firestore）と再マージ
            basic_df = st.session_state.get("pm_basic_df") or (
                load_basic_info_from_firestore(current_user_id) if current_user_id else pd.DataFrame(columns=BASIC_COLUMNS)
            )
            merged_df_latest = merge_master_and_basic(save_df, basic_df)
            st.session_state["pm_merged_df"] = merged_df_latest

            st.success("物件マスタをスプレッドシートに保存しました。")
        except Exception:
            # save_master_to_sheets 内でエラー表示済み
            pass
