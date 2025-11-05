# tabs/tab6_harigami.py
from __future__ import annotations

import io
import os
import re
import zipfile
from datetime import date, datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# 既存ユーティリティ（他タブと同様の依存前提）
from utils.helpers import default_fetch_window_years
from calendar_utils import fetch_all_events  # (service, calendar_id, timeMin, timeMax) を想定
from utils.harigami_generator import (
    PLACEHOLDERS,
    DEFAULT_TEMPLATE_MAP,
    extract_tags_from_description,
    build_replacements_from_event,
    generate_docx_from_template_like,
)

JST = timezone(timedelta(hours=9))


def _to_rfc3339(dt: datetime) -> str:
    # Google Calendar API 用（ISO/RFC3339, UTC）
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=JST)
    return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _daterange_defaults() -> Tuple[date, date]:
    today = datetime.now(JST).date()
    years = default_fetch_window_years() if callable(default_fetch_window_years) else 1
    start = today
    end = today + timedelta(days=30)  # 既定：本日から1か月先
    return start, end


def _safe_summary(event: dict) -> str:
    return (event.get("summary") or "").strip() or "無題イベント"


def _is_all_day(event: dict) -> bool:
    # all-day 判定（必要なら除外）：description 生成前提では許可してもOK
    start = event.get("start", {})
    return "date" in start and "dateTime" not in start


def _extract_work_type(description: str) -> Optional[str]:
    # [作業タイプ: 点検] / [作業タイプ：検査] / 全角括弧・全角コロンにも対応
    if not description:
        return None
    pat = re.compile(
        r"[［\[]\s*作業タイプ\s*[：:]\s*([^\]］\r\n]+?)\s*[］\]]",
        flags=re.IGNORECASE,
    )
    m = pat.search(description)
    if m:
        return m.group(1).strip()
    return None


def _pick_template_for_work_type(
    work_type: Optional[str],
    template_dir: str,
    uploaded_overrides: Dict[str, bytes],
) -> Tuple[str, Optional[io.BytesIO]]:
    """
    テンプレート選択：
    1) アップロードオーバーライドがあれば最優先
    2) /templates 既定マップから選択（work_type がキーに無い場合は default）
    戻り値: (logical_key, file_like or None) ; file_like は BytesIO（アップロード時のみ）
    """
    logical_key = "default"
    if work_type:
        key = work_type.strip()
        if key in uploaded_overrides:
            return key, io.BytesIO(uploaded_overrides[key])
        if key in DEFAULT_TEMPLATE_MAP:
            fs_path = os.path.join(template_dir, DEFAULT_TEMPLATE_MAP[key])
            if os.path.exists(fs_path):
                return key, None
    # fallback: default
    if "default" in uploaded_overrides:
        return "default", io.BytesIO(uploaded_overrides["default"])
    return "default", None


def render_tab6_harigami(service, editable_calendar_options: Dict[str, str]):
    st.subheader("📄 貼紙生成（Googleカレンダー → Word）")

    if not editable_calendar_options:
        st.error("編集可能なカレンダーが見つかりません。認可設定を確認してください。")
        return

    # === 期間選択 ===
    st.markdown("**1. 期間を選択**（本日から1か月先が既定）")
    default_start, default_end = _daterange_defaults()
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("開始日", value=default_start, key="harigami_start_date")
    with col2:
        end_date = st.date_input("終了日", value=default_end, key="harigami_end_date")

    if start_date > end_date:
        st.error("開始日が終了日より後になっています。")
        return

    # === カレンダー選択 ===
    st.markdown("**2. カレンダーを選択**（複数選択可）")
    selected_names = st.multiselect(
        "対象カレンダー",
        options=list(editable_calendar_options.keys()),
        default=list(editable_calendar_options.keys())[:1],
        key="harigami_calendar_names",
    )
    if not selected_names:
        st.info("対象カレンダーを選択してください。")
        return

    # === テンプレート ===
    st.markdown("**3. テンプレート設定**（デフォルト＋任意アップロードの両対応）")
    st.caption("Description 内の `[作業タイプ: 点検]` などでテンプレートを自動切り替えます。未指定は `default.docx`。")

    template_dir = "templates"
    os.makedirs(template_dir, exist_ok=True)

    with st.expander("テンプレートをアップロードで上書き（任意）", expanded=False):
        st.write("以下キー名でアップロードすると既定より優先して使用されます。")
        st.write("利用可能キー：`default`, `点検`, `検査`, `有償工事`, `無償工事`（必要なら増やせます）")

        uploaded_overrides: Dict[str, bytes] = {}
        keys = ["default", "点検", "検査", "有償工事", "無償工事"]
        cols = st.columns(len(keys))
        for i, k in enumerate(keys):
            with cols[i]:
                f = st.file_uploader(f"{k}", type=["docx"], key=f"harigami_tpl_{k}")
                if f is not None:
                    uploaded_overrides[k] = f.read()
        st.session_state["harigami_uploaded_tpl"] = uploaded_overrides

    uploaded_overrides = st.session_state.get("harigami_uploaded_tpl", {})

    # === 実行 ===
    if st.button("4. Word文書を生成（ZIPダウンロード）", type="primary"):
        with st.spinner("カレンダーから予定を取得 → Word 生成中..."):
            start_dt = datetime.combine(start_date, datetime.min.time()).replace(tzinfo=JST)
            # カレンダーの終了は一日終わりの直後まで含む
            end_dt = datetime.combine(end_date + timedelta(days=1), datetime.min.time()).replace(tzinfo=JST)

            all_events: List[dict] = []
            for name in selected_names:
                cal_id = editable_calendar_options[name]
                events = fetch_all_events(
                    service=service,
                    calendar_id=cal_id,
                    timeMin=_to_rfc3339(start_dt),
                    timeMax=_to_rfc3339(end_dt),
                ) or []
                # カレンダー名保持（出力名などで使いたい場合）
                for e in events:
                    e["_calendar_name"] = name
                    e["_calendar_id"] = cal_id
                all_events.extend(events)

            if not all_events:
                st.warning("対象期間にイベントが見つかりませんでした。")
                return

            # 生成対象の抽出：作業タイプを持つイベントを優先（未指定は default で生成）
            # ※要件に合わせて「作業タイプ必須」にするなら以下の条件を変える
            results: List[Tuple[str, bytes]] = []  # (ファイル名, バイナリ)
            progress = st.progress(0.0)
            total = len(all_events)
            done = 0

            for ev in all_events:
                done += 1
                progress.progress(done / total)

                description = ev.get("description") or ""
                summary = _safe_summary(ev)
                work_type = _extract_work_type(description)

                # テンプレート決定
                key, upload_like = _pick_template_for_work_type(work_type, template_dir, uploaded_overrides)
                if upload_like is None:
                    # 既定ファイルパス
                    logical = key if key in DEFAULT_TEMPLATE_MAP else "default"
                    tpl_path = os.path.join(template_dir, DEFAULT_TEMPLATE_MAP[logical])
                else:
                    tpl_path = upload_like  # BytesIO

                # Descriptionから各タグ抽出（作業指示書 / 管理番号等）
                tags = extract_tags_from_description(description)

                # 置換データ生成
                try:
                    replacements = build_replacements_from_event(ev, summary=summary, tags=tags)
                except Exception:
                    # 不正データなどはスキップ
                    continue

                # DOCX生成
                try:
                    out_name, out_bytes = generate_docx_from_template_like(
                        template_like=tpl_path,
                        replacements=replacements,
                        safe_title=summary,
                    )
                    results.append((out_name, out_bytes))
                except Exception:
                    continue

            if not results:
                st.warning("生成対象イベントがありません（置換用データ不足/テンプレ選択不可など）。")
                return

            # ZIP化
            mem_zip = io.BytesIO()
            with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, blob in results:
                    zf.writestr(fname, blob)
            mem_zip.seek(0)

            st.success(f"🎉 貼紙（Word）生成完了：{len(results)}件")
            st.download_button(
                "ZIPをダウンロード",
                data=mem_zip.getvalue(),
                file_name="harigami_documents.zip",
                mime="application/zip",
            )

    st.caption("※ Descriptionの `[作業タイプ: 〇〇]` `[作業指示書: 1234567]` `[管理番号: A1234]` に対応。テンプレが無い作業タイプは default を使用。")
