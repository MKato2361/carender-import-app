# =====================================================================
# tab2_register.py の修正箇所 2点
# =====================================================================

# -------------------------------------------------------
# 修正① _save_description_settings（並び順も一緒に保存）
# -------------------------------------------------------
def _save_description_settings(user_id: str):
    """説明欄設定（列の選択＋並び順）を保存するコールバック"""
    desc_key       = f"description_selector_register_{user_id}"
    desc_order_key = f"description_order_register_{user_id}"

    if desc_key not in st.session_state:
        return

    description_columns_pool = st.session_state.get("description_columns_pool", [])
    selected = st.session_state[desc_key]
    valid_selected = [col for col in selected if col in description_columns_pool]

    # 並び順キャッシュを選択済み列に合わせて同期
    current_order = st.session_state.get(desc_order_key, [])
    current_order = [c for c in current_order if c in valid_selected]
    for c in valid_selected:
        if c not in current_order:
            current_order.append(c)
    st.session_state[desc_order_key] = current_order

    # 選択列と並び順をまとめて Firestore へ保存
    set_user_setting(user_id, "description_columns_selected", current_order)
    st.toast("✅ 説明欄の設定を保存しました", icon="💾")


# -------------------------------------------------------
# 修正② expander 内の説明欄UI（元のコードと置き換え）
# -------------------------------------------------------
# ▼ 置き換え前（元コード）
# -----------
#             description_columns_pool = st.session_state.get("description_columns_pool", [])
#             saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
#             default_selection = [col for col in saved_description_cols if col in description_columns_pool]
#
#             desc_key = f"description_selector_register_{user_id}"
#
#             if desc_key not in st.session_state:
#                 st.session_state[desc_key] = list(default_selection)
#             else:
#                 st.session_state[desc_key] = [c for c in st.session_state[desc_key] if c in description_columns_pool]
#
#             description_columns = st.multiselect(
#                 "説明欄に含める列（複数選択可）",
#                 description_columns_pool,
#                 key=desc_key,
#                 on_change=_save_description_settings,
#                 args=(user_id,),
#             )
#             description_columns = st.session_state.get(desc_key, [])
# -----------

# ▼ 置き換え後（↓ これをそのまま元のコードと差し替え）
# -----------
            description_columns_pool = st.session_state.get("description_columns_pool", [])
            saved_description_cols = get_user_setting(user_id, "description_columns_selected") or []
            default_selection = [col for col in saved_description_cols if col in description_columns_pool]

            desc_key       = f"description_selector_register_{user_id}"
            desc_order_key = f"description_order_register_{user_id}"

            # multiselect の初期化
            if desc_key not in st.session_state:
                st.session_state[desc_key] = list(default_selection)
            else:
                st.session_state[desc_key] = [c for c in st.session_state[desc_key] if c in description_columns_pool]

            st.multiselect(
                "説明欄に含める列（複数選択可）",
                description_columns_pool,
                key=desc_key,
                on_change=_save_description_settings,
                args=(user_id,),
            )
            description_columns_selected = st.session_state.get(desc_key, [])

            # 並び替えテーブル（選択済みの列がある場合のみ表示）
            if description_columns_selected:
                st.caption("↕️ ドラッグして説明欄の列の順番を変更できます")

                # 並び順の初期化（保存済み順序 → multiselect順 にフォールバック）
                current_order = st.session_state.get(desc_order_key, [])
                current_order = [c for c in current_order if c in description_columns_selected]
                for c in description_columns_selected:
                    if c not in current_order:
                        current_order.append(c)

                order_df = pd.DataFrame({"列名（説明欄への出力順）": current_order})

                edited_order_df = st.data_editor(
                    order_df,
                    num_rows="fixed",
                    hide_index=False,
                    use_container_width=True,
                    column_config={
                        "列名（説明欄への出力順）": st.column_config.SelectboxColumn(
                            "列名（説明欄への出力順）",
                            options=description_columns_selected,
                            required=True,
                        )
                    },
                    key=f"{desc_order_key}_editor",
                )

                # 編集後の順序を取得（重複・未選択列を除外して確定）
                new_order = edited_order_df["列名（説明欄への出力順）"].tolist()
                new_order = list(dict.fromkeys([c for c in new_order if c in description_columns_selected]))
                st.session_state[desc_order_key] = new_order
                description_columns = new_order
            else:
                st.session_state.pop(desc_order_key, None)
                description_columns = []
# -----------
