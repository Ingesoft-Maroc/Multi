import streamlit as st
import pandas as pd
import io


# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

def load_excel(uploaded_file) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df.columns = df.columns.astype(str)
        return df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        return None


def reset_reconciler():
    keys = [k for k in st.session_state if k.startswith("rec_")]
    for k in keys:
        del st.session_state[k]


def step_badge(n: int, label: str, active: bool):
    color = "#6366f1" if active else "#374151"
    txt_color = "#fff" if active else "#9ca3af"
    st.markdown(
        f"""<span style="background:{color};color:{txt_color};
        border-radius:20px;padding:3px 14px;font-size:0.85rem;
        font-weight:600;margin-right:6px;">Étape {n} — {label}</span>""",
        unsafe_allow_html=True,
    )


def find_matching_rows(df_b: pd.DataFrame, key_b: str, val_a: str,
                       match_mode: str, case_sensitive: bool) -> pd.DataFrame:
    """Return all rows in df_b whose key_b matches val_a according to match rules."""
    col = df_b[key_b].astype(str)
    if match_mode == "parfait":
        if case_sensitive:
            mask = col == val_a
        else:
            mask = col.str.lower() == val_a.lower()
    else:  # normal = contains
        if case_sensitive:
            mask = col.str.contains(val_a, regex=False, na=False)
        else:
            mask = col.str.contains(val_a, case=False, regex=False, na=False)
    return df_b[mask]


def build_concat_value(matched_rows: pd.DataFrame, cols: list[str],
                       col_sep: str, row_sep: str) -> str:
    """For all matched rows, build a concatenated string of selected columns."""
    row_parts = []
    for _, row in matched_rows.iterrows():
        cell_parts = []
        for c in cols:
            v = str(row.get(c, ""))
            if v not in ("nan", "None", ""):
                cell_parts.append(v)
        if cell_parts:
            row_parts.append(col_sep.join(cell_parts))
    return row_sep.join(row_parts)


# ─────────────────────────────────────────────
#  Main entry point
# ─────────────────────────────────────────────

def run_reconciler():
    st.markdown("## 🔗 Réconciliateur Excel")
    st.markdown(
        "Importez deux fichiers Excel, définissez la clé et le mode de rapprochement, "
        "puis configurez vos colonnes de sortie."
    )
    st.markdown("---")

    step = st.session_state.get("rec_step", 1)

    col_steps = st.columns(4)
    labels = ["Import", "Clé & Mode", "Colonnes de sortie", "Résultat"]
    for i, (c, lbl) in enumerate(zip(col_steps, labels), start=1):
        with c:
            step_badge(i, lbl, active=(i == step))

    st.markdown("<br>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    #  ÉTAPE 1 — Import
    # ══════════════════════════════════════════
    if step == 1:
        st.markdown("### 📂 Étape 1 — Importez vos deux fichiers")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Fichier de base (référence)**")
            file_a = st.file_uploader(
                "Fichier A", type=["xlsx", "xls"], key="rec_upload_a",
                label_visibility="collapsed"
            )
        with col_b:
            st.markdown("**Fichier à rapprocher**")
            file_b = st.file_uploader(
                "Fichier B", type=["xlsx", "xls"], key="rec_upload_b",
                label_visibility="collapsed"
            )

        if file_a and file_b:
            df_a = load_excel(file_a)
            df_b = load_excel(file_b)

            if df_a is not None and df_b is not None:
                st.session_state["rec_df_a"] = df_a
                st.session_state["rec_df_b"] = df_b
                st.session_state["rec_name_a"] = file_a.name
                st.session_state["rec_name_b"] = file_b.name

                col_a, col_b = st.columns(2)
                with col_a:
                    st.success(f"✅ {file_a.name}")
                    st.caption(f"{len(df_a)} lignes · {len(df_a.columns)} colonnes")
                    with st.expander("Aperçu"):
                        st.dataframe(df_a.head(5), use_container_width=True)
                with col_b:
                    st.success(f"✅ {file_b.name}")
                    st.caption(f"{len(df_b)} lignes · {len(df_b.columns)} colonnes")
                    with st.expander("Aperçu"):
                        st.dataframe(df_b.head(5), use_container_width=True)

                st.markdown("")
                if st.button("Suivant →", type="primary"):
                    st.session_state["rec_step"] = 2
                    st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 2 — Clé & Mode de rapprochement
    # ══════════════════════════════════════════
    elif step == 2:
        df_a: pd.DataFrame = st.session_state["rec_df_a"]
        df_b: pd.DataFrame = st.session_state["rec_df_b"]
        name_a = st.session_state["rec_name_a"]
        name_b = st.session_state["rec_name_b"]

        st.markdown("### 🔑 Étape 2 — Clé de rapprochement & mode de match")

        # ── Key columns ───────────────────────
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**Colonne clé dans {name_a} (base)**")
            key_a = st.selectbox(
                "Clé A", options=list(df_a.columns),
                key="rec_key_a", label_visibility="collapsed"
            )
        with col_b:
            st.markdown(f"**Colonne clé dans {name_b} (à rapprocher)**")
            key_b = st.selectbox(
                "Clé B", options=list(df_b.columns),
                key="rec_key_b", label_visibility="collapsed"
            )

        if key_a and key_b:
            col_a, col_b = st.columns(2)
            with col_a:
                sample_a = df_a[key_a].dropna().astype(str).unique()[:6]
                st.caption(f"Exemples : {', '.join(sample_a)}")
            with col_b:
                sample_b = df_b[key_b].dropna().astype(str).unique()[:6]
                st.caption(f"Exemples : {', '.join(sample_b)}")

        st.markdown("---")

        # ── Match mode ────────────────────────
        st.markdown("**Mode de rapprochement**")

        match_mode = st.radio(
            "Mode",
            options=["Match parfait", "Match normal (contient)"],
            key="rec_match_mode_radio",
            horizontal=True,
            label_visibility="collapsed",
            help=(
                "**Match parfait** : la valeur clé doit être identique.\n\n"
                "**Match normal** : la valeur du fichier B doit *contenir* la valeur clé du fichier A. "
                "Ex : clé = 'Xavier' → matche 'il est là Xavier'."
            ),
        )

        case_opt = st.radio(
            "Casse",
            options=["Ignorer la casse (XAVier = xavier)", "Respecter la casse (XAVier ≠ xavier)"],
            key="rec_case_radio",
            horizontal=True,
            label_visibility="collapsed",
        )
        case_sensitive = "Respecter" in case_opt

        # ── Live preview ──────────────────────
        if key_a and key_b:
            mm = "parfait" if "parfait" in match_mode else "normal"
            sample_key = str(df_a[key_a].dropna().iloc[0]) if len(df_a[key_a].dropna()) > 0 else ""
            if sample_key:
                matched_preview = find_matching_rows(df_b, key_b, sample_key, mm, case_sensitive)
                st.info(
                    f"🔍 Exemple avec **\"{sample_key}\"** : "
                    f"**{len(matched_preview)}** ligne(s) trouvée(s) dans {name_b}"
                )

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["rec_step"] = 1
                st.rerun()
        with c2:
            if st.button("Suivant →", type="primary"):
                st.session_state["rec_key_a_val"] = key_a
                st.session_state["rec_key_b_val"] = key_b
                st.session_state["rec_match_mode"] = "parfait" if "parfait" in match_mode else "normal"
                st.session_state["rec_case_sensitive"] = case_sensitive
                if "rec_special_cols" not in st.session_state:
                    st.session_state["rec_special_cols"] = []
                if "rec_output_cols_a" not in st.session_state:
                    st.session_state["rec_output_cols_a"] = list(df_a.columns)
                st.session_state["rec_step"] = 3
                st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 3 — Colonnes de sortie
    # ══════════════════════════════════════════
    elif step == 3:
        df_a: pd.DataFrame = st.session_state["rec_df_a"]
        df_b: pd.DataFrame = st.session_state["rec_df_b"]
        name_a = st.session_state["rec_name_a"]
        name_b = st.session_state["rec_name_b"]
        match_mode = st.session_state["rec_match_mode"]
        case_sensitive = st.session_state["rec_case_sensitive"]

        st.markdown("### 🗂️ Étape 3 — Colonnes de sortie")

        mode_lbl = "parfait" if match_mode == "parfait" else "normal (contient)"
        casse_lbl = "casse respectée" if case_sensitive else "casse ignorée"
        st.caption(f"Mode : **{mode_lbl}** · {casse_lbl}")

        # ── Colonnes du fichier A ─────────────
        st.markdown(f"#### 📄 Colonnes de **{name_a}** à afficher")
        all_cols_a = list(df_a.columns)
        selected_a = st.multiselect(
            "Colonnes A",
            options=all_cols_a,
            default=st.session_state.get("rec_output_cols_a", all_cols_a),
            key="rec_ms_cols_a",
            label_visibility="collapsed",
        )

        st.markdown("---")

        # ── Colonnes spéciales (concaténation) ─
        st.markdown(f"#### 🔧 Colonnes de sortie depuis **{name_b}** (concaténation)")
        st.markdown(
            "Chaque colonne spéciale regroupe les valeurs de **toutes les lignes "
            f"correspondantes** dans {name_b}, concaténées selon vos séparateurs."
        )

        special_cols = st.session_state.get("rec_special_cols", [])

        for idx, spec in enumerate(special_cols):
            with st.expander(f"🔧 **{spec['name']}**", expanded=True):
                spec["name"] = st.text_input(
                    "Nom de la colonne",
                    value=spec["name"],
                    key=f"rec_spec_name_{idx}",
                )
                spec["cols"] = st.multiselect(
                    "Colonnes à concaténer (dans l'ordre, par ligne matchée)",
                    options=list(df_b.columns),
                    default=[c for c in spec.get("cols", []) if c in df_b.columns],
                    key=f"rec_spec_cols_{idx}",
                )
                c1, c2 = st.columns(2)
                with c1:
                    spec["col_sep"] = st.text_input(
                        "Séparateur entre colonnes (au sein d'une ligne)",
                        value=spec.get("col_sep", " | "),
                        key=f"rec_spec_colsep_{idx}",
                    )
                with c2:
                    spec["row_sep"] = st.text_input(
                        "Séparateur entre lignes matchées",
                        value=spec.get("row_sep", " // "),
                        key=f"rec_spec_rowsep_{idx}",
                    )

                # Live preview
                key_a = st.session_state["rec_key_a_val"]
                key_b = st.session_state["rec_key_b_val"]
                if spec["cols"] and len(df_a[key_a].dropna()) > 0:
                    sample_val = str(df_a[key_a].dropna().iloc[0])
                    matched = find_matching_rows(df_b, key_b, sample_val, match_mode, case_sensitive)
                    preview_val = build_concat_value(matched, spec["cols"], spec["col_sep"], spec["row_sep"])
                    st.caption(
                        f"Aperçu pour **\"{sample_val}\"** → "
                        f"`{preview_val[:120]}{'...' if len(preview_val) > 120 else ''}`"
                    )

                if st.button("🗑️ Supprimer", key=f"rec_del_spec_{idx}"):
                    special_cols.pop(idx)
                    st.session_state["rec_special_cols"] = special_cols
                    st.rerun()

        if st.button("➕ Ajouter une colonne spéciale"):
            special_cols.append({
                "name": f"Résultat_{len(special_cols) + 1}",
                "cols": [],
                "col_sep": " | ",
                "row_sep": " // ",
            })
            st.session_state["rec_special_cols"] = special_cols
            st.rerun()

        st.session_state["rec_special_cols"] = special_cols

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["rec_step"] = 2
                st.rerun()
        with c2:
            if st.button("Voir le résultat →", type="primary"):
                st.session_state["rec_output_cols_a"] = selected_a
                st.session_state["rec_step"] = 4
                st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 4 — Résultat
    # ══════════════════════════════════════════
    elif step == 4:
        df_a: pd.DataFrame = st.session_state["rec_df_a"].copy()
        df_b: pd.DataFrame = st.session_state["rec_df_b"].copy()
        name_a = st.session_state["rec_name_a"]
        name_b = st.session_state["rec_name_b"]
        key_a = st.session_state["rec_key_a_val"]
        key_b = st.session_state["rec_key_b_val"]
        match_mode = st.session_state["rec_match_mode"]
        case_sensitive = st.session_state["rec_case_sensitive"]
        output_cols_a = st.session_state.get("rec_output_cols_a", list(df_a.columns))
        special_cols = st.session_state.get("rec_special_cols", [])

        mode_lbl = "Match parfait" if match_mode == "parfait" else "Match normal (contient)"
        casse_lbl = "casse respectée" if case_sensitive else "casse ignorée"

        st.markdown("### ✅ Étape 4 — Résultat du rapprochement")
        st.caption(f"Mode : **{mode_lbl}** · {casse_lbl}")

        try:
            # ── Select A columns ──────────────────
            cols_a = [c for c in output_cols_a if c in df_a.columns]
            if key_a not in cols_a:
                cols_a = [key_a] + cols_a
            df_result = df_a[cols_a].copy()

            # ── Build special columns for each row ─
            spec_col_names = []
            for spec in special_cols:
                if not spec["cols"]:
                    continue
                col_name = spec["name"]
                spec_col_names.append(col_name)
                values = []
                for _, row_a in df_a.iterrows():
                    val_a = str(row_a[key_a])
                    matched = find_matching_rows(df_b, key_b, val_a, match_mode, case_sensitive)
                    concat_val = build_concat_value(
                        matched, spec["cols"], spec["col_sep"], spec["row_sep"]
                    )
                    values.append(concat_val)
                df_result[col_name] = values

            # ── Stats ─────────────────────────────
            no_match_mask = pd.Series([False] * len(df_result))
            if spec_col_names:
                no_match_mask = df_result[spec_col_names].apply(
                    lambda col: col == "", axis=0
                ).all(axis=1)

            matched_count = int((~no_match_mask).sum())

            col1, col2, col3 = st.columns(3)
            col1.metric("Lignes dans le fichier de base", len(df_a))
            col2.metric("Lignes dans le fichier B", len(df_b))
            col3.metric("Lignes avec correspondance", matched_count)

            if no_match_mask.any():
                st.caption(f"🟡 {int(no_match_mask.sum())} ligne(s) sans correspondance (surlignées en jaune)")

            st.markdown("---")

            # ── Display with highlight ────────────
            def highlight_unmatched(row):
                idx = row.name
                if spec_col_names and no_match_mask.iloc[idx]:
                    return ["background-color: #fef3c7; color:#92400e"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_result.style.apply(highlight_unmatched, axis=1),
                use_container_width=True,
                height=450,
            )

            # ── Download ──────────────────────────
            st.markdown("---")
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Résultat")
                    if no_match_mask.any():
                        df_result[no_match_mask].to_excel(
                            writer, index=False, sheet_name="Sans correspondance"
                        )
                st.download_button(
                    label="⬇️ Télécharger (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="reconciliation_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

            with col_dl2:
                csv = df_result.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="⬇️ Télécharger (.csv)",
                    data=csv,
                    file_name="reconciliation_result.csv",
                    mime="text/csv",
                )

        except Exception as e:
            st.error(f"Erreur lors du rapprochement : {e}")

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Modifier"):
                st.session_state["rec_step"] = 3
                st.rerun()
        with c2:
            if st.button("🔄 Nouveau rapprochement"):
                reset_reconciler()
                st.rerun()
