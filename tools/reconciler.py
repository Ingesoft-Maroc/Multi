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


# ─────────────────────────────────────────────
#  Main entry point
# ─────────────────────────────────────────────

def run_reconciler():
    st.markdown("## 🔗 Réconciliateur Excel")
    st.markdown(
        "Importez deux fichiers Excel, définissez la clé de rapprochement "
        "et configurez votre sortie colonne par colonne."
    )
    st.markdown("---")

    # ── Step progress tracker ──────────────────
    step = st.session_state.get("rec_step", 1)

    col_steps = st.columns(4)
    labels = ["Import", "Clé de rapprochement", "Colonnes de sortie", "Résultat"]
    for i, (c, lbl) in enumerate(zip(col_steps, labels), start=1):
        with c:
            step_badge(i, lbl, active=(i == step))

    st.markdown("<br>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    #  ÉTAPE 1 — Import des fichiers
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
    #  ÉTAPE 2 — Clé de rapprochement
    # ══════════════════════════════════════════
    elif step == 2:
        df_a: pd.DataFrame = st.session_state["rec_df_a"]
        df_b: pd.DataFrame = st.session_state["rec_df_b"]
        name_a = st.session_state["rec_name_a"]
        name_b = st.session_state["rec_name_b"]

        st.markdown("### 🔑 Étape 2 — Clé de rapprochement")
        st.markdown(
            "Choisissez la colonne qui servira de **clé commune** entre les deux fichiers."
        )

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

        # Preview of key values
        if key_a and key_b:
            col_a, col_b = st.columns(2)
            with col_a:
                sample_a = df_a[key_a].dropna().astype(str).unique()[:8]
                st.caption(f"Exemples : {', '.join(sample_a)}")
            with col_b:
                sample_b = df_b[key_b].dropna().astype(str).unique()[:8]
                st.caption(f"Exemples : {', '.join(sample_b)}")

            # Quick match stats
            vals_a = set(df_a[key_a].astype(str))
            vals_b = set(df_b[key_b].astype(str))
            matched = len(vals_a & vals_b)
            total_a = len(vals_a)
            st.info(
                f"🔍 **{matched}** valeurs communes trouvées sur **{total_a}** "
                f"valeurs uniques dans le fichier de base."
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
                # init output config defaults
                if "rec_output_cols_a" not in st.session_state:
                    st.session_state["rec_output_cols_a"] = list(df_a.columns)
                if "rec_special_cols" not in st.session_state:
                    st.session_state["rec_special_cols"] = []
                st.session_state["rec_step"] = 3
                st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 3 — Configuration de la sortie
    # ══════════════════════════════════════════
    elif step == 3:
        df_a: pd.DataFrame = st.session_state["rec_df_a"]
        df_b: pd.DataFrame = st.session_state["rec_df_b"]
        name_a = st.session_state["rec_name_a"]
        name_b = st.session_state["rec_name_b"]
        key_a = st.session_state["rec_key_a_val"]
        key_b = st.session_state["rec_key_b_val"]

        st.markdown("### 🗂️ Étape 3 — Colonnes de sortie")
        st.markdown(
            "Configurez ce que vous voulez voir dans le résultat final."
        )

        # ── Part A : colonnes du fichier de base ─
        st.markdown(f"#### 📄 Colonnes de **{name_a}** (fichier de base)")
        all_cols_a = list(df_a.columns)
        selected_a = st.multiselect(
            "Colonnes à afficher",
            options=all_cols_a,
            default=st.session_state.get("rec_output_cols_a", all_cols_a),
            key="rec_ms_cols_a",
        )

        st.markdown("---")

        # ── Part B : colonnes du fichier B ────────
        st.markdown(f"#### 📄 Colonnes de **{name_b}** (fichier à rapprocher)")

        mode = st.radio(
            "Mode d'affichage",
            options=["Colonnes individuelles", "Colonne spéciale (concaténation)"],
            key="rec_mode_b",
            horizontal=True,
        )

        if mode == "Colonnes individuelles":
            all_cols_b = list(df_b.columns)
            selected_b_cols = st.multiselect(
                "Colonnes à afficher",
                options=all_cols_b,
                default=st.session_state.get("rec_output_cols_b", all_cols_b[:3]),
                key="rec_ms_cols_b",
            )
            st.session_state["rec_output_cols_b"] = selected_b_cols
            st.session_state["rec_special_cols"] = []
            st.session_state["rec_concat_sep"] = ""
            st.session_state["rec_concat_name"] = ""

        else:  # Colonne spéciale
            st.markdown("Créez une ou plusieurs **colonnes de concaténation** personnalisées.")

            # existing special columns
            special_cols = st.session_state.get("rec_special_cols", [])

            for idx, spec in enumerate(special_cols):
                with st.expander(f"🔧 Colonne spéciale : **{spec['name']}**", expanded=False):
                    spec["name"] = st.text_input(
                        "Nom de la colonne",
                        value=spec["name"],
                        key=f"rec_spec_name_{idx}",
                    )
                    spec["cols"] = st.multiselect(
                        "Colonnes à concaténer (dans l'ordre)",
                        options=list(df_b.columns),
                        default=[c for c in spec["cols"] if c in df_b.columns],
                        key=f"rec_spec_cols_{idx}",
                    )
                    spec["sep"] = st.text_input(
                        "Séparateur",
                        value=spec.get("sep", " | "),
                        key=f"rec_spec_sep_{idx}",
                    )
                    if st.button("🗑️ Supprimer", key=f"rec_del_spec_{idx}"):
                        special_cols.pop(idx)
                        st.session_state["rec_special_cols"] = special_cols
                        st.rerun()

            if st.button("➕ Ajouter une colonne spéciale"):
                special_cols.append({
                    "name": f"Concat_{len(special_cols)+1}",
                    "cols": [],
                    "sep": " | ",
                })
                st.session_state["rec_special_cols"] = special_cols
                st.rerun()

            st.session_state["rec_special_cols"] = special_cols
            st.session_state["rec_output_cols_b"] = []

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
        output_cols_a = st.session_state.get("rec_output_cols_a", list(df_a.columns))
        output_cols_b = st.session_state.get("rec_output_cols_b", [])
        special_cols = st.session_state.get("rec_special_cols", [])
        mode = st.session_state.get("rec_mode_b", "Colonnes individuelles")

        st.markdown("### ✅ Étape 4 — Résultat du rapprochement")

        try:
            # ── Build B subset ────────────────────
            if mode == "Colonnes individuelles":
                cols_to_take_b = [key_b] + [c for c in output_cols_b if c != key_b]
                df_b_sub = df_b[cols_to_take_b].copy()
            else:
                # Build special columns on df_b
                df_b_sub = df_b[[key_b]].copy()
                for spec in special_cols:
                    if spec["cols"]:
                        df_b_sub[spec["name"]] = (
                            df_b[spec["cols"]]
                            .astype(str)
                            .apply(lambda row: spec["sep"].join(
                                v for v in row if v not in ("nan", "None", "")
                            ), axis=1)
                        )

            # ── Select A columns ──────────────────
            cols_to_take_a = list(dict.fromkeys(
                [c for c in output_cols_a if c in df_a.columns]
            ))
            if key_a not in cols_to_take_a:
                cols_to_take_a = [key_a] + cols_to_take_a
            df_a_sub = df_a[cols_to_take_a].copy()

            # ── Merge ─────────────────────────────
            df_b_sub = df_b_sub.rename(columns={key_b: key_a})
            result = df_a_sub.merge(df_b_sub, on=key_a, how="left")

            # ── Stats ─────────────────────────────
            matched = result[
                result.iloc[:, result.columns.get_loc(key_a) + 1]
                if len(result.columns) > result.columns.get_loc(key_a) + 1
                else result[key_a]
            ].notna() if False else None

            matched_count = result.dropna(
                subset=[c for c in result.columns if c != key_a], how="all"
            ).shape[0]

            col1, col2, col3 = st.columns(3)
            col1.metric("Lignes dans le fichier de base", len(df_a))
            col2.metric("Lignes dans le fichier B", len(df_b))
            col3.metric("Lignes réconciliées", matched_count)

            st.markdown("---")

            # Highlight unmatched (all NaN from B side)
            b_extra_cols = [c for c in result.columns if c not in cols_to_take_a]
            if b_extra_cols:
                unmatched_mask = result[b_extra_cols].isnull().all(axis=1)
                st.caption(
                    f"🟡 {unmatched_mask.sum()} ligne(s) sans correspondance dans {name_b} "
                    f"(affichées en jaune)"
                )

                def highlight_unmatched(row):
                    if b_extra_cols and all(pd.isnull(row.get(c)) for c in b_extra_cols):
                        return ["background-color: #fef3c7; color:#92400e"] * len(row)
                    return [""] * len(row)

                st.dataframe(
                    result.style.apply(highlight_unmatched, axis=1),
                    use_container_width=True,
                    height=450,
                )
            else:
                st.dataframe(result, use_container_width=True, height=450)

            # ── Download ──────────────────────────
            st.markdown("---")
            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    result.to_excel(writer, index=False, sheet_name="Résultat")
                    if b_extra_cols:
                        unmatched = result[result[b_extra_cols].isnull().all(axis=1)]
                        if not unmatched.empty:
                            unmatched.to_excel(writer, index=False, sheet_name="Non réconciliés")
                st.download_button(
                    label="⬇️ Télécharger le résultat (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="reconciliation_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

            with col_dl2:
                csv = result.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="⬇️ Télécharger en CSV",
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
