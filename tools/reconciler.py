import streamlit as st
import pandas as pd
import io

# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

def load_excel(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df.columns = df.columns.astype(str)
        return df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        return None


def reset_reconciler():
    for k in [k for k in st.session_state if k.startswith("rec_")]:
        del st.session_state[k]


def step_badge(n, label, active):
    color = "#6366f1" if active else "#374151"
    txt = "#fff" if active else "#9ca3af"
    st.markdown(
        f'<span style="background:{color};color:{txt};border-radius:20px;'
        f'padding:3px 14px;font-size:0.85rem;font-weight:600;margin-right:6px;">'
        f'Étape {n} — {label}</span>',
        unsafe_allow_html=True,
    )


def find_matching_rows(df_b, key_b, val_a, match_mode, case_sensitive):
    col = df_b[key_b].astype(str)
    if match_mode == "parfait":
        mask = col == val_a if case_sensitive else col.str.lower() == val_a.lower()
    else:
        mask = (col.str.contains(val_a, regex=False, na=False) if case_sensitive
                else col.str.contains(val_a, case=False, regex=False, na=False))
    return df_b[mask]


def build_concat_value(matched_rows, cols, col_sep, row_sep):
    parts = []
    for _, row in matched_rows.iterrows():
        cell = col_sep.join(
            str(row[c]) for c in cols
            if c in row and str(row[c]) not in ("nan", "None", "")
        )
        if cell:
            parts.append(cell)
    return row_sep.join(parts)


def build_sum_value(matched_rows, col):
    try:
        vals = pd.to_numeric(matched_rows[col], errors="coerce").dropna()
        return vals.sum() if len(vals) > 0 else ""
    except Exception:
        return ""


def evaluate_condition(row, rules, else_output):
    for rule in rules:
        left_val = str(row.get(rule.get("left", ""), ""))
        right_val = (str(row.get(rule.get("right", ""), ""))
                     if rule.get("right_type") == "col"
                     else rule.get("right", ""))
        op = rule.get("op", "=")
        try:
            if op == "=":
                cond = left_val == right_val
            elif op == "≠":
                cond = left_val != right_val
            elif op == "contient":
                cond = right_val.lower() in left_val.lower()
            elif op == "ne contient pas":
                cond = right_val.lower() not in left_val.lower()
            elif op == ">":
                cond = float(left_val) > float(right_val)
            elif op == "<":
                cond = float(left_val) < float(right_val)
            elif op == ">=":
                cond = float(left_val) >= float(right_val)
            elif op == "<=":
                cond = float(left_val) <= float(right_val)
            else:
                cond = False
        except Exception:
            cond = False
        if cond:
            return rule.get("output", "")
    return else_output


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def run_reconciler():
    st.markdown("## 🔗 Réconciliateur Excel")
    st.markdown("Rapprochez deux fichiers Excel avec matching avancé et colonnes configurables.")
    st.markdown("---")

    step = st.session_state.get("rec_step", 1)
    labels = ["Import", "Clé & Mode", "Colonnes de sortie", "Résultat"]
    cols_s = st.columns(4)
    for i, (c, lbl) in enumerate(zip(cols_s, labels), 1):
        with c:
            step_badge(i, lbl, i == step)
    st.markdown("<br>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    #  ÉTAPE 1 — Import
    # ══════════════════════════════════════════
    if step == 1:
        st.markdown("### 📂 Étape 1 — Importez vos deux fichiers")
        ca, cb = st.columns(2)
        with ca:
            st.markdown("**Fichier de base (référence)**")
            file_a = st.file_uploader("A", type=["xlsx", "xls"], key="rec_upload_a",
                                      label_visibility="collapsed")
        with cb:
            st.markdown("**Fichier à rapprocher**")
            file_b = st.file_uploader("B", type=["xlsx", "xls"], key="rec_upload_b",
                                      label_visibility="collapsed")

        if file_a and file_b:
            df_a, df_b = load_excel(file_a), load_excel(file_b)
            if df_a is not None and df_b is not None:
                st.session_state.update({
                    "rec_df_a": df_a, "rec_df_b": df_b,
                    "rec_name_a": file_a.name, "rec_name_b": file_b.name,
                })
                ca2, cb2 = st.columns(2)
                with ca2:
                    st.success(f"✅ {file_a.name}")
                    st.caption(f"{len(df_a)} lignes · {len(df_a.columns)} colonnes")
                    with st.expander("Aperçu"):
                        st.dataframe(df_a.head(5), use_container_width=True)
                with cb2:
                    st.success(f"✅ {file_b.name}")
                    st.caption(f"{len(df_b)} lignes · {len(df_b.columns)} colonnes")
                    with st.expander("Aperçu"):
                        st.dataframe(df_b.head(5), use_container_width=True)
                st.markdown("")
                if st.button("Suivant →", type="primary"):
                    st.session_state["rec_step"] = 2
                    st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 2 — Clé & Mode
    # ══════════════════════════════════════════
    elif step == 2:
        df_a = st.session_state["rec_df_a"]
        df_b = st.session_state["rec_df_b"]
        name_a, name_b = st.session_state["rec_name_a"], st.session_state["rec_name_b"]

        st.markdown("### 🔑 Étape 2 — Clé & mode de rapprochement")

        ca, cb = st.columns(2)
        with ca:
            st.markdown(f"**Colonne clé dans {name_a}**")
            key_a = st.selectbox("Clé A", list(df_a.columns), key="rec_key_a",
                                 label_visibility="collapsed")
        with cb:
            st.markdown(f"**Colonne clé dans {name_b}**")
            key_b = st.selectbox("Clé B", list(df_b.columns), key="rec_key_b",
                                 label_visibility="collapsed")

        if key_a and key_b:
            ca2, cb2 = st.columns(2)
            with ca2:
                st.caption("Exemples : " + ", ".join(
                    df_a[key_a].dropna().astype(str).unique()[:6]))
            with cb2:
                st.caption("Exemples : " + ", ".join(
                    df_b[key_b].dropna().astype(str).unique()[:6]))

        st.markdown("---")
        st.markdown("**Mode de rapprochement**")

        match_mode = st.radio(
            "Mode", ["Match parfait", "Match normal (contient)"],
            key="rec_match_mode_radio", horizontal=True, label_visibility="collapsed",
            help="**Parfait** : valeur identique. **Normal** : le champ B contient la clé A.",
        )
        case_opt = st.radio(
            "Casse",
            ["Ignorer la casse  (Xavier = XAVIER)", "Respecter la casse  (Xavier ≠ XAVIER)"],
            key="rec_case_radio", horizontal=True, label_visibility="collapsed",
        )
        case_sensitive = "Respecter" in case_opt
        mm = "parfait" if "parfait" in match_mode else "normal"

        if key_a and key_b and len(df_a[key_a].dropna()) > 0:
            sample = str(df_a[key_a].dropna().iloc[0])
            matched = find_matching_rows(df_b, key_b, sample, mm, case_sensitive)
            st.info(f"🔍 Exemple avec **\"{sample}\"** → **{len(matched)}** ligne(s) dans {name_b}")

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["rec_step"] = 1; st.rerun()
        with c2:
            if st.button("Suivant →", type="primary"):
                st.session_state.update({
                    "rec_key_a_val": key_a, "rec_key_b_val": key_b,
                    "rec_match_mode": mm, "rec_case_sensitive": case_sensitive,
                })
                if "rec_special_cols" not in st.session_state:
                    st.session_state["rec_special_cols"] = []
                if "rec_output_cols_a" not in st.session_state:
                    st.session_state["rec_output_cols_a"] = list(df_a.columns)
                if "rec_exclusions" not in st.session_state:
                    st.session_state["rec_exclusions"] = []
                st.session_state["rec_step"] = 3; st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 3 — Config sortie
    # ══════════════════════════════════════════
    elif step == 3:
        df_a = st.session_state["rec_df_a"]
        df_b = st.session_state["rec_df_b"]
        name_a, name_b = st.session_state["rec_name_a"], st.session_state["rec_name_b"]
        key_a = st.session_state["rec_key_a_val"]
        key_b = st.session_state["rec_key_b_val"]
        match_mode = st.session_state["rec_match_mode"]
        case_sensitive = st.session_state["rec_case_sensitive"]

        st.markdown("### 🗂️ Étape 3 — Configuration de la sortie")

        # ── FILTRES D'EXCLUSION ───────────────
        with st.expander("🚫 Filtres d'exclusion (optionnel)", expanded=False):
            st.markdown(
                "Exclure des lignes du **fichier de base** selon la valeur d'une colonne. "
                "Les lignes exclues n'apparaissent pas dans le résultat."
            )
            exclusions = st.session_state.get("rec_exclusions", [])

            for idx, excl in enumerate(exclusions):
                c1, c2 = st.columns([2, 3])
                with c1:
                    excl["col"] = st.selectbox(
                        f"Colonne {idx+1}", list(df_a.columns),
                        index=list(df_a.columns).index(excl["col"])
                              if excl["col"] in df_a.columns else 0,
                        key=f"rec_excl_col_{idx}",
                    )
                with c2:
                    unique_vals = sorted(df_a[excl["col"]].dropna().astype(str).unique().tolist())
                    excl["values"] = st.multiselect(
                        f"Valeurs à exclure",
                        options=unique_vals,
                        default=[v for v in excl.get("values", []) if v in unique_vals],
                        key=f"rec_excl_vals_{idx}",
                    )
                if st.button("🗑️ Supprimer ce filtre", key=f"rec_del_excl_{idx}"):
                    exclusions.pop(idx)
                    st.session_state["rec_exclusions"] = exclusions
                    st.rerun()
                st.markdown("---")

            if st.button("➕ Ajouter un filtre d'exclusion"):
                exclusions.append({"col": list(df_a.columns)[0], "values": []})
                st.session_state["rec_exclusions"] = exclusions
                st.rerun()
            st.session_state["rec_exclusions"] = exclusions

        st.markdown("---")

        # ── COLONNES FICHIER A ────────────────
        st.markdown(f"#### 📄 Colonnes de **{name_a}** à afficher")
        all_cols_a = list(df_a.columns)
        selected_a = st.multiselect(
            "Colonnes A", all_cols_a,
            default=st.session_state.get("rec_output_cols_a", all_cols_a),
            key="rec_ms_cols_a", label_visibility="collapsed",
        )

        st.markdown("---")

        # ── COLONNES SPÉCIALES ────────────────
        st.markdown(f"#### 🔧 Colonnes spéciales depuis **{name_b}**")
        st.markdown(
            "Trois types disponibles : **Concaténation**, **Somme**, **Condition**."
        )

        special_cols = st.session_state.get("rec_special_cols", [])

        # Names of already-defined special cols + selected A cols (for condition references)
        available_ref_cols = list(selected_a) + [s["name"] for s in special_cols]

        for idx, spec in enumerate(special_cols):
            stype = spec.get("type", "concat")
            type_icon = {"concat": "🔗", "sum": "➕", "condition": "🔀"}.get(stype, "🔧")
            with st.expander(f"{type_icon} **{spec['name']}** ({stype})", expanded=True):

                spec["name"] = st.text_input(
                    "Nom de la colonne", value=spec["name"], key=f"rec_spec_name_{idx}"
                )

                # ── CONCAT ──
                if stype == "concat":
                    spec["cols"] = st.multiselect(
                        "Colonnes à concaténer (par ligne matchée)", list(df_b.columns),
                        default=[c for c in spec.get("cols", []) if c in df_b.columns],
                        key=f"rec_spec_cols_{idx}",
                    )
                    c1, c2 = st.columns(2)
                    with c1:
                        spec["col_sep"] = st.text_input(
                            "Séparateur entre colonnes", value=spec.get("col_sep", " | "),
                            key=f"rec_spec_colsep_{idx}",
                        )
                    with c2:
                        spec["row_sep"] = st.text_input(
                            "Séparateur entre lignes matchées", value=spec.get("row_sep", " // "),
                            key=f"rec_spec_rowsep_{idx}",
                        )
                    if spec["cols"] and len(df_a[key_a].dropna()) > 0:
                        sv = str(df_a[key_a].dropna().iloc[0])
                        m = find_matching_rows(df_b, key_b, sv, match_mode, case_sensitive)
                        prev = build_concat_value(m, spec["cols"], spec["col_sep"], spec["row_sep"])
                        st.caption(f"Aperçu pour \"{sv}\" → `{prev[:100]}{'...' if len(prev)>100 else ''}`")

                # ── SUM ──
                elif stype == "sum":
                    num_cols = list(df_b.columns)
                    spec["sum_col"] = st.selectbox(
                        "Colonne numérique à sommer", num_cols,
                        index=num_cols.index(spec["sum_col"])
                              if spec.get("sum_col") in num_cols else 0,
                        key=f"rec_spec_sumcol_{idx}",
                    )
                    if spec.get("sum_col") and len(df_a[key_a].dropna()) > 0:
                        sv = str(df_a[key_a].dropna().iloc[0])
                        m = find_matching_rows(df_b, key_b, sv, match_mode, case_sensitive)
                        total = build_sum_value(m, spec["sum_col"])
                        st.caption(f"Aperçu pour \"{sv}\" → somme = `{total}`")

                # ── CONDITION ──
                elif stype == "condition":
                    ref_cols_before = list(selected_a) + [
                        special_cols[j]["name"] for j in range(idx)
                    ]
                    if not ref_cols_before:
                        st.warning("Définissez d'abord des colonnes de référence (fichier A ou colonnes spéciales précédentes).")
                    else:
                        st.markdown("**Règles (évaluées dans l'ordre — première qui matche)**")
                        rules = spec.get("rules", [])

                        for ri, rule in enumerate(rules):
                            rc1, rc2, rc3, rc4, rc5 = st.columns([2, 1, 1, 2, 2])
                            with rc1:
                                rule["left"] = st.selectbox(
                                    "Colonne gauche", ref_cols_before,
                                    index=ref_cols_before.index(rule["left"])
                                          if rule.get("left") in ref_cols_before else 0,
                                    key=f"rec_rule_left_{idx}_{ri}",
                                    label_visibility="collapsed",
                                )
                            with rc2:
                                rule["op"] = st.selectbox(
                                    "Op", ["=", "≠", "contient", "ne contient pas", ">", "<", ">=", "<="],
                                    index=["=", "≠", "contient", "ne contient pas", ">", "<", ">=", "<="]
                                          .index(rule.get("op", "=")),
                                    key=f"rec_rule_op_{idx}_{ri}",
                                    label_visibility="collapsed",
                                )
                            with rc3:
                                rule["right_type"] = st.selectbox(
                                    "Type droite", ["valeur", "col"],
                                    index=0 if rule.get("right_type", "valeur") == "valeur" else 1,
                                    key=f"rec_rule_rtype_{idx}_{ri}",
                                    label_visibility="collapsed",
                                )
                            with rc4:
                                if rule["right_type"] == "col":
                                    rule["right"] = st.selectbox(
                                        "Col droite", ref_cols_before,
                                        index=ref_cols_before.index(rule["right"])
                                              if rule.get("right") in ref_cols_before else 0,
                                        key=f"rec_rule_right_{idx}_{ri}",
                                        label_visibility="collapsed",
                                    )
                                else:
                                    rule["right"] = st.text_input(
                                        "Valeur", value=rule.get("right", ""),
                                        key=f"rec_rule_right_{idx}_{ri}",
                                        label_visibility="collapsed",
                                        placeholder="valeur à comparer",
                                    )
                            with rc5:
                                rule["output"] = st.text_input(
                                    "→ Écrire", value=rule.get("output", ""),
                                    key=f"rec_rule_out_{idx}_{ri}",
                                    label_visibility="collapsed",
                                    placeholder="texte si vrai",
                                )
                            if st.button("✕", key=f"rec_del_rule_{idx}_{ri}"):
                                rules.pop(ri)
                                spec["rules"] = rules
                                st.session_state["rec_special_cols"] = special_cols
                                st.rerun()

                        spec["rules"] = rules

                        if st.button("➕ Ajouter une règle", key=f"rec_add_rule_{idx}"):
                            rules.append({"left": ref_cols_before[0], "op": "=",
                                          "right_type": "valeur", "right": "", "output": ""})
                            spec["rules"] = rules
                            st.session_state["rec_special_cols"] = special_cols
                            st.rerun()

                        spec["else_output"] = st.text_input(
                            "Sinon (ELSE) → écrire",
                            value=spec.get("else_output", ""),
                            key=f"rec_spec_else_{idx}",
                        )

                if st.button("🗑️ Supprimer cette colonne", key=f"rec_del_spec_{idx}"):
                    special_cols.pop(idx)
                    st.session_state["rec_special_cols"] = special_cols
                    st.rerun()

        # ── Add buttons ───────────────────────
        st.markdown("")
        ca, cb, cc = st.columns(3)
        with ca:
            if st.button("➕ Concaténation"):
                special_cols.append({"type": "concat", "name": f"Concat_{len(special_cols)+1}",
                                     "cols": [], "col_sep": " | ", "row_sep": " // "})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()
        with cb:
            if st.button("➕ Somme"):
                special_cols.append({"type": "sum", "name": f"Somme_{len(special_cols)+1}",
                                     "sum_col": list(df_b.columns)[0] if df_b.columns.any() else ""})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()
        with cc:
            if st.button("➕ Condition"):
                special_cols.append({"type": "condition", "name": f"Condition_{len(special_cols)+1}",
                                     "rules": [], "else_output": ""})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()

        st.session_state["rec_special_cols"] = special_cols

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["rec_step"] = 2; st.rerun()
        with c2:
            if st.button("Voir le résultat →", type="primary"):
                st.session_state["rec_output_cols_a"] = selected_a
                st.session_state["rec_step"] = 4; st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 4 — Résultat
    # ══════════════════════════════════════════
    elif step == 4:
        df_a = st.session_state["rec_df_a"].copy()
        df_b = st.session_state["rec_df_b"].copy()
        name_a, name_b = st.session_state["rec_name_a"], st.session_state["rec_name_b"]
        key_a = st.session_state["rec_key_a_val"]
        key_b = st.session_state["rec_key_b_val"]
        match_mode = st.session_state["rec_match_mode"]
        case_sensitive = st.session_state["rec_case_sensitive"]
        output_cols_a = st.session_state.get("rec_output_cols_a", list(df_a.columns))
        special_cols = st.session_state.get("rec_special_cols", [])
        exclusions = st.session_state.get("rec_exclusions", [])

        st.markdown("### ✅ Étape 4 — Résultat")
        mode_lbl = "Match parfait" if match_mode == "parfait" else "Match normal (contient)"
        casse_lbl = "casse respectée" if case_sensitive else "casse ignorée"
        st.caption(f"Mode : **{mode_lbl}** · {casse_lbl}")

        try:
            # ── Apply exclusion filters ───────────
            mask_keep = pd.Series([True] * len(df_a), index=df_a.index)
            for excl in exclusions:
                col_e = excl.get("col")
                vals_e = excl.get("values", [])
                if col_e and vals_e:
                    mask_keep &= ~df_a[col_e].astype(str).isin(vals_e)

            excluded_count = int((~mask_keep).sum())
            df_a_filtered = df_a[mask_keep].copy().reset_index(drop=True)

            # ── Build base result from A ──────────
            cols_a = [c for c in output_cols_a if c in df_a_filtered.columns]
            if key_a not in cols_a:
                cols_a = [key_a] + cols_a
            df_result = df_a_filtered[cols_a].copy()

            # ── Compute concat / sum columns ──────
            no_match_flags = []
            for spec in special_cols:
                stype = spec.get("type", "concat")
                if stype == "condition":
                    continue  # computed later
                col_name = spec["name"]
                values = []
                for _, row_a in df_a_filtered.iterrows():
                    val_a = str(row_a[key_a])
                    matched = find_matching_rows(df_b, key_b, val_a, match_mode, case_sensitive)
                    if stype == "concat":
                        v = build_concat_value(matched, spec.get("cols", []),
                                               spec.get("col_sep", " | "),
                                               spec.get("row_sep", " // "))
                    elif stype == "sum":
                        v = build_sum_value(matched, spec.get("sum_col", ""))
                    else:
                        v = ""
                    values.append(v)
                df_result[col_name] = values
                no_match_flags.append(col_name)

            # ── Compute condition columns ─────────
            for spec in special_cols:
                if spec.get("type") != "condition":
                    continue
                col_name = spec["name"]
                df_result[col_name] = df_result.apply(
                    lambda row: evaluate_condition(row, spec.get("rules", []),
                                                   spec.get("else_output", "")),
                    axis=1,
                )

            # ── Stats ─────────────────────────────
            if no_match_flags:
                no_match_mask = df_result[no_match_flags].apply(
                    lambda col: col.astype(str).isin(["", "nan", "0", "0.0"]),
                ).all(axis=1)
            else:
                no_match_mask = pd.Series([False] * len(df_result))

            matched_count = int((~no_match_mask).sum())

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Lignes fichier de base", len(df_a))
            m2.metric("Exclusions appliquées", excluded_count)
            m3.metric("Lignes traitées", len(df_a_filtered))
            m4.metric("Avec correspondance", matched_count)

            if no_match_mask.any():
                st.caption(f"🟡 {int(no_match_mask.sum())} ligne(s) sans correspondance (surlignées)")

            st.markdown("---")

            def highlight_row(row):
                if no_match_mask.iloc[row.name]:
                    return ["background-color:#fef3c7;color:#92400e"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_result.style.apply(highlight_row, axis=1),
                use_container_width=True, height=450,
            )

            # ── Download ──────────────────────────
            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Résultat")
                    if no_match_mask.any():
                        df_result[no_match_mask].to_excel(
                            writer, index=False, sheet_name="Sans correspondance")
                    if excluded_count > 0:
                        df_a[~mask_keep].to_excel(
                            writer, index=False, sheet_name="Exclusions")
                st.download_button("⬇️ Télécharger (.xlsx)", buf.getvalue(),
                                   "reconciliation_result.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   type="primary")
            with c2:
                st.download_button("⬇️ Télécharger (.csv)",
                                   df_result.to_csv(index=False).encode("utf-8-sig"),
                                   "reconciliation_result.csv", "text/csv")

        except Exception as e:
            st.error(f"Erreur lors du rapprochement : {e}")
            import traceback
            st.code(traceback.format_exc())

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Modifier"):
                st.session_state["rec_step"] = 3; st.rerun()
        with c2:
            if st.button("🔄 Nouveau rapprochement"):
                reset_reconciler(); st.rerun()
