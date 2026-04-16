import streamlit as st
import pandas as pd
import io
import json


# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

def load_excel(uploaded_file, sheet=0):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df.columns = df.columns.astype(str)
        return df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        return None


def get_sheets(uploaded_file):
    try:
        xl = pd.ExcelFile(uploaded_file)
        return xl.sheet_names
    except Exception:
        return []


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


def sum_matched_col(matched_rows, col):
    try:
        vals = pd.to_numeric(matched_rows[col], errors="coerce").dropna()
        return float(vals.sum()) if len(vals) > 0 else None
    except Exception:
        return None


def _apply_col_filters(df_b, col_filters):
    df = df_b.copy()
    for f in col_filters:
        col_f, vals_f = f.get("col"), f.get("values", [])
        if col_f and col_f in df.columns and vals_f:
            df = df[~df[col_f].astype(str).isin(vals_f)]
    return df


def evaluate_condition(row, rules, else_output):
    for rule in rules:
        left_val = str(row.get(rule.get("left", ""), ""))
        right_val = (str(row.get(rule.get("right", ""), ""))
                     if rule.get("right_type") == "col" else rule.get("right", ""))
        op = rule.get("op", "=")
        try:
            if op == "=":             cond = left_val == right_val
            elif op == "≠":           cond = left_val != right_val
            elif op == "contient":    cond = right_val.lower() in left_val.lower()
            elif op == "ne contient pas": cond = right_val.lower() not in left_val.lower()
            elif op == ">":           cond = float(left_val) > float(right_val)
            elif op == "<":           cond = float(left_val) < float(right_val)
            elif op == ">=":          cond = float(left_val) >= float(right_val)
            elif op == "<=":          cond = float(left_val) <= float(right_val)
            else: cond = False
        except Exception:
            cond = False
        if cond:
            return rule.get("output", "")
    return else_output


def evaluate_formula(row, terms, df_b, key_b, key_a_col, match_mode, case_sensitive):
    result = None
    for i, term in enumerate(terms):
        src = term.get("source", "val")
        op = term.get("op", "+")
        if src == "col_result":
            raw = row.get(term.get("col", ""), None)
            try:
                val = float(raw) if raw not in (None, "", "nan", "None") else None
            except (ValueError, TypeError):
                val = None
        elif src == "sum_matched":
            col_b = term.get("col", "")
            val_a_key = str(row.get(key_a_col, ""))
            df_b_f = _apply_col_filters(df_b, term.get("col_filters", []))
            matched = find_matching_rows(df_b_f, key_b, val_a_key, match_mode, case_sensitive)
            val = sum_matched_col(matched, col_b)
        else:
            try:
                val = float(term.get("val", 0))
            except (ValueError, TypeError):
                val = None
        if val is None:
            continue
        if result is None or i == 0:
            result = val
        else:
            try:
                if op == "+":   result += val
                elif op == "-": result -= val
                elif op == "*": result *= val
                elif op == "/": result = result / val if val != 0 else None
            except Exception:
                result = None
    return result if result is not None else ""


# ── Config save/load ─────────────────────────
CONFIG_KEYS = [
    "rec_key_a_val", "rec_key_b_val", "rec_match_mode", "rec_case_sensitive",
    "rec_output_cols_a", "rec_special_cols", "rec_exclusions",
]

def export_config():
    cfg = {k.replace("rec_", ""): st.session_state.get(k) for k in CONFIG_KEYS}
    return json.dumps(cfg, ensure_ascii=False, indent=2)


def import_config(cfg: dict):
    for k in CONFIG_KEYS:
        short = k.replace("rec_", "")
        if short in cfg:
            st.session_state[k] = cfg[short]


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
        st.markdown("### 📂 Étape 1 — Importez vos données")

        # ── Mode d'import ─────────────────────
        st.markdown("**Mode d'import**")
        input_mode = st.radio(
            "Mode", ["📄 Deux fichiers Excel séparés", "📑 Un fichier Excel multi-feuilles"],
            horizontal=True, key="rec_input_mode", label_visibility="collapsed",
        )

        df_a, df_b, name_a, name_b = None, None, None, None

        if input_mode == "📄 Deux fichiers Excel séparés":
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
                df_a = load_excel(file_a)
                df_b = load_excel(file_b)
                name_a, name_b = file_a.name, file_b.name

        else:  # Multi-feuilles
            st.markdown("**Fichier Excel (plusieurs feuilles)**")
            multi_file = st.file_uploader("Fichier", type=["xlsx", "xls"],
                                          key="rec_upload_multi", label_visibility="collapsed")
            if multi_file:
                file_bytes = multi_file.read()
                sheets = get_sheets(io.BytesIO(file_bytes))
                if sheets:
                    ca, cb = st.columns(2)
                    with ca:
                        st.markdown("**Feuille de base (référence)**")
                        sheet_a = st.selectbox("Feuille A", sheets, key="rec_sheet_a",
                                               label_visibility="collapsed")
                    with cb:
                        st.markdown("**Feuille à rapprocher**")
                        sheet_b = st.selectbox("Feuille B", sheets,
                                               index=min(1, len(sheets) - 1),
                                               key="rec_sheet_b", label_visibility="collapsed")
                    df_a = load_excel(io.BytesIO(file_bytes), sheet=sheet_a)
                    df_b = load_excel(io.BytesIO(file_bytes), sheet=sheet_b)
                    name_a = f"{multi_file.name} › {sheet_a}"
                    name_b = f"{multi_file.name} › {sheet_b}"

        if df_a is not None and df_b is not None:
            st.session_state.update({
                "rec_df_a": df_a, "rec_df_b": df_b,
                "rec_name_a": name_a, "rec_name_b": name_b,
            })
            ca2, cb2 = st.columns(2)
            with ca2:
                st.success(f"✅ {name_a}")
                st.caption(f"{len(df_a)} lignes · {len(df_a.columns)} colonnes")
                with st.expander("Aperçu"):
                    st.dataframe(df_a.head(5), use_container_width=True)
            with cb2:
                st.success(f"✅ {name_b}")
                st.caption(f"{len(df_b)} lignes · {len(df_b.columns)} colonnes")
                with st.expander("Aperçu"):
                    st.dataframe(df_b.head(5), use_container_width=True)

            # ── Charger une configuration ─────
            st.markdown("---")
            with st.expander("📂 Charger une configuration sauvegardée", expanded=False):
                st.caption("Importez un fichier .json généré par HeroTool pour restaurer tous vos paramètres.")
                cfg_file = st.file_uploader("Configuration (.json)", type=["json"],
                                            key="rec_cfg_upload", label_visibility="collapsed")
                if cfg_file:
                    try:
                        cfg = json.load(cfg_file)
                        if st.button("✅ Appliquer cette configuration", type="primary"):
                            import_config(cfg)
                            if "rec_exclusions" not in st.session_state:
                                st.session_state["rec_exclusions"] = []
                            st.session_state["rec_step"] = 4
                            st.rerun()
                    except Exception as e:
                        st.error(f"Impossible de lire la configuration : {e}")

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
                st.caption("Exemples : " + ", ".join(df_a[key_a].dropna().astype(str).unique()[:6]))
            with cb2:
                st.caption("Exemples : " + ", ".join(df_b[key_b].dropna().astype(str).unique()[:6]))

        st.markdown("---")
        st.markdown("**Mode de rapprochement**")
        match_mode = st.radio("Mode", ["Match parfait", "Match normal (contient)"],
                              key="rec_match_mode_radio", horizontal=True,
                              label_visibility="collapsed",
                              help="**Parfait** : valeur identique. **Normal** : le champ B contient la clé A.")
        case_opt = st.radio("Casse",
                            ["Ignorer la casse  (Xavier = XAVIER)", "Respecter la casse  (Xavier ≠ XAVIER)"],
                            key="rec_case_radio", horizontal=True, label_visibility="collapsed")
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

        # ── FILTRES D'EXCLUSION (global) ─────
        with st.expander("🚫 Filtres d'exclusion globaux sur " + name_b, expanded=False):
            st.caption("Ces lignes de " + name_b + " sont exclues pour toutes les colonnes.")
            exclusions = st.session_state.get("rec_exclusions", [])
            for idx, excl in enumerate(exclusions):
                fc1, fc2 = st.columns([2, 3])
                with fc1:
                    excl["col"] = st.selectbox(
                        f"Colonne #{idx+1}", list(df_b.columns),
                        index=list(df_b.columns).index(excl["col"])
                              if excl.get("col") in df_b.columns else 0,
                        key=f"rec_excl_col_{idx}",
                    )
                with fc2:
                    unique_vals = sorted(df_b[excl["col"]].dropna().astype(str).unique().tolist())
                    excl["values"] = st.multiselect(
                        "Valeurs à exclure", options=unique_vals,
                        default=[v for v in excl.get("values", []) if v in unique_vals],
                        key=f"rec_excl_vals_{idx}",
                    )
                if st.button("✕ Supprimer", key=f"rec_del_excl_{idx}"):
                    exclusions.pop(idx)
                    st.session_state["rec_exclusions"] = exclusions; st.rerun()
                st.markdown("---")
            if st.button("➕ Ajouter un filtre global"):
                exclusions.append({"col": list(df_b.columns)[0], "values": []})
                st.session_state["rec_exclusions"] = exclusions; st.rerun()
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

        special_cols = st.session_state.get("rec_special_cols", [])

        for idx, spec in enumerate(special_cols):
            stype = spec.get("type", "concat")
            icon = {"concat": "🔗", "calcul": "🧮", "condition": "🔀"}.get(stype, "🔧")
            with st.expander(f"{icon} **{spec['name']}** — {stype}", expanded=True):

                spec["name"] = st.text_input(
                    "Nom", value=spec["name"], key=f"rec_spec_name_{idx}")

                # ── CONCAT ───────────────────
                if stype == "concat":
                    spec["cols"] = st.multiselect(
                        "Colonnes à concaténer", list(df_b.columns),
                        default=[c for c in spec.get("cols", []) if c in df_b.columns],
                        key=f"rec_spec_cols_{idx}",
                    )
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        spec["col_sep"] = st.text_input(
                            "Séparateur entre colonnes", value=spec.get("col_sep", " | "),
                            key=f"rec_spec_colsep_{idx}")
                    with cc2:
                        spec["row_sep"] = st.text_input(
                            "Séparateur entre lignes matchées", value=spec.get("row_sep", " // "),
                            key=f"rec_spec_rowsep_{idx}")

                # ── CALCUL ───────────────────
                elif stype == "calcul":
                    ref_cols = list(selected_a) + [special_cols[j]["name"] for j in range(idx)]
                    terms = spec.get("terms", [])
                    st.markdown("**Termes du calcul** (évalués de gauche à droite)")
                    st.caption("Sources : *Col résultat* = colonne A ou spéciale précédente · *Somme matchée* = somme des lignes correspondantes dans " + name_b + " · *Valeur fixe* = nombre")

                    for ti, term in enumerate(terms):
                        tc = st.columns([1, 2, 3, 1])
                        with tc[0]:
                            if ti == 0:
                                st.markdown("<div style='padding-top:8px;text-align:center;color:#9ca3af'>départ</div>", unsafe_allow_html=True)
                            else:
                                term["op"] = st.selectbox(
                                    "Op", ["+", "-", "*", "/"],
                                    index=["+", "-", "*", "/"].index(term.get("op", "+")),
                                    key=f"rec_term_op_{idx}_{ti}", label_visibility="collapsed")
                        with tc[1]:
                            src_opts = ["col résultat", "somme matchée", "valeur fixe"]
                            src_map = {"col_result": "col résultat", "sum_matched": "somme matchée", "val": "valeur fixe"}
                            src_rev = {v: k for k, v in src_map.items()}
                            current_src = src_map.get(term.get("source", "val"), "valeur fixe")
                            chosen_src = st.selectbox("Source", src_opts,
                                index=src_opts.index(current_src),
                                key=f"rec_term_src_{idx}_{ti}", label_visibility="collapsed")
                            term["source"] = src_rev[chosen_src]
                        with tc[2]:
                            if term["source"] == "col_result":
                                opts = ref_cols if ref_cols else ["(aucune)"]
                                term["col"] = st.selectbox("Col", opts,
                                    index=opts.index(term["col"]) if term.get("col") in opts else 0,
                                    key=f"rec_term_col_{idx}_{ti}", label_visibility="collapsed")
                            elif term["source"] == "sum_matched":
                                b_cols = list(df_b.columns)
                                term["col"] = st.selectbox("Col B", b_cols,
                                    index=b_cols.index(term["col"]) if term.get("col") in b_cols else 0,
                                    key=f"rec_term_col_{idx}_{ti}", label_visibility="collapsed")
                            else:
                                term["val"] = st.text_input("Valeur", value=str(term.get("val", "0")),
                                    key=f"rec_term_val_{idx}_{ti}", label_visibility="collapsed")
                        with tc[3]:
                            if st.button("✕", key=f"rec_del_term_{idx}_{ti}"):
                                terms.pop(ti)
                                spec["terms"] = terms
                                st.session_state["rec_special_cols"] = special_cols; st.rerun()

                    spec["terms"] = terms
                    if st.button("➕ Ajouter un terme", key=f"rec_add_term_{idx}"):
                        terms.append({"op": "+", "source": "val", "val": "0", "col": ""})
                        spec["terms"] = terms
                        st.session_state["rec_special_cols"] = special_cols; st.rerun()

                # ── CONDITION ────────────────
                elif stype == "condition":
                    ref_cols = list(selected_a) + [special_cols[j]["name"] for j in range(idx)]
                    if not ref_cols:
                        st.warning("Définissez d'abord des colonnes de référence.")
                    else:
                        st.markdown("**Règles IF/ELIF/ELSE** (première qui matche)")
                        rules = spec.get("rules", [])
                        for ri, rule in enumerate(rules):
                            rc1, rc2, rc3, rc4, rc5 = st.columns([2, 1, 1, 2, 2])
                            with rc1:
                                rule["left"] = st.selectbox("G", ref_cols,
                                    index=ref_cols.index(rule["left"]) if rule.get("left") in ref_cols else 0,
                                    key=f"rec_rule_left_{idx}_{ri}", label_visibility="collapsed")
                            with rc2:
                                ops = ["=", "≠", "contient", "ne contient pas", ">", "<", ">=", "<="]
                                rule["op"] = st.selectbox("Op", ops,
                                    index=ops.index(rule.get("op", "=")),
                                    key=f"rec_rule_op_{idx}_{ri}", label_visibility="collapsed")
                            with rc3:
                                rule["right_type"] = st.selectbox("T", ["valeur", "col"],
                                    index=0 if rule.get("right_type", "valeur") == "valeur" else 1,
                                    key=f"rec_rule_rtype_{idx}_{ri}", label_visibility="collapsed")
                            with rc4:
                                if rule["right_type"] == "col":
                                    rule["right"] = st.selectbox("D", ref_cols,
                                        index=ref_cols.index(rule["right"]) if rule.get("right") in ref_cols else 0,
                                        key=f"rec_rule_right_{idx}_{ri}", label_visibility="collapsed")
                                else:
                                    rule["right"] = st.text_input("Val", value=rule.get("right", ""),
                                        key=f"rec_rule_right_{idx}_{ri}", label_visibility="collapsed",
                                        placeholder="valeur")
                            with rc5:
                                rule["output"] = st.text_input("→", value=rule.get("output", ""),
                                    key=f"rec_rule_out_{idx}_{ri}", label_visibility="collapsed",
                                    placeholder="écrire")
                            if st.button("✕", key=f"rec_del_rule_{idx}_{ri}"):
                                rules.pop(ri)
                                spec["rules"] = rules
                                st.session_state["rec_special_cols"] = special_cols; st.rerun()
                        spec["rules"] = rules
                        if st.button("➕ Ajouter une règle", key=f"rec_add_rule_{idx}"):
                            rules.append({"left": ref_cols[0], "op": "=",
                                          "right_type": "valeur", "right": "", "output": ""})
                            spec["rules"] = rules
                            st.session_state["rec_special_cols"] = special_cols; st.rerun()
                        spec["else_output"] = st.text_input("ELSE → écrire",
                            value=spec.get("else_output", ""), key=f"rec_spec_else_{idx}")

                # ── FILTRES PAR COLONNE ───────
                st.markdown("---")
                with st.expander(f"🚫 Filtres sur {name_b} *(cette colonne uniquement)*", expanded=False):
                    col_filters = spec.get("col_filters", [])
                    for fi, cf in enumerate(col_filters):
                        fc1, fc2 = st.columns([2, 3])
                        with fc1:
                            cf["col"] = st.selectbox(f"Col #{fi+1}", list(df_b.columns),
                                index=list(df_b.columns).index(cf["col"]) if cf.get("col") in df_b.columns else 0,
                                key=f"rec_cf_col_{idx}_{fi}")
                        with fc2:
                            uv = sorted(df_b[cf["col"]].dropna().astype(str).unique().tolist())
                            cf["values"] = st.multiselect("Valeurs à exclure", options=uv,
                                default=[v for v in cf.get("values", []) if v in uv],
                                key=f"rec_cf_vals_{idx}_{fi}")
                        if st.button("✕", key=f"rec_del_cf_{idx}_{fi}"):
                            col_filters.pop(fi)
                            spec["col_filters"] = col_filters
                            st.session_state["rec_special_cols"] = special_cols; st.rerun()
                    if st.button("➕ Ajouter un filtre", key=f"rec_add_cf_{idx}"):
                        col_filters.append({"col": list(df_b.columns)[0], "values": []})
                        spec["col_filters"] = col_filters
                        st.session_state["rec_special_cols"] = special_cols; st.rerun()
                    spec["col_filters"] = col_filters

                if st.button("🗑️ Supprimer cette colonne", key=f"rec_del_spec_{idx}"):
                    special_cols.pop(idx)
                    st.session_state["rec_special_cols"] = special_cols; st.rerun()

        # ── Boutons d'ajout ───────────────────
        st.markdown("")
        ca, cb, cc = st.columns(3)
        with ca:
            if st.button("➕ Concaténation"):
                special_cols.append({"type": "concat", "name": f"Concat_{len(special_cols)+1}",
                                     "cols": [], "col_sep": " | ", "row_sep": " // ", "col_filters": []})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()
        with cb:
            if st.button("➕ Calcul"):
                special_cols.append({"type": "calcul", "name": f"Calcul_{len(special_cols)+1}",
                                     "terms": [], "col_filters": []})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()
        with cc:
            if st.button("➕ Condition"):
                special_cols.append({"type": "condition", "name": f"Condition_{len(special_cols)+1}",
                                     "rules": [], "else_output": "", "col_filters": []})
                st.session_state["rec_special_cols"] = special_cols; st.rerun()

        st.session_state["rec_special_cols"] = special_cols

        # ── Sauvegarde config ─────────────────
        st.markdown("---")
        st.session_state["rec_output_cols_a"] = selected_a
        st.download_button("💾 Sauvegarder la configuration", data=export_config(),
                           file_name="herotool_config.json", mime="application/json")

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
        st.caption(f"Mode : **{mode_lbl}** · {'casse respectée' if case_sensitive else 'casse ignorée'}")

        try:
            mask_keep_b = pd.Series([True] * len(df_b), index=df_b.index)
            for excl in exclusions:
                col_e, vals_e = excl.get("col"), excl.get("values", [])
                if col_e and col_e in df_b.columns and vals_e:
                    mask_keep_b &= ~df_b[col_e].astype(str).isin(vals_e)
            excluded_count = int((~mask_keep_b).sum())
            df_b = df_b[mask_keep_b].copy().reset_index(drop=True)
            df_a_filtered = df_a.copy().reset_index(drop=True)

            cols_a = [c for c in output_cols_a if c in df_a_filtered.columns]
            if key_a not in cols_a:
                cols_a = [key_a] + cols_a
            df_result = df_a_filtered[cols_a].copy()

            no_match_flags = []
            for spec in special_cols:
                stype = spec.get("type", "concat")
                if stype == "condition":
                    continue
                col_name = spec["name"]
                df_b_spec = _apply_col_filters(df_b, spec.get("col_filters", []))
                values = []
                for row_idx, (_, row_a) in enumerate(df_a_filtered.iterrows()):
                    val_a_key = str(row_a[key_a])
                    matched = find_matching_rows(df_b_spec, key_b, val_a_key, match_mode, case_sensitive)
                    if stype == "concat":
                        v = build_concat_value(matched, spec.get("cols", []),
                                               spec.get("col_sep", " | "),
                                               spec.get("row_sep", " // "))
                    elif stype == "calcul":
                        partial_row = df_result.iloc[row_idx].to_dict() if row_idx < len(df_result) else {}
                        v = evaluate_formula(partial_row, spec.get("terms", []),
                                             df_b_spec, key_b, key_a, match_mode, case_sensitive)
                    else:
                        v = ""
                    values.append(v)
                df_result[col_name] = values
                no_match_flags.append(col_name)

            for spec in special_cols:
                if spec.get("type") != "condition":
                    continue
                col_name = spec["name"]
                df_result[col_name] = df_result.apply(
                    lambda row: evaluate_condition(row, spec.get("rules", []),
                                                   spec.get("else_output", "")), axis=1)

            if no_match_flags:
                no_match_mask = df_result[no_match_flags].apply(
                    lambda col: col.astype(str).isin(["", "nan", "None"])).all(axis=1)
            else:
                no_match_mask = pd.Series([False] * len(df_result))

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Lignes fichier de base", len(df_a))
            m2.metric(f"Exclusions dans {name_b}", excluded_count)
            m3.metric("Lignes traitées", len(df_a_filtered))
            m4.metric("Avec correspondance", int((~no_match_mask).sum()))

            if no_match_mask.any():
                st.caption(f"🟡 {int(no_match_mask.sum())} ligne(s) sans correspondance")

            st.markdown("---")

            def highlight_row(row):
                if no_match_mask.iloc[row.name]:
                    return ["background-color:#fef3c7;color:#92400e"] * len(row)
                return [""] * len(row)

            st.dataframe(df_result.style.apply(highlight_row, axis=1),
                         use_container_width=True, height=450)

            st.markdown("---")
            dc1, dc2, dc3 = st.columns(3)
            with dc1:
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Résultat")
                    if no_match_mask.any():
                        df_result[no_match_mask].to_excel(writer, index=False, sheet_name="Sans correspondance")
                    if excluded_count > 0:
                        st.session_state["rec_df_b"][~mask_keep_b].to_excel(writer, index=False, sheet_name="Exclusions B")
                st.download_button("⬇️ Télécharger (.xlsx)", buf.getvalue(),
                                   "reconciliation_result.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   type="primary")
            with dc2:
                st.download_button("⬇️ Télécharger (.csv)",
                                   df_result.to_csv(index=False).encode("utf-8-sig"),
                                   "reconciliation_result.csv", "text/csv")
            with dc3:
                st.download_button("💾 Sauvegarder config", export_config(),
                                   "herotool_config.json", "application/json")

        except Exception as e:
            st.error(f"Erreur : {e}")
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
