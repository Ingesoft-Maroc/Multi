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
        st.error(f"Erreur lecture fichier : {e}")
        return None


def get_sheets(file_bytes):
    try:
        return pd.ExcelFile(file_bytes).sheet_names
    except Exception:
        return []


def reset_lettrage():
    for k in [k for k in st.session_state if k.startswith("let_")]:
        del st.session_state[k]


LET_CONFIG_KEYS = [
    "let_criteria", "let_col_name", "let_fmt", "let_prefix",
    "let_one_to_one", "let_unmatched", "let_cols_a_out", "let_cols_b_out",
]

def export_let_config():
    cfg = {k.replace("let_", ""): st.session_state.get(k) for k in LET_CONFIG_KEYS}
    return json.dumps(cfg, ensure_ascii=False, indent=2)

def import_let_config(cfg: dict):
    for k in LET_CONFIG_KEYS:
        short = k.replace("let_", "")
        if short in cfg:
            st.session_state[k] = cfg[short]


def step_badge(n, label, active):
    color = "#8b5cf6" if active else "#374151"
    txt = "#fff" if active else "#9ca3af"
    st.markdown(
        f'<span style="background:{color};color:{txt};border-radius:20px;'
        f'padding:3px 14px;font-size:0.85rem;font-weight:600;margin-right:6px;">'
        f'Étape {n} — {label}</span>',
        unsafe_allow_html=True,
    )


def generate_label(n: int, fmt: str, prefix: str) -> str:
    """Generate A, B, ..., Z, AA, AB, ... or 001, 002, ... style labels."""
    if fmt == "numérique":
        return f"{prefix}{n+1:04d}"
    # Alphabétique
    result = ""
    idx = n
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return prefix + result


def row_matches_criteria(row_a, row_b, criteria):
    """Return True if row_a and row_b match on ALL criteria."""
    for crit in criteria:
        col_a = crit.get("col_a", "")
        col_b = crit.get("col_b", "")
        mtype = crit.get("match_type", "exact")
        case_s = crit.get("case_sensitive", False)

        val_a = str(row_a.get(col_a, "")) if col_a else ""
        val_b = str(row_b.get(col_b, "")) if col_b else ""

        if val_a in ("nan", "None", ""):
            return False

        if mtype == "exact":
            match = (val_a == val_b) if case_s else (val_a.lower() == val_b.lower())
        elif mtype == "contient (A dans B)":
            match = (val_a in val_b) if case_s else (val_a.lower() in val_b.lower())
        elif mtype == "contient (B dans A)":
            match = (val_b in val_a) if case_s else (val_b.lower() in val_a.lower())
        elif mtype == "numérique exact":
            try:
                match = float(val_a) == float(val_b)
            except Exception:
                match = False
        elif mtype == "numérique tolérance %":
            tol = float(crit.get("tolerance", 0))
            try:
                fa, fb = float(val_a), float(val_b)
                base = abs(fa) if fa != 0 else 1
                match = abs(fa - fb) / base * 100 <= tol
            except Exception:
                match = False
        elif mtype == "numérique tolérance abs":
            tol = float(crit.get("tolerance", 0))
            try:
                match = abs(float(val_a) - float(val_b)) <= tol
            except Exception:
                match = False
        else:
            match = False

        if not match:
            return False
    return True


def run_lettrage_engine(df_a, df_b, criteria, one_to_one, unmatched_marker,
                        fmt, prefix, col_name):
    """Assign letter labels to matched rows between df_a and df_b."""
    labels_a = [""] * len(df_a)
    labels_b = [""] * len(df_b)
    used_b = set()
    counter = 0

    for i, (_, row_a) in enumerate(df_a.iterrows()):
        matched_b = []
        for j, (_, row_b) in enumerate(df_b.iterrows()):
            if one_to_one and j in used_b:
                continue
            if row_matches_criteria(row_a, row_b, criteria):
                matched_b.append(j)

        if matched_b:
            label = generate_label(counter, fmt, prefix)
            counter += 1
            labels_a[i] = label
            for j in matched_b:
                labels_b[j] = label
                if one_to_one:
                    used_b.add(j)

    result_a = df_a.copy()
    result_b = df_b.copy()
    result_a[col_name] = labels_a
    result_b[col_name] = labels_b

    if unmatched_marker:
        result_a[col_name] = result_a[col_name].replace("", unmatched_marker)
        result_b[col_name] = result_b[col_name].replace("", unmatched_marker)

    return result_a, result_b, counter


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def run_lettrage():
    st.markdown("## 🏷️ Lettrage")
    st.markdown(
        "Attribuez des codes de lettrage (A, A / B, B / ...) aux lignes correspondantes "
        "entre deux sources de données selon vos critères."
    )
    st.markdown("---")

    step = st.session_state.get("let_step", 1)
    labels = ["Import", "Critères", "Options", "Résultat"]
    cols_s = st.columns(4)
    for i, (c, lbl) in enumerate(zip(cols_s, labels), 1):
        with c:
            step_badge(i, lbl, i == step)
    st.markdown("<br>", unsafe_allow_html=True)

    # ══════════════════════════════════════════
    #  ÉTAPE 1 — Import
    # ══════════════════════════════════════════
    if step == 1:
        st.markdown("### 📂 Étape 1 — Sources de données")

        input_mode = st.radio(
            "Mode", ["📄 Deux fichiers Excel séparés", "📑 Un fichier Excel multi-feuilles"],
            horizontal=True, key="let_input_mode", label_visibility="collapsed",
        )

        df_a, df_b, name_a, name_b = None, None, None, None

        if input_mode == "📄 Deux fichiers Excel séparés":
            ca, cb = st.columns(2)
            with ca:
                st.markdown("**Fichier / Feuille A**")
                file_a = st.file_uploader("A", type=["xlsx", "xls"], key="let_upload_a",
                                          label_visibility="collapsed")
            with cb:
                st.markdown("**Fichier / Feuille B**")
                file_b = st.file_uploader("B", type=["xlsx", "xls"], key="let_upload_b",
                                          label_visibility="collapsed")
            if file_a:
                df_a = load_excel(file_a); name_a = file_a.name
            if file_b:
                df_b = load_excel(file_b); name_b = file_b.name

        else:
            st.markdown("**Fichier Excel (plusieurs feuilles)**")
            multi_file = st.file_uploader("Fichier", type=["xlsx", "xls"],
                                          key="let_upload_multi", label_visibility="collapsed")
            if multi_file:
                file_bytes = multi_file.read()
                sheets = get_sheets(io.BytesIO(file_bytes))
                if sheets:
                    ca, cb = st.columns(2)
                    with ca:
                        st.markdown("**Feuille A**")
                        sheet_a = st.selectbox("Feuille A", sheets, key="let_sheet_a",
                                               label_visibility="collapsed")
                    with cb:
                        st.markdown("**Feuille B**")
                        sheet_b = st.selectbox("Feuille B", sheets,
                                               index=min(1, len(sheets) - 1),
                                               key="let_sheet_b", label_visibility="collapsed")
                    df_a = load_excel(io.BytesIO(file_bytes), sheet=sheet_a)
                    df_b = load_excel(io.BytesIO(file_bytes), sheet=sheet_b)
                    name_a = f"{multi_file.name} › {sheet_a}"
                    name_b = f"{multi_file.name} › {sheet_b}"

        if df_a is not None and df_b is not None:
            st.session_state.update({
                "let_df_a": df_a, "let_df_b": df_b,
                "let_name_a": name_a, "let_name_b": name_b,
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
                st.caption("Importez un fichier .json généré par HeroTool Lettrage pour restaurer tous vos paramètres.")
                cfg_file = st.file_uploader("Configuration (.json)", type=["json"],
                                            key="let_cfg_upload", label_visibility="collapsed")
                if cfg_file:
                    try:
                        cfg = json.load(cfg_file)
                        if st.button("✅ Appliquer cette configuration", type="primary"):
                            import_let_config(cfg)
                            st.session_state["let_step"] = 4
                            st.rerun()
                    except Exception as e:
                        st.error(f"Impossible de lire la configuration : {e}")

            st.markdown("")
            if st.button("Suivant →", type="primary"):
                if "let_criteria" not in st.session_state:
                    st.session_state["let_criteria"] = []
                st.session_state["let_step"] = 2; st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 2 — Critères de lettrage
    # ══════════════════════════════════════════
    elif step == 2:
        df_a = st.session_state["let_df_a"]
        df_b = st.session_state["let_df_b"]
        name_a, name_b = st.session_state["let_name_a"], st.session_state["let_name_b"]

        st.markdown("### 🔑 Étape 2 — Critères de correspondance")
        st.markdown(
            "Définissez **un ou plusieurs critères** (tous doivent être vrais simultanément — logique ET). "
            "Deux lignes sont lettrées ensemble si elles vérifient tous les critères."
        )

        MATCH_TYPES = [
            "exact",
            "contient (A dans B)",
            "contient (B dans A)",
            "numérique exact",
            "numérique tolérance %",
            "numérique tolérance abs",
        ]

        criteria = st.session_state.get("let_criteria", [])

        for ci, crit in enumerate(criteria):
            with st.expander(f"📌 Critère {ci+1}", expanded=True):
                cc1, cc2 = st.columns(2)
                with cc1:
                    st.markdown(f"**Colonne dans {name_a}**")
                    crit["col_a"] = st.selectbox(
                        "Col A", list(df_a.columns),
                        index=list(df_a.columns).index(crit["col_a"])
                              if crit.get("col_a") in df_a.columns else 0,
                        key=f"let_crit_cola_{ci}", label_visibility="collapsed")
                with cc2:
                    st.markdown(f"**Colonne dans {name_b}**")
                    crit["col_b"] = st.selectbox(
                        "Col B", list(df_b.columns),
                        index=list(df_b.columns).index(crit["col_b"])
                              if crit.get("col_b") in df_b.columns else 0,
                        key=f"let_crit_colb_{ci}", label_visibility="collapsed")

                cm1, cm2, cm3 = st.columns([2, 2, 1])
                with cm1:
                    crit["match_type"] = st.selectbox(
                        "Type de correspondance", MATCH_TYPES,
                        index=MATCH_TYPES.index(crit["match_type"])
                              if crit.get("match_type") in MATCH_TYPES else 0,
                        key=f"let_crit_mtype_{ci}")
                with cm2:
                    if "tolérance" in crit.get("match_type", ""):
                        unit = "%" if "%" in crit["match_type"] else "valeur absolue"
                        crit["tolerance"] = st.number_input(
                            f"Tolérance ({unit})", min_value=0.0, value=float(crit.get("tolerance", 0)),
                            key=f"let_crit_tol_{ci}")
                    elif crit.get("match_type") in ["exact", "contient (A dans B)", "contient (B dans A)"]:
                        crit["case_sensitive"] = st.checkbox(
                            "Respecter la casse", value=crit.get("case_sensitive", False),
                            key=f"let_crit_case_{ci}")

                with cm3:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️", key=f"let_del_crit_{ci}"):
                        criteria.pop(ci)
                        st.session_state["let_criteria"] = criteria; st.rerun()

                # Preview
                if crit.get("col_a") and crit.get("col_b"):
                    sample_a = df_a[crit["col_a"]].dropna().iloc[0] if len(df_a[crit["col_a"]].dropna()) > 0 else "—"
                    sample_b = df_b[crit["col_b"]].dropna().iloc[0] if len(df_b[crit["col_b"]].dropna()) > 0 else "—"
                    match_test = row_matches_criteria(
                        {crit["col_a"]: sample_a}, {crit["col_b"]: sample_b}, [crit])
                    st.caption(f"Test rapide : A=`{sample_a}` vs B=`{sample_b}` → {'✅ match' if match_test else '❌ pas de match'}")

        st.session_state["let_criteria"] = criteria

        if st.button("➕ Ajouter un critère"):
            criteria.append({
                "col_a": list(df_a.columns)[0],
                "col_b": list(df_b.columns)[0],
                "match_type": "exact",
                "case_sensitive": False,
                "tolerance": 0,
            })
            st.session_state["let_criteria"] = criteria; st.rerun()

        if not criteria:
            st.info("Ajoutez au moins un critère de correspondance.")

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["let_step"] = 1; st.rerun()
        with c2:
            if criteria and st.button("Suivant →", type="primary"):
                st.session_state["let_step"] = 3; st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 3 — Options de lettrage
    # ══════════════════════════════════════════
    elif step == 3:
        df_a = st.session_state["let_df_a"]
        df_b = st.session_state["let_df_b"]
        name_a, name_b = st.session_state["let_name_a"], st.session_state["let_name_b"]

        st.markdown("### ⚙️ Étape 3 — Options de lettrage")

        # ── SECTION 1 : Colonne de lettrage ───
        with st.expander("🏷️ Colonne de lettrage", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                col_name = st.text_input(
                    "Nom de la colonne",
                    value=st.session_state.get("let_col_name", "Lettrage"),
                    key="let_col_name_w",
                    help="Ce nom sera ajouté comme nouvelle colonne dans les deux sources.")
            with c2:
                fmt_opts = ["Alphabétique — A, B, C, ..., Z, AA, AB...", "Numérique — 0001, 0002, 0003..."]
                fmt_saved = st.session_state.get("let_fmt", "alphabétique")
                fmt_default = 1 if fmt_saved == "numérique" else 0
                fmt = st.selectbox("Format des codes", fmt_opts, index=fmt_default, key="let_fmt_w")
            with c3:
                prefix = st.text_input(
                    "Préfixe (optionnel)",
                    value=st.session_state.get("let_prefix", ""),
                    key="let_prefix_w",
                    placeholder="ex : LT-, 2024-, REF-",
                    help="Ajouté devant chaque code. Exemple : préfixe 'LT-' → LT-A, LT-B, LT-0001...")

            fmt_clean = "numérique" if "Numérique" in fmt else "alphabétique"
            preview_labels = [generate_label(i, fmt_clean, prefix) for i in range(5)]
            st.caption(f"Aperçu des codes générés : **{' · '.join(preview_labels)} · ...**")

        st.markdown("")

        # ── SECTION 2 : Mode de correspondance ─
        with st.expander("🔄 Mode de correspondance entre A et B", expanded=True):
            mode_opts = [
                "1-à-plusieurs  —  une ligne B peut être lettrée par plusieurs lignes A",
                "1-à-1  —  une ligne B ne peut être lettrée qu'une seule fois (par la première ligne A qui la matche)",
            ]
            saved_oto = st.session_state.get("let_one_to_one", True)
            mode_default = 1 if saved_oto else 0
            mode_choice = st.selectbox("Mode", mode_opts, index=mode_default, key="let_mode_w",
                                       label_visibility="collapsed")
            one_to_one = "1-à-1" in mode_choice

            if one_to_one:
                st.info("Mode 1-à-1 : pour chaque ligne de A (dans l'ordre), on cherche des lignes de B disponibles. "
                        "Une fois qu'une ligne B est lettrée, elle est réservée et ne peut plus être prise par une autre ligne A.")
            else:
                st.info("Mode 1-à-plusieurs : une même ligne de B peut recevoir plusieurs lettres différentes "
                        "si elle correspond à plusieurs lignes de A. Utile quand une écriture B fait référence à plusieurs A.")

        st.markdown("")

        # ── SECTION 3 : Lignes sans correspondance ─
        with st.expander("❓ Lignes sans correspondance", expanded=True):
            preset_opts = ["(vide — cellule vide)", "NL", "—", "?", "Non lettré", "Personnalisé..."]
            saved_marker = st.session_state.get("let_unmatched", "")
            if saved_marker in ["NL", "—", "?", "Non lettré"]:
                preset_default = preset_opts.index(saved_marker)
            elif saved_marker == "":
                preset_default = 0
            else:
                preset_default = preset_opts.index("Personnalisé...")

            c1, c2 = st.columns(2)
            with c1:
                marker_choice = st.selectbox(
                    "Marqueur pour les lignes non lettrées", preset_opts,
                    index=preset_default, key="let_marker_preset_w",
                    help="Valeur affichée dans la colonne de lettrage pour les lignes sans correspondance.")
            with c2:
                if marker_choice == "Personnalisé...":
                    unmatched_marker = st.text_input(
                        "Valeur personnalisée", value=saved_marker if preset_default == len(preset_opts)-1 else "",
                        key="let_unmatched_w", placeholder="ex : HORS SCOPE, NA, ...")
                elif marker_choice == "(vide — cellule vide)":
                    unmatched_marker = ""
                    st.caption("Les lignes non lettrées auront une cellule vide.")
                else:
                    unmatched_marker = marker_choice
                    st.caption(f'Les lignes non lettrées afficheront : **"{unmatched_marker}"**')

        st.markdown("")

        # ── SECTION 4 : Colonnes de sortie ────
        with st.expander("📋 Colonnes à inclure dans la sortie", expanded=False):
            st.caption("Choisissez quelles colonnes de chaque source apparaissent dans le résultat.")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**{name_a}**")
                cols_a_out = st.multiselect(
                    f"Colonnes A", list(df_a.columns),
                    default=st.session_state.get("let_cols_a_out", list(df_a.columns)),
                    key="let_cols_a_out_w", label_visibility="collapsed")
            with c2:
                st.markdown(f"**{name_b}**")
                cols_b_out = st.multiselect(
                    f"Colonnes B", list(df_b.columns),
                    default=st.session_state.get("let_cols_b_out", list(df_b.columns)),
                    key="let_cols_b_out_w", label_visibility="collapsed")

        # ── Sauvegarde config ─────────────────
        st.markdown("---")
        if st.session_state.get("let_criteria"):
            st.session_state["let_col_name"] = col_name
            st.session_state["let_fmt"] = fmt_clean
            st.session_state["let_prefix"] = prefix
            st.session_state["let_one_to_one"] = one_to_one
            st.session_state["let_unmatched"] = unmatched_marker
            st.session_state["let_cols_a_out"] = cols_a_out
            st.session_state["let_cols_b_out"] = cols_b_out
            st.download_button("💾 Sauvegarder la configuration", data=export_let_config(),
                               file_name="herotool_lettrage_config.json", mime="application/json")

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Retour"):
                st.session_state["let_step"] = 2; st.rerun()
        with c2:
            if st.button("Lettrer →", type="primary"):
                st.session_state["let_col_name"] = col_name
                st.session_state["let_fmt"] = fmt_clean
                st.session_state["let_prefix"] = prefix
                st.session_state["let_one_to_one"] = one_to_one
                st.session_state["let_unmatched"] = unmatched_marker
                st.session_state["let_cols_a_out"] = cols_a_out
                st.session_state["let_cols_b_out"] = cols_b_out
                st.session_state["let_step"] = 4; st.rerun()

    # ══════════════════════════════════════════
    #  ÉTAPE 4 — Résultat
    # ══════════════════════════════════════════
    elif step == 4:
        df_a = st.session_state["let_df_a"].copy()
        df_b = st.session_state["let_df_b"].copy()
        name_a, name_b = st.session_state["let_name_a"], st.session_state["let_name_b"]
        criteria = st.session_state.get("let_criteria", [])
        col_name = st.session_state.get("let_col_name", "Lettrage")
        fmt = st.session_state.get("let_fmt", "alphabétique")
        prefix = st.session_state.get("let_prefix", "")
        one_to_one = st.session_state.get("let_one_to_one", True)
        unmatched_marker = st.session_state.get("let_unmatched", "")
        cols_a_out = st.session_state.get("let_cols_a_out", list(df_a.columns))
        cols_b_out = st.session_state.get("let_cols_b_out", list(df_b.columns))

        st.markdown("### ✅ Étape 4 — Résultat du lettrage")

        with st.spinner("Calcul du lettrage en cours..."):
            result_a, result_b, nb_letters = run_lettrage_engine(
                df_a[cols_a_out], df_b[cols_b_out],
                criteria, one_to_one, unmatched_marker, fmt, prefix, col_name
            )

        # ── Stats ─────────────────────────────
        lettered_a = int((result_a[col_name] != (unmatched_marker or "")).sum())
        lettered_b = int((result_b[col_name] != (unmatched_marker or "")).sum())

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Groupes de lettrage", nb_letters)
        m2.metric(f"Lignes lettrées dans A", lettered_a)
        m3.metric(f"Lignes lettrées dans B", lettered_b)
        m4.metric("Non lettrées (A)", len(result_a) - lettered_a)

        st.markdown("---")

        tab_a, tab_b, tab_combined = st.tabs([f"📄 {name_a}", f"📄 {name_b}", "🔀 Vue combinée"])

        with tab_a:
            # Move lettrage col to front
            cols_order = [col_name] + [c for c in result_a.columns if c != col_name]
            st.dataframe(result_a[cols_order], use_container_width=True, height=400)

        with tab_b:
            cols_order = [col_name] + [c for c in result_b.columns if c != col_name]
            st.dataframe(result_b[cols_order], use_container_width=True, height=400)

        with tab_combined:
            st.caption("Toutes les lignes de A et B, triées par code de lettrage.")
            ra = result_a.copy(); ra["__source__"] = name_a
            rb = result_b.copy(); rb["__source__"] = name_b
            combined = pd.concat([ra, rb], ignore_index=True).sort_values(
                col_name, na_position="last")
            st.dataframe(combined, use_container_width=True, height=400)

        # ── Téléchargement ────────────────────
        st.markdown("---")
        dc1, dc2, dc3, dc4 = st.columns(4)

        with dc1:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                result_a[[col_name] + [c for c in result_a.columns if c != col_name]].to_excel(
                    writer, index=False, sheet_name="Lettrage A")
                result_b[[col_name] + [c for c in result_b.columns if c != col_name]].to_excel(
                    writer, index=False, sheet_name="Lettrage B")
                combined.to_excel(writer, index=False, sheet_name="Vue combinée")
            st.download_button("⬇️ Télécharger (.xlsx)", buf.getvalue(),
                               "lettrage_result.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary")
        with dc2:
            csv_a = result_a.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ CSV Feuille A", csv_a, "lettrage_A.csv", "text/csv")
        with dc3:
            csv_b = result_b.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ CSV Feuille B", csv_b, "lettrage_B.csv", "text/csv")
        with dc4:
            st.download_button("💾 Sauvegarder config", export_let_config(),
                               "herotool_lettrage_config.json", "application/json")

        st.markdown("")
        c1, c2 = st.columns([1, 5])
        with c1:
            if st.button("← Modifier"):
                st.session_state["let_step"] = 3; st.rerun()
        with c2:
            if st.button("🔄 Nouveau lettrage"):
                reset_lettrage(); st.rerun()
