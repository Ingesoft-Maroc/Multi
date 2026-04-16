import streamlit as st

st.set_page_config(
    page_title="HeroTool",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- CSS ----------
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background: #1a1a2e;
    }
    [data-testid="stSidebar"] * {
        color: #e0e0e0 !important;
    }
    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        background: linear-gradient(135deg, #6366f1, #8b5cf6, #ec4899);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0;
    }
    .hero-subtitle {
        color: #9ca3af;
        font-size: 1rem;
        margin-top: 0;
    }
    .tool-card {
        border: 1px solid #374151;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 0.5rem 0;
        cursor: pointer;
        transition: all 0.2s;
    }
    .tool-card:hover {
        border-color: #6366f1;
        background: #1e1b4b22;
    }
    div[data-testid="stButton"] > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s;
    }
</style>
""", unsafe_allow_html=True)

# ---------- Session state ----------
if "active_tool" not in st.session_state:
    st.session_state.active_tool = None

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("## ⚡ HeroTool")
    st.markdown("---")
    st.markdown("### 🧰 Mini-outils")

    tools = {
        "reconciler": ("🔗", "Réconciliateur"),
        "lettrage":   ("🏷️", "Lettrage"),
    }

    for key, (icon, label) in tools.items():
        if st.button(f"{icon} {label}", key=f"nav_{key}", use_container_width=True):
            st.session_state.active_tool = key

    st.markdown("---")
    if st.button("📖 Manuel d'utilisation", key="nav_manual", use_container_width=True):
        st.session_state.active_tool = "manual"

    if st.session_state.active_tool:
        if st.button("🏠 Accueil", use_container_width=True):
            st.session_state.active_tool = None

# ---------- Main content ----------
if st.session_state.active_tool is None:
    st.markdown('<p class="hero-title">⚡ HeroTool</p>', unsafe_allow_html=True)
    st.markdown('<p class="hero-subtitle">Votre boîte à outils pour les tâches répétitives</p>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### Choisissez un outil :")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="tool-card">
            <h3>🔗 Réconciliateur</h3>
            <p style="color:#9ca3af;">Rapprochez deux sources Excel avec matching avancé,
            colonnes calculées, conditions et sauvegarde de configuration.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ouvrir →", key="open_reconciler"):
            st.session_state.active_tool = "reconciler"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="tool-card">
            <h3>🏷️ Lettrage</h3>
            <p style="color:#9ca3af;">Attribuez des codes de lettrage (A/A, B/B...)
            aux lignes correspondantes entre deux feuilles selon vos critères.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ouvrir →", key="open_lettrage"):
            st.session_state.active_tool = "lettrage"
            st.rerun()

    with col3:
        st.markdown("""
        <div class="tool-card">
            <h3>📖 Manuel</h3>
            <p style="color:#9ca3af;">Guide complet d'utilisation de HeroTool
            avec exemples et conseils pratiques.</p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Lire →", key="open_manual"):
            st.session_state.active_tool = "manual"
            st.rerun()

elif st.session_state.active_tool == "reconciler":
    from tools.reconciler import run_reconciler
    run_reconciler()

elif st.session_state.active_tool == "lettrage":
    from tools.lettrage import run_lettrage
    run_lettrage()

elif st.session_state.active_tool == "manual":
    st.markdown("## 📖 Manuel d'utilisation — HeroTool")
    st.markdown("---")
    try:
        with open("MANUEL.txt", "r", encoding="utf-8") as f:
            content = f.read()
        st.text_area("", value=content, height=700, label_visibility="collapsed")
        with open("MANUEL.txt", "rb") as f:
            st.download_button("⬇️ Télécharger le manuel (.txt)", f.read(),
                               "MANUEL_HeroTool.txt", "text/plain")
    except FileNotFoundError:
        st.error("Fichier MANUEL.txt introuvable.")
