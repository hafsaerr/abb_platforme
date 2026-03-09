# app.py - OPCVM Manager – Albarid Bank
# Run: streamlit run app.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date
import io
import base64
from pathlib import Path

# ─── AUTH CREDENTIALS ─────────────────────────────────────────────────────────
AUTH_USERNAME = "opcvmABB"
AUTH_PASSWORD = "albarid2026"

from config import COLORS, CATEGORY_CODES
from storage import save_upload, load_all_keys, load_by_key, load_latest, get_historical_series, count_entries
from data_processor import process_sfin_file
from analytics import get_market_df, get_abb_df, compute_tableau1, compute_tableau2, compute_tableau3
from exporter import generate_excel

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="OPCVM Manager – Albarid Bank",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto",
)

# ─── CUSTOM CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
/* ── Global background ──────────────────────────── */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"],
.main, section.main,
[data-testid="stMainBlockContainer"] {
    background: #100804 !important;
    color: #EDE0D4 !important;
}

/* ── ALL text → beige/cream ─────────────────────── */
*, p, span, div, label, li, td, th, h1, h2, h3, h4, h5, h6 {
    color: #EDE0D4 !important;
}

/* ── Sidebar ────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1E0A03 0%, #130602 100%) !important;
    border-right: 1px solid rgba(200,80,30,0.25) !important;
}
[data-testid="stSidebar"] * { color: #EDE0D4 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(200,80,30,0.3) !important;
    color: #EDE0D4 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    border: 2px dashed rgba(200,80,30,0.5) !important;
    border-radius: 8px !important;
    background: rgba(255,255,255,0.03) !important;
}
/* Text inside file uploader (white/light background) → black */
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] p,
[data-testid="stFileUploaderDropzone"] span,
[data-testid="stFileUploaderDropzone"] small,
[data-testid="stFileUploaderDropzone"] p { color: #111111 !important; }
[data-testid="stFileUploaderDropzone"] button { color: #111111 !important; border-color: #444 !important; }
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] [data-baseweb="input"] {
    background: rgba(255,255,255,0.06) !important;
    color: #EDE0D4 !important;
    border: 1px solid rgba(200,80,30,0.3) !important;
}
/* Date picker in sidebar */
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background: rgba(255,255,255,0.06) !important;
    color: #EDE0D4 !important;
}

/* ── Sidebar nav buttons (custom class) ─────────── */
[data-testid="stSidebar"] .nav-btn > button {
    background: transparent !important;
    border: none !important;
    border-radius: 0 !important;
    color: #C8A888 !important;
    font-size: 17px !important;
    font-weight: 600 !important;
    letter-spacing: 0.5px !important;
    text-align: center !important;
    width: 100% !important;
    padding: 10px 0 !important;
    box-shadow: none !important;
    transition: color 0.2s, background 0.2s !important;
}
[data-testid="stSidebar"] .nav-btn > button:hover {
    background: rgba(200,80,30,0.12) !important;
    color: #FF9060 !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] .nav-btn-active > button {
    background: rgba(200,80,30,0.22) !important;
    color: #FFFFFF !important;
    border-left: 3px solid #C8501E !important;
    border-radius: 6px !important;
    font-weight: 800 !important;
}

/* ── Cards ──────────────────────────────────────── */
.card {
    background: rgba(255,255,255,0.05);
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 16px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.5);
    border-left: 4px solid #C8501E;
}

/* ── Section headers ────────────────────────────── */
.section-header {
    background: linear-gradient(90deg, #C8501E 0%, #7A2E08 100%);
    color: #FFFFFF !important;
    padding: 9px 18px;
    border-radius: 6px;
    font-size: 16px;
    font-weight: 700;
    margin-bottom: 14px;
    letter-spacing: 0.5px;
}
.section-header * { color: #FFFFFF !important; }

/* ── KPI boxes ──────────────────────────────────── */
.kpi-box {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(200,80,30,0.3);
    border-top: 3px solid #C8501E;
    border-radius: 8px;
    padding: 14px 16px;
    text-align: center;
}
.kpi-title { font-size: 11px; color: #A89070 !important; font-weight: 600;
             text-transform: uppercase; letter-spacing: 0.5px; }
.kpi-value { font-size: 22px; font-weight: 800; }

/* ── Tabs ───────────────────────────────────────── */
[data-baseweb="tab-list"] { background: rgba(0,0,0,0.3) !important; border-radius: 8px; }
[data-baseweb="tab"]      { color: #C8A888 !important; font-size: 14px !important; }
[aria-selected="true"]    { color: #FF9060 !important; border-bottom: 2px solid #C8501E !important; }

/* ── Alert / info / success ─────────────────────── */
.stAlert   { background: rgba(200,80,30,0.10) !important; border: 1px solid rgba(200,80,30,0.35) !important; }
.stInfo    { background: rgba(80,120,200,0.10) !important; border: 1px solid rgba(80,120,200,0.3) !important; }
.stSuccess { background: rgba(60,160,60,0.10) !important; border: 1px solid rgba(60,160,60,0.3) !important; }
[data-testid="stAlertContainer"] * { color: #EDE0D4 !important; }

/* ── Expander ───────────────────────────────────── */
[data-testid="stExpander"] {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(200,80,30,0.2) !important;
    border-radius: 8px !important;
}

/* ── Dataframe ──────────────────────────────────── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── Buttons (action) ───────────────────────────── */
.stButton > button {
    background: linear-gradient(90deg, #C8501E, #7A2E08) !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px !important;
    padding: 8px 18px !important;
}
.stButton > button:hover {
    background: linear-gradient(90deg, #E0621E, #9A3E10) !important;
}

/* ── Download button ────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #1A6B2E, #0D4019) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}

/* ── Selectbox / date inputs (main area) ─────────── */
[data-baseweb="select"] > div {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(200,80,30,0.3) !important;
    color: #EDE0D4 !important;
}
[data-baseweb="popover"] { background: #1E0A03 !important; }
[data-baseweb="menu"]    { background: #1E0A03 !important; }
[data-baseweb="option"]  { background: #1E0A03 !important; color: #EDE0D4 !important; }
[data-baseweb="option"]:hover { background: rgba(200,80,30,0.25) !important; }

/* ── White-bg elements → black text ─────────────── */
/* Login form inputs */
[data-testid="stForm"] input { color: #111111 !important; background: #FFFFFF !important; }
/* Sidebar date input */
[data-testid="stSidebar"] [data-testid="stDateInput"] input { color: #111111 !important; background: #FFFFFF !important; }
/* Selectbox selected value text */
[data-baseweb="select"] [data-baseweb="input"] input { color: #111111 !important; }
[data-baseweb="select"] span { color: #111111 !important; }
/* Browse files button text */
[data-testid="stFileUploaderDropzone"] * { color: #111111 !important; }

/* ── Spinner / caption ──────────────────────────── */
.stCaption { color: #8A6A50 !important; }

/* ── Hide Streamlit branding ────────────────────── */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
header    { visibility: hidden; }

/* ── Mobile responsive ──────────────────────────── */
@media (max-width: 768px) {
    /* Login: make all columns full-width so form is visible */
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
        min-width: 90vw !important;
        flex: 1 1 90vw !important;
    }
    /* Shrink sidebar nav font on small screens */
    [data-testid="stSidebar"] .nav-btn > button {
        font-size: 15px !important;
        padding: 8px 0 !important;
    }
    /* KPI boxes: allow wrapping */
    .kpi-box { margin-bottom: 10px; }
    /* Dataframe: enable horizontal scroll */
    [data-testid="stDataFrame"] { overflow-x: auto !important; }
    /* Reduce main padding */
    [data-testid="stMainBlockContainer"] {
        padding-left: 12px !important;
        padding-right: 12px !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ─── LOGO HELPER ──────────────────────────────────────────────────────────────

def _logo_b64() -> str:
    logo_path = Path(__file__).parent / "assets" / "ALBARID.png"
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


# ─── LOGIN PAGE ───────────────────────────────────────────────────────────────

def render_login():
    """Full-screen login page. Sets st.session_state['authenticated'] = True on success."""

    logo_b64 = _logo_b64()
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" '
        f'style="width:90px;height:90px;object-fit:contain;border-radius:16px;'
        f'box-shadow:0 4px 20px rgba(0,0,0,0.5);margin-bottom:12px;" />'
        if logo_b64 else
        '<div style="font-size:48px;margin-bottom:12px;">🏦</div>'
    )

    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="stForm"] {
        background: rgba(10,5,2,0.75) !important;
        border: 1px solid rgba(200,80,30,0.25) !important;
        border-radius: 20px !important;
        padding: 10px 20px 20px 20px !important;
    }
    [data-testid="stForm"] input {
        background: rgba(255,255,255,0.08) !important;
        color: #EDE0D4 !important;
        border: 1px solid rgba(200,80,30,0.3) !important;
        border-radius: 8px !important;
    }
    [data-testid="stForm"] label { color: #C8A888 !important; font-weight: 600 !important; }
    [data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(90deg,#C8501E,#7A2E08) !important;
        color: white !important; border: none !important;
        border-radius: 10px !important; font-weight: 700 !important;
        font-size: 15px !important; padding: 10px !important;
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # Centered layout — uses CSS max-width so it works on both mobile and desktop
    st.markdown("""
    <style>
    .login-wrapper {
        max-width: 420px;
        margin: 0 auto;
        padding: 10px;
    }
    </style>
    <div class="login-wrapper">
    """, unsafe_allow_html=True)

    # Card header (logo + title)
    st.markdown(f"""
    <div style="text-align:center; padding: 30px 0 20px 0;">
        {logo_html}
        <div style="font-size:26px;font-weight:900;color:#FFFFFF;letter-spacing:0.5px;margin-bottom:4px;">
            Suivi des OPCVM
        </div>
        <div style="font-size:16px;font-weight:700;color:#F5C518;margin-bottom:4px;">
            Al Barid Bank
        </div>
        <div style="font-size:12px;color:#A89888;margin-bottom:20px;">
            Connexion sécurisée à la plateforme interne
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Username", placeholder="opcvmABB", key="login_user")
        password = st.text_input("Code", type="password", placeholder="••••••••", key="login_pass")
        submit   = st.form_submit_button("Se connecter", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    if submit:
        if username == AUTH_USERNAME and password == AUTH_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Identifiants incorrects. Réessayez.")


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

def render_sidebar():
    with st.sidebar:
        # Logo / Brand
        logo_b64 = _logo_b64()
        if logo_b64:
            st.markdown(f"""
            <div style="text-align:center;padding:14px 0 6px 0;">
                <img src="data:image/png;base64,{logo_b64}"
                     style="width:70px;height:70px;object-fit:contain;border-radius:12px;
                            box-shadow:0 2px 12px rgba(0,0,0,0.5);" />
            </div>
            """, unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center; padding: 6px 0 16px 0;">
            <div style="font-size:15px; font-weight:900; color:#F5C518; letter-spacing:2px;">Al Barid Bank</div>
            <div style="font-size:11px; color:#8A7060; margin-top:2px;">OPCVM Manager</div>
            <hr style="border-color:rgba(200,80,30,0.3); margin:12px 0 8px 0;">
        </div>
        """, unsafe_allow_html=True)

        # Navigation – styled buttons, no emojis, centered
        st.markdown("<div style='margin-bottom:4px;'></div>", unsafe_allow_html=True)
        nav_items = ["Accueil", "OCT", "OMLT", "Diversifié", "Export"]
        current_page = st.session_state.get("_nav_active", "Accueil")
        for item in nav_items:
            is_active = (current_page == item)
            css_class = "nav-btn-active" if is_active else "nav-btn"
            with st.container():
                st.markdown(f'<div class="{css_class}">', unsafe_allow_html=True)
                if st.button(item, key=f"nav_{item}", use_container_width=True):
                    st.session_state["_nav_active"] = item
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        page = st.session_state.get("_nav_active", "Accueil")

        st.markdown("<hr style='border-color:rgba(200,80,30,0.2); margin:14px 0;'>", unsafe_allow_html=True)

        # Upload section
        st.markdown("<span style='font-size:13px;font-weight:700;color:#FFFFFF;'>Charger un fichier ASFIM</span>",
                    unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Fichier Excel ASFIM",
            type=["xlsx", "xls"],
            key="sfin_uploader",
            label_visibility="collapsed",
        )

        periodicity = st.selectbox(
            "Périodicité",
            ["Quotidien", "Hebdomadaire"],
            key="w_periodicity",
        )

        report_date = st.date_input(
            "Date du rapport",
            value=date.today(),
            key="w_report_date",
        )

        process_btn = st.button("⚡ Traiter le fichier", use_container_width=True)

        # Previous uploads
        st.markdown("<hr style='border-color:rgba(200,80,30,0.2); margin:14px 0;'>", unsafe_allow_html=True)
        st.markdown("<span style='font-size:13px;font-weight:700;color:#FFFFFF;'>Fichiers enregistrés</span>",
                    unsafe_allow_html=True)
        keys = load_all_keys()
        if keys:
            key_labels = [f"{k['date']} – {k['periodicity'].upper()}" for k in keys]
            sel_label = st.selectbox("Charger depuis l'historique",
                                     ["(nouveau fichier)"] + key_labels,
                                     key="history_select")
            if sel_label != "(nouveau fichier)":
                sel_idx = key_labels.index(sel_label)
                st.session_state["_selected_hist_key"] = keys[sel_idx]["key"]
        else:
            st.caption("Aucun fichier enregistré.")
            st.session_state["_selected_hist_key"] = None

        st.markdown(f"<div style='color:#7A5A40;font-size:11px;margin-top:8px;'>{count_entries()} rapport(s) en mémoire</div>",
                    unsafe_allow_html=True)

        # Logout
        st.markdown("<hr style='border-color:rgba(200,80,30,0.2);margin:16px 0 8px 0;'>", unsafe_allow_html=True)
        if st.button("Déconnexion", use_container_width=True, key="logout_btn"):
            st.session_state["authenticated"] = False
            st.rerun()

    return page, uploaded, periodicity, str(report_date), process_btn


# ─── PROCESS FILE ─────────────────────────────────────────────────────────────

def process_and_store(uploaded, periodicity: str, report_date: date):
    with st.spinner("Traitement du fichier en cours…"):
        df_full, warnings = process_sfin_file(uploaded)
    if df_full.empty:
        st.error("Impossible de lire le fichier. Vérifiez le format.")
        for w in warnings:
            st.warning(w)
        return
    for w in warnings:
        st.warning(w)

    # Persist to disk
    save_upload(report_date, periodicity, df_full)
    st.session_state["df_full"]            = df_full
    st.session_state["active_periodicity"] = periodicity
    st.session_state["active_report_date"] = report_date.isoformat()
    st.success(f"✅ Fichier traité et sauvegardé ({len(df_full)} fonds détectés).")


def load_state_from_history(hist_key: str):
    entry = load_by_key(hist_key)
    if entry:
        st.session_state["df_full"]            = entry["df"]
        st.session_state["active_periodicity"] = entry["periodicity"]
        st.session_state["active_report_date"] = entry["date"]


def get_active_data() -> tuple[pd.DataFrame | None, str, str]:
    df  = st.session_state.get("df_full", None)
    prd = st.session_state.get("active_periodicity", "Quotidien")
    dt  = st.session_state.get("active_report_date", date.today().isoformat())
    return df, prd, dt


# ─── CATEGORY PAGE ────────────────────────────────────────────────────────────

def render_category_page(category: str):
    df, periodicity, report_date_str = get_active_data()

    st.markdown(f"""
    <div class="section-header">
        📊 Analyse {category} – {periodicity.upper()} – {report_date_str}
    </div>
    """, unsafe_allow_html=True)

    if df is None:
        st.info("Veuillez charger un fichier ASFIM depuis la barre latérale.")
        return

    df_market = get_market_df(df, category)
    df_abb    = get_abb_df(df, category, periodicity)

    if df_market.empty:
        st.warning(f"Aucun fonds {category} trouvé dans le fichier chargé.")
        return
    if df_abb.empty:
        st.warning(f"Aucun fonds ABB {category} ({periodicity}) trouvé dans ce fichier.")
        st.info(f"Codes attendus: {', '.join(CATEGORY_CODES[category][periodicity.lower()][:5])}…")
        return

    t1    = compute_tableau1(df_abb)
    synth = compute_tableau2(df_market, df_abb, category)
    t3    = compute_tableau3(df_market, df_abb, synth)

    # KPIs
    n_abb    = len(df_abb)
    n_market = len(df_market)
    avg_1j   = df_abb["perf_1j"].mean()
    avg_ytd  = df_abb["ytd"].mean()

    col1, col2, col3, col4 = st.columns(4)
    _kpi(col1, "Fonds ABB", str(n_abb), f"sur {n_market} du marché")
    _kpi(col2, "Perf. quotidienne moy.", f"{avg_1j:+.3f}%" if not np.isnan(avg_1j) else "N/A", "ABB")
    _kpi(col3, "YTD moyen ABB", f"{avg_ytd:+.3f}%" if not np.isnan(avg_ytd) else "N/A",
         "vert = positif", color="green" if not np.isnan(avg_ytd) and avg_ytd > 0 else "red")
    _kpi(col4, "YTD vs Marché",
         f"{(avg_ytd - synth.attrs.get('ytd_moyen_marche', np.nan)):+.3f}%"
         if not np.isnan(synth.attrs.get("ytd_moyen_marche", np.nan)) else "N/A",
         "ABB - marché")

    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["📋 Tableau 1 – Fonds", "📊 Tableau 2 – Synthèse", "🏆 Tableau 3 – Ranking", "📈 Courbe historique"]
    )

    # ── Tab 1
    with tab1:
        st.markdown('<div class="section-header">Fonds ABB</div>', unsafe_allow_html=True)
        display_t1 = t1.copy()
        display_t1.columns = [
            "Code ISIN", "OPCVM", "Société de Gestion", "Souscripteurs",
            "Périodicité VL", "Actif Net", "Classification",
            "Perf. quotidienne", "YTD",
        ][:len(display_t1.columns)]
        _style_t1(display_t1)

    # ── Tab 2
    with tab2:
        st.markdown('<div class="section-header">Synthèse du marché</div>', unsafe_allow_html=True)
        _style_t2(synth, category)

    # ── Tab 3
    with tab3:
        st.markdown('<div class="section-header">Classement relatif</div>', unsafe_allow_html=True)
        if t3.empty:
            st.info("Données insuffisantes pour le classement.")
        else:
            _style_t3(t3)

    # ── Tab 4
    with tab4:
        st.markdown('<div class="section-header">Évolution historique</div>', unsafe_allow_html=True)
        _render_history_chart(df_abb, periodicity, category)

    # ── Download button
    st.markdown("---")
    report_date = date.fromisoformat(report_date_str) if isinstance(report_date_str, str) else report_date_str
    xls_bytes = generate_excel(report_date, periodicity, category, t1, synth, t3)
    fname = f"ABB_{category}_{periodicity.upper()}_{report_date_str.replace('-','')}.xlsx"
    st.download_button(
        label=f"⬇️  Télécharger le rapport Excel – {category}",
        data=xls_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ─── TABLE STYLING ────────────────────────────────────────────────────────────

def _style_t1(df: pd.DataFrame):
    def color_pct(val, col_name=""):
        if not isinstance(val, (int, float)):
            try:
                val = float(str(val).replace("%", "").replace(",", ".").replace("+", ""))
            except Exception:
                return ""
        if np.isnan(val):
            return ""
        return "background-color: #92D050" if val > 0 else ("background-color: #FF6B6B" if val < 0 else "")

    pct_cols = [c for c in df.columns if c in ("Perf. quotidienne", "YTD")]

    styler = df.style
    for col in pct_cols:
        col_idx = df.columns.get_loc(col)
        styler = styler.applymap(lambda v: color_pct(v), subset=[col])

    # Format pct cols
    fmt = {c: "{:+.3f}%" for c in pct_cols}
    fmt["Actif Net"] = "{:,.2f}" if "Actif Net" in df.columns else None
    fmt = {k: v for k, v in fmt.items() if v and k in df.columns}

    # Apply header style
    styler = (
        styler
        .format(fmt, na_rep="N/A")
        .set_table_styles([
            {"selector": "thead th",
             "props": [("background-color", "#C8501E"), ("color", "white"),
                       ("font-weight", "bold"), ("text-align", "center")]},
            {"selector": "tbody tr:nth-child(even)",
             "props": [("background-color", "#F5F5F5")]},
        ])
        .hide(axis="index")
    )
    st.dataframe(styler, use_container_width=True, height=min(400, 50 + 35 * len(df)))


def _style_t2(synth: pd.DataFrame, category: str):
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"""
        <div style="background:#FAD7A0;padding:8px 14px;border-radius:6px;
                    font-weight:700;text-align:center;margin-bottom:8px;">
            {category} MARCHÉ
        </div>
        """, unsafe_allow_html=True)
        for _, row in synth.iterrows():
            ind = row["Indicateur"]
            val = row["Valeur"]
            det = row.get("Détail", "")
            # Colour value
            val_color = "#000"
            if isinstance(val, str) and "%" in val:
                try:
                    v = float(val.replace("%", "").replace("+", ""))
                    val_color = "#006400" if v > 0 else ("#CC0000" if v < 0 else "#000")
                except Exception:
                    pass
            st.markdown(f"""
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:6px 12px;margin-bottom:4px;border-radius:4px;
                        background:#F9F9F9;border:1px solid #E0E0E0;">
                <span style="font-weight:600;color:#333;">{ind}</span>
                <span style="font-weight:700;color:{val_color};">{val}</span>
                <span style="color:#006400;font-weight:600;">{det}</span>
            </div>
            """, unsafe_allow_html=True)


def _style_t3(t3: pd.DataFrame):
    pct_cols = ["Écart vs Plus perf.", "Écart vs Moyenne", "Écart vs Moins perf.",
                "YTD", "YTD Marché Moyen", "Écart vs Marché"]

    def color_cell(val, col):
        num = _to_float(val)
        if num is None:
            return ""
        if col in ("Écart vs Plus perf.", "Écart vs Moins perf."):
            return "background-color: #FFF2CC"
        if col in ("Écart vs Moyenne", "Écart vs Marché"):
            return "background-color: #92D050" if num > 0 else ("background-color: #FF6B6B" if num < 0 else "")
        if col in ("YTD", "YTD Marché Moyen"):
            return "background-color: #92D050" if num > 0 else "background-color: #FF6B6B"
        return ""

    def color_position(val):
        if val == "Top 25%":  return "background-color: #92D050; font-weight:bold"
        if val == "Bas 25%":  return "background-color: #FF6B6B; font-weight:bold"
        return "background-color: #FFF2CC"

    def color_note(val):
        try:
            n = int(val)
            if n == 10: return "background-color: #92D050; font-weight:bold"
            if n == 5:  return "background-color: #FFF2CC; font-weight:bold"
            return "background-color: #FF6B6B; font-weight:bold"
        except Exception:
            return ""

    fmt = {c: "{:+.3f}%" for c in pct_cols if c in t3.columns}

    styler = (
        t3.style
        .format(fmt, na_rep="N/A")
        .set_table_styles([
            {"selector": "thead th",
             "props": [("background-color", "#C8501E"), ("color", "white"),
                       ("font-weight", "bold"), ("text-align", "center"),
                       ("font-size", "12px")]},
            {"selector": "tbody td",
             "props": [("font-size", "12px"), ("text-align", "center")]},
        ])
        .hide(axis="index")
    )
    for col in pct_cols:
        if col in t3.columns:
            styler = styler.applymap(lambda v, c=col: color_cell(v, c), subset=[col])
    if "Position" in t3.columns:
        styler = styler.applymap(color_position, subset=["Position"])
    if "Note /10" in t3.columns:
        styler = styler.applymap(color_note, subset=["Note /10"])

    st.dataframe(styler, use_container_width=True, height=min(500, 60 + 38 * len(t3)))


def _render_history_chart(df_abb: pd.DataFrame, periodicity: str, category: str):
    isin_list = df_abb["isin"].tolist()
    opcvm_names = dict(zip(df_abb["isin"], df_abb["opcvm"]))

    selected_isin = st.selectbox(
        "Sélectionner un OPCVM",
        options=isin_list,
        format_func=lambda x: opcvm_names.get(x, x),
        key=f"hist_select_{category}",
    )

    hist = get_historical_series(selected_isin, periodicity)
    if hist.empty or len(hist) < 2:
        st.info("Données historiques insuffisantes. Téléchargez plusieurs rapports pour voir l'évolution.")
        return

    fig = go.Figure()

    # YTD line
    fig.add_trace(go.Scatter(
        x=hist["date"], y=hist["ytd"],
        mode="lines+markers",
        name="YTD (%)",
        line=dict(color="#C8501E", width=2.5),
        marker=dict(size=7, color="#C8501E"),
    ))

    # Perf 1J bar
    fig.add_trace(go.Bar(
        x=hist["date"], y=hist["perf_1j"],
        name="Perf. quotidienne (%)",
        marker_color=[("#92D050" if v and v > 0 else "#FF6B6B") for v in hist["perf_1j"]],
        opacity=0.7,
        yaxis="y2",
    ))

    fig.update_layout(
        title=dict(text=f"Évolution – {opcvm_names.get(selected_isin, selected_isin)}", font_size=15),
        xaxis=dict(title="Date", showgrid=True, gridcolor="#EEE"),
        yaxis=dict(title="YTD (%)", showgrid=True, gridcolor="#EEE",
                   tickformat=".3f", ticksuffix="%"),
        yaxis2=dict(title="Perf. quot. (%)", overlaying="y", side="right",
                    tickformat=".3f", ticksuffix="%"),
        legend=dict(orientation="h", y=1.08),
        plot_bgcolor="white",
        paper_bgcolor="white",
        height=400,
        hovermode="x unified",
        margin=dict(l=40, r=40, t=60, b=40),
    )
    st.plotly_chart(fig, use_container_width=True)


# ─── EXPORT PAGE ──────────────────────────────────────────────────────────────

def render_export_page():
    st.markdown('<div class="section-header">📤 Export – Rapport Excel complet</div>', unsafe_allow_html=True)

    df, periodicity, report_date_str = get_active_data()
    report_date = date.fromisoformat(report_date_str) if isinstance(report_date_str, str) else date.today()

    if df is None:
        st.info("Veuillez d'abord charger un fichier ASFIM.")
        return

    st.markdown(f"**Fichier actif :** {report_date_str} – {periodicity.upper()}")
    st.markdown(f"**Total fonds dans le fichier :** {len(df)}")

    categories = ["OCT", "OMLT", "DIVERSIFIE"]

    st.markdown("---")
    st.subheader("Aperçu par catégorie")

    cols = st.columns(3)
    for i, cat in enumerate(categories):
        df_abb = get_abb_df(df, cat, periodicity)
        with cols[i]:
            n = len(df_abb)
            st.markdown(f"""
            <div class="card">
                <b>{cat}</b><br>
                <span style="font-size:24px;font-weight:700;color:#C8501E;">{n}</span>
                <span style="font-size:13px;color:#666;"> fonds ABB</span>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")

    # Export all categories in one workbook
    if st.button("📥 Générer le rapport Excel complet (toutes catégories)", use_container_width=True):
        from openpyxl import Workbook as WB
        from exporter import _build_sheet1, _build_sheet2, _build_sheet3

        wb = WB()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        for cat in categories:
            df_market = get_market_df(df, cat)
            df_abb    = get_abb_df(df, cat, periodicity)
            if df_abb.empty:
                continue
            t1    = compute_tableau1(df_abb)
            synth = compute_tableau2(df_market, df_abb, cat)
            t3    = compute_tableau3(df_market, df_abb, synth)

            # Rename sheets with category prefix
            _build_sheet1(wb, t1, cat, report_date, periodicity)
            _build_sheet2(wb, synth, cat)
            _build_sheet3(wb, t3, cat)

        buf = io.BytesIO()
        wb.save(buf)
        fname = f"ABB_OPCVM_COMPLET_{periodicity.upper()}_{report_date_str.replace('-','')}.xlsx"
        st.download_button(
            "⬇️  Télécharger le rapport complet",
            data=buf.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    st.subheader("Export par catégorie")
    sel_cat = st.selectbox("Catégorie", categories, key="export_cat")
    df_market = get_market_df(df, sel_cat)
    df_abb    = get_abb_df(df, sel_cat, periodicity)
    if not df_abb.empty:
        t1    = compute_tableau1(df_abb)
        synth = compute_tableau2(df_market, df_abb, sel_cat)
        t3    = compute_tableau3(df_market, df_abb, synth)
        xls   = generate_excel(report_date, periodicity, sel_cat, t1, synth, t3)
        fname = f"ABB_{sel_cat}_{periodicity.upper()}_{report_date_str.replace('-','')}.xlsx"
        st.download_button(
            f"⬇️  Télécharger – {sel_cat}",
            data=xls,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning(f"Aucun fonds ABB {sel_cat} ({periodicity}) dans le fichier actif.")


# ─── HOME PAGE ────────────────────────────────────────────────────────────────

def render_home():
    st.markdown("""
    <div style="text-align:center;padding:20px 0 10px 0;">
        <div style="font-size:36px;font-weight:900;color:#C8501E;">ABB – OPCVM Manager</div>
        <div style="font-size:16px;color:#666;margin-top:6px;">
            Suivi quotidien & hebdomadaire des OPCVM Albarid Bank
        </div>
    </div>
    """, unsafe_allow_html=True)

    df, periodicity, report_date_str = get_active_data()

    if df is not None:
        st.success(f"✅ Fichier actif : **{report_date_str}** – {periodicity.upper()} – {len(df)} fonds")

        st.markdown("### Résumé rapide")
        cols = st.columns(3)
        for i, cat in enumerate(["OCT", "OMLT", "DIVERSIFIE"]):
            df_abb = get_abb_df(df, cat, periodicity)
            df_mkt = get_market_df(df, cat)
            avg_ytd = df_abb["ytd"].mean() if not df_abb.empty else np.nan
            with cols[i]:
                color = "#92D050" if not np.isnan(avg_ytd) and avg_ytd > 0 else "#FF6B6B"
                st.markdown(f"""
                <div class="card">
                    <div style="font-size:18px;font-weight:700;color:#C8501E;">{cat}</div>
                    <div>{len(df_abb)} fonds ABB / {len(df_mkt)} marché</div>
                    <div style="font-size:20px;font-weight:700;color:{color};">
                        YTD moy: {f'{avg_ytd:+.3f}%' if not np.isnan(avg_ytd) else 'N/A'}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🗺️ Comment utiliser l'application")
    else:
        st.info("Chargez un fichier ASFIM Excel depuis la barre latérale pour commencer.")
        st.markdown("---")
        st.markdown("### 🗺️ Guide d'utilisation")

    with st.expander("📖 Instructions", expanded=df is None):
        st.markdown("""
**1. Charger un fichier ASFIM**
- Dans la barre latérale, cliquez sur *Charger un fichier ASFIM*
- Sélectionnez le fichier Excel publié par l'ASFIM
- Choisissez la périodicité : **Quotidien** ou **Hebdomadaire**
- Indiquez la date du rapport puis cliquez **⚡ Traiter le fichier**

**2. Consulter les tableaux d'analyse**
- Naviguez via la barre latérale vers **OCT**, **OMLT** ou **Diversifié**
- Chaque page affiche :
  - **Tableau 1** – Fonds ABB avec performance quotidienne et YTD
  - **Tableau 2** – Synthèse du marché (moyennes, médiane, best/worst)
  - **Tableau 3** – Classement relatif (rang, position, notes /10)
  - **Courbe historique** – Évolution sur tous les rapports enregistrés

**3. Exporter**
- Page **Export** → télécharger le rapport Excel formaté (identique aux exports ABB)
- Le fichier contient 3 onglets : Fonds, Synthèse, Ranking

**4. Mémoire persistante**
- Chaque fichier chargé est enregistré automatiquement sur le disque
- Rechargeable depuis la liste *Fichiers enregistrés* même après redémarrage
        """)

    st.markdown("---")
    n = count_entries()
    st.markdown(f"**{n} rapport(s) enregistré(s)** en mémoire permanente. "
                "L'application conserve l'historique entre les sessions.")


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _kpi(col, title: str, value: str, sub: str = "", color: str = "default"):
    val_color = "#92D050" if color == "green" else ("#FF6B6B" if color == "red" else "#EDE0D4")
    with col:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value" style="color:{val_color};">{value}</div>
            <div style="font-size:11px;color:#AAA;">{sub}</div>
        </div>
        """, unsafe_allow_html=True)


def _to_float(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return None if np.isnan(val) else float(val)
    try:
        return float(str(val).replace(",", ".").replace("%", "").replace("+", "").strip())
    except Exception:
        return None


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    # ── Authentication gate ────────────────────────────────────────────────────
    if not st.session_state.get("authenticated", False):
        render_login()
        return

    page, uploaded, periodicity, report_date_str, process_btn = render_sidebar()

    # Auto-load from history selection
    hist_key = st.session_state.get("_selected_hist_key")
    if hist_key and st.session_state.get("_last_loaded_key") != hist_key:
        load_state_from_history(hist_key)
        st.session_state["_last_loaded_key"] = hist_key

    # Process uploaded file
    if process_btn:
        if uploaded is not None:
            report_date = date.fromisoformat(report_date_str)
            process_and_store(uploaded, periodicity, report_date)
        else:
            st.sidebar.warning("Veuillez sélectionner un fichier Excel.")

    # Auto-load latest if nothing in session
    if st.session_state.get("df_full") is None:
        latest = load_latest()
        if latest:
            st.session_state["df_full"]            = latest["df"]
            st.session_state["active_periodicity"] = latest["periodicity"]
            st.session_state["active_report_date"] = latest["date"]

    # Route to page
    if page == "Accueil":
        render_home()
    elif page == "OCT":
        render_category_page("OCT")
    elif page == "OMLT":
        render_category_page("OMLT")
    elif page == "Diversifié":
        render_category_page("DIVERSIFIE")
    elif page == "Export":
        render_export_page()
    else:
        render_home()


if __name__ == "__main__":
    main()
