"""
╔══════════════════════════════════════════════════════════════╗
║        VaR ANALYTICS SUITE  ·  Streamlit App v4.1           ║
║        Corrections v4.1 :                                    ║
║          · Import LineCollection déplacé en tête             ║
║          · Optional[bytes] (compat Python 3.9+)             ║
║          · Sidebar navigation entièrement refaite           ║
║          · Animations fond renforcées                        ║
║          · Cache Markowitz                                   ║
╚══════════════════════════════════════════════════════════════╝

Lancement :
    pip install streamlit yfinance pandas numpy scipy openpyxl reportlab matplotlib
    streamlit run app.py
"""

# ── FIX 1 : imports groupés en tête ──────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection          # FIX 1 — était dans la fonction
from typing import Optional, Dict, Tuple                   # FIX 2 — compat Python 3.9+
from scipy import stats
from scipy.optimize import minimize
import io, warnings
warnings.filterwarnings("ignore")

try:
    import yfinance as yf
    HAS_YF = True
except ImportError:
    HAS_YF = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak, Image)
    from reportlab.lib.colors import HexColor
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG + CSS
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="VaR Analytics Suite",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── FIX 3 : CSS sidebar entièrement revu + animations renforcées ─────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

/* ── Tokens ────────────────────────────────────────────── */
:root {
  --bg-void:      #07090F;
  --bg-deep:      #0C0F1A;
  --bg-surface:   #111827;
  --bg-raised:    #1a2235;
  --bg-glass:     rgba(17,24,37,0.80);
  --cyan:         #00D4FF;
  --cyan-dim:     rgba(0,212,255,0.14);
  --cyan-glow:    rgba(0,212,255,0.30);
  --amber:        #F0B429;
  --amber-dim:    rgba(240,180,41,0.12);
  --risk-red:     #FF4D6D;
  --risk-dim:     rgba(255,77,109,0.12);
  --signal:       #00C896;
  --signal-dim:   rgba(0,200,150,0.12);
  --txt-hi:       #E8EDF5;
  --txt-lo:       #7A8BA8;
  --txt-muted:    #3D4F6B;
  --border:       rgba(0,212,255,0.14);
  --border-hi:    rgba(0,212,255,0.30);
  --r-sm: 6px; --r-md: 10px; --r-lg: 16px;
  --shadow: 0 4px 24px rgba(0,0,0,0.55), 0 1px 0 rgba(255,255,255,0.03) inset;
  --font-ui:   'Outfit', sans-serif;
  --font-data: 'DM Mono', monospace;
}

/* ── Reset ─────────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: var(--font-ui) !important;
    -webkit-font-smoothing: antialiased;
}

/* ── App background (fond fixe, z-index bas) ─────────── */
.stApp {
    background: var(--bg-void) !important;
    position: relative;
}

/* ══════════════════════════════════════════════════════
   SIDEBAR — FIX COMPLET
   ══════════════════════════════════════════════════════ */

/* Conteneur sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0a0d17 0%, #0c1020 100%) !important;
    border-right: 1px solid var(--border) !important;
    min-width: 240px !important;
    z-index: 999 !important;
}

/* Barre de couleur en haut de la sidebar */
[data-testid="stSidebar"] > div:first-child::before {
    content: '';
    display: block;
    height: 3px;
    background: linear-gradient(90deg, var(--cyan), var(--amber), var(--cyan));
    background-size: 200% 100%;
    animation: shimmer 3s linear infinite;
    margin-bottom: 4px;
}
@keyframes shimmer {
    0%   { background-position: 200% 0; }
    100% { background-position: -200% 0; }
}

/* Textes dans la sidebar */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown {
    color: var(--txt-lo) !important;
    font-size: 12px !important;
    font-family: var(--font-ui) !important;
}

/* Selectbox dans la sidebar */
[data-testid="stSidebar"] .stSelectbox > div > div {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-hi) !important;
    border-radius: var(--r-sm) !important;
    color: var(--cyan) !important;
    font-family: var(--font-ui) !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    box-shadow: 0 0 12px var(--cyan-dim) !important;
}
[data-testid="stSidebar"] .stSelectbox svg {
    color: var(--cyan) !important;
    fill: var(--cyan) !important;
}

/* Bouton de réouverture sidebar (chevron) — TOUJOURS VISIBLE */
[data-testid="stSidebarCollapsedControl"] {
    visibility: visible !important;
    opacity: 1 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    background: var(--bg-surface) !important;
    border: 1px solid var(--border-hi) !important;
    border-left: none !important;
    border-radius: 0 var(--r-sm) var(--r-sm) 0 !important;
    width: 26px !important;
    height: 48px !important;
    position: fixed !important;
    top: 50% !important;
    transform: translateY(-50%) !important;
    z-index: 9999 !important;
    box-shadow: 4px 0 16px var(--cyan-dim) !important;
    transition: all 0.2s !important;
    cursor: pointer !important;
}
[data-testid="stSidebarCollapsedControl"]:hover {
    background: var(--bg-raised) !important;
    border-color: var(--cyan) !important;
    box-shadow: 4px 0 24px var(--cyan-glow) !important;
    width: 30px !important;
}
[data-testid="stSidebarCollapsedControl"] button {
    background: transparent !important;
    border: none !important;
    color: var(--cyan) !important;
    width: 100% !important;
    height: 100% !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
}
[data-testid="stSidebarCollapsedControl"] svg {
    color: var(--cyan) !important;
    fill: var(--cyan) !important;
    width: 12px !important;
    height: 12px !important;
}

/* Bouton fermeture sidebar (X en haut) */
[data-testid="stSidebar"] button[kind="headerNoPadding"],
[data-testid="stSidebar"] [data-testid="baseButton-headerNoPadding"] {
    color: var(--txt-lo) !important;
    opacity: 0.7 !important;
}
[data-testid="stSidebar"] button[kind="headerNoPadding"]:hover,
[data-testid="stSidebar"] [data-testid="baseButton-headerNoPadding"]:hover {
    color: var(--cyan) !important;
    opacity: 1 !important;
}

/* ── Titres ─────────────────────────────────────────────── */
h1 {
    font-family: var(--font-ui) !important;
    font-size: 2rem !important;
    font-weight: 800 !important;
    letter-spacing: -0.5px !important;
    background: linear-gradient(120deg, #fff 0%, var(--cyan) 50%, var(--amber) 100%) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
    line-height: 1.15 !important;
    margin-bottom: 4px !important;
}
h2 { color: var(--txt-hi) !important; font-size: 1.1rem !important; font-weight: 600 !important; }
h3 { color: var(--amber) !important; font-size: 0.95rem !important; font-weight: 600 !important; }

/* ── Metric cards ───────────────────────────────────────── */
[data-testid="metric-container"] {
    background: var(--bg-glass) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-md) !important;
    padding: 16px 18px 14px !important;
    box-shadow: var(--shadow) !important;
    position: relative; overflow: hidden;
    transition: border-color 0.25s, box-shadow 0.25s !important;
    animation: fadeUp 0.4s cubic-bezier(.4,0,.2,1) both;
}
[data-testid="metric-container"]:hover {
    border-color: var(--border-hi) !important;
    box-shadow: var(--shadow), 0 0 20px var(--cyan-dim) !important;
}
[data-testid="metric-container"]::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent, var(--cyan) 40%, var(--amber) 70%, transparent);
    opacity: 0.7;
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    color: var(--txt-lo) !important;
    font-size: 9.5px !important;
    font-family: var(--font-data) !important;
    text-transform: uppercase !important;
    letter-spacing: 1.2px !important;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: var(--cyan) !important;
    font-size: 1.5rem !important;
    font-family: var(--font-data) !important;
    font-weight: 500 !important;
}
[data-testid="metric-container"] [data-testid="stMetricDelta"] {
    font-family: var(--font-data) !important;
    font-size: 10px !important;
}

/* ── Boutons ────────────────────────────────────────────── */
.stButton > button {
    background: var(--bg-raised) !important;
    color: var(--txt-hi) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-sm) !important;
    font-family: var(--font-ui) !important;
    font-weight: 500 !important;
    font-size: 13px !important;
    padding: 9px 20px !important;
    transition: all 0.18s !important;
}
.stButton > button:hover {
    border-color: var(--cyan) !important;
    color: var(--cyan) !important;
    box-shadow: 0 0 16px var(--cyan-dim) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0088aa, var(--cyan)) !important;
    color: var(--bg-void) !important;
    border: none !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 18px var(--cyan-dim) !important;
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 6px 26px var(--cyan-glow) !important;
    transform: translateY(-2px) !important;
    color: var(--bg-void) !important;
}

/* ── Inputs ─────────────────────────────────────────────── */
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-sm) !important;
    color: var(--txt-hi) !important;
}
.stSelectbox > div > div:focus-within,
.stMultiSelect > div > div:focus-within {
    border-color: var(--cyan) !important;
    box-shadow: 0 0 0 3px var(--cyan-dim) !important;
}
.stNumberInput > div > div > input,
.stDateInput > div > div > input {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-sm) !important;
    color: var(--txt-hi) !important;
    font-family: var(--font-data) !important;
}
[data-testid="stRadio"] label {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-sm) !important;
    padding: 5px 12px !important;
    color: var(--txt-lo) !important;
    font-size: 12px !important;
    transition: all 0.15s !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    border-color: var(--cyan) !important;
    color: var(--cyan) !important;
    background: var(--cyan-dim) !important;
}

/* ── DataFrames ─────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--r-md) !important;
    overflow: hidden !important;
    box-shadow: var(--shadow) !important;
}
[data-testid="stDataFrame"] th {
    background: var(--bg-surface) !important;
    color: var(--txt-lo) !important;
    font-family: var(--font-data) !important;
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
}
[data-testid="stDataFrame"] td {
    font-family: var(--font-data) !important;
    font-size: 12px !important;
    color: var(--txt-hi) !important;
}

/* ── Expander ───────────────────────────────────────────── */
.streamlit-expanderHeader {
    background: var(--bg-surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--r-sm) !important;
    color: var(--txt-lo) !important;
    font-size: 13px !important;
}
.streamlit-expanderContent {
    border: 1px solid var(--border) !important;
    border-top: none !important;
    background: var(--bg-deep) !important;
}

hr { border: none !important; border-top: 1px solid var(--border) !important; }

/* ── Composants custom ──────────────────────────────────── */
.var-card {
    background: var(--bg-glass);
    border: 1px solid var(--border);
    border-radius: var(--r-md);
    padding: 18px 20px 16px;
    margin-bottom: 12px;
    position: relative; overflow: hidden;
    transition: border-color 0.2s, transform 0.2s, box-shadow 0.2s;
    animation: fadeUp 0.35s cubic-bezier(.4,0,.2,1) both;
}
.var-card:hover {
    border-color: var(--border-hi);
    transform: translateY(-2px);
    box-shadow: var(--shadow), 0 0 20px var(--cyan-dim);
}
.var-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent, var(--cyan) 35%, var(--amber) 70%, transparent);
    opacity: 0.6;
}
.var-card-title  { font-size:10px; color:var(--txt-muted); text-transform:uppercase; letter-spacing:1.8px; font-family:var(--font-data); margin-bottom:8px; }
.var-card-value  { font-size:26px; font-weight:300; color:var(--risk-red); font-family:var(--font-data); letter-spacing:-1px; line-height:1.1; }
.var-card-pct    { font-size:11px; color:var(--txt-lo); font-family:var(--font-data); margin-top:3px; }
.var-card-es     { font-size:12px; color:rgba(255,77,109,0.65); font-family:var(--font-data); margin-top:6px; padding-top:6px; border-top:1px solid var(--risk-dim); }

.badge-rec {
    display:inline-flex; align-items:center;
    background:linear-gradient(135deg,var(--cyan),#0099bb);
    color:var(--bg-void); font-size:8px; font-weight:700;
    padding:2px 8px; border-radius:20px; letter-spacing:0.8px;
    font-family:var(--font-data); margin-left:8px;
    box-shadow:0 2px 8px var(--cyan-dim);
}

.section-header {
    display:flex; align-items:center; gap:10px;
    margin:28px 0 14px; font-size:11px; font-weight:600;
    color:var(--txt-lo); text-transform:uppercase;
    letter-spacing:1.5px; font-family:var(--font-data);
}
.section-header::before {
    content:''; display:inline-block;
    width:3px; height:16px; flex-shrink:0;
    background:linear-gradient(180deg,var(--cyan),var(--amber));
    border-radius:2px;
}
.section-header::after {
    content:''; flex:1; height:1px;
    background:linear-gradient(90deg, var(--border), transparent);
}

.info-box {
    background:rgba(0,212,255,0.04);
    border:1px solid rgba(0,212,255,0.15);
    border-left:3px solid var(--cyan);
    border-radius:0 var(--r-sm) var(--r-sm) 0;
    padding:14px 18px; margin:14px 0;
    font-size:12.5px; color:var(--txt-lo);
    line-height:1.75; font-family:var(--font-ui);
}

.stress-card {
    background:var(--bg-glass);
    border:1px solid rgba(255,77,109,0.20);
    border-radius:var(--r-md);
    padding:16px 18px; margin-bottom:10px;
    position:relative; overflow:hidden;
    transition:transform 0.18s, border-color 0.18s;
    animation:fadeUp 0.35s cubic-bezier(.4,0,.2,1) both;
}
.stress-card:hover { border-color:rgba(255,77,109,0.45); transform:translateY(-1px); }
.stress-title { font-size:9.5px; color:rgba(255,77,109,0.8); font-family:var(--font-data); text-transform:uppercase; letter-spacing:1.5px; margin-bottom:8px; }
.stress-val   { font-size:24px; font-weight:300; color:var(--risk-red); font-family:var(--font-data); letter-spacing:-0.8px; }

.markowitz-card {
    background:rgba(0,200,150,0.04);
    border:1px solid rgba(0,200,150,0.18);
    border-radius:var(--r-md);
    padding:16px 18px; margin-bottom:10px;
    transition:transform 0.18s, border-color 0.18s;
    animation:fadeUp 0.35s cubic-bezier(.4,0,.2,1) both;
}
.markowitz-card:hover { border-color:rgba(0,200,150,0.40); transform:translateY(-1px); }

/* ── Animations ─────────────────────────────────────────── */
@keyframes fadeUp {
    from { opacity:0; transform:translateY(14px); }
    to   { opacity:1; transform:translateY(0); }
}
@keyframes haloFloat1 {
    0%   { transform:translate(-5%,-8%)  scale(1.00); }
    50%  { transform:translate(12%,5%)   scale(1.08); }
    100% { transform:translate(-5%,-8%)  scale(1.00); }
}
@keyframes haloFloat2 {
    0%   { transform:translate(5%,8%)   scale(1.00); }
    50%  { transform:translate(-10%,-6%)  scale(1.06); }
    100% { transform:translate(5%,8%)   scale(1.00); }
}
@keyframes haloFloat3 {
    0%   { transform:translate(0,0)     scale(1.00); }
    50%  { transform:translate(8%,-5%)  scale(1.04); }
    100% { transform:translate(0,0)     scale(1.00); }
}
@keyframes gridPulse {
    0%,100% { opacity:0.20; }
    50%     { opacity:0.38; }
}
@keyframes scanMove {
    0%   { background-position: 0 0; }
    100% { background-position: 0 4px; }
}

/* Suppression éléments indésirables */
#MainMenu, footer, [data-testid="stToolbar"] { display:none !important; }
[data-testid="stHeader"] { background:transparent !important; }
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:var(--bg-deep); }
::-webkit-scrollbar-thumb { background:var(--border); border-radius:10px; }
::-webkit-scrollbar-thumb:hover { background:var(--cyan); }
</style>
""", unsafe_allow_html=True)

# ── FIX 4 : Fond animé — couche fixe avec z-index NÉGATIF ────────────────────
# Injecté après le CSS pour s'assurer que le DOM est prêt
st.markdown("""
<div id="animated-bg" style="
    position:fixed; inset:0; z-index:0; pointer-events:none; overflow:hidden;">

  <!-- Halos lumineux -->
  <div style="
    position:absolute; width:700px; height:700px;
    top:-10%; left:-8%;
    background:radial-gradient(circle at 40% 40%, rgba(0,212,255,0.13) 0%, transparent 65%);
    border-radius:50%; filter:blur(70px);
    animation:haloFloat1 22s ease-in-out infinite;">
  </div>
  <div style="
    position:absolute; width:600px; height:600px;
    bottom:-8%; right:-5%;
    background:radial-gradient(circle at 55% 55%, rgba(240,180,41,0.10) 0%, transparent 65%);
    border-radius:50%; filter:blur(80px);
    animation:haloFloat2 28s ease-in-out infinite;">
  </div>
  <div style="
    position:absolute; width:500px; height:500px;
    top:35%; left:45%;
    background:radial-gradient(circle at 50% 50%, rgba(157,127,234,0.09) 0%, transparent 65%);
    border-radius:50%; filter:blur(90px);
    animation:haloFloat3 34s ease-in-out infinite;">
  </div>

  <!-- Grille de points -->
  <div style="
    position:absolute; inset:0;
    background-image:radial-gradient(circle, rgba(0,212,255,0.18) 1px, transparent 1px);
    background-size:54px 54px;
    animation:gridPulse 6s ease-in-out infinite;">
  </div>

  <!-- Scanlines horizontales -->
  <div style="
    position:absolute; inset:0;
    background:repeating-linear-gradient(
      0deg, transparent, transparent 3px,
      rgba(0,212,255,0.018) 3px, rgba(0,212,255,0.018) 4px);
    animation:scanMove 8s linear infinite;">
  </div>
</div>

<!-- Contenu au-dessus du fond -->
<style>
  [data-testid="stAppViewContainer"] > section,
  [data-testid="stMainBlockContainer"] {
    position:relative; z-index:1;
  }
  [data-testid="stSidebar"] {
    position:relative !important; z-index:999 !important;
  }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# DONNÉES & CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ACTIFS_DISPONIBLES = {
    "Apple (AAPL)":        "AAPL",
    "Microsoft (MSFT)":    "MSFT",
    "LVMH (MC.PA)":        "MC.PA",
    "TotalEnergies (TTE)": "TTE.PA",
    "BNP Paribas (BNP)":   "BNP.PA",
    "Nestlé (NESN.SW)":    "NESN.SW",
    "SAP (SAP)":           "SAP",
    "Airbus (AIR.PA)":     "AIR.PA",
    "Tesla (TSLA)":        "TSLA",
    "Amazon (AMZN)":       "AMZN",
    "Nvidia (NVDA)":       "NVDA",
    "Safran (SAF.PA)":     "SAF.PA",
    "L'Oréal (OR.PA)":     "OR.PA",
    "ASML (ASML.AS)":      "ASML.AS",
    "Hermès (RMS.PA)":     "RMS.PA",
}

SECTEURS = {
    "AAPL":"Technologie","MSFT":"Technologie","MC.PA":"Luxe",
    "TTE.PA":"Énergie","BNP.PA":"Finance","NESN.SW":"Conso.",
    "SAP":"Technologie","AIR.PA":"Aéronaut.","TSLA":"Auto/Tech",
    "AMZN":"Commerce","NVDA":"Technologie","SAF.PA":"Aéronaut.",
    "OR.PA":"Beauté","ASML.AS":"Technologie","RMS.PA":"Luxe",
}

SCENARIOS_STRESS = {
    "Crise 2008 (Lehman)": {
        "description": "Faillite de Lehman Brothers",
        "choc_marche": -0.0469, "choc_vol": 3.8,
        "date": "15 sept. 2008",
    },
    "Flash Crash 2010": {
        "description": "Effondrement éclair DJIA",
        "choc_marche": -0.034, "choc_vol": 2.5,
        "date": "6 mai 2010",
    },
    "Brexit 2016": {
        "description": "Référendum Brexit",
        "choc_marche": -0.0281, "choc_vol": 2.1,
        "date": "24 juin 2016",
    },
    "COVID-19 2020": {
        "description": "Pire séance COVID",
        "choc_marche": -0.0598, "choc_vol": 4.2,
        "date": "16 mars 2020",
    },
    "Krach Taux 2022": {
        "description": "Remontée brutale Fed",
        "choc_marche": -0.0312, "choc_vol": 2.0,
        "date": "T1 2022",
    },
    "Scénario −3σ": {
        "description": "Choc extrême calibré",
        "choc_marche": None, "choc_vol": 3.0,
        "date": "Hypothétique",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT DES DONNÉES
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def telecharger_donnees(tickers: list, date_debut, date_fin) -> pd.DataFrame:
    if not HAS_YF:
        return pd.DataFrame()
    data = yf.download(tickers, start=str(date_debut), end=str(date_fin),
                       auto_adjust=True, progress=False)
    if isinstance(data.columns, pd.MultiIndex):
        prix = data["Close"].copy() if "Close" in data.columns.get_level_values(0) \
               else data.xs(data.columns.get_level_values(0)[0], axis=1, level=0)
    else:
        if "Close" in data.columns:
            prix = data[["Close"]].copy()
            if len(tickers) == 1:
                prix.columns = tickers
        else:
            prix = data.copy()
    if isinstance(prix, pd.Series):
        prix = prix.to_frame()
    return prix.dropna(axis=1, how="all")


def donnees_simulation(tickers: list, n_days: int = 1500) -> pd.DataFrame:
    np.random.seed(42)
    n = len(tickers)
    corr = np.eye(n) + 0.4 * (np.ones((n, n)) - np.eye(n))
    L = np.linalg.cholesky(corr)
    z = np.random.randn(n_days, n) @ L.T
    lr = np.full(n, 0.0004) + np.full(n, 0.012) * z
    lr[n_days // 2: n_days // 2 + 20] *= 3.5
    prices = 100 * np.exp(np.cumsum(lr, axis=0))
    dates = pd.bdate_range(end="2024-12-31", periods=n_days)
    return pd.DataFrame(prices, index=dates, columns=tickers)


# ══════════════════════════════════════════════════════════════════════════════
# OPTIMISATION MARKOWITZ — FIX 5 : cache via hash array
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def optimiser_portefeuille_cached(
    rend_array: np.ndarray,
    columns: tuple,
    rf: float
) -> dict:
    """Version cacheable — reçoit un array numpy au lieu d'un DataFrame."""
    rendements = pd.DataFrame(rend_array, columns=list(columns))
    return _optimiser_portefeuille(rendements, rf)


def _optimiser_portefeuille(rendements: pd.DataFrame, rf: float = 0.03) -> dict:
    mu  = rendements.mean().values * 252
    cov = rendements.cov().values * 252
    n   = len(mu)

    def port_stats(w):
        r = float(w @ mu)
        v = float(np.sqrt(w @ cov @ w))
        s = (r - rf) / v if v > 0 else 0.0
        return r, v, s

    cons   = [{"type": "eq", "fun": lambda w: np.sum(w) - 1}]
    bounds = [(0.0, 1.0)] * n
    w0     = np.ones(n) / n

    def neg_sharpe(w):
        r, v, _ = port_stats(w)
        return -(r - rf) / v if v > 0 else 1e10

    res_s = minimize(neg_sharpe, w0, method="SLSQP", bounds=bounds, constraints=cons)
    w_s   = res_s.x if res_s.success else w0

    res_v = minimize(lambda w: w @ cov @ w, w0, method="SLSQP", bounds=bounds, constraints=cons)
    w_v   = res_v.x if res_v.success else w0

    # Frontière efficiente
    ret_range = np.linspace(mu.min(), mu.max(), 50)
    fv, fr = [], []
    for target in ret_range:
        c2 = cons + [{"type": "eq", "fun": lambda w, t=target: w @ mu - t}]
        r2 = minimize(lambda w: w @ cov @ w, w0, method="SLSQP", bounds=bounds, constraints=c2)
        if r2.success:
            fv.append(float(np.sqrt(r2.fun)))
            fr.append(target)

    tickers = list(rendements.columns)
    return {
        "tickers":  tickers, "mu": mu, "cov": cov,
        "sharpe":   {"weights": w_s,          "stats": port_stats(w_s), "label": "Sharpe Max."},
        "minvar":   {"weights": w_v,          "stats": port_stats(w_v), "label": "Variance Min."},
        "equi":     {"weights": np.ones(n)/n, "stats": port_stats(np.ones(n)/n), "label": "Équipondéré"},
        "frontier": (fv, fr),
    }


# ══════════════════════════════════════════════════════════════════════════════
# STRESS-TESTING
# ══════════════════════════════════════════════════════════════════════════════

def appliquer_stress(port_r: np.ndarray, pv: float,
                     scenario: dict, var_results: dict,
                     alpha: float = 0.99) -> dict:
    mu  = port_r.mean()
    sig = port_r.std()
    choc = scenario["choc_marche"] if scenario["choc_marche"] is not None else mu - 3 * sig
    pnl_stress   = choc * pv
    sigma_stress = sig * scenario["choc_vol"]
    z            = stats.norm.ppf(1 - alpha)
    var_stress   = -z * sigma_stress * pv
    var_normal   = var_results.get("Variance-Covariance", {}).get(alpha, {}).get("VaR", 0)
    return {
        "pnl_stress": pnl_stress,
        "var_stress":  var_stress,
        "var_normal":  var_normal,
        "ratio":       var_stress / var_normal if var_normal > 0 else np.nan,
        "choc":        choc,
        "choc_vol":    scenario["choc_vol"],
    }


# ══════════════════════════════════════════════════════════════════════════════
# MOTEUR VaR — 7 MÉTHODES
# ══════════════════════════════════════════════════════════════════════════════

class VaREngine:
    def __init__(self, r: np.ndarray, pv: float = 10_000_000, horizon: int = 1):
        self.r = r; self.pv = pv; self.horizon = horizon

    def historique(self, alpha: float) -> dict:
        q  = np.percentile(self.r, (1 - alpha) * 100)
        es = self.r[self.r <= q].mean() if (self.r <= q).any() else q
        return {"VaR": -q*self.pv*np.sqrt(self.horizon),
                "ES":  -es*self.pv*np.sqrt(self.horizon), "VaR_pct": -q}

    def variance_covariance(self, alpha: float) -> dict:
        mu, sig = self.r.mean(), self.r.std()
        z   = stats.norm.ppf(1 - alpha)
        var = -(mu + z * sig) * np.sqrt(self.horizon)
        es  = (sig * stats.norm.pdf(z) / (1 - alpha) - mu) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"μ={mu*100:.4f}%, σ={sig*100:.4f}%"}

    def riskmetrics(self, alpha: float, lam: float = 0.94) -> dict:
        sig2 = np.zeros(len(self.r)); sig2[0] = self.r[0]**2
        for t in range(1, len(self.r)):
            sig2[t] = lam * sig2[t-1] + (1 - lam) * self.r[t-1]**2
        st_ = np.sqrt(sig2[-1]); z = stats.norm.ppf(1 - alpha)
        return {"VaR": -z*st_*np.sqrt(self.horizon)*self.pv,
                "ES":  st_*stats.norm.pdf(z)/(1-alpha)*np.sqrt(self.horizon)*self.pv,
                "VaR_pct": -z*st_*np.sqrt(self.horizon),
                "params": f"λ=0.94, σ_t={st_*100:.4f}%"}

    def cornish_fisher(self, alpha: float) -> dict:
        mu, sig = self.r.mean(), self.r.std()
        s = float(stats.skew(self.r)); k = float(stats.kurtosis(self.r))
        z = stats.norm.ppf(1 - alpha)
        z_cf = z + (z**2-1)*s/6 + (z**3-3*z)*k/24 - (2*z**3-5*z)*s**2/36
        var = -(mu + z_cf * sig) * np.sqrt(self.horizon)
        es  = (sig * stats.norm.pdf(z_cf) / (1 - alpha) - mu) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"z_CF={z_cf:.4f}, skew={s:.3f}, kurt={k:.3f}"}

    def _fit_garch(self) -> Tuple[float, float, float, np.ndarray]:
        r = self.r
        def neg_ll(p):
            w, a, b = p
            if w <= 0 or a <= 0 or b <= 0 or a + b >= 1: return 1e10
            s2 = np.zeros(len(r)); s2[0] = np.var(r)
            for t in range(1, len(r)):
                s2[t] = w + a * r[t-1]**2 + b * s2[t-1]
            return 0.5 * np.sum(np.log(2*np.pi*s2) + r**2/s2)
        try:
            res = minimize(neg_ll, [1e-6, 0.08, 0.89], method="L-BFGS-B",
                           bounds=[(1e-8,None),(0.001,0.3),(0.5,0.999)])
            w, a, b = res.x
        except Exception:
            w, a, b = 5e-7, 0.09, 0.90
        s2 = np.zeros(len(r)); s2[0] = np.var(r)
        for t in range(1, len(r)):
            s2[t] = w + a * r[t-1]**2 + b * s2[t-1]
        return w, a, b, s2

    def garch(self, alpha: float) -> dict:
        w, a, b, s2 = self._fit_garch()
        sf = np.sqrt(w + a * self.r[-1]**2 + b * s2[-1])
        z  = stats.norm.ppf(1 - alpha)
        return {"VaR": -z*sf*np.sqrt(self.horizon)*self.pv,
                "ES":  sf*stats.norm.pdf(z)/(1-alpha)*np.sqrt(self.horizon)*self.pv,
                "VaR_pct": -z*sf*np.sqrt(self.horizon),
                "params": f"ω={w:.2e}, α={a:.3f}, β={b:.3f}"}

    def _fit_gpd(self, exc: np.ndarray) -> Tuple[float, float]:
        def neg_ll(p):
            xi, beta = p
            if beta <= 0: return 1e10
            u_ = exc / beta
            if xi != 0:
                if np.any(1 + xi * u_ <= 0): return 1e10
                return len(u_)*np.log(beta) + (1+1/xi)*np.sum(np.log(1+xi*u_))
            return len(u_)*np.log(beta) + np.sum(u_)
        try:
            res = minimize(neg_ll, [0.1, np.std(exc)], method="L-BFGS-B",
                           bounds=[(-0.5, 0.5),(1e-6, None)])
            return tuple(res.x) if res.success else (0.1, float(np.std(exc)))
        except Exception:
            return 0.1, float(np.std(exc))

    def tve(self, alpha: float) -> dict:
        losses = -self.r
        u = np.percentile(losses, 90)
        exc = losses[losses > u] - u
        xi, beta = self._fit_gpd(exc)
        n_u, n = len(exc), len(losses)
        p = 1 - alpha
        var = u + (beta/xi)*((n/n_u*p)**(-xi)-1) if xi != 0 else u + beta*np.log(n/n_u*p)
        var = max(var, 0.0)
        es  = (var + beta - xi*u) / (1 - xi) if xi < 1 else var * 2
        return {"VaR": var*self.pv*np.sqrt(self.horizon),
                "ES":  es*self.pv*np.sqrt(self.horizon), "VaR_pct": var,
                "params": f"ξ={xi:.3f}, β={beta:.4f}, u={u:.4f}"}

    def tve_garch(self, alpha: float) -> dict:
        w, a, b, s2 = self._fit_garch()
        resid = self.r / np.sqrt(s2)
        losses_r = -resid
        u_r = np.percentile(losses_r, 90)
        exc_r = losses_r[losses_r > u_r] - u_r
        xi, beta_r = self._fit_gpd(exc_r)
        n_u, n = len(exc_r), len(losses_r)
        p = 1 - alpha
        var_r = u_r + (beta_r/xi)*((n/n_u*p)**(-xi)-1) if xi != 0 else u_r + beta_r*np.log(n/n_u*p)
        var_r = max(var_r, 0.0)
        es_r  = (var_r + beta_r - xi*u_r) / (1 - xi) if xi < 1 else var_r * 2
        sf = np.sqrt(w + a*self.r[-1]**2 + b*s2[-1])
        var = var_r * sf * np.sqrt(self.horizon)
        es  = es_r  * sf * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"GARCH+GPD, ξ={xi:.3f}, σ_f={sf*100:.4f}%"}

    def compute_all(self, alphas=(0.95, 0.99)) -> dict:
        methods = {
            "Historique":          self.historique,
            "Variance-Covariance": self.variance_covariance,
            "RiskMetrics":         self.riskmetrics,
            "Cornish-Fisher":      self.cornish_fisher,
            "GARCH(1,1)":          self.garch,
            "TVE (POT)":           self.tve,
            "TVE-GARCH":           self.tve_garch,
        }
        results = {}
        for name, fn in methods.items():
            results[name] = {}
            for a in alphas:
                try:
                    results[name][a] = fn(a)
                except Exception as e:
                    results[name][a] = {"VaR": np.nan, "ES": np.nan,
                                        "VaR_pct": np.nan, "params": str(e)}
        return results


# ══════════════════════════════════════════════════════════════════════════════
# BACKTESTING
# ══════════════════════════════════════════════════════════════════════════════

def kupiec_test(r: np.ndarray, var_pct: float, alpha: float) -> dict:
    exc = (r < -var_pct).astype(int)
    N, T = int(exc.sum()), len(exc)
    if T == 0:
        return {"LR": np.nan, "p_value": np.nan, "valid": False, "N": 0, "T": 0, "rate": 0}
    p0    = 1 - alpha
    p_hat = N / T
    if p_hat == 0:
        lr = -2 * T * np.log(1 - p0)
    elif p_hat == 1:
        lr = -2 * N * np.log(p0)
    else:
        lr = -2 * (T*np.log(1-p0) + N*np.log(p0)
                   - N*np.log(p_hat) - (T-N)*np.log(1-p_hat))
    pv = 1 - stats.chi2.cdf(max(lr, 0), df=1)
    return {"LR": lr, "p_value": pv, "valid": pv > 0.05,
            "N": N, "T": T, "rate": p_hat, "expected": p0}


def christoffersen_test(r: np.ndarray, var_pct: float) -> dict:
    exc = (r < -var_pct).astype(int)
    n00 = np.sum((exc[:-1]==0) & (exc[1:]==0))
    n01 = np.sum((exc[:-1]==0) & (exc[1:]==1))
    n10 = np.sum((exc[:-1]==1) & (exc[1:]==0))
    n11 = np.sum((exc[:-1]==1) & (exc[1:]==1))
    pi01 = n01 / (n00+n01+1e-10)
    pi11 = n11 / (n10+n11+1e-10)
    pi   = (n01+n11) / (n00+n01+n10+n11+1e-10)
    try:
        lr = -2*(
            (n00+n10)*np.log(max(1-pi,1e-15)) + (n01+n11)*np.log(max(pi,1e-15))
            - n00*np.log(max(1-pi01,1e-15)) - n01*np.log(max(pi01,1e-15))
            - n10*np.log(max(1-pi11,1e-15)) - n11*np.log(max(pi11,1e-15))
        )
    except Exception:
        lr = np.nan
    pv = 1 - stats.chi2.cdf(max(lr, 0), df=1) if not np.isnan(lr) else np.nan
    return {"LR_ind": lr, "p_value_ind": pv,
            "valid": bool(pv > 0.05) if pv is not None and not np.isnan(pv) else False}


# ══════════════════════════════════════════════════════════════════════════════
# GRAPHIQUES — PLT_DARK sans rgba()
# ══════════════════════════════════════════════════════════════════════════════

PLT_DARK = {
    "figure.facecolor": "#07090F", "axes.facecolor": "#0C0F1A",
    "axes.edgecolor":   "#1a2235", "axes.labelcolor": "#7A8BA8",
    "xtick.color":      "#7A8BA8", "ytick.color":     "#7A8BA8",
    "grid.color":       "#111827", "text.color":      "#E8EDF5",
    "legend.facecolor": "#111827", "legend.edgecolor":"#1a2235",
}

PALETTE = ["#00D4FF","#F0B429","#00C896","#FF4D6D","#9D7FEA",
           "#FF7849","#4CC9F0","#F72585","#06D6A0","#FFD166",
           "#B5E48C","#C77DFF","#00B4D8","#FFAA5B","#FF9B54"]


def _spine_style(ax):
    for sp in ax.spines.values():
        sp.set_edgecolor("#1a2235")


def fig_perf(rendements: pd.DataFrame, tickers: list) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 5.5),
                                        gridspec_kw={"height_ratios": [3, 1]})
        fig.patch.set_alpha(0)
        port_r = rendements.mean(axis=1)
        cumul  = (1 + port_r).cumprod()
        ax1.fill_between(range(len(cumul)), cumul.values, 1,
                         where=(cumul.values >= 1), alpha=0.12,
                         color="#00D4FF", interpolate=True)
        ax1.fill_between(range(len(cumul)), cumul.values, 1,
                         where=(cumul.values < 1), alpha=0.12,
                         color="#FF4D6D", interpolate=True)
        ax1.plot(cumul.values, color="#00D4FF", lw=2, label="Portefeuille", zorder=4)
        for i, t in enumerate(tickers[:5]):
            if t in rendements.columns:
                ax1.plot((1+rendements[t]).cumprod().values,
                         color=PALETTE[(i+1) % len(PALETTE)],
                         lw=0.9, alpha=0.5, label=t, zorder=3)
        ax1.axhline(1, color="#F0B429", lw=0.8, ls="--", alpha=0.5)
        ax1.set_title("Performance cumulée", fontsize=11, color="#F0B429", pad=8, fontweight="bold")
        ax1.legend(fontsize=7.5, loc="upper left", facecolor="#111827",
                   edgecolor="#1a2235", labelcolor="#d0daea")
        ax1.grid(True, alpha=0.15, ls="--"); ax1.tick_params(labelsize=8.5)
        _spine_style(ax1)
        cols_bar = ["#1db87a" if v >= 0 else "#e05252" for v in port_r]
        ax2.bar(range(len(port_r)), port_r * 100, color=cols_bar, alpha=0.75, width=1, lw=0)
        ax2.axhline(0, color="#F0B429", lw=0.6)
        ax2.set_title("Rendements journaliers (%)", fontsize=9, color="#8aa0bc", pad=4)
        ax2.tick_params(labelsize=7.5); ax2.grid(True, alpha=0.15, ls="--")
        _spine_style(ax2); plt.tight_layout()
        return fig


def fig_correlation(rendements: pd.DataFrame) -> plt.Figure:
    from matplotlib.colors import LinearSegmentedColormap
    corr = rendements.corr()
    tickers = list(corr.columns)
    n = len(tickers)
    sz = max(6, n * 1.1 + 1.5)
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(sz, sz * 0.82))
        fig.patch.set_alpha(0)
        cmap = LinearSegmentedColormap.from_list("vc", ["#e05252","#111827","#1db87a"], N=256)
        im = ax.imshow(corr.values, cmap=cmap, vmin=-1, vmax=1, aspect="auto")
        cbar = plt.colorbar(im, ax=ax, shrink=0.82, pad=0.02)
        cbar.set_label("Corrélation", fontsize=9, color="#8aa0bc")
        cbar.ax.yaxis.set_tick_params(color="#8aa0bc", labelsize=8)
        plt.setp(cbar.ax.yaxis.get_ticklabels(), color="#8aa0bc")
        ax.set_xticks(range(n)); ax.set_yticks(range(n))
        ax.set_xticklabels(tickers, rotation=40, ha="right", fontsize=8.5)
        ax.set_yticklabels(tickers, fontsize=8.5)
        for i in range(n):
            for j in range(n):
                v = corr.values[i, j]
                ax.text(j, i, f"{v:.2f}", ha="center", va="center",
                        fontsize=7.5, color="white" if abs(v) > 0.45 else "#d0daea",
                        fontweight="bold")
        ax.set_title("Matrice de Corrélation", fontsize=12, color="#F0B429", pad=10, fontweight="bold")
        _spine_style(ax); plt.tight_layout()
        return fig


def fig_frontiere_efficiente(opt: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(11, 5.5))
        fig.patch.set_alpha(0)
        fv, fr = opt["frontier"]
        if fv:
            vf = [v * 100 for v in fv]; rf_ = [r * 100 for r in fr]
            # FIX 1 : LineCollection importé en tête de fichier
            pts  = np.array([vf, rf_]).T.reshape(-1, 1, 2)
            segs = np.concatenate([pts[:-1], pts[1:]], axis=1)
            lc   = LineCollection(segs, cmap="cool", linewidth=2.8, zorder=3)
            lc.set_array(np.linspace(0, 1, len(segs)))
            ax.add_collection(lc)
            ax.fill_between(vf, rf_, min(rf_), alpha=0.07, color="#00D4FF", zorder=1)
        mu = opt["mu"]; cov = opt["cov"]; tickers = opt["tickers"]
        for i, t in enumerate(tickers):
            vi = np.sqrt(cov[i, i]) * 100; mi = mu[i] * 100
            col = PALETTE[i % len(PALETTE)]
            ax.scatter(vi, mi, s=55, color=col, zorder=6, alpha=0.85,
                       edgecolors="#07090F", lw=0.8)
            ax.annotate(t, (vi, mi), fontsize=7.5, color=col,
                        fontweight="bold", xytext=(5, 4), textcoords="offset points")
        for key, col, mkr, sz, lbl in [
            ("sharpe","#F0B429","*", 280,"Sharpe Max."),
            ("minvar", "#00C896","v", 140,"Variance Min."),
            ("equi",   "#9D7FEA","D", 120,"Équipondéré"),
        ]:
            r, v, s = opt[key]["stats"]
            ax.scatter(v*100, r*100, s=sz*2.5, color=col, alpha=0.18, zorder=7, marker=mkr)
            ax.scatter(v*100, r*100, s=sz, color=col, zorder=8, marker=mkr,
                       edgecolors="white", lw=0.9,
                       label=f"{lbl}  (Sharpe {s:.2f})")
            ax.annotate(lbl, (v*100, r*100), fontsize=8.5, color=col, fontweight="bold",
                        xytext=(9, 5), textcoords="offset points",
                        bbox=dict(boxstyle="round,pad=0.25", fc="#07090F", ec=col, alpha=0.75, lw=0.8))
        ax.set_xlabel("Volatilité annualisée (%)", fontsize=10, labelpad=8)
        ax.set_ylabel("Rendement annualisé (%)", fontsize=10, labelpad=8)
        ax.set_title("Frontière Efficiente de Markowitz", fontsize=13,
                     color="#F0B429", pad=12, fontweight="bold")
        ax.legend(fontsize=8.5, loc="lower right", framealpha=0.85,
                  facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax.tick_params(labelsize=8.5); ax.grid(True, alpha=0.18, ls="--")
        _spine_style(ax); plt.tight_layout()
        return fig


def fig_poids_portefeuilles(opt: dict) -> plt.Figure:
    tickers = opt["tickers"]
    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1, 3, figsize=(13, 4.5))
        fig.patch.set_alpha(0)
        for ax, key, col in zip(axes,
                                 ["sharpe", "minvar", "equi"],
                                 ["#F0B429", "#00C896", "#9D7FEA"]):
            w    = opt[key]["weights"]
            mask = w > 0.005
            w_f  = w[mask]
            t_f  = [tickers[i] for i, m in enumerate(mask) if m]
            idx  = [i for i, m in enumerate(mask) if m]
            pie_colors = [PALETTE[i % len(PALETTE)] for i in idx]
            wedges, texts, autotexts = ax.pie(
                w_f, labels=t_f,
                autopct=lambda p: f"{p:.1f}%" if p > 3 else "",
                colors=pie_colors, startangle=90, pctdistance=0.72,
                wedgeprops={"edgecolor": "#07090F", "linewidth": 1.8},
                textprops={"fontsize": 7.5},
            )
            for txt in texts: txt.set_color("#d0daea"); txt.set_fontsize(7.5)
            for at in autotexts:
                at.set_fontsize(7); at.set_color("#07090F"); at.set_fontweight("bold")
            circle = plt.Circle((0,0), 0.42, color="#07090F", zorder=10)
            ring   = plt.Circle((0,0), 0.44, color=col, alpha=0.18, zorder=9)
            ax.add_patch(circle); ax.add_patch(ring)
            r, v, s = opt[key]["stats"]
            ax.text(0, 0.08, f"{s:.2f}", ha="center", va="center",
                    fontsize=14, fontweight="bold", color=col, zorder=11)
            ax.text(0,-0.14, "Sharpe",  ha="center", va="center",
                    fontsize=7, color="#8aa0bc", zorder=11)
            ax.set_title(f"{opt[key]['label']}\nRdt {r*100:.1f}%  ·  Vol {v*100:.1f}%",
                         fontsize=9, color=col, pad=8, fontweight="bold")
        plt.tight_layout()
        return fig


def fig_var_comparaison(var_results: dict, conf: float, pv: float) -> plt.Figure:
    methods = list(var_results.keys())
    vars_   = [var_results[m][conf]["VaR"] / 1000 for m in methods]
    ess_    = [var_results[m][conf]["ES"]  / 1000 for m in methods]
    short   = [m.replace("Variance-Covariance","VCV")
                .replace("Cornish-Fisher","C-Fisher")
                .replace("RiskMetrics","RiskM.") for m in methods]
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(12, 4.5))
        fig.patch.set_alpha(0)
        x = np.arange(len(methods)); w = 0.38
        bars1 = ax.bar(x-w/2, vars_, w, color=PALETTE[:len(methods)],
                       alpha=0.9, label="VaR", edgecolor="#07090F", lw=0.6)
        ax.bar(x+w/2, ess_, w, color=PALETTE[:len(methods)],
               alpha=0.42, label="ES (CVaR)", edgecolor="#07090F", lw=0.6, hatch="//")
        for bar, col in zip(bars1, PALETTE[:len(methods)]):
            h = bar.get_height()
            ax.text(bar.get_x()+bar.get_width()/2, h+max(vars_)*0.015,
                    f"{h:.0f}k", ha="center", va="bottom", fontsize=7.5,
                    color=col, fontweight="bold")
        ax.set_xticks(x); ax.set_xticklabels(short, rotation=22, ha="right", fontsize=9)
        ax.set_ylabel("k€", fontsize=10, labelpad=6)
        ax.grid(axis="y", alpha=0.18, ls="--")
        ax.set_title(f"VaR & Expected Shortfall — Niveau {conf*100:.0f}%  ·  "
                     f"Portefeuille {pv/1e6:.0f} M€",
                     fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax.legend(fontsize=9, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        _spine_style(ax); ax.tick_params(labelsize=8.5); plt.tight_layout()
        return fig


def fig_distribution(r: np.ndarray, var_results: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(12, 4.5))
        fig.patch.set_alpha(0)
        ax.hist(r[r >= 0]*100, bins=55, density=True, color="#00C896",
                alpha=0.45, edgecolor="#07090F", lw=0.2, label="Rdt ≥ 0")
        ax.hist(r[r <  0]*100, bins=55, density=True, color="#00D4FF",
                alpha=0.55, edgecolor="#07090F", lw=0.2, label="Rdt < 0")
        mu_, sig_ = r.mean(), r.std()
        xr = np.linspace(r.min(), r.max(), 400)
        ax.plot(xr*100, stats.norm.pdf(xr, mu_, sig_)/100,
                color="#F0B429", lw=2.2, ls="--", label="N(μ,σ)", zorder=5)
        for meth, (col, ls) in [("Historique",("#e05252","--")),
                                  ("GARCH(1,1)",("#a855f7","-.")),
                                  ("TVE-GARCH", ("#f97316",":"))]:
            if meth in var_results:
                p = var_results[meth].get(0.99,{}).get("VaR_pct")
                if p and not np.isnan(p):
                    ax.axvline(-p*100, color=col, lw=1.8, ls=ls,
                               label=f"VaR 99% {meth}", zorder=6)
        ax.set_xlabel("Rendement journalier (%)", fontsize=10, labelpad=6)
        ax.set_ylabel("Densité", fontsize=10, labelpad=6)
        ax.set_title("Distribution des rendements  ·  VaR 99% superposée",
                     fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax.legend(fontsize=8, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax.grid(True, alpha=0.15, ls="--"); _spine_style(ax)
        ax.tick_params(labelsize=8.5); plt.tight_layout()
        return fig


def fig_backtesting(bt_data: dict) -> plt.Figure:
    methods = list(bt_data.keys())
    short   = [m.replace("Variance-Covariance","VCV")
                .replace("Cornish-Fisher","C-Fisher")
                .replace("RiskMetrics","RiskM.") for m in methods]
    p95 = [bt_data[m].get(0.95, {}).get("p_value", 0) for m in methods]
    p99 = [bt_data[m].get(0.99, {}).get("p_value", 0) for m in methods]
    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1, 2, figsize=(12, 4))
        fig.patch.set_alpha(0)
        for ax, pvals, title in zip(axes, [p95, p99],
                                    ["Kupiec 95%","Kupiec 99%"]):
            for xi, pv_ in enumerate(pvals):
                col = "#1db87a" if pv_ > 0.05 else "#e05252"
                ax.bar(xi, pv_, 0.6, color=col, alpha=0.85, edgecolor="#07090F", lw=0.7)
                ax.text(xi, pv_+0.008, f"{pv_:.3f}", ha="center", va="bottom",
                        fontsize=7.5, color=col, fontweight="bold")
            ax.axhline(0.05, color="#F0B429", lw=2, ls="--", label="Seuil α=5%")
            ax.axhspan(0, 0.05, alpha=0.06, color="#FF4D6D")
            ax.set_xticks(range(len(short)))
            ax.set_xticklabels(short, rotation=28, ha="right", fontsize=8.5)
            ax.set_ylabel("p-value", fontsize=10, labelpad=6)
            ax.set_ylim(0, max(max(pvals)*1.2, 0.15))
            ax.set_title(title, fontsize=11, color="#F0B429", pad=8, fontweight="bold")
            ax.legend(fontsize=9, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
            ax.grid(axis="y", alpha=0.15, ls="--"); _spine_style(ax)
            ax.tick_params(labelsize=8.5)
        plt.tight_layout()
        return fig


def fig_stress_comparaison(stress_results: dict, pv: float) -> plt.Figure:
    scenarios = list(stress_results.keys())
    pnls    = [stress_results[s]["pnl_stress"] / 1000 for s in scenarios]
    vars_nm = [stress_results[s]["var_normal"]  / 1000 for s in scenarios]
    ratios  = [stress_results[s]["ratio"] for s in scenarios]
    short   = [s[:20] for s in scenarios]
    x = np.arange(len(scenarios))
    with plt.rc_context(PLT_DARK):
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 4.5))
        fig.patch.set_alpha(0)
        bars1 = ax1.barh(x-0.2, [abs(p) for p in pnls], 0.38,
                         color="#FF4D6D", alpha=0.85, label="P&L stressé")
        ax1.barh(x+0.2, vars_nm, 0.38, color="#00D4FF", alpha=0.65, label="VaR normale 99%")
        for bar in bars1:
            w_ = bar.get_width()
            ax1.text(w_+max(vars_nm)*0.02, bar.get_y()+bar.get_height()/2,
                     f"{w_:.0f}k", va="center", fontsize=7.5,
                     color="#FF4D6D", fontweight="bold")
        ax1.set_yticks(x); ax1.set_yticklabels(short, fontsize=8.5)
        ax1.set_xlabel("k€", fontsize=9, labelpad=6)
        ax1.set_title("Impact P&L vs VaR normale", fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax1.legend(fontsize=8.5, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax1.grid(axis="x", alpha=0.18, ls="--"); _spine_style(ax1)

        cols_r = ["#e05252" if r > 2 else "#F0B429" if r > 1.5 else "#1db87a" for r in ratios]
        for xi, (r_, col) in enumerate(zip(ratios, cols_r)):
            ax2.bar(xi, r_, 0.6, color=col, alpha=0.85, edgecolor="#07090F", lw=0.8)
            ax2.text(xi, r_+0.04, f"×{r_:.2f}", ha="center", va="bottom",
                     fontsize=8, fontweight="bold", color=col)
        ax2.axhline(1.0, color="#00C896", lw=1.8, ls="--", label="VaR normale (×1)")
        ax2.axhline(2.0, color="#FF4D6D", lw=1.4, ls=":", label="Seuil alerte (×2)")
        if max(ratios) > 2:
            ax2.axhspan(2.0, max(ratios)*1.15, alpha=0.05, color="#FF4D6D")
        ax2.set_xticks(x); ax2.set_xticklabels(short, rotation=30, ha="right", fontsize=8)
        ax2.set_ylabel("Ratio VaR Stressée / Normale", fontsize=9, labelpad=6)
        ax2.set_title("Multiplicateurs de stress", fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax2.legend(fontsize=8.5, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax2.grid(axis="y", alpha=0.18, ls="--"); _spine_style(ax2)
        plt.tight_layout()
        return fig


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL — FIX 2 : Optional[bytes]
# ══════════════════════════════════════════════════════════════════════════════

def generer_excel(rendements_df: pd.DataFrame, var_results: dict,
                   bt_results: dict, pv: float,
                   opt_results: Optional[dict] = None) -> Optional[bytes]:
    if not HAS_XLSX:
        return None
    wb = Workbook()
    NAVY, BLUE, WHITE, LGRAY = "1B2A4A", "2E5FA3", "FFFFFF", "F0F4FA"

    def th(ws, r, c, v, bg=NAVY, fg=WHITE):
        cell = ws.cell(r, c, v)
        cell.font      = Font(name="Calibri", bold=True, color=fg, size=10)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="CCCCCC")
        cell.border = Border(left=s, right=s, top=s, bottom=s)

    def td(ws, r, c, v, bg=WHITE):
        cell = ws.cell(r, c, v)
        cell.font      = Font(name="Calibri", size=9)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="DDDDDD")
        cell.border = Border(left=s, right=s, top=s, bottom=s)

    # ── Feuille 1 : Résumé VaR ────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "Résumé VaR"
    ws1.merge_cells("A1:F1")
    c = ws1["A1"]; c.value = "RAPPORT VaR — SYNTHÈSE"
    c.font = Font(name="Calibri", bold=True, color=WHITE, size=14)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30
    for j, h in enumerate(["Méthode","VaR 95% (€)","VaR 95% (%)","VaR 99% (€)","VaR 99% (%)","ES 99% (€)"], 1):
        th(ws1, 2, j, h, bg=BLUE)
    for i, (m, res) in enumerate(var_results.items(), 3):
        r95 = res.get(0.95, {}); r99 = res.get(0.99, {})
        bg  = LGRAY if i % 2 == 0 else WHITE
        td(ws1,i,1,m,bg=bg)
        td(ws1,i,2,round(r95.get("VaR",0),0),bg=bg)
        td(ws1,i,3,f"{r95.get('VaR_pct',0)*100:.3f}%",bg=bg)
        td(ws1,i,4,round(r99.get("VaR",0),0),bg=bg)
        td(ws1,i,5,f"{r99.get('VaR_pct',0)*100:.3f}%",bg=bg)
        td(ws1,i,6,round(r99.get("ES",0),0),bg=bg)
    for w_, col in zip([26,16,12,16,12,16],["A","B","C","D","E","F"]):
        ws1.column_dimensions[col].width = w_

    # ── Feuille 2 : Backtesting ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Backtesting")
    for j, h in enumerate(["Méthode","CL","Exc.","T","Taux obs.","Taux att.",
                             "Kupiec LR","Kupiec p","Kupiec OK","CC LR","CC p","CC OK"], 1):
        th(ws2, 1, j, h, bg=BLUE)
    row = 2
    for m, alphas in bt_results.items():
        for a, res in alphas.items():
            bg = LGRAY if row % 2 == 0 else WHITE
            k  = res.get("kupiec", {}); cc = res.get("cc", {})
            td(ws2,row,1,m,bg=bg); td(ws2,row,2,f"{a*100:.0f}%",bg=bg)
            td(ws2,row,3,k.get("N",""),bg=bg); td(ws2,row,4,k.get("T",""),bg=bg)
            td(ws2,row,5,f"{k.get('rate',0)*100:.2f}%",bg=bg)
            td(ws2,row,6,f"{(1-a)*100:.2f}%",bg=bg)
            td(ws2,row,7,round(k.get("LR",0),3),bg=bg)
            td(ws2,row,8,round(k.get("p_value",0),4),bg=bg)
            td(ws2,row,9,"OUI" if k.get("valid") else "NON",bg=bg)
            td(ws2,row,10,round(cc.get("LR_ind",0),3),bg=bg)
            td(ws2,row,11,round(cc.get("p_value_ind",0),4),bg=bg)
            td(ws2,row,12,"OUI" if cc.get("valid") else "NON",bg=bg)
            row += 1
    for w_, col in zip([26,6,8,8,10,10,10,10,10,10,10,10],
                        list("ABCDEFGHIJKL")):
        ws2.column_dimensions[col].width = w_

    # ── Feuille 3 : Markowitz (optionnel) ─────────────────────────────────────
    if opt_results:
        ws3 = wb.create_sheet("Markowitz")
        for j, h in enumerate(["Actif","Sharpe Max. (%)","Variance Min. (%)","Équipondéré (%)"], 1):
            th(ws3, 1, j, h, bg=BLUE)
        for i, ticker in enumerate(opt_results["tickers"], 2):
            bg = LGRAY if i % 2 == 0 else WHITE
            td(ws3,i,1,ticker,bg=bg)
            td(ws3,i,2,f"{opt_results['sharpe']['weights'][i-2]*100:.1f}%",bg=bg)
            td(ws3,i,3,f"{opt_results['minvar']['weights'][i-2]*100:.1f}%",bg=bg)
            td(ws3,i,4,f"{opt_results['equi']['weights'][i-2]*100:.1f}%",bg=bg)
        rs = len(opt_results["tickers"]) + 3
        for j, h in enumerate(["Statistiques","Sharpe Max.","Variance Min.","Équipondéré"], 1):
            th(ws3, rs, j, h, bg=NAVY)
        for j, key in enumerate(["sharpe","minvar","equi"], 2):
            r_, v_, s_ = opt_results[key]["stats"]
            td(ws3,rs+1,j,f"{r_*100:.2f}%"); td(ws3,rs+2,j,f"{v_*100:.2f}%"); td(ws3,rs+3,j,f"{s_:.3f}")
        for r_, lbl in zip([rs+1,rs+2,rs+3],["Rdt Annualisé","Vol. Annualisée","Ratio Sharpe"]):
            td(ws3,r_,1,lbl)
        for col in ["A","B","C","D"]:
            ws3.column_dimensions[col].width = 18

    # ── Feuille 4 : Données ───────────────────────────────────────────────────
    ws4 = wb.create_sheet("Données")
    th(ws4,1,1,"Date",bg=BLUE); th(ws4,1,2,"Rdt Portfolio (%)",bg=BLUE)
    r_port = rendements_df.mean(axis=1).tail(250)
    for i, (d, v) in enumerate(r_port.items(), 2):
        bg = LGRAY if i % 2 == 0 else WHITE
        td(ws4,i,1,d.strftime("%d/%m/%Y"),bg=bg)
        td(ws4,i,2,round(v*100,4),bg=bg)
    ws4.column_dimensions["A"].width = 14
    ws4.column_dimensions["B"].width = 18

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT PDF — FIX 2 : Optional[bytes]
# ══════════════════════════════════════════════════════════════════════════════

def generer_pdf(var_results: dict, bt_results: dict, pv: float,
                fig_var=None, fig_dist=None) -> Optional[bytes]:
    if not HAS_PDF:
        return None
    from datetime import date
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             leftMargin=2*cm, rightMargin=2*cm,
                             topMargin=1.5*cm, bottomMargin=2*cm)
    NAVY_C  = HexColor("#1B2A4A"); BLUE_C = HexColor("#2E5FA3")
    GOLD_C  = HexColor("#C9A84C"); LGRAY_C= HexColor("#F0F4FA")

    def ps(name, **kw): return ParagraphStyle(name, **kw)
    s_title = ps("t", fontName="Helvetica-Bold", fontSize=22,
                 textColor=HexColor("#FFFFFF"), alignment=TA_CENTER, spaceAfter=4)
    s_sub   = ps("s", fontName="Helvetica-Oblique", fontSize=11,
                 textColor=GOLD_C, alignment=TA_CENTER, spaceAfter=6)
    s_h2    = ps("h2", fontName="Helvetica-Bold", fontSize=12, textColor=NAVY_C,
                 spaceBefore=12, spaceAfter=6)
    s_body  = ps("b", fontName="Helvetica", fontSize=9,
                 textColor=HexColor("#3A3A3A"), alignment=TA_JUSTIFY,
                 spaceAfter=6, leading=14)

    def tbl_s():
        return TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLUE_C),("TEXTCOLOR",(0,0),(-1,0),HexColor("#FFFFFF")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8.5),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[HexColor("#FFFFFF"),LGRAY_C]),
            ("GRID",(0,0),(-1,-1),0.3,HexColor("#CCCCCC")),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ])

    story = [Spacer(1, 2*cm)]
    cover = Table([[Paragraph("RAPPORT DE GESTION DES RISQUES", s_title)]], [15*cm])
    cover.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY_C),
                                ("TOPPADDING",(0,0),(-1,-1),18),("BOTTOMPADDING",(0,0),(-1,-1),18)]))
    story.append(cover); story.append(Spacer(1, 0.3*cm))
    cover2 = Table([[Paragraph("Value at Risk · 7 méthodes · Backtesting · Stress-Testing · Markowitz", s_sub)]],[15*cm])
    cover2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),HexColor("#243B60")),
                                  ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(cover2); story.append(Spacer(1, 1.5*cm))
    info = [["Valeur portefeuille",f"{pv:,.0f} €"],["Horizon","1 jour ouvré"],
            ["Niveaux de confiance","95% et 99%"],
            ["Date de production", date.today().strftime("%d/%m/%Y")]]
    t_info = Table([[Paragraph(k,s_body),Paragraph(v,s_body)] for k,v in info],[5*cm,10*cm])
    t_info.setStyle(TableStyle([
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[HexColor("#FFFFFF"),LGRAY_C]),
        ("GRID",(0,0),(-1,-1),0.3,HexColor("#DDDDDD")),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
    ]))
    story.append(t_info); story.append(PageBreak())

    story.append(Paragraph("1. RÉSULTATS DE LA VALUE AT RISK", s_h2))
    story.append(HRFlowable(width="100%", thickness=1, color=GOLD_C, spaceAfter=8))
    var_data = [["Méthode","VaR 95% (€)","VaR 95%(%)","VaR 99% (€)","VaR 99%(%)","ES 99% (€)"]]
    for m, res in var_results.items():
        r95, r99 = res.get(0.95,{}), res.get(0.99,{})
        var_data.append([m,
            f"{r95.get('VaR',0):,.0f} €",  f"{r95.get('VaR_pct',0)*100:.3f}%",
            f"{r99.get('VaR',0):,.0f} €",  f"{r99.get('VaR_pct',0)*100:.3f}%",
            f"{r99.get('ES',0):,.0f} €"])
    t_var = Table(var_data, [3.5*cm,2.5*cm,2*cm,2.5*cm,2*cm,2.5*cm])
    t_var.setStyle(tbl_s()); story.append(t_var); story.append(Spacer(1, 0.5*cm))
    if fig_var:
        buf2 = io.BytesIO(); fig_var.savefig(buf2, format="png", dpi=120, bbox_inches="tight")
        buf2.seek(0); story.append(Image(buf2, width=15*cm, height=5*cm))
        plt.close(fig_var)
    story.append(PageBreak())

    story.append(Paragraph("2. BACKTESTING", s_h2))
    story.append(HRFlowable(width="100%", thickness=1, color=GOLD_C, spaceAfter=8))
    story.append(Paragraph(
        "Test de Kupiec (POF) : H₀ → fréquence observée = 1−α. "
        "Test de Christoffersen : indépendance temporelle des exceptions. "
        "p > 0.05 → modèle non rejeté.", s_body))
    bt_data = [["Méthode","CL","Exc.","Taux obs.","Kupiec p","Kupiec","CC p","CC"]]
    for m, alphas in bt_results.items():
        for a, res in alphas.items():
            k = res.get("kupiec",{}); cc = res.get("cc",{})
            bt_data.append([m, f"{a*100:.0f}%", str(k.get("N","")),
                f"{k.get('rate',0)*100:.2f}%",
                f"{k.get('p_value',0):.4f}", "OUI" if k.get("valid") else "NON",
                f"{cc.get('p_value_ind',0):.4f}", "OUI" if cc.get("valid") else "NON"])
    t_bt = Table(bt_data, [3.5*cm,1.2*cm,1.3*cm,1.8*cm,1.8*cm,1.5*cm,1.8*cm,1.5*cm])
    t_bt.setStyle(tbl_s()); story.append(t_bt); story.append(Spacer(1, 0.5*cm))
    if fig_dist:
        buf3 = io.BytesIO(); fig_dist.savefig(buf3, format="png", dpi=120, bbox_inches="tight")
        buf3.seek(0); story.append(Image(buf3, width=15*cm, height=5*cm))
        plt.close(fig_dist)

    doc.build(story); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

for key in ["prix","rendements","var_results","bt_results","actifs_choisis",
            "pv","opt_results","stress_results"]:
    if key not in st.session_state:
        st.session_state[key] = None

if "actifs_selection" not in st.session_state:
    st.session_state["actifs_selection"] = [
        "Apple (AAPL)", "Microsoft (MSFT)", "LVMH (MC.PA)",
        "TotalEnergies (TTE)", "BNP Paribas (BNP)"
    ]


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — FIX 3 : navigation robuste avec st.radio + style custom
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div style='padding:16px 4px 8px;'>
      <div style='font-size:15px;font-weight:800;
           background:linear-gradient(90deg,#00D4FF,#F0B429);
           -webkit-background-clip:text;-webkit-text-fill-color:transparent;
           letter-spacing:1px;margin-bottom:2px'>
        📉 VaR Analytics Suite
      </div>
      <div style='font-size:10px;color:#3D4F6B;font-family:DM Mono,monospace'>
        v4.1 · Département Risque
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr style='border-color:rgba(0,212,255,0.12);margin:4px 0 12px'/>",
                unsafe_allow_html=True)

    # Navigation par radio buttons (plus robuste que selectbox sur certains navigateurs)
    PAGES = {
        "🏠  Accueil":        "accueil",
        "🏦  Portefeuille":   "portfolio",
        "📐  Optimisation":   "optim",
        "📉  Calcul VaR":     "var",
        "🧪  Backtesting":    "backtest",
        "🔥  Stress-Testing": "stress",
        "📊  Reporting":      "reporting",
    }
    page_label = st.radio(
        "Navigation",
        list(PAGES.keys()),
        label_visibility="collapsed",
        key="nav_radio",
    )
    menu = PAGES[page_label]

    st.markdown("<hr style='border-color:rgba(0,212,255,0.12);margin:12px 0'/>",
                unsafe_allow_html=True)

    # Mini-résumé portefeuille
    if st.session_state["rendements"] is not None:
        r_ = st.session_state["rendements"].mean(axis=1)
        ann_r_ = r_.mean() * 252
        ann_v_ = r_.std()  * np.sqrt(252)
        sharpe_ = (r_.mean() - 0.03/252) / r_.std() * np.sqrt(252)
        sim_tag = "" if HAS_YF else " <span style='color:#F0B429'>(sim.)</span>"
        st.markdown(f"""
        <div style='background:rgba(0,212,255,0.04);border:1px solid rgba(0,212,255,0.12);
             border-radius:8px;padding:12px 14px;margin-bottom:10px'>
          <div style='font-size:9px;color:#3D4F6B;font-family:DM Mono,monospace;
               text-transform:uppercase;letter-spacing:1px;margin-bottom:8px'>
            Portefeuille chargé{sim_tag}
          </div>
          <div style='font-size:11px;font-family:DM Mono,monospace;line-height:2'>
            <span style='color:#7A8BA8'>Rdt ann. :</span>
            <span style='color:#00C896;font-weight:600'>+{ann_r_*100:.2f}%</span><br>
            <span style='color:#7A8BA8'>Vol. ann. :</span>
            <span style='color:#E8EDF5'>{ann_v_*100:.2f}%</span><br>
            <span style='color:#7A8BA8'>Sharpe   :</span>
            <span style='color:#F0B429;font-weight:600'>{sharpe_:.3f}</span>
          </div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='font-size:10px;color:#3D4F6B;line-height:1.9;font-family:DM Mono,monospace'>
    <span style='color:#F0B429'>■</span> 7 méthodes VaR<br>
    <span style='color:#00D4FF'>■</span> Tests Kupiec + CC<br>
    <span style='color:#00C896'>■</span> Frontière Markowitz<br>
    <span style='color:#FF4D6D'>■</span> 6 scénarios stress<br>
    <span style='color:#9D7FEA'>■</span> Export Excel + PDF
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════

if menu == "accueil":
    st.title("VaR Analytics Suite")
    st.markdown("""
    <div class='info-box'>
    Progiciel professionnel de calcul, comparaison et validation de la
    <b>Value at Risk</b> sur un portefeuille d'actions.
    Développé selon les standards <b>Bâle III/IV</b>.
    Version 4.1 — optimisation Markowitz, stress-testing et ES analytique.
    </div>""", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Méthodes VaR",    "7",           "Complètes")
    with c2: st.metric("Tests backtest",  "2",           "Kupiec + CC")
    with c3: st.metric("Scénarios Stress","6",           "Historiques")
    with c4: st.metric("Export",          "Excel + PDF", "4 feuilles")

    st.markdown("<div class='section-header'>Fonctionnalités</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**📥 Données de marché**
- Téléchargement Yahoo Finance (15 actifs)
- Simulation intégrée hors connexion
- Matrice de corrélation interactive

**📐 Optimisation Markowitz** *(v4)*
- Frontière efficiente complète
- Sharpe Max · Variance Min · Équipondéré
- Allocation exportable Excel
        """)
    with c2:
        st.markdown("""
**🔥 Stress-Testing** *(v4)*
- 6 scénarios historiques (2008, COVID…)
- Multiplicateurs de risque
- Seuils d'alerte Bâle III

**📊 Reporting**
- Excel 4 feuilles + PDF exécutif
- Graphiques intégrés
- Prêt Direction Risque
        """)

    st.markdown("<div class='section-header'>Équipe Projet</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    membres = ["Anta Mbaye", "Harlem D. Adjagba", "Ecclésiaste Gnargo", "Wariol G. Kopangoye"]
    for col, m in zip(cols, membres):
        with col:
            st.markdown(f"""
            <div class='var-card' style='text-align:center;padding:14px'>
              <div style='font-size:22px;margin-bottom:6px'>👤</div>
              <div style='font-size:11px;font-weight:600;color:#E8EDF5'>{m}</div>
            </div>""", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align:center;margin-top:14px;font-size:11px;color:#3D4F6B;font-family:DM Mono,monospace'>
    Double diplôme M2 IFIM · Ing 3 MACS — Mathématiques Appliquées au Calcul Scientifique
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : PORTEFEUILLE
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "portfolio":
    st.title("Construction du Portefeuille")

    st.markdown("<div class='section-header'>Sélection des actifs</div>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        actifs_choisis = st.multiselect(
            "Actifs financiers", list(ACTIFS_DISPONIBLES.keys()),
            default=st.session_state["actifs_selection"],
            key="actifs_selection",
        )
    with col2:
        pv_m = st.number_input("Valeur portefeuille (M€)", 0.1, 1000.0, 10.0, 0.5)
        pv   = pv_m * 1_000_000

    col3, col4, col5 = st.columns(3)
    with col3: date_debut = st.date_input("Date de début", value=pd.to_datetime("2019-01-01"))
    with col4: date_fin   = st.date_input("Date de fin",   value=pd.to_datetime("today"))
    with col5: source     = st.radio("Source", ["Yahoo Finance","Simulation"], horizontal=True)

    btn_col, _ = st.columns([1, 3])
    with btn_col:
        btn = st.button("▶  Charger les données", type="primary", use_container_width=True)

    if btn:
        if len(actifs_choisis) < 2:
            st.warning("Sélectionnez au moins 2 actifs.")
        elif date_debut >= date_fin:
            st.warning("La date de début doit être antérieure à la date de fin.")
        else:
            tickers = [ACTIFS_DISPONIBLES[a] for a in actifs_choisis]
            with st.spinner("Chargement des données…"):
                if source == "Yahoo Finance" and HAS_YF:
                    prix = telecharger_donnees(tickers, date_debut, date_fin)
                    if prix.empty:
                        st.warning("Aucune donnée, passage en simulation.")
                        prix = donnees_simulation(tickers)
                else:
                    prix = donnees_simulation(tickers)
                rend = prix.pct_change(fill_method=None).dropna(how="all")
                rend = rend.dropna(axis=1, how="any")
                prix = prix[rend.columns]
            st.session_state.update({
                "prix": prix, "rendements": rend,
                "actifs_choisis": list(rend.columns), "pv": pv,
                "var_results": None, "bt_results": None,
                "opt_results": None, "stress_results": None,
            })
            st.success(f"✅ {len(rend.columns)} actif(s) — {len(rend)} jours.")

    if st.session_state["rendements"] is not None:
        rend = st.session_state["rendements"]
        prix = st.session_state["prix"]
        port_r = rend.mean(axis=1)
        ann_r  = port_r.mean() * 252
        ann_v  = port_r.std()  * np.sqrt(252)
        sh_    = (port_r.mean() - 0.03/252) / port_r.std() * np.sqrt(252)
        mdd_   = float(((1+port_r).cumprod()/(1+port_r).cumprod().cummax()-1).min())

        st.markdown("<div class='section-header'>Statistiques du portefeuille</div>", unsafe_allow_html=True)
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Rdt Annualisé",  f"+{ann_r*100:.2f}%")
        c2.metric("Vol. Annuelle",   f"{ann_v*100:.2f}%")
        c3.metric("Sharpe",          f"{sh_:.3f}")
        c4.metric("Max Drawdown",    f"{mdd_*100:.2f}%")
        c5.metric("Skewness",        f"{stats.skew(port_r):.4f}")
        c6.metric("Kurtosis (exc)",  f"{stats.kurtosis(port_r):.4f}")

        st.markdown("<div class='section-header'>Performance & Rendements</div>", unsafe_allow_html=True)
        f1 = fig_perf(rend, list(rend.columns))
        st.pyplot(f1, use_container_width=True); plt.close(f1)

        st.markdown("<div class='section-header'>Matrice de Corrélation</div>", unsafe_allow_html=True)
        f2 = fig_correlation(rend)
        st.pyplot(f2, use_container_width=True); plt.close(f2)

        st.markdown("<div class='section-header'>Statistiques individuelles</div>", unsafe_allow_html=True)
        df_s = pd.DataFrame({
            "Secteur":       [SECTEURS.get(t,"—") for t in rend.columns],
            "Rdt moy. (%)":  (rend.mean()*252*100).round(2),
            "Vol. ann. (%)": (rend.std()*np.sqrt(252)*100).round(2),
            "Skewness":      rend.apply(lambda c: round(float(stats.skew(c)), 4)),
            "Kurtosis":      rend.apply(lambda c: round(float(stats.kurtosis(c)), 4)),
            "Min (%)":       (rend.min()*100).round(3),
            "Max (%)":       (rend.max()*100).round(3),
        }, index=rend.columns)
        df_s.index.name = "Ticker"
        st.dataframe(df_s, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : OPTIMISATION MARKOWITZ
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "optim":
    st.title("Optimisation de Portefeuille — Markowitz")

    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille.")
        st.stop()

    rend = st.session_state["rendements"]
    st.markdown("""
    <div class='info-box'>
    <b>Théorie Moderne du Portefeuille (Markowitz, 1952)</b> — La frontière efficiente
    représente les portefeuilles offrant le <b>meilleur rendement pour un risque donné</b>.
    Trois allocations optimales sont calculées : Sharpe Max, Variance Min, Équipondéré.
    </div>""", unsafe_allow_html=True)

    col_rf, col_btn, _ = st.columns([1, 1, 2])
    with col_rf:
        rf = st.number_input("Taux sans risque (%/an)", 0.0, 10.0, 3.0, 0.1) / 100
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        opt_btn = st.button("▶  Optimiser", type="primary", use_container_width=True)

    if opt_btn:
        with st.spinner("Calcul de la frontière efficiente…"):
            # FIX 5 : cache via tuple columns + numpy array
            opt = optimiser_portefeuille_cached(
                rend.values, tuple(rend.columns), rf
            )
            st.session_state["opt_results"] = opt
        st.success("✅ Optimisation terminée.")

    if st.session_state["opt_results"]:
        opt = st.session_state["opt_results"]

        st.markdown("<div class='section-header'>Frontière Efficiente</div>", unsafe_allow_html=True)
        f3 = fig_frontiere_efficiente(opt)
        st.pyplot(f3, use_container_width=True); plt.close(f3)

        st.markdown("<div class='section-header'>Allocations optimales</div>", unsafe_allow_html=True)
        cols = st.columns(3)
        for col, (key, color) in zip(cols, [("sharpe","#F0B429"),("minvar","#00C896"),("equi","#9D7FEA")]):
            p = opt[key]; r_, v_, s_ = p["stats"]
            with col:
                st.markdown(f"""
                <div class='markowitz-card'>
                  <div style='font-size:11px;color:{color};font-family:DM Mono,monospace;
                       text-transform:uppercase;letter-spacing:1px;margin-bottom:8px'>{p['label']}</div>
                  <div style='font-size:13px;color:#E8EDF5;line-height:2;font-family:DM Mono,monospace'>
                  📈 Rdt ann. : <b style='color:{color}'>{r_*100:.2f}%</b><br>
                  📊 Vol. ann. : <b style='color:#E8EDF5'>{v_*100:.2f}%</b><br>
                  ⭐ Sharpe   : <b style='color:{color}'>{s_:.3f}</b>
                  </div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Répartition des poids</div>", unsafe_allow_html=True)
        f4 = fig_poids_portefeuilles(opt)
        st.pyplot(f4, use_container_width=True); plt.close(f4)

        st.markdown("<div class='section-header'>Tableau des poids</div>", unsafe_allow_html=True)
        df_w = pd.DataFrame({
            "Actif":             opt["tickers"],
            "Sharpe Max. (%)":   [f"{w*100:.1f}%" for w in opt["sharpe"]["weights"]],
            "Variance Min. (%)": [f"{w*100:.1f}%" for w in opt["minvar"]["weights"]],
            "Équipondéré (%)":   [f"{w*100:.1f}%" for w in opt["equi"]["weights"]],
        }).set_index("Actif")
        st.dataframe(df_w, use_container_width=True)

        with st.expander("ℹ️ Hypothèses & Limites"):
            st.markdown("""
- **Positions longues uniquement** (w ≥ 0)
- **Rendements supposés stationnaires** sur la période d'estimation
- **Pas de coûts de transaction**
- **Recommandation** : coupler l'allocation Sharpe Max avec la VaR TVE-GARCH
            """)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : CALCUL VaR
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "var":
    st.title("Calcul de la Value at Risk")

    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille.")
        st.stop()

    rend   = st.session_state["rendements"]
    pv     = st.session_state["pv"] or 10_000_000
    port_r = rend.mean(axis=1).values

    st.markdown("<div class='section-header'>Paramètres de calcul</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: horizon = st.slider("Horizon (jours)", 1, 10, 1)
    with c2:
        conf_opts = st.multiselect("Niveaux de confiance",
                                    [0.90, 0.95, 0.975, 0.99],
                                    default=[0.95, 0.99],
                                    format_func=lambda x: f"{x*100:.1f}%")
    with c3:
        methods_sel = st.multiselect("Méthodes",
                                      ["Historique","Variance-Covariance","RiskMetrics",
                                       "Cornish-Fisher","GARCH(1,1)","TVE (POT)","TVE-GARCH"],
                                      default=["Historique","Variance-Covariance",
                                               "RiskMetrics","Cornish-Fisher",
                                               "GARCH(1,1)","TVE (POT)","TVE-GARCH"])

    btn_c, _ = st.columns([1, 3])
    with btn_c:
        calc_btn = st.button("▶  Calculer les 7 VaR", type="primary", use_container_width=True)

    if calc_btn:
        if not conf_opts:
            st.warning("Sélectionnez au moins un niveau de confiance.")
        else:
            with st.spinner("Calcul en cours…"):
                engine     = VaREngine(port_r, pv, horizon)
                var_res    = engine.compute_all(tuple(sorted(conf_opts)))
                var_res    = {k: v for k, v in var_res.items() if k in methods_sel}
                st.session_state["var_results"] = var_res
            st.success(f"✅ {len(var_res)} méthodes × {len(conf_opts)} niveaux calculés.")

    if st.session_state["var_results"]:
        var_res    = st.session_state["var_results"]
        alphas_    = sorted(list(list(var_res.values())[0].keys()))
        alpha_disp = st.select_slider("Afficher pour :", options=alphas_,
                                       format_func=lambda x: f"{x*100:.0f}%",
                                       value=alphas_[-1])

        st.markdown("<div class='section-header'>Résultats par méthode</div>", unsafe_allow_html=True)
        cols = st.columns(min(len(var_res), 4))
        for i, (method, res) in enumerate(var_res.items()):
            r   = res.get(alpha_disp, {})
            var_v = r.get("VaR", np.nan); pct_v = r.get("VaR_pct", np.nan); es_v = r.get("ES", np.nan)
            rec = '<span class="badge-rec">★ Recommandé</span>' if method == "TVE-GARCH" else ""
            with cols[i % len(cols)]:
                st.markdown(f"""
                <div class='var-card'>
                  <div class='var-card-title'>{method}{rec}</div>
                  <div class='var-card-value'>{var_v/1000:.1f} k€</div>
                  <div class='var-card-pct'>VaR {alpha_disp*100:.0f}% · {pct_v*100:.3f}%</div>
                  <div class='var-card-es'>ES : {es_v/1000:.1f} k€</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Tableau comparatif</div>", unsafe_allow_html=True)
        rows = []
        for m, res in var_res.items():
            row = {"Méthode": m}
            for a in alphas_:
                r = res.get(a, {})
                row[f"VaR {a*100:.0f}%(€)"]  = f"{r.get('VaR',0):,.0f}"
                row[f"VaR {a*100:.0f}%(%)"]  = f"{r.get('VaR_pct',0)*100:.3f}%"
                row[f"ES {a*100:.0f}%(€)"]   = f"{r.get('ES',0):,.0f}"
            row["Paramètres"] = list(var_res[m].values())[0].get("params","")
            rows.append(row)
        st.dataframe(pd.DataFrame(rows).set_index("Méthode"), use_container_width=True)

        st.markdown("<div class='section-header'>Graphique comparatif</div>", unsafe_allow_html=True)
        f5 = fig_var_comparaison(var_res, alpha_disp, pv)
        st.pyplot(f5, use_container_width=True); plt.close(f5)

        st.markdown("<div class='section-header'>Distribution des rendements</div>", unsafe_allow_html=True)
        f6 = fig_distribution(port_r, var_res)
        st.pyplot(f6, use_container_width=True); plt.close(f6)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : BACKTESTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "backtest":
    st.title("Backtesting des modèles de VaR")

    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR.")
        st.stop()

    rend     = st.session_state["rendements"]
    var_res  = st.session_state["var_results"]
    port_r   = rend.mean(axis=1).values
    alphas_  = sorted(list(list(var_res.values())[0].keys()))

    st.markdown("""
    <div class='info-box'>
    <b>Test de Kupiec (POF)</b> — fréquence d'exceptions vs niveau de confiance déclaré. LR ~ χ²(1).<br><br>
    <b>Test de Christoffersen (CC)</b> — indépendance temporelle des exceptions.<br><br>
    <span style='color:#00C896'>✅ p > 5%</span> → validé &nbsp;
    <span style='color:#FF4D6D'>❌ p ≤ 5%</span> → rejeté
    </div>""", unsafe_allow_html=True)

    btn_c, _ = st.columns([1, 3])
    with btn_c:
        bt_btn = st.button("▶  Lancer le backtesting", type="primary", use_container_width=True)

    if bt_btn:
        with st.spinner("Backtesting en cours…"):
            bt_res = {}
            for method, res in var_res.items():
                bt_res[method] = {}
                for a in alphas_:
                    vp = res.get(a, {}).get("VaR_pct", np.nan)
                    if np.isnan(vp): continue
                    bt_res[method][a] = {
                        "kupiec": kupiec_test(port_r, vp, a),
                        "cc":     christoffersen_test(port_r, vp),
                    }
            st.session_state["bt_results"] = bt_res
        st.success("✅ Backtesting terminé.")

    if st.session_state["bt_results"]:
        bt_res = st.session_state["bt_results"]
        rows = []
        for m, alphas in bt_res.items():
            for a, res in alphas.items():
                k, cc = res["kupiec"], res["cc"]
                rows.append({
                    "Méthode":   m,
                    "CL":        f"{a*100:.0f}%",
                    "Exceptions":k["N"], "T": k["T"],
                    "Taux obs.": f"{k['rate']*100:.2f}%",
                    "Taux att.": f"{(1-a)*100:.2f}%",
                    "Kupiec LR": round(k["LR"],3),
                    "Kupiec p":  round(k["p_value"],4),
                    "Kupiec ✓":  "✅" if k["valid"] else "❌",
                    "CC p":      round(cc.get("p_value_ind",0),4),
                    "CC ✓":      "✅" if cc.get("valid") else "❌",
                })
        st.dataframe(pd.DataFrame(rows).set_index("Méthode"), use_container_width=True)

        st.markdown("<div class='section-header'>p-values Kupiec</div>", unsafe_allow_html=True)
        f7 = fig_backtesting({m: {a: {"p_value": bt_res[m][a]["kupiec"]["p_value"]}
                                   for a in bt_res[m]} for m in bt_res})
        st.pyplot(f7, use_container_width=True); plt.close(f7)

        st.markdown("<div class='section-header'>Verdict synthétique</div>", unsafe_allow_html=True)
        alpha_v = alphas_[-1]
        cols_v  = st.columns(len(bt_res))
        for col, (m, alphas) in zip(cols_v, bt_res.items()):
            res_ = alphas.get(alpha_v, {}); k = res_.get("kupiec",{}); cc = res_.get("cc",{})
            k_ok, cc_ok = k.get("valid",False), cc.get("valid",False)
            score = "✅ Validé" if k_ok and cc_ok else ("⚠️ Partiel" if k_ok or cc_ok else "❌ Rejeté")
            color = "#00C896" if k_ok and cc_ok else ("#F0B429" if k_ok or cc_ok else "#FF4D6D")
            with col:
                st.markdown(f"""
                <div class='var-card' style='text-align:center'>
                  <div class='var-card-title'>{m}</div>
                  <div style='font-size:15px;font-weight:700;color:{color};margin:6px 0'>{score}</div>
                  <div style='font-size:10px;color:#7A8BA8;font-family:DM Mono,monospace'>
                  Kupiec p={k.get('p_value',0):.4f}<br>CC p={cc.get('p_value_ind',0):.4f}
                  </div>
                </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : STRESS-TESTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "stress":
    st.title("Stress-Testing — Scénarios Historiques")

    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille.")
        st.stop()

    rend      = st.session_state["rendements"]
    var_res   = st.session_state["var_results"] or {}
    pv        = st.session_state["pv"] or 10_000_000
    port_r    = rend.mean(axis=1).values

    st.markdown("""
    <div class='info-box'>
    Le <b>stress-testing</b> applique des chocs déterministes calibrés sur des crises
    historiques réelles. Exigé par <b>Bâle III/IV</b> et les superviseurs (EBA, BCE).
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-header'>Sélection des scénarios</div>", unsafe_allow_html=True)
    sc_choisis = st.multiselect("Scénarios", list(SCENARIOS_STRESS.keys()),
                                 default=list(SCENARIOS_STRESS.keys()))

    c_a, c_b, _ = st.columns([1, 1, 2])
    with c_a:
        alpha_st = st.selectbox("Niveau VaR", [0.95, 0.99],
                                 format_func=lambda x: f"{x*100:.0f}%", index=1)
    with c_b:
        st.markdown("<br>", unsafe_allow_html=True)
        st_btn = st.button("▶  Lancer le stress-test", type="primary", use_container_width=True)

    if st_btn and sc_choisis:
        with st.spinner("Calcul des scénarios…"):
            sr = {s: appliquer_stress(port_r, pv, SCENARIOS_STRESS[s], var_res, alpha_st)
                  for s in sc_choisis}
            st.session_state["stress_results"] = sr
        st.success(f"✅ {len(sr)} scénario(s) calculé(s).")

    if st.session_state["stress_results"]:
        sr = st.session_state["stress_results"]

        st.markdown("<div class='section-header'>Résultats par scénario</div>", unsafe_allow_html=True)
        sc_list = list(sr.items())
        for i in range(0, len(sc_list), 3):
            cols = st.columns(3)
            for col, (sc_name, s) in zip(cols, sc_list[i:i+3]):
                ratio = s["ratio"]
                alert = "🔴" if (not np.isnan(ratio) and ratio > 2) else \
                        ("🟡" if (not np.isnan(ratio) and ratio > 1.5) else "🟢")
                rtxt  = f"×{ratio:.2f}" if not np.isnan(ratio) else "—"
                with col:
                    st.markdown(f"""
                    <div class='stress-card'>
                      <div class='stress-title'>{alert} {sc_name}</div>
                      <div style='font-size:10px;color:#7A8BA8;margin-bottom:8px'>
                        {SCENARIOS_STRESS.get(sc_name,{}).get('date','')}
                      </div>
                      <div class='stress-val'>{s['pnl_stress']/1000:+.0f} k€</div>
                      <div style='font-size:11px;color:#7A8BA8;margin-top:4px;font-family:DM Mono,monospace'>
                        VaR stressée : <b style='color:#e87373'>{s['var_stress']/1000:.0f} k€</b><br>
                        Ratio : <b style='color:#F0B429'>{rtxt}</b><br>
                        Choc : <b style='color:#FF4D6D'>{s['choc']*100:.2f}%</b>
                      </div>
                    </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Analyse comparative</div>", unsafe_allow_html=True)
        f8 = fig_stress_comparaison(sr, pv)
        st.pyplot(f8, use_container_width=True); plt.close(f8)

        st.markdown("<div class='section-header'>Tableau de synthèse</div>", unsafe_allow_html=True)
        rows_st = []
        for sc_name, s in sr.items():
            ratio = s["ratio"]
            rows_st.append({
                "Scénario":         sc_name,
                "Choc (%)":         f"{s['choc']*100:.2f}%",
                "P&L stressé (k€)": f"{s['pnl_stress']/1000:+.1f}",
                "VaR stressée (k€)":f"{s['var_stress']/1000:.1f}",
                "VaR normale (k€)": f"{s['var_normal']/1000:.1f}",
                "Ratio ×":          f"{ratio:.2f}" if not np.isnan(ratio) else "—",
                "Statut":           "🔴 ALERTE" if (not np.isnan(ratio) and ratio>2)
                                    else ("🟡 VIGILANCE" if (not np.isnan(ratio) and ratio>1.5)
                                    else "🟢 OK"),
            })
        st.dataframe(pd.DataFrame(rows_st).set_index("Scénario"), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : REPORTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "reporting":
    st.title("Génération des Rapports")

    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR.")
        st.stop()

    rend       = st.session_state["rendements"]
    var_res    = st.session_state["var_results"]
    bt_res     = st.session_state["bt_results"] or {}
    pv         = st.session_state["pv"] or 10_000_000
    opt_res    = st.session_state.get("opt_results")
    port_r     = rend.mean(axis=1).values
    alphas_    = sorted(list(list(var_res.values())[0].keys()))
    alpha_99   = max(alphas_)
    best       = "TVE-GARCH" if "TVE-GARCH" in var_res else list(var_res.keys())[-1]

    st.markdown("""
    <div class='info-box'>
    Générez un <b>rapport Excel</b> (4 feuilles) et un <b>rapport PDF exécutif</b>
    avec graphiques. Conformes aux exigences Bâle III/IV.
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-header'>Aperçu</div>", unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Méthode recommandée", best)
    c2.metric(f"VaR 99%",  f"{var_res[best][alpha_99]['VaR']/1000:.1f} k€")
    c3.metric(f"ES 99%",   f"{var_res[best][alpha_99]['ES']/1000:.1f} k€")
    n_ok    = sum(1 for m in bt_res for a in bt_res[m] if bt_res[m][a]["kupiec"]["valid"]) if bt_res else 0
    n_total = sum(len(bt_res[m]) for m in bt_res) if bt_res else 0
    c4.metric("Backtests validés", f"{n_ok}/{n_total}" if n_total else "—")

    bt_fmt = {m: {a: res for a, res in alphas.items()} for m, alphas in bt_res.items()}

    st.markdown("<div class='section-header'>Téléchargements</div>", unsafe_allow_html=True)
    col_x, col_p = st.columns(2)

    with col_x:
        st.markdown("""
        <div class='var-card'>
          <div class='var-card-title'>📊 Rapport Excel</div>
          <div style='font-size:12px;color:#E8EDF5;margin:8px 0;line-height:1.6'>
          4 feuilles : Résumé VaR · Backtesting · Markowitz · Données<br>
          Formatage professionnel, prêt pour la Direction.
          </div>
        </div>""", unsafe_allow_html=True)
        if HAS_XLSX:
            with st.spinner("Génération Excel…"):
                xlsx = generer_excel(rend, var_res, bt_fmt, pv, opt_res)
            if xlsx:
                st.download_button("⬇  Télécharger Excel",
                                   data=xlsx, file_name="VaR_Report_v41.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            st.warning("pip install openpyxl")

    with col_p:
        st.markdown("""
        <div class='var-card'>
          <div class='var-card-title'>📄 Rapport PDF</div>
          <div style='font-size:12px;color:#E8EDF5;margin:8px 0;line-height:1.6'>
          Couverture · VaR · Backtesting · Graphiques<br>
          Mise en page exécutive, prêt à envoyer.
          </div>
        </div>""", unsafe_allow_html=True)
        if HAS_PDF:
            with st.spinner("Génération PDF…"):
                fv_ = fig_var_comparaison(var_res, alpha_99, pv)
                fd_ = fig_distribution(port_r, var_res)
                pdf = generer_pdf(var_res, bt_fmt, pv, fv_, fd_)
            if pdf:
                st.download_button("⬇  Télécharger PDF",
                                   data=pdf, file_name="VaR_Risk_Report_v41.pdf",
                                   mime="application/pdf",
                                   use_container_width=True)
        else:
            st.warning("pip install reportlab")

    with st.expander("📦 requirements.txt — Déploiement Streamlit Cloud"):
        st.code("""streamlit>=1.32
yfinance>=0.2.36
pandas>=2.0
numpy>=1.26
scipy>=1.12
matplotlib>=3.8
openpyxl>=3.1
reportlab>=4.1""", language="text")
        st.markdown("""
**Déploiement :**
1. Pousser `app.py` + `requirements.txt` sur GitHub
2. [share.streamlit.io](https://share.streamlit.io) → New app → Deploy
        """)
