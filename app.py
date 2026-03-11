"""
╔══════════════════════════════════════════════════════════════╗
║        VaR ANALYTICS SUITE  ·  Streamlit App v4.0           ║
║        Département Gestion des Risques de Marché            ╠
║        Améliorations v4.0 :                                  ║
║          · Optimisation de portefeuille (Markowitz)          ║
║          · Stress-Testing (scénarios historiques)            ║
║          · ES analytique TVE-GARCH                           ║
║          · Matrice de corrélation interactive                ║
╚══════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from scipy import stats
from scipy.optimize import minimize, differential_evolution
import io, os, warnings
warnings.filterwarnings("ignore")

# ── Imports optionnels ────────────────────────────────────────────────────────
try:
    import yfinance as yf
    HAS_YF = True
except ImportError:
    HAS_YF = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak, Image)
    from reportlab.lib.colors import HexColor
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG STREAMLIT + CSS
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="VaR Analytics Suite",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] { font-family: 'Sora', sans-serif !important; }
.stApp { background: linear-gradient(135deg, #0B1628 0%, #0d1e38 60%, #091422 100%); }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0a1525 0%, #0c1a30 100%) !important;
    border-right: 1px solid rgba(201,168,76,0.15) !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span { color: #8aa0bc !important; font-size: 12px !important; }

h1 {
    font-family: 'Sora', sans-serif !important;
    background: linear-gradient(90deg, #C9A84C, #e8c56a) !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    font-size: 1.9rem !important; font-weight: 700 !important;
}
h2 { color: #d0daea !important; font-size: 1.15rem !important; font-weight: 600 !important; }
h3 { color: #C9A84C !important; font-size: 1rem !important; font-weight: 600 !important; }

[data-testid="metric-container"] {
    background: rgba(17,31,53,0.85) !important;
    border: 1px solid rgba(201,168,76,0.18) !important;
    border-radius: 10px !important;
    padding: 14px 16px !important;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3) !important;
    position: relative; overflow: hidden;
}
[data-testid="metric-container"]::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, #C9A84C, #2E6FD4);
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    color: #8aa0bc !important; font-size: 10px !important;
    text-transform: uppercase; letter-spacing: 1px;
    font-family: 'JetBrains Mono', monospace !important;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #C9A84C !important; font-size: 1.4rem !important;
    font-family: 'JetBrains Mono', monospace !important; font-weight: 700 !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid rgba(201,168,76,0.15) !important;
    border-radius: 8px !important;
}

.stButton > button {
    background: linear-gradient(135deg, #1a3060, #243d78) !important;
    color: #C9A84C !important; border: 1px solid rgba(201,168,76,0.4) !important;
    border-radius: 8px !important; font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important; font-size: 13px !important;
    padding: 8px 22px !important; transition: all 0.2s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #243d78, #2E6FD4) !important;
    border-color: #C9A84C !important;
    box-shadow: 0 4px 16px rgba(201,168,76,0.25) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #C9A84C, #e8c56a) !important;
    color: #0B1628 !important; border-color: transparent !important;
}

.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: rgba(17,31,53,0.9) !important;
    border: 1px solid rgba(201,168,76,0.25) !important;
    border-radius: 8px !important; color: #d0daea !important;
}

.stSlider > div > div > div > div {
    background: linear-gradient(90deg, #C9A84C, #2E6FD4) !important;
}

.streamlit-expanderHeader {
    background: rgba(17,31,53,0.6) !important;
    border: 1px solid rgba(201,168,76,0.15) !important;
    border-radius: 8px !important; color: #d0daea !important;
}

.var-card {
    background: rgba(17,31,53,0.85);
    border: 1px solid rgba(201,168,76,0.18);
    border-radius: 10px; padding: 16px 20px;
    margin-bottom: 12px; position: relative; overflow: hidden;
}
.var-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, #C9A84C, #2E6FD4);
}
.var-card-title { font-size: 11px; color: #8aa0bc; text-transform: uppercase;
    letter-spacing: 1.5px; font-family: 'JetBrains Mono', monospace; margin-bottom: 6px; }
.var-card-value { font-size: 22px; font-weight: 700; color: #e05252;
    font-family: 'JetBrains Mono', monospace; }
.var-card-pct { font-size: 11px; color: #8aa0bc; font-family: 'JetBrains Mono', monospace; }
.var-card-es { font-size: 13px; color: #e87373; font-family: 'JetBrains Mono', monospace; margin-top: 4px; }
.badge-rec {
    display: inline-block; background: #C9A84C; color: #0B1628;
    font-size: 9px; font-weight: 700; padding: 2px 7px; border-radius: 4px;
    letter-spacing: 1px; font-family: 'JetBrains Mono', monospace; margin-left: 8px;
}
.section-header {
    border-left: 3px solid #C9A84C; padding-left: 12px;
    margin: 24px 0 12px; font-size: 14px; font-weight: 600; color: #d0daea;
}
.info-box {
    background: rgba(46,111,212,0.08); border: 1px solid rgba(46,111,212,0.25);
    border-radius: 8px; padding: 14px 18px; margin: 12px 0;
    font-size: 12.5px; color: #c0d0e8; line-height: 1.7;
}
.stress-card {
    background: rgba(224,82,82,0.08); border: 1px solid rgba(224,82,82,0.25);
    border-radius: 10px; padding: 14px 18px; margin-bottom: 10px;
}
.stress-title { font-size: 11px; color: #e87373; font-family: 'JetBrains Mono', monospace;
    text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
.stress-val { font-size: 20px; font-weight: 700; color: #e05252;
    font-family: 'JetBrains Mono', monospace; }
.markowitz-card {
    background: rgba(29,184,122,0.07); border: 1px solid rgba(29,184,122,0.25);
    border-radius: 10px; padding: 14px 18px; margin-bottom: 10px;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# DONNÉES & CACHE
# ══════════════════════════════════════════════════════════════════════════════

ACTIFS_DISPONIBLES = {
    "Apple (AAPL)":         "AAPL",
    "Microsoft (MSFT)":     "MSFT",
    "LVMH (MC.PA)":         "MC.PA",
    "TotalEnergies (TTE)":  "TTE.PA",
    "BNP Paribas (BNP)":    "BNP.PA",
    "Nestlé (NESN.SW)":     "NESN.SW",
    "SAP (SAP)":            "SAP",
    "Airbus (AIR.PA)":      "AIR.PA",
    "Tesla (TSLA)":         "TSLA",
    "Amazon (AMZN)":        "AMZN",
    "Nvidia (NVDA)":        "NVDA",
    "Safran (SAF.PA)":      "SAF.PA",
    "L'Oréal (OR.PA)":      "OR.PA",
    "ASML (ASML.AS)":       "ASML.AS",
    "Hermès (RMS.PA)":      "RMS.PA",
}

SECTEURS = {
    "AAPL": "Technologie", "MSFT": "Technologie", "MC.PA": "Luxe",
    "TTE.PA": "Énergie",   "BNP.PA": "Finance",   "NESN.SW": "Conso.",
    "SAP": "Technologie",  "AIR.PA": "Aéronaut.", "TSLA": "Auto/Tech",
    "AMZN": "Commerce",    "NVDA": "Technologie", "SAF.PA": "Aéronaut.",
    "OR.PA": "Beauté",     "ASML.AS": "Technologie", "RMS.PA": "Luxe",
}

# Scénarios de stress historiques (chocs journaliers approximatifs)
SCENARIOS_STRESS = {
    "Crise Financière 2008 (Lehman, 15/09/2008)": {
        "description": "Faillite de Lehman Brothers — pire journée de la crise des subprimes",
        "choc_marche": -0.0469,
        "choc_vol": 3.8,
        "date": "15 sept. 2008",
        "couleur": "#e05252"
    },
    "Flash Crash 2010 (06/05/2010)": {
        "description": "Effondrement éclair de 1000 pts du DJIA en quelques minutes",
        "choc_marche": -0.034,
        "choc_vol": 2.5,
        "date": "6 mai 2010",
        "couleur": "#f97316"
    },
    "Brexit (24/06/2016)": {
        "description": "Référendum Brexit — choc sur marchés européens et GBP",
        "choc_marche": -0.0281,
        "choc_vol": 2.1,
        "date": "24 juin 2016",
        "couleur": "#a855f7"
    },
    "COVID-19 (16/03/2020)": {
        "description": "Pire séance COVID — annonce confinements mondiaux",
        "choc_marche": -0.0598,
        "choc_vol": 4.2,
        "date": "16 mars 2020",
        "couleur": "#e05252"
    },
    "Krach Obligataire 2022 (Q1)": {
        "description": "Remontée brutale des taux Fed — pire Q1 depuis 1973",
        "choc_marche": -0.0312,
        "choc_vol": 2.0,
        "date": "T1 2022",
        "couleur": "#C9A84C"
    },
    "Scénario Adverse (−3σ)": {
        "description": "Choc extrême calibré à −3 écarts-types journaliers",
        "choc_marche": None,  # calculé dynamiquement
        "choc_vol": 3.0,
        "date": "Hypothétique",
        "couleur": "#8b5cf6"
    },
}

@st.cache_data(show_spinner=False)
def telecharger_donnees(tickers: list, date_debut, date_fin) -> pd.DataFrame:
    if not HAS_YF:
        return pd.DataFrame()
    data = yf.download(tickers, start=str(date_debut), end=str(date_fin),
                       auto_adjust=True, progress=False)
    if isinstance(data.columns, pd.MultiIndex):
        if "Close" in data.columns.get_level_values(0):
            prix = data["Close"].copy()
        else:
            prix = data.xs(data.columns.get_level_values(0)[0], axis=1, level=0)
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
    mu_d  = np.full(n, 0.0004)
    sig_d = np.full(n, 0.012)
    corr  = np.eye(n) + 0.4 * (np.ones((n,n)) - np.eye(n))
    L = np.linalg.cholesky(corr)
    z = np.random.randn(n_days, n) @ L.T
    lr = mu_d + sig_d * z
    mid = n_days // 2
    lr[mid:mid+20] *= 3.5
    prices = 100 * np.exp(np.cumsum(lr, axis=0))
    dates = pd.bdate_range(end="2024-12-31", periods=n_days)
    return pd.DataFrame(prices, index=dates, columns=tickers)


# ══════════════════════════════════════════════════════════════════════════════
# OPTIMISATION DE PORTEFEUILLE — MARKOWITZ
# ══════════════════════════════════════════════════════════════════════════════

def optimiser_portefeuille(rendements: pd.DataFrame, rf: float = 0.03) -> dict:
    """Calcule la frontière efficiente et 3 portefeuilles optimaux."""
    mu = rendements.mean().values * 252
    cov = rendements.cov().values * 252
    n = len(mu)
    rf_d = rf / 252

    def port_stats(w):
        ret = w @ mu
        vol = np.sqrt(w @ cov @ w)
        sharpe = (ret - rf) / vol if vol > 0 else 0
        return ret, vol, sharpe

    constraints = [{"type": "eq", "fun": lambda w: np.sum(w) - 1}]
    bounds = [(0.0, 1.0)] * n
    w0 = np.ones(n) / n

    # 1. Sharpe Maximum
    def neg_sharpe(w):
        r, v, _ = port_stats(w)
        return -(r - rf) / v if v > 0 else 1e10

    res_sharpe = minimize(neg_sharpe, w0, method='SLSQP',
                          bounds=bounds, constraints=constraints)
    w_sharpe = res_sharpe.x if res_sharpe.success else w0

    # 2. Variance Minimum
    def port_var(w):
        return w @ cov @ w

    res_minvar = minimize(port_var, w0, method='SLSQP',
                          bounds=bounds, constraints=constraints)
    w_minvar = res_minvar.x if res_minvar.success else w0

    # 3. Équipondéré
    w_equi = np.ones(n) / n

    # Frontière efficiente
    ret_range = np.linspace(mu.min(), mu.max(), 60)
    frontier_vols, frontier_rets = [], []
    for target_ret in ret_range:
        cons = [{"type": "eq", "fun": lambda w: np.sum(w) - 1},
                {"type": "eq", "fun": lambda w, r=target_ret: w @ mu - r}]
        res = minimize(port_var, w0, method='SLSQP', bounds=bounds, constraints=cons)
        if res.success:
            frontier_vols.append(np.sqrt(res.fun))
            frontier_rets.append(target_ret)

    tickers = list(rendements.columns)
    results = {
        "tickers": tickers,
        "mu": mu,
        "cov": cov,
        "sharpe": {
            "weights": w_sharpe,
            "stats": port_stats(w_sharpe),
            "label": "Sharpe Max."
        },
        "minvar": {
            "weights": w_minvar,
            "stats": port_stats(w_minvar),
            "label": "Variance Min."
        },
        "equi": {
            "weights": w_equi,
            "stats": port_stats(w_equi),
            "label": "Équipondéré"
        },
        "frontier": (frontier_vols, frontier_rets)
    }
    return results


def fig_frontiere_efficiente(opt_results: dict) -> plt.Figure:
    PLT_DARK = {
        "figure.facecolor": "#0d1e38", "axes.facecolor": "#0d1e38",
        "axes.edgecolor": "#2a3f5f", "axes.labelcolor": "#8aa0bc",
        "xtick.color": "#8aa0bc", "ytick.color": "#8aa0bc",
        "grid.color": "#1a2d48", "text.color": "#d0daea",
        "legend.facecolor": "#111f35", "legend.edgecolor": "#2a3f5f",
    }
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(10, 5))
        fv, fr = opt_results["frontier"]
        if fv:
            ax.plot([v*100 for v in fv], [r*100 for r in fr],
                    color="#2E6FD4", lw=2.5, label="Frontière efficiente")
            ax.fill_betweenx([r*100 for r in fr], [v*100 for v in fv],
                             alpha=0.08, color="#2E6FD4")

        # Actifs individuels
        mu = opt_results["mu"]
        cov = opt_results["cov"]
        tickers = opt_results["tickers"]
        for i, t in enumerate(tickers):
            vol_i = np.sqrt(cov[i, i]) * 100
            mu_i  = mu[i] * 100
            ax.scatter(vol_i, mu_i, s=60, color="#8aa0bc", zorder=5, alpha=0.7)
            ax.annotate(t, (vol_i, mu_i), fontsize=7, color="#8aa0bc",
                        xytext=(4, 2), textcoords="offset points")

        # 3 portefeuilles
        styles = [
            ("sharpe", "#C9A84C", "★", 180),
            ("minvar", "#1db87a", "▼", 140),
            ("equi",   "#a855f7", "●", 100),
        ]
        for key, col, marker, size in styles:
            p = opt_results[key]
            r, v, s = p["stats"]
            ax.scatter(v*100, r*100, s=size, color=col, zorder=10,
                       marker=marker if marker != "★" else "*",
                       label=f"{p['label']} (Sharpe={s:.2f})")
            ax.annotate(p["label"], (v*100, r*100), fontsize=8, color=col,
                        fontweight="bold", xytext=(6, 4), textcoords="offset points")

        ax.set_xlabel("Volatilité annualisée (%)", fontsize=9)
        ax.set_ylabel("Rendement annualisé (%)", fontsize=9)
        ax.set_title("Frontière Efficiente de Markowitz", fontsize=11,
                     color="#C9A84C", pad=10)
        ax.legend(fontsize=8, loc="lower right")
        ax.grid(True, alpha=0.2)
        plt.tight_layout()
        return fig


def fig_poids_portefeuilles(opt_results: dict) -> plt.Figure:
    PLT_DARK = {
        "figure.facecolor": "#0d1e38", "axes.facecolor": "#0d1e38",
        "axes.edgecolor": "#2a3f5f", "axes.labelcolor": "#8aa0bc",
        "xtick.color": "#8aa0bc", "ytick.color": "#8aa0bc",
        "grid.color": "#1a2d48", "text.color": "#d0daea",
        "legend.facecolor": "#111f35", "legend.edgecolor": "#2a3f5f",
    }
    tickers = opt_results["tickers"]
    keys = ["sharpe", "minvar", "equi"]
    labels = ["Sharpe Max.", "Variance Min.", "Équipondéré"]
    colors_ = ["#C9A84C", "#1db87a", "#a855f7"]

    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1, 3, figsize=(12, 4))
        for ax, key, label, col in zip(axes, keys, labels, colors_):
            w = opt_results[key]["weights"]
            mask = w > 0.005
            w_f = w[mask]
            t_f = [tickers[i] for i, m in enumerate(mask) if m]
            wedges, texts, autotexts = ax.pie(
                w_f, labels=t_f, autopct='%1.1f%%',
                colors=[col] + [f"{col}99", f"{col}77", f"{col}55",
                                "#2E6FD4", "#2E6FD499", "#2E6FD477"][:len(w_f)-1],
                startangle=90, pctdistance=0.75,
                wedgeprops={"edgecolor": "#0d1e38", "linewidth": 1.5}
            )
            for text in texts: text.set_fontsize(7); text.set_color("#d0daea")
            for at in autotexts: at.set_fontsize(6.5); at.set_color("#0d1e38"); at.set_fontweight("bold")
            r, v, s = opt_results[key]["stats"]
            ax.set_title(f"{label}\nRdt: {r*100:.1f}% · Vol: {v*100:.1f}% · Sharpe: {s:.2f}",
                         fontsize=8.5, color=col, pad=6)
        plt.tight_layout()
        return fig


# ══════════════════════════════════════════════════════════════════════════════
# STRESS-TESTING
# ══════════════════════════════════════════════════════════════════════════════

def appliquer_stress(rendements_port: np.ndarray, pv: float,
                     scenario: dict, var_results: dict, alpha: float = 0.99) -> dict:
    mu   = rendements_port.mean()
    sig  = rendements_port.std()
    choc = scenario["choc_marche"] if scenario["choc_marche"] is not None else mu - 3*sig

    # P&L stressé
    pnl_stress = choc * pv

    # VaR stressée (vol multipliée)
    sigma_stress = sig * scenario["choc_vol"]
    z = stats.norm.ppf(1 - alpha)
    var_stress_vc = -z * sigma_stress * pv

    # Comparaison avec VaR normale
    var_normal = var_results.get("Variance-Covariance", {}).get(alpha, {}).get("VaR", 0)

    # Pire perte historique sur la période
    pire_perte = rendements_port.min() * pv

    return {
        "pnl_stress": pnl_stress,
        "var_stress":  var_stress_vc,
        "var_normal":  var_normal,
        "ratio":       var_stress_vc / var_normal if var_normal > 0 else np.nan,
        "pire_perte":  pire_perte,
        "choc":        choc,
        "choc_vol":    scenario["choc_vol"],
    }


def fig_stress_comparaison(stress_results: dict, pv: float) -> plt.Figure:
    PLT_DARK = {
        "figure.facecolor": "#0d1e38", "axes.facecolor": "#0d1e38",
        "axes.edgecolor": "#2a3f5f", "axes.labelcolor": "#8aa0bc",
        "xtick.color": "#8aa0bc", "ytick.color": "#8aa0bc",
        "grid.color": "#1a2d48", "text.color": "#d0daea",
        "legend.facecolor": "#111f35", "legend.edgecolor": "#2a3f5f",
    }
    with plt.rc_context(PLT_DARK):
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4))

        scenarios = list(stress_results.keys())
        pnls    = [stress_results[s]["pnl_stress"]/1000 for s in scenarios]
        vars_st = [stress_results[s]["var_stress"]/1000  for s in scenarios]
        vars_nm = [stress_results[s]["var_normal"]/1000  for s in scenarios]

        short_sc = [s.split("(")[0].strip()[:22] for s in scenarios]
        x = np.arange(len(scenarios))

        ax1.barh(x, [abs(p) for p in pnls], color="#e05252", alpha=0.8, label="P&L stressé")
        ax1.barh(x, vars_nm, color="#2E6FD4", alpha=0.5, label="VaR normale 99%")
        ax1.set_yticks(x); ax1.set_yticklabels(short_sc, fontsize=8)
        ax1.set_xlabel("k€", fontsize=9)
        ax1.set_title("Impact P&L vs VaR normale", fontsize=10, color="#C9A84C", pad=6)
        ax1.legend(fontsize=8); ax1.grid(axis="x", alpha=0.2)

        ratios = [stress_results[s]["ratio"] for s in scenarios]
        colors_r = ["#e05252" if r > 2 else "#C9A84C" if r > 1.5 else "#1db87a" for r in ratios]
        ax2.bar(short_sc, ratios, color=colors_r, alpha=0.85)
        ax2.axhline(1.0, color="#1db87a", lw=1.5, ls="--", label="VaR normale (ratio=1)")
        ax2.axhline(2.0, color="#e05252", lw=1.2, ls=":", label="Seuil d'alerte (×2)")
        ax2.set_xticks(range(len(short_sc)))
        ax2.set_xticklabels(short_sc, rotation=30, ha="right", fontsize=7.5)
        ax2.set_ylabel("Ratio VaR Stressée / VaR Normale", fontsize=9)
        ax2.set_title("Multiplicateurs de stress", fontsize=10, color="#C9A84C", pad=6)
        ax2.legend(fontsize=8); ax2.grid(axis="y", alpha=0.2)

        plt.tight_layout()
        return fig


# ══════════════════════════════════════════════════════════════════════════════
# MOTEUR VAR — 7 MÉTHODES (avec ES analytique TVE-GARCH)
# ══════════════════════════════════════════════════════════════════════════════

class VaREngine:
    def __init__(self, rendements: np.ndarray, pv: float = 10_000_000, horizon: int = 1):
        self.r       = rendements
        self.pv      = pv
        self.horizon = horizon
        self.n       = len(rendements)

    def historique(self, alpha: float) -> dict:
        q   = np.percentile(self.r, (1-alpha)*100)
        exc = self.r[self.r <= q]
        es  = exc.mean() if len(exc) > 0 else q
        return {"VaR": -q*self.pv*np.sqrt(self.horizon),
                "ES":  -es*self.pv*np.sqrt(self.horizon),
                "VaR_pct": -q}

    def variance_covariance(self, alpha: float) -> dict:
        mu, sigma = self.r.mean(), self.r.std()
        z   = stats.norm.ppf(1-alpha)
        var = -(mu + z*sigma) * np.sqrt(self.horizon)
        es  = (sigma*stats.norm.pdf(z)/(1-alpha) - mu) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"μ={mu*100:.4f}%, σ={sigma*100:.4f}%"}

    def riskmetrics(self, alpha: float, lam: float = 0.94) -> dict:
        r = self.r
        sig2 = np.zeros(len(r)); sig2[0] = r[0]**2
        for t in range(1, len(r)):
            sig2[t] = lam*sig2[t-1] + (1-lam)*r[t-1]**2
        sigma_t = np.sqrt(sig2[-1])
        z = stats.norm.ppf(1-alpha)
        var = -z * sigma_t * np.sqrt(self.horizon)
        es  = sigma_t * stats.norm.pdf(z) / (1-alpha) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"λ=0.94, σ_t={sigma_t*100:.4f}%"}

    def cornish_fisher(self, alpha: float) -> dict:
        mu, sigma = self.r.mean(), self.r.std()
        s = float(stats.skew(self.r))
        k = float(stats.kurtosis(self.r))
        z = stats.norm.ppf(1-alpha)
        z_cf = z + (z**2-1)*s/6 + (z**3-3*z)*k/24 - (2*z**3-5*z)*s**2/36
        var = -(mu + z_cf*sigma) * np.sqrt(self.horizon)
        es  = (sigma*(stats.norm.pdf(z_cf)/(1-alpha)) - mu) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"z_CF={z_cf:.4f}, skew={s:.3f}, kurt={k:.3f}"}

    def _fit_garch(self):
        r = self.r
        def neg_ll(p):
            w, a, b = p
            if w<=0 or a<=0 or b<=0 or a+b>=1: return 1e10
            sig2 = np.zeros(len(r)); sig2[0] = np.var(r)
            for t in range(1,len(r)):
                sig2[t] = w + a*r[t-1]**2 + b*sig2[t-1]
            return 0.5*np.sum(np.log(2*np.pi*sig2) + r**2/sig2)
        try:
            res = minimize(neg_ll, [1e-6,0.08,0.89], method='L-BFGS-B',
                           bounds=[(1e-8,None),(0.001,0.3),(0.5,0.999)])
            w, a, b = res.x
        except Exception:
            w, a, b = 5e-7, 0.09, 0.90
        sig2 = np.zeros(len(r)); sig2[0] = np.var(r)
        for t in range(1,len(r)):
            sig2[t] = w + a*r[t-1]**2 + b*sig2[t-1]
        return w, a, b, sig2

    def garch(self, alpha: float) -> dict:
        w, a, b, sig2 = self._fit_garch()
        r = self.r
        sigma_f = np.sqrt(w + a*r[-1]**2 + b*sig2[-1])
        z   = stats.norm.ppf(1-alpha)
        var = -z * sigma_f * np.sqrt(self.horizon)
        es  = sigma_f * stats.norm.pdf(z) / (1-alpha) * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"ω={w:.2e}, α={a:.3f}, β={b:.3f}"}

    def _fit_gpd(self, excesses: np.ndarray):
        """Ajustement GPD par MLE."""
        def neg_ll_gpd(p):
            xi, beta = p
            if beta <= 0: return 1e10
            u_ = excesses / beta
            if xi != 0:
                if np.any(1+xi*u_ <= 0): return 1e10
                return len(u_)*np.log(beta) + (1+1/xi)*np.sum(np.log(1+xi*u_))
            return len(u_)*np.log(beta) + np.sum(u_)
        try:
            res = minimize(neg_ll_gpd, [0.1, np.std(excesses)], method='L-BFGS-B',
                           bounds=[(-0.5,0.5),(1e-6,None)])
            return res.x if res.success else np.array([0.1, np.std(excesses)])
        except Exception:
            return np.array([0.1, np.std(excesses)])

    def tve(self, alpha: float) -> dict:
        losses = -self.r
        u = np.percentile(losses, 90)
        exc = losses[losses > u] - u
        xi, beta = self._fit_gpd(exc)
        n_u, n = len(exc), len(losses)
        p = 1 - alpha
        if xi != 0:
            var = u + (beta/xi)*((n/n_u*p)**(-xi)-1)
        else:
            var = u + beta*np.log(n/n_u*p)
        var = max(var, 0)
        # ES analytique GPD
        if xi < 1:
            es = (var + beta - xi*u) / (1 - xi)
        else:
            es = var * 2
        return {"VaR": var*self.pv*np.sqrt(self.horizon),
                "ES":  es*self.pv*np.sqrt(self.horizon),
                "VaR_pct": var,
                "params": f"ξ={xi:.3f}, β={beta:.4f}, u={u:.4f}"}

    def tve_garch(self, alpha: float) -> dict:
        """TVE-GARCH avec ES analytique depuis la GPD."""
        r = self.r
        w, a, b, sig2 = self._fit_garch()
        # Résidus standardisés
        resid = r / np.sqrt(sig2)
        # Ajustement GPD sur les résidus
        losses_r = -resid
        u_r = np.percentile(losses_r, 90)
        exc_r = losses_r[losses_r > u_r] - u_r
        xi, beta_r = self._fit_gpd(exc_r)
        n_u, n = len(exc_r), len(losses_r)
        p = 1 - alpha
        if xi != 0:
            var_r = u_r + (beta_r/xi)*((n/n_u*p)**(-xi)-1)
        else:
            var_r = u_r + beta_r*np.log(n/n_u*p)
        var_r = max(var_r, 0)
        # ES analytique GPD
        if xi < 1:
            es_r = (var_r + beta_r - xi*u_r) / (1 - xi)
        else:
            es_r = var_r * 2
        # Volatilité prévisionnelle GARCH
        sigma_f = np.sqrt(w + a*r[-1]**2 + b*sig2[-1])
        var = var_r * sigma_f * np.sqrt(self.horizon)
        es  = es_r  * sigma_f * np.sqrt(self.horizon)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": var,
                "params": f"GARCH+GPD hybride, ξ={xi:.3f}, σ_f={sigma_f*100:.4f}%"}

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

def kupiec_test(rendements: np.ndarray, var_pct: float, alpha: float) -> dict:
    exc = (rendements < -var_pct).astype(int)
    N, T = int(exc.sum()), len(exc)
    if T == 0: return {"LR": np.nan, "p_value": np.nan, "valid": False, "N": 0, "T": 0}
    p0 = 1 - alpha
    p_hat = N / T
    if p_hat == 0:
        lr = -2 * T * np.log(1 - p0)
    elif p_hat == 1:
        lr = -2 * N * np.log(p0)
    else:
        lr = -2 * (T*np.log(1-p0) + N*np.log(p0)
                   - N*np.log(p_hat) - (T-N)*np.log(1-p_hat))
    pv = 1 - stats.chi2.cdf(max(lr,0), df=1)
    return {"LR": lr, "p_value": pv, "valid": pv > 0.05,
            "N": N, "T": T, "rate": p_hat, "expected": p0}


def christoffersen_test(rendements: np.ndarray, var_pct: float) -> dict:
    exc = (rendements < -var_pct).astype(int)
    n00 = np.sum((exc[:-1]==0) & (exc[1:]==0))
    n01 = np.sum((exc[:-1]==0) & (exc[1:]==1))
    n10 = np.sum((exc[:-1]==1) & (exc[1:]==0))
    n11 = np.sum((exc[:-1]==1) & (exc[1:]==1))
    pi01 = n01 / (n00+n01+1e-10)
    pi11 = n11 / (n10+n11+1e-10)
    pi   = (n01+n11) / (n00+n01+n10+n11+1e-10)
    try:
        lr = -2*(
            (n00+n10)*np.log(max(1-pi,1e-15))+(n01+n11)*np.log(max(pi,1e-15))
            - n00*np.log(max(1-pi01,1e-15)) - n01*np.log(max(pi01,1e-15))
            - n10*np.log(max(1-pi11,1e-15)) - n11*np.log(max(pi11,1e-15))
        )
    except Exception:
        lr = np.nan
    pv = 1 - stats.chi2.cdf(max(lr,0), df=1) if not np.isnan(lr) else np.nan
    return {"LR_ind": lr, "p_value_ind": pv,
            "valid": pv > 0.05 if not np.isnan(pv) else False}


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

PLT_DARK = {
    "figure.facecolor": "#0d1e38", "axes.facecolor": "#0d1e38",
    "axes.edgecolor": "#2a3f5f", "axes.labelcolor": "#8aa0bc",
    "xtick.color": "#8aa0bc", "ytick.color": "#8aa0bc",
    "grid.color": "#1a2d48", "text.color": "#d0daea",
    "legend.facecolor": "#111f35", "legend.edgecolor": "#2a3f5f",
}

def fig_perf(rendements: pd.DataFrame, tickers: list) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, (ax1, ax2) = plt.subplots(2,1, figsize=(11,5),
                                        gridspec_kw={"height_ratios":[3,1]})
        colors = ["#2E6FD4","#C9A84C","#1db87a","#e05252","#a855f7","#f97316","#06b6d4","#ec4899"]
        port_r = rendements.mean(axis=1)
        cumul  = (1+port_r).cumprod()
        ax1.fill_between(range(len(cumul)), cumul, 1, alpha=0.1, color="#2E6FD4")
        ax1.plot(cumul.values, color="#2E6FD4", lw=1.8, label="Portefeuille")
        for i,t in enumerate(tickers[:4]):
            if t in rendements.columns:
                c = (1+rendements[t]).cumprod()
                ax1.plot(c.values, color=colors[(i+1)%len(colors)], lw=0.8, alpha=0.5, label=t)
        ax1.axhline(1, color="#C9A84C", lw=0.7, ls="--", alpha=0.6)
        ax1.set_title("Performance cumulée", fontsize=10, color="#C9A84C", pad=6)
        ax1.legend(fontsize=7, loc="upper left"); ax1.grid(True, alpha=0.3)
        ax1.tick_params(labelsize=8)
        col_bars = ["#1db87a" if v>=0 else "#e05252" for v in port_r]
        ax2.bar(range(len(port_r)), port_r*100, color=col_bars, alpha=0.7, width=1)
        ax2.axhline(0, color="#C9A84C", lw=0.5)
        ax2.set_title("Rendements journaliers (%)", fontsize=9, color="#8aa0bc", pad=4)
        ax2.tick_params(labelsize=7); ax2.grid(True, alpha=0.2)
        plt.tight_layout(); return fig


def fig_correlation(rendements: pd.DataFrame) -> plt.Figure:
    corr = rendements.corr()
    tickers = list(corr.columns)
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(min(10, len(tickers)*1.2+2),
                                        min(8, len(tickers)*1.0+2)))
        cmap = plt.cm.RdYlGn
        im = ax.imshow(corr.values, cmap=cmap, vmin=-1, vmax=1, aspect="auto")
        plt.colorbar(im, ax=ax, shrink=0.8, label="Corrélation")
        ax.set_xticks(range(len(tickers))); ax.set_yticks(range(len(tickers)))
        ax.set_xticklabels(tickers, rotation=45, ha="right", fontsize=8)
        ax.set_yticklabels(tickers, fontsize=8)
        for i in range(len(tickers)):
            for j in range(len(tickers)):
                v = corr.values[i, j]
                ax.text(j, i, f"{v:.2f}", ha="center", va="center",
                        fontsize=7, color="black" if abs(v) < 0.5 else "white",
                        fontweight="bold")
        ax.set_title("Matrice de Corrélation", fontsize=11, color="#C9A84C", pad=8)
        plt.tight_layout(); return fig


def fig_var_comparaison(var_results: dict, conf: float, pv: float) -> plt.Figure:
    methods = list(var_results.keys())
    vars_   = [var_results[m][conf]["VaR"]/1000 for m in methods]
    ess_    = [var_results[m][conf]["ES"]/1000  for m in methods]
    colors  = ["#2E6FD4","#C9A84C","#1db87a","#e05252","#a855f7","#f97316","#06b6d4"]
    short   = [m.replace("Variance-Covariance","VCV").replace("Cornish-Fisher","C-Fisher")
               .replace("RiskMetrics","RiskM.") for m in methods]
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(11,4))
        x = np.arange(len(methods)); w = 0.38
        b1 = ax.bar(x-w/2, vars_, w, color=colors, alpha=0.85, label="VaR")
        b2 = ax.bar(x+w/2, ess_,  w, color=colors, alpha=0.45, label="ES")
        for b in b1:
            ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.5,
                    f"{b.get_height():.0f}k", ha="center", va="bottom",
                    fontsize=7, color="#e05252", fontweight="bold")
        ax.set_xticks(x); ax.set_xticklabels(short, rotation=25, ha="right", fontsize=8)
        ax.set_ylabel("k€", fontsize=9); ax.grid(axis="y", alpha=0.25)
        ax.set_title(f"VaR & ES — {conf*100:.0f}% — Portefeuille {pv/1e6:.0f}M€",
                     fontsize=10, color="#C9A84C", pad=8)
        ax.legend(fontsize=8); plt.tight_layout(); return fig


def fig_distribution(rendements: np.ndarray, var_results: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(11,4))
        ax.hist(rendements*100, bins=70, density=True, color="#2E6FD4",
                alpha=0.55, edgecolor="#1a3060", lw=0.3, label="Rendements obs.")
        mu, sig = rendements.mean(), rendements.std()
        x = np.linspace(rendements.min(), rendements.max(), 300)
        ax.plot(x*100, stats.norm.pdf(x,mu,sig)/100, color="#C9A84C",
                lw=2, ls="--", label="N(μ,σ)")
        colors_v = {"Historique":"#e05252","GARCH(1,1)":"#a855f7","TVE-GARCH":"#f97316"}
        for meth, col in colors_v.items():
            if meth in var_results:
                p = var_results[meth].get(0.99,{}).get("VaR_pct")
                if p and not np.isnan(p):
                    ax.axvline(-p*100, color=col, lw=1.5, ls=":",
                               label=f"VaR 99% {meth}")
        ax.set_xlabel("Rendement journalier (%)", fontsize=9)
        ax.set_ylabel("Densité", fontsize=9)
        ax.set_title("Distribution des rendements & VaR 99%", fontsize=10, color="#C9A84C", pad=8)
        ax.legend(fontsize=7); ax.grid(True, alpha=0.2)
        plt.tight_layout(); return fig


def fig_backtesting(bt_results: dict) -> plt.Figure:
    methods = list(bt_results.keys())
    short   = [m.replace("Variance-Covariance","VCV").replace("Cornish-Fisher","C-Fisher")
               .replace("RiskMetrics","RiskM.") for m in methods]
    p95 = [bt_results[m][0.95]["p_value"] for m in methods if 0.95 in bt_results[m]]
    p99 = [bt_results[m][0.99]["p_value"] for m in methods if 0.99 in bt_results[m]]
    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1,2, figsize=(11,3.5))
        for ax, pvals, title in zip(axes, [p95, p99], ["Kupiec 95%","Kupiec 99%"]):
            cols = ["#1db87a" if p>0.05 else "#e05252" for p in pvals]
            ax.bar(short[:len(pvals)], pvals, color=cols, alpha=0.85)
            ax.axhline(0.05, color="#C9A84C", lw=1.8, ls="--", label="Seuil 5%")
            ax.set_xticks(range(len(short[:len(pvals)])))
            ax.set_xticklabels(short[:len(pvals)], rotation=30, ha="right", fontsize=7.5)
            ax.set_ylabel("p-value", fontsize=9); ax.grid(axis="y", alpha=0.2)
            ax.set_title(title, fontsize=10, color="#C9A84C", pad=6)
            ax.legend(fontsize=8)
        plt.tight_layout(); return fig


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def generer_excel(rendements_df: pd.DataFrame, var_results: dict,
                   bt_results: dict, pv: float, opt_results: dict = None) -> bytes | None:
    if not HAS_XLSX: return None
    wb = Workbook()
    NAVY, BLUE, GOLD = "1B2A4A", "2E5FA3", "C9A84C"
    WHITE, LGRAY = "FFFFFF", "F0F4FA"

    def th(ws, r, c, v, bg=NAVY, fg=WHITE):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Calibri", bold=True, color=fg, size=10)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="CCCCCC")
        cell.border = Border(left=s, right=s, top=s, bottom=s)
        return cell

    def td(ws, r, c, v, bg=WHITE):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Calibri", size=9)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="DDDDDD")
        cell.border = Border(left=s, right=s, top=s, bottom=s)
        return cell

    # Feuille 1 : Résumé VaR
    ws1 = wb.active; ws1.title = "📊 Résumé VaR"
    ws1.merge_cells("A1:F1")
    c = ws1["A1"]; c.value = "RAPPORT VaR — SYNTHÈSE DES RÉSULTATS"
    c.font = Font(name="Calibri", bold=True, color=WHITE, size=14)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30
    headers = ["Méthode","VaR 95% (€)","VaR 95% (%)","VaR 99% (€)","VaR 99% (%)","ES 99% (€)"]
    for j,h in enumerate(headers,1): th(ws1, 2, j, h, bg=BLUE)
    for i,(m,res) in enumerate(var_results.items(), 3):
        r95 = res.get(0.95,{}); r99 = res.get(0.99,{})
        bg = LGRAY if i%2==0 else WHITE
        td(ws1,i,1,m,bg=bg)
        td(ws1,i,2,round(r95.get("VaR",0),0),bg=bg)
        td(ws1,i,3,f"{r95.get('VaR_pct',0)*100:.3f}%",bg=bg)
        td(ws1,i,4,round(r99.get("VaR",0),0),bg=bg)
        td(ws1,i,5,f"{r99.get('VaR_pct',0)*100:.3f}%",bg=bg)
        td(ws1,i,6,round(r99.get("ES",0),0),bg=bg)
    for w,col in zip([26,16,12,16,12,16],["A","B","C","D","E","F"]):
        ws1.column_dimensions[col].width = w

    # Feuille 2 : Backtesting
    ws2 = wb.create_sheet("🧪 Backtesting")
    hdrs = ["Méthode","CL","Exceptions","T","Taux obs.","Taux att.",
            "Kupiec LR","Kupiec p","Kupiec OK","CC LR","CC p","CC OK"]
    for j,h in enumerate(hdrs,1): th(ws2,1,j,h,bg=BLUE)
    row = 2
    for m, alphas in bt_results.items():
        for a, res in alphas.items():
            bg = LGRAY if row%2==0 else WHITE
            k = res.get("kupiec",{}); cc = res.get("cc",{})
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
    for w,col in zip([26,6,12,8,10,10,10,10,10,10,10,10],
                     ["A","B","C","D","E","F","G","H","I","J","K","L"]):
        ws2.column_dimensions[col].width = w

    # Feuille 3 : Markowitz (si disponible)
    if opt_results:
        ws3 = wb.create_sheet("📐 Markowitz")
        th(ws3,1,1,"Actif",bg=BLUE); th(ws3,1,2,"Sharpe Max. (%)",bg=BLUE)
        th(ws3,1,3,"Variance Min. (%)",bg=BLUE); th(ws3,1,4,"Équipondéré (%)",bg=BLUE)
        for i, ticker in enumerate(opt_results["tickers"], 2):
            bg = LGRAY if i%2==0 else WHITE
            td(ws3,i,1,ticker,bg=bg)
            td(ws3,i,2,f"{opt_results['sharpe']['weights'][i-2]*100:.1f}%",bg=bg)
            td(ws3,i,3,f"{opt_results['minvar']['weights'][i-2]*100:.1f}%",bg=bg)
            td(ws3,i,4,f"{opt_results['equi']['weights'][i-2]*100:.1f}%",bg=bg)
        row_stats = len(opt_results["tickers"]) + 3
        th(ws3,row_stats,1,"Statistiques",bg=NAVY,fg=WHITE)
        th(ws3,row_stats,2,"Sharpe Max.",bg=NAVY,fg=WHITE)
        th(ws3,row_stats,3,"Variance Min.",bg=NAVY,fg=WHITE)
        th(ws3,row_stats,4,"Équipondéré",bg=NAVY,fg=WHITE)
        for j,key in enumerate(["sharpe","minvar","equi"],2):
            r,v,s = opt_results[key]["stats"]
            td(ws3,row_stats+1,j,f"{r*100:.2f}%")
            td(ws3,row_stats+2,j,f"{v*100:.2f}%")
            td(ws3,row_stats+3,j,f"{s:.3f}")
        for r_,label in zip([row_stats+1,row_stats+2,row_stats+3],
                             ["Rdt Annualisé","Vol. Annualisée","Ratio Sharpe"]):
            td(ws3,r_,1,label)
        ws3.column_dimensions["A"].width = 18
        for col in ["B","C","D"]: ws3.column_dimensions[col].width = 18

    # Feuille 4 : Rendements
    ws4 = wb.create_sheet("📋 Données")
    r_port = rendements_df.mean(axis=1).tail(250)
    for j,h in enumerate(["Date","Rdt Portfolio (%)"],1): th(ws4,1,j,h,bg=BLUE)
    for i,(d,v) in enumerate(r_port.items(),2):
        bg = LGRAY if i%2==0 else WHITE
        td(ws4,i,1,d.strftime("%d/%m/%Y"),bg=bg)
        td(ws4,i,2,round(v*100,4),bg=bg)
    ws4.column_dimensions["A"].width = 14
    ws4.column_dimensions["B"].width = 18

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT PDF
# ══════════════════════════════════════════════════════════════════════════════

def generer_pdf(var_results: dict, bt_results: dict, metrics: dict, pv: float,
                fig_var=None, fig_dist=None) -> bytes | None:
    if not HAS_PDF: return None
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=1.5*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    NAVY_C  = HexColor("#1B2A4A"); BLUE_C = HexColor("#2E5FA3")
    GOLD_C  = HexColor("#C9A84C"); LGRAY_C= HexColor("#F0F4FA")

    S = lambda name, **kw: ParagraphStyle(name, **kw)
    s_title = S("t", fontName="Helvetica-Bold", fontSize=22, textColor=HexColor("#FFFFFF"),
                alignment=TA_CENTER, spaceAfter=4)
    s_sub   = S("s", fontName="Helvetica-Oblique", fontSize=11, textColor=GOLD_C,
                alignment=TA_CENTER, spaceAfter=6)
    s_h2    = S("h2", fontName="Helvetica-Bold", fontSize=12, textColor=NAVY_C,
                spaceBefore=12, spaceAfter=6)
    s_body  = S("b", fontName="Helvetica", fontSize=9, textColor=HexColor("#3A3A3A"),
                alignment=TA_JUSTIFY, spaceAfter=6, leading=14)

    def tbl_style():
        return TableStyle([
            ("BACKGROUND",(0,0),(-1,0),BLUE_C),("TEXTCOLOR",(0,0),(-1,0),HexColor("#FFFFFF")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8.5),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[HexColor("#FFFFFF"),LGRAY_C]),
            ("GRID",(0,0),(-1,-1),0.3,HexColor("#CCCCCC")),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ])

    story = []
    story.append(Spacer(1,2*cm))
    cover = Table([[Paragraph("RAPPORT DE GESTION DES RISQUES", s_title)]], [15*cm])
    cover.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY_C),
                                ("TOPPADDING",(0,0),(-1,-1),18),("BOTTOMPADDING",(0,0),(-1,-1),18)]))
    story.append(cover)
    story.append(Spacer(1,0.3*cm))
    cover2 = Table([[Paragraph("Value at Risk — 7 méthodes · Backtesting · Stress-Testing · Markowitz", s_sub)]],[15*cm])
    cover2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),HexColor("#243B60")),
                                  ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(cover2)
    story.append(Spacer(1,1.5*cm))

    from datetime import date
    info = [["Valeur portefeuille", f"{pv:,.0f} €"],
            ["Horizon", "1 jour ouvré"],
            ["Niveaux de confiance", "95% et 99%"],
            ["Date de production", date.today().strftime("%d/%m/%Y")]]
    t_info = Table([[Paragraph(k,s_body),Paragraph(v,s_body)] for k,v in info], [5*cm,10*cm])
    t_info.setStyle(TableStyle([("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                                  ("ROWBACKGROUNDS",(0,0),(-1,-1),[HexColor("#FFFFFF"),LGRAY_C]),
                                  ("GRID",(0,0),(-1,-1),0.3,HexColor("#DDDDDD")),
                                  ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
    story.append(t_info); story.append(PageBreak())

    story.append(Paragraph("1. RÉSULTATS DE LA VALUE AT RISK", s_h2))
    story.append(HRFlowable(width="100%",thickness=1,color=GOLD_C,spaceAfter=8))
    var_data = [["Méthode","VaR 95% (€)","VaR 95% (%)","VaR 99% (€)","VaR 99% (%)","ES 99% (€)"]]
    for m, res in var_results.items():
        r95, r99 = res.get(0.95,{}), res.get(0.99,{})
        var_data.append([m,
            f"{r95.get('VaR',0):,.0f} €", f"{r95.get('VaR_pct',0)*100:.3f}%",
            f"{r99.get('VaR',0):,.0f} €", f"{r99.get('VaR_pct',0)*100:.3f}%",
            f"{r99.get('ES',0):,.0f} €"])
    t_var = Table(var_data, [3.5*cm,2.5*cm,2*cm,2.5*cm,2*cm,2.5*cm])
    t_var.setStyle(tbl_style()); story.append(t_var); story.append(Spacer(1,0.5*cm))

    if fig_var:
        buf_img = io.BytesIO(); fig_var.savefig(buf_img, format="png", dpi=120, bbox_inches="tight"); buf_img.seek(0)
        story.append(Image(buf_img, width=15*cm, height=5*cm))
        plt.close(fig_var)
    story.append(PageBreak())

    story.append(Paragraph("2. BACKTESTING", s_h2))
    story.append(HRFlowable(width="100%",thickness=1,color=GOLD_C,spaceAfter=8))
    story.append(Paragraph(
        "Test de Kupiec (POF) : H₀ → fréquence observée = 1−α. "
        "Test de Christoffersen : teste l'indépendance temporelle des exceptions. "
        "p > 0.05 → modèle non rejeté.", s_body))
    bt_data = [["Méthode","CL","Exceptions","Taux obs.","Kupiec p","Kupiec","CC p","CC"]]
    for m, alphas in bt_results.items():
        for a, res in alphas.items():
            k = res.get("kupiec",{}); cc = res.get("cc",{})
            bt_data.append([m, f"{a*100:.0f}%", str(k.get("N","")),
                f"{k.get('rate',0)*100:.2f}%",
                f"{k.get('p_value',0):.4f}", "OUI" if k.get("valid") else "NON",
                f"{cc.get('p_value_ind',0):.4f}", "OUI" if cc.get("valid") else "NON"])
    t_bt = Table(bt_data, [3.5*cm,1.2*cm,1.5*cm,1.8*cm,1.8*cm,1.5*cm,1.8*cm,1.5*cm])
    t_bt.setStyle(tbl_style()); story.append(t_bt); story.append(Spacer(1,0.5*cm))

    if fig_dist:
        buf_img2 = io.BytesIO(); fig_dist.savefig(buf_img2, format="png", dpi=120, bbox_inches="tight"); buf_img2.seek(0)
        story.append(Image(buf_img2, width=15*cm, height=5*cm))
        plt.close(fig_dist)

    doc.build(story)
    buf.seek(0); return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

for key in ["prix","rendements","var_results","bt_results","actifs_choisis","pv","opt_results","stress_results"]:
    if key not in st.session_state:
        st.session_state[key] = None


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### 📉 VaR Analytics Suite")
    st.markdown("<div style='font-size:10px;color:#8aa0bc;font-family:JetBrains Mono,monospace;margin-bottom:16px'>v4.0 · Département Risque</div>", unsafe_allow_html=True)
    st.divider()

    menu = st.selectbox(
        "Navigation",
        ["🏠 Accueil", "🏦 Portefeuille", "📐 Optimisation", "📉 Calcul VaR",
         "🧪 Backtesting", "🔥 Stress-Testing", "📊 Reporting"],
        label_visibility="collapsed"
    )
    st.divider()

    if st.session_state["rendements"] is not None:
        r = st.session_state["rendements"].mean(axis=1)
        ann_r = r.mean() * 252
        ann_v = r.std() * np.sqrt(252)
        sharpe = (r.mean() - 0.03/252) / r.std() * np.sqrt(252)
        st.markdown("**Portefeuille chargé**")
        st.markdown(f"""
        <div style='font-size:11px;font-family:JetBrains Mono,monospace;line-height:1.9'>
        <span style='color:#8aa0bc'>Rdt ann. :</span> <span style='color:#1db87a'>+{ann_r*100:.2f}%</span><br>
        <span style='color:#8aa0bc'>Vol. ann. :</span> <span style='color:#d0daea'>{ann_v*100:.2f}%</span><br>
        <span style='color:#8aa0bc'>Sharpe   :</span> <span style='color:#C9A84C'>{sharpe:.3f}</span>
        </div>""", unsafe_allow_html=True)
        st.divider()

    st.markdown("""
    <div style='font-size:10px;color:#8aa0bc;line-height:1.8'>
    <b style='color:#C9A84C'>7 méthodes VaR</b><br>
    · Historique · VCV · RiskMetrics<br>
    · Cornish-Fisher · GARCH(1,1)<br>
    · TVE (POT) · TVE-GARCH<br><br>
    <b style='color:#C9A84C'>Nouveautés v4.0</b><br>
    · Optimisation Markowitz<br>
    · Stress-Testing (6 scénarios)<br>
    · ES analytique TVE-GARCH<br>
    · Matrice corrélation
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════

if menu == "🏠 Accueil":
    st.title("VaR Analytics Suite")
    st.markdown("""
    <div class='info-box'>
    Progiciel professionnel de calcul, comparaison et validation de la <b>Value at Risk</b>
    sur un portefeuille d'actions. Développé selon les standards <b>Bâle III/IV</b>.
    Version 4.0 : optimisation Markowitz, stress-testing et ES analytique intégrés.
    </div>""", unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("Méthodes VaR", "7", "Complètes")
    with col2: st.metric("Tests backtest", "2", "Kupiec + CC")
    with col3: st.metric("Scénarios Stress", "6", "Historiques")
    with col4: st.metric("Export", "Excel + PDF", "4 feuilles")

    st.markdown("<div class='section-header'>Fonctionnalités</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        **📥 Données de marché**
        - Téléchargement automatique Yahoo Finance
        - 15 actifs prédéfinis + simulation intégrée
        - Matrice de corrélation interactive

        **📐 Optimisation Markowitz** *(nouveau)*
        - Frontière efficiente complète
        - 3 portefeuilles : Sharpe Max, Variance Min, Équipondéré
        - Allocation recommandée exportable
        """)
    with c2:
        st.markdown("""
        **🔥 Stress-Testing** *(nouveau)*
        - 6 scénarios historiques (2008, COVID, Brexit…)
        - Comparaison VaR normale vs stressée
        - Multiplicateurs et seuils d'alerte

        **📊 Reporting amélioré**
        - Excel 4 feuilles (+ feuille Markowitz)
        - PDF exécutif avec graphiques
        - Prêt Direction Risque Bâle III/IV
        """)

    st.markdown("<div class='section-header'>Équipe Projet</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    membres = ["Anta Mbaye", "Harlem D. Adjagba", "Ecclésiaste Gnargo", "Wariol G. Kopangoye"]
    for col, m in zip(cols, membres):
        with col:
            st.markdown(f"""
            <div class='var-card' style='text-align:center;padding:14px'>
            <div style='font-size:22px;margin-bottom:6px'>👤</div>
            <div style='font-size:11px;font-weight:600;color:#d0daea'>{m}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='text-align:center;margin-top:16px;font-size:11px;color:#8aa0bc'>
    Double diplôme M2 IFIM · Ing 3 MACS — Mathématiques Appliquées au Calcul Scientifique
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : PORTEFEUILLE
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "🏦 Portefeuille":
    st.title("Construction du Portefeuille")

    st.markdown("<div class='section-header'>Sélection des actifs</div>", unsafe_allow_html=True)
    col1, col2 = st.columns([2,1])
    with col1:
        actifs_choisis = st.multiselect(
            "Actifs financiers",
            list(ACTIFS_DISPONIBLES.keys()),
            default=["Apple (AAPL)", "Microsoft (MSFT)", "LVMH (MC.PA)",
                     "TotalEnergies (TTE)", "BNP Paribas (BNP)"],
        )
    with col2:
        pv_millions = st.number_input("Valeur du portefeuille (M€)", min_value=0.1, max_value=1000.0, value=10.0, step=0.5)
        pv = pv_millions * 1_000_000

    col3, col4, col5 = st.columns(3)
    with col3: date_debut = st.date_input("Date de début", value=pd.to_datetime("2019-01-01"))
    with col4: date_fin   = st.date_input("Date de fin",   value=pd.to_datetime("today"))
    with col5:
        source = st.radio("Source données", ["Yahoo Finance", "Simulation"], horizontal=True)

    btn_col, _ = st.columns([1,3])
    with btn_col:
        btn = st.button("▶  Charger les données", type="primary", use_container_width=True)

    if btn:
        if len(actifs_choisis) < 2:
            st.warning("Sélectionnez au moins 2 actifs.")
        elif date_debut >= date_fin:
            st.warning("La date de début doit être antérieure à la date de fin.")
        else:
            tickers = [ACTIFS_DISPONIBLES[a] for a in actifs_choisis]
            with st.spinner("Chargement des données en cours…"):
                if source == "Yahoo Finance" and HAS_YF:
                    prix = telecharger_donnees(tickers, date_debut, date_fin)
                    if prix.empty:
                        st.warning("Aucune donnée récupérée, passage en simulation.")
                        prix = donnees_simulation(tickers)
                else:
                    prix = donnees_simulation(tickers)
                rendements = prix.pct_change(fill_method=None).dropna(how="all")
                rendements = rendements.dropna(axis=1, how="any")
                prix = prix[rendements.columns]
                tickers_valides = list(rendements.columns)

            st.session_state.update({
                "prix": prix, "rendements": rendements,
                "actifs_choisis": tickers_valides, "pv": pv,
                "var_results": None, "bt_results": None,
                "opt_results": None, "stress_results": None
            })
            st.success(f"✅ {len(tickers_valides)} actif(s) chargés — {len(rendements)} jours de données.")

    if st.session_state["rendements"] is not None:
        rendements = st.session_state["rendements"]
        prix       = st.session_state["prix"]

        st.markdown("<div class='section-header'>Statistiques du portefeuille</div>", unsafe_allow_html=True)
        port_r = rendements.mean(axis=1)
        ann_r  = port_r.mean() * 252
        ann_v  = port_r.std()  * np.sqrt(252)
        sharpe = (port_r.mean() - 0.03/252) / port_r.std() * np.sqrt(252)
        skew   = float(stats.skew(port_r))
        kurt   = float(stats.kurtosis(port_r))
        mdd    = float(((1+port_r).cumprod() / (1+port_r).cumprod().cummax() - 1).min())

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Rdt Annualisé",  f"+{ann_r*100:.2f}%")
        c2.metric("Vol. Annuelle",   f"{ann_v*100:.2f}%")
        c3.metric("Sharpe",          f"{sharpe:.3f}")
        c4.metric("Max Drawdown",    f"{mdd*100:.2f}%")
        c5.metric("Skewness",        f"{skew:.4f}")
        c6.metric("Kurtosis (exc)",  f"{kurt:.4f}")

        st.markdown("<div class='section-header'>Performance & Rendements</div>", unsafe_allow_html=True)
        fig = fig_perf(rendements, list(rendements.columns))
        st.pyplot(fig, use_container_width=True); plt.close(fig)

        st.markdown("<div class='section-header'>Matrice de Corrélation</div>", unsafe_allow_html=True)
        fig_corr = fig_correlation(rendements)
        st.pyplot(fig_corr, use_container_width=True); plt.close(fig_corr)

        st.markdown("<div class='section-header'>Statistiques individuelles</div>", unsafe_allow_html=True)
        stats_df = pd.DataFrame({
            "Ticker":       rendements.columns,
            "Secteur":      [SECTEURS.get(t,"—") for t in rendements.columns],
            "Rdt moy. (%)": (rendements.mean()*252*100).round(2),
            "Vol. ann. (%)": (rendements.std()*np.sqrt(252)*100).round(2),
            "Skewness":     rendements.apply(lambda c: round(stats.skew(c),4)),
            "Kurtosis":     rendements.apply(lambda c: round(stats.kurtosis(c),4)),
            "Min (%)":      (rendements.min()*100).round(3),
            "Max (%)":      (rendements.max()*100).round(3),
        }).set_index("Ticker")
        st.dataframe(stats_df, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : OPTIMISATION MARKOWITZ  (NOUVEAU)
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "📐 Optimisation":
    st.title("Optimisation de Portefeuille — Markowitz")

    if st.session_state["rendements"] is None:
        st.info("💡 Commencez par charger un portefeuille dans la page **Portefeuille**.")
        st.stop()

    rendements = st.session_state["rendements"]

    st.markdown("""
    <div class='info-box'>
    <b style='color:#C9A84C'>Théorie Moderne du Portefeuille (Markowitz, 1952)</b> — La frontière
    efficiente représente l'ensemble des portefeuilles offrant le <b>meilleur rendement pour
    un niveau de risque donné</b>. Trois allocations optimales sont proposées : maximisation
    du ratio de Sharpe, minimisation de la variance, et équipondération.
    </div>""", unsafe_allow_html=True)

    col_rf, col_btn, _ = st.columns([1,1,2])
    with col_rf:
        rf = st.number_input("Taux sans risque (%/an)", min_value=0.0, max_value=10.0,
                              value=3.0, step=0.1) / 100
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        opt_btn = st.button("▶  Optimiser le portefeuille", type="primary", use_container_width=True)

    if opt_btn:
        with st.spinner("Calcul de la frontière efficiente…"):
            opt_results = optimiser_portefeuille(rendements, rf)
            st.session_state["opt_results"] = opt_results
        st.success("✅ Optimisation terminée.")

    if st.session_state["opt_results"]:
        opt = st.session_state["opt_results"]

        st.markdown("<div class='section-header'>Frontière Efficiente</div>", unsafe_allow_html=True)
        fig_fe = fig_frontiere_efficiente(opt)
        st.pyplot(fig_fe, use_container_width=True); plt.close(fig_fe)

        st.markdown("<div class='section-header'>Allocations optimales</div>", unsafe_allow_html=True)
        cols = st.columns(3)
        styles_ = [("sharpe","#C9A84C"), ("minvar","#1db87a"), ("equi","#a855f7")]
        for col, (key, color) in zip(cols, styles_):
            p = opt[key]
            r, v, s = p["stats"]
            with col:
                st.markdown(f"""
                <div class='markowitz-card'>
                  <div style='font-size:11px;color:{color};font-family:JetBrains Mono,monospace;
                       text-transform:uppercase;letter-spacing:1px;margin-bottom:8px'>{p['label']}</div>
                  <div style='font-size:13px;color:#d0daea;line-height:2;font-family:JetBrains Mono,monospace'>
                  📈 Rdt ann. : <b style='color:{color}'>{r*100:.2f}%</b><br>
                  📊 Vol. ann. : <b style='color:#d0daea'>{v*100:.2f}%</b><br>
                  ⭐ Sharpe   : <b style='color:{color}'>{s:.3f}</b>
                  </div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Répartition des poids</div>", unsafe_allow_html=True)
        fig_w = fig_poids_portefeuilles(opt)
        st.pyplot(fig_w, use_container_width=True); plt.close(fig_w)

        st.markdown("<div class='section-header'>Tableau des poids détaillés</div>", unsafe_allow_html=True)
        df_poids = pd.DataFrame({
            "Actif": opt["tickers"],
            "Sharpe Max. (%)":   [f"{w*100:.1f}%" for w in opt["sharpe"]["weights"]],
            "Variance Min. (%)": [f"{w*100:.1f}%" for w in opt["minvar"]["weights"]],
            "Équipondéré (%)":   [f"{w*100:.1f}%" for w in opt["equi"]["weights"]],
        }).set_index("Actif")
        st.dataframe(df_poids, use_container_width=True)

        with st.expander("ℹ️ Hypothèses & Limites du modèle"):
            st.markdown("""
            - **Positions longues uniquement** (contrainte w ≥ 0)
            - **Rendements supposés normaux** — en pratique, les rendements présentent des queues épaisses
            - **Paramètres historiques** — la frontière est sensible à la période d'estimation
            - **Pas de coûts de transaction** — les poids théoriques ignorent les frais de rééquilibrage
            - **Recommandation** : combiner l'allocation Sharpe Max avec la VaR TVE-GARCH pour un pilotage complet du risque
            """)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : VAR
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "📉 Calcul VaR":
    st.title("Calcul de la Value at Risk")

    if st.session_state["rendements"] is None:
        st.info("💡 Commencez par charger un portefeuille dans la page **Portefeuille**.")
        st.stop()

    rendements = st.session_state["rendements"]
    pv         = st.session_state["pv"] or 10_000_000
    port_r     = rendements.mean(axis=1).values

    st.markdown("<div class='section-header'>Paramètres de calcul</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: horizon = st.slider("Horizon (jours)", 1, 10, 1)
    with c2:
        conf_options = st.multiselect("Niveaux de confiance",
                                       [0.90, 0.95, 0.975, 0.99],
                                       default=[0.95, 0.99],
                                       format_func=lambda x: f"{x*100:.1f}%")
    with c3:
        methodes_sel = st.multiselect("Méthodes à calculer",
                                       ["Historique","Variance-Covariance","RiskMetrics",
                                        "Cornish-Fisher","GARCH(1,1)","TVE (POT)","TVE-GARCH"],
                                       default=["Historique","Variance-Covariance",
                                                "RiskMetrics","Cornish-Fisher",
                                                "GARCH(1,1)","TVE (POT)","TVE-GARCH"])

    btn2, _ = st.columns([1,3])
    with btn2:
        calc_btn = st.button("▶  Calculer les 7 VaR", type="primary", use_container_width=True)

    if calc_btn:
        if not conf_options:
            st.warning("Sélectionnez au moins un niveau de confiance.")
        else:
            with st.spinner("Calcul en cours…"):
                engine = VaREngine(port_r, pv, horizon)
                var_results = engine.compute_all(tuple(sorted(conf_options)))
                var_results = {k:v for k,v in var_results.items() if k in methodes_sel}
                st.session_state["var_results"] = var_results
            st.success(f"✅ VaR calculée — {len(var_results)} méthodes × {len(conf_options)} niveaux.")

    if st.session_state["var_results"]:
        var_results = st.session_state["var_results"]
        alphas_used = sorted(list(list(var_results.values())[0].keys()))
        alpha_display = st.select_slider("Afficher pour :",
                                          options=alphas_used,
                                          format_func=lambda x: f"{x*100:.0f}%",
                                          value=alphas_used[-1])

        st.markdown("<div class='section-header'>Résultats par méthode</div>", unsafe_allow_html=True)
        METHODE_RECOMMANDEE = "TVE-GARCH"
        cols = st.columns(min(len(var_results), 4))
        for i, (method, res) in enumerate(var_results.items()):
            r = res.get(alpha_display, {})
            var_val = r.get("VaR", np.nan)
            pct_val = r.get("VaR_pct", np.nan)
            es_val  = r.get("ES",  np.nan)
            rec = f'<span class="badge-rec">★ Recommandé</span>' if method == METHODE_RECOMMANDEE else ""
            with cols[i % len(cols)]:
                st.markdown(f"""
                <div class='var-card'>
                  <div class='var-card-title'>{method}{rec}</div>
                  <div class='var-card-value'>{var_val/1000:.1f} k€</div>
                  <div class='var-card-pct'>VaR {alpha_display*100:.0f}% · {pct_val*100:.3f}%</div>
                  <div class='var-card-es'>ES : {es_val/1000:.1f} k€</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Tableau comparatif</div>", unsafe_allow_html=True)
        rows = []
        for m, res in var_results.items():
            row = {"Méthode": m}
            for a in alphas_used:
                r = res.get(a, {})
                row[f"VaR {a*100:.0f}% (€)"]  = f"{r.get('VaR',0):,.0f}"
                row[f"VaR {a*100:.0f}% (%)"]  = f"{r.get('VaR_pct',0)*100:.3f}%"
                row[f"ES {a*100:.0f}% (€)"]   = f"{r.get('ES',0):,.0f}"
            row["Paramètres"] = list(var_results[m].values())[0].get("params","")
            rows.append(row)
        df_var = pd.DataFrame(rows).set_index("Méthode")
        st.dataframe(df_var, use_container_width=True)

        st.markdown("<div class='section-header'>Graphique comparatif</div>", unsafe_allow_html=True)
        fig_v = fig_var_comparaison(var_results, alpha_display, pv)
        st.pyplot(fig_v, use_container_width=True); plt.close(fig_v)

        st.markdown("<div class='section-header'>Distribution des rendements</div>", unsafe_allow_html=True)
        fig_d = fig_distribution(port_r, var_results)
        st.pyplot(fig_d, use_container_width=True); plt.close(fig_d)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : BACKTESTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "🧪 Backtesting":
    st.title("Backtesting des modèles de VaR")

    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR dans la page **Calcul VaR**.")
        st.stop()

    rendements   = st.session_state["rendements"]
    var_results  = st.session_state["var_results"]
    port_r       = rendements.mean(axis=1).values
    alphas_used  = sorted(list(list(var_results.values())[0].keys()))

    st.markdown("""
    <div class='info-box'>
    <b style='color:#C9A84C'>Test de Kupiec (POF)</b> — Vérifie si la fréquence d'exceptions
    est statistiquement conforme au niveau de confiance déclaré. <b>LR ~ χ²(1).</b><br><br>
    <b style='color:#C9A84C'>Test de Christoffersen (CC)</b> — Teste l'indépendance temporelle
    des exceptions. Un clustering signale un modèle insensible aux chocs.<br><br>
    <span style='color:#1db87a'>✅ p-value &gt; 5%</span> → modèle validé &nbsp;
    <span style='color:#e05252'>❌ p-value ≤ 5%</span> → modèle rejeté
    </div>""", unsafe_allow_html=True)

    btn3, _ = st.columns([1,3])
    with btn3:
        bt_btn = st.button("▶  Lancer le backtesting", type="primary", use_container_width=True)

    if bt_btn:
        with st.spinner("Backtesting en cours…"):
            bt_results = {}
            for method, res in var_results.items():
                bt_results[method] = {}
                for a in alphas_used:
                    var_pct = res.get(a, {}).get("VaR_pct", np.nan)
                    if np.isnan(var_pct): continue
                    k  = kupiec_test(port_r, var_pct, a)
                    cc = christoffersen_test(port_r, var_pct)
                    bt_results[method][a] = {"kupiec": k, "cc": cc}
            st.session_state["bt_results"] = bt_results
        st.success("✅ Backtesting terminé.")

    if st.session_state["bt_results"]:
        bt_results = st.session_state["bt_results"]
        rows = []
        for m, alphas in bt_results.items():
            for a, res in alphas.items():
                k, cc = res["kupiec"], res["cc"]
                rows.append({
                    "Méthode": m, "CL": f"{a*100:.0f}%",
                    "Exceptions": k["N"], "T": k["T"],
                    "Taux obs.": f"{k['rate']*100:.2f}%",
                    "Taux att.": f"{(1-a)*100:.2f}%",
                    "Kupiec LR": round(k["LR"],3),
                    "Kupiec p":  round(k["p_value"],4),
                    "Kupiec ✓":  "✅ OK" if k["valid"] else "❌",
                    "CC p":      round(cc.get("p_value_ind",0),4),
                    "CC ✓":      "✅ OK" if cc.get("valid") else "❌",
                })
        st.dataframe(pd.DataFrame(rows).set_index("Méthode"), use_container_width=True)

        st.markdown("<div class='section-header'>Graphique p-values (Kupiec)</div>", unsafe_allow_html=True)
        fig_bt = fig_backtesting({m: {a: {"p_value": bt_results[m][a]["kupiec"]["p_value"]}
                                       for a in bt_results[m]} for m in bt_results})
        st.pyplot(fig_bt, use_container_width=True); plt.close(fig_bt)

        st.markdown("<div class='section-header'>Verdict synthétique</div>", unsafe_allow_html=True)
        alpha_v = alphas_used[-1]
        cols_v = st.columns(len(bt_results))
        for col, (m, alphas) in zip(cols_v, bt_results.items()):
            res = alphas.get(alpha_v, {})
            k, cc = res.get("kupiec",{}), res.get("cc",{})
            k_ok, cc_ok = k.get("valid",False), cc.get("valid",False)
            score = "✅ Validé" if k_ok and cc_ok else ("⚠️ Partiel" if k_ok or cc_ok else "❌ Rejeté")
            color = "#1db87a" if k_ok and cc_ok else ("#C9A84C" if k_ok or cc_ok else "#e05252")
            with col:
                st.markdown(f"""
                <div class='var-card' style='text-align:center'>
                  <div class='var-card-title'>{m}</div>
                  <div style='font-size:15px;font-weight:700;color:{color};margin:6px 0'>{score}</div>
                  <div style='font-size:10px;color:#8aa0bc;font-family:JetBrains Mono,monospace'>
                  Kupiec p={k.get('p_value',0):.4f}<br>CC p={cc.get('p_value_ind',0):.4f}
                  </div>
                </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : STRESS-TESTING  (NOUVEAU)
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "🔥 Stress-Testing":
    st.title("Stress-Testing — Scénarios Historiques")

    if st.session_state["rendements"] is None:
        st.info("💡 Commencez par charger un portefeuille dans la page **Portefeuille**.")
        st.stop()

    rendements  = st.session_state["rendements"]
    var_results = st.session_state["var_results"]
    pv          = st.session_state["pv"] or 10_000_000
    port_r      = rendements.mean(axis=1).values

    if var_results is None:
        st.warning("⚠️ Les VaR ne sont pas encore calculées. Les comparaisons seront limitées.")
        var_results = {}

    st.markdown("""
    <div class='info-box'>
    Le <b>stress-testing</b> mesure l'impact de scénarios de marché extrêmes sur le portefeuille.
    Contrairement à la VaR (probabiliste), les stress-tests appliquent des <b>chocs déterministes</b>
    calibrés sur des crises historiques réelles. Exigé par <b>Bâle III/IV</b> et les superviseurs (EBA, BCE).
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-header'>Sélection des scénarios</div>", unsafe_allow_html=True)
    scenarios_choisis = st.multiselect(
        "Scénarios à analyser",
        list(SCENARIOS_STRESS.keys()),
        default=list(SCENARIOS_STRESS.keys()),
    )

    c_alpha, c_btn, _ = st.columns([1,1,2])
    with c_alpha:
        alpha_stress = st.selectbox("Niveau VaR de référence", [0.95, 0.99],
                                     format_func=lambda x: f"{x*100:.0f}%", index=1)
    with c_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        stress_btn = st.button("▶  Lancer le stress-test", type="primary", use_container_width=True)

    if stress_btn and scenarios_choisis:
        with st.spinner("Calcul des scénarios de stress…"):
            stress_results = {}
            for sc_name in scenarios_choisis:
                sc = SCENARIOS_STRESS[sc_name]
                stress_results[sc_name] = appliquer_stress(
                    port_r, pv, sc, var_results, alpha=alpha_stress
                )
            st.session_state["stress_results"] = stress_results
        st.success(f"✅ {len(stress_results)} scénario(s) calculé(s).")

    if st.session_state["stress_results"]:
        stress_results = st.session_state["stress_results"]

        st.markdown("<div class='section-header'>Résultats par scénario</div>", unsafe_allow_html=True)
        cols_per_row = 3
        sc_list = list(stress_results.items())
        for row_start in range(0, len(sc_list), cols_per_row):
            cols = st.columns(cols_per_row)
            for col, (sc_name, sr) in zip(cols, sc_list[row_start:row_start+cols_per_row]):
                sc_info = SCENARIOS_STRESS.get(sc_name, {})
                ratio = sr["ratio"]
                ratio_txt = f"×{ratio:.2f}" if not np.isnan(ratio) else "—"
                alert = "🔴" if (not np.isnan(ratio) and ratio > 2) else ("🟡" if (not np.isnan(ratio) and ratio > 1.5) else "🟢")
                with col:
                    st.markdown(f"""
                    <div class='stress-card'>
                      <div class='stress-title'>{alert} {sc_name.split("(")[0].strip()}</div>
                      <div style='font-size:10px;color:#8aa0bc;margin-bottom:8px'>{sc_info.get('date','')}</div>
                      <div class='stress-val'>{sr['pnl_stress']/1000:+.0f} k€</div>
                      <div style='font-size:11px;color:#8aa0bc;margin-top:4px;font-family:JetBrains Mono,monospace'>
                        P&L stressé<br>
                        VaR stressée : <b style='color:#e87373'>{sr['var_stress']/1000:.0f} k€</b><br>
                        Ratio : <b style='color:#C9A84C'>{ratio_txt}</b> vs VaR normale<br>
                        Choc marché : <b style='color:#e05252'>{sr['choc']*100:.2f}%</b>
                      </div>
                    </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Analyse comparative</div>", unsafe_allow_html=True)
        fig_st = fig_stress_comparaison(stress_results, pv)
        st.pyplot(fig_st, use_container_width=True); plt.close(fig_st)

        st.markdown("<div class='section-header'>Tableau de synthèse</div>", unsafe_allow_html=True)
        rows_st = []
        for sc_name, sr in stress_results.items():
            sc_info = SCENARIOS_STRESS.get(sc_name, {})
            ratio = sr["ratio"]
            alert = "🔴 ALERTE" if (not np.isnan(ratio) and ratio > 2) else (
                    "🟡 VIGILANCE" if (not np.isnan(ratio) and ratio > 1.5) else "🟢 OK")
            rows_st.append({
                "Scénario":          sc_name.split("(")[0].strip(),
                "Date":              sc_info.get("date",""),
                "Choc marché (%)":   f"{sr['choc']*100:.2f}%",
                "P&L stressé (k€)":  f"{sr['pnl_stress']/1000:+.1f}",
                "VaR stressée (k€)": f"{sr['var_stress']/1000:.1f}",
                "VaR normale (k€)":  f"{sr['var_normal']/1000:.1f}",
                "Ratio ×":           f"{ratio:.2f}" if not np.isnan(ratio) else "—",
                "Statut":            alert,
            })
        st.dataframe(pd.DataFrame(rows_st).set_index("Scénario"), use_container_width=True)

        with st.expander("ℹ️ Méthodologie des stress-tests"):
            st.markdown("""
            **Choc de marché** : rendement journalier appliqué au portefeuille équipondéré.

            **VaR stressée** : recalculée avec la volatilité multipliée par le facteur de choc du scénario,
            selon la méthode Variance-Covariance : VaR_stress = z_{α} × σ_stress × PV

            **Ratio** : VaR_stressée / VaR_normale → mesure l'amplification du risque en période de crise.

            **Seuils d'alerte** :
            - 🟢 Ratio < 1.5 → Impact limité
            - 🟡 1.5 ≤ Ratio < 2.0 → Vigilance renforcée
            - 🔴 Ratio ≥ 2.0 → Alerte — révision des limites recommandée
            """)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE : REPORTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "📊 Reporting":
    st.title("Génération des Rapports")

    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR dans la page **Calcul VaR**.")
        st.stop()

    rendements  = st.session_state["rendements"]
    var_results = st.session_state["var_results"]
    bt_results  = st.session_state["bt_results"] or {}
    pv          = st.session_state["pv"] or 10_000_000
    opt_results = st.session_state.get("opt_results")
    port_r      = rendements.mean(axis=1).values
    alphas_used = sorted(list(list(var_results.values())[0].keys()))

    st.markdown("""
    <div class='info-box'>
    Générez automatiquement un <b>rapport Excel</b> de suivi opérationnel (4 feuilles)
    et un <b>rapport PDF</b> exécutif avec graphiques. Conformes aux exigences Bâle III/IV.
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='section-header'>Aperçu des résultats</div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    alpha_99 = max(alphas_used)
    best = "TVE-GARCH" if "TVE-GARCH" in var_results else list(var_results.keys())[-1]
    var_best = var_results[best][alpha_99]["VaR"]
    es_best  = var_results[best][alpha_99]["ES"]
    n_ok = sum(1 for m in bt_results for a in bt_results[m]
               if bt_results[m][a]["kupiec"]["valid"]) if bt_results else 0
    n_total = sum(len(bt_results[m]) for m in bt_results) if bt_results else 0

    c1.metric("Méthode recommandée", best)
    c2.metric(f"VaR 99% ({best})", f"{var_best/1000:.1f} k€")
    c3.metric(f"ES 99% ({best})",  f"{es_best/1000:.1f} k€")
    c4.metric("Backtests validés",  f"{n_ok}/{n_total}" if n_total else "—")

    st.markdown("<div class='section-header'>Téléchargements</div>", unsafe_allow_html=True)
    col_xlsx, col_pdf = st.columns(2)

    bt_fmt = {}
    for m, alphas in bt_results.items():
        bt_fmt[m] = {}
        for a, res in alphas.items():
            bt_fmt[m][a] = res

    with col_xlsx:
        st.markdown("""
        <div class='var-card'>
          <div class='var-card-title'>📊 Rapport Excel</div>
          <div style='font-size:12px;color:#d0daea;margin:8px 0;line-height:1.6'>
          4 feuilles : Résumé VaR · Backtesting · Markowitz · Données<br>
          Formatage professionnel, tableaux structurés, prêt pour la Direction.
          </div>
        </div>""", unsafe_allow_html=True)
        if HAS_XLSX:
            with st.spinner("Génération Excel…"):
                xlsx_bytes = generer_excel(rendements, var_results, bt_fmt, pv, opt_results)
            if xlsx_bytes:
                st.download_button("⬇  Télécharger le rapport Excel",
                                   data=xlsx_bytes, file_name="VaR_Report_v4.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            st.warning("openpyxl non installé : pip install openpyxl")

    with col_pdf:
        st.markdown("""
        <div class='var-card'>
          <div class='var-card-title'>📄 Rapport PDF</div>
          <div style='font-size:12px;color:#d0daea;margin:8px 0;line-height:1.6'>
          Couverture · VaR · Backtesting · Graphiques intégrés<br>
          Mise en page exécutive, prêt à imprimer / envoyer.
          </div>
        </div>""", unsafe_allow_html=True)
        if HAS_PDF:
            with st.spinner("Génération PDF…"):
                fv  = fig_var_comparaison(var_results, alpha_99, pv)
                fd  = fig_distribution(port_r, var_results)
                pdf_bytes = generer_pdf(var_results, bt_fmt, {}, pv, fv, fd)
            if pdf_bytes:
                st.download_button("⬇  Télécharger le rapport PDF",
                                   data=pdf_bytes, file_name="VaR_Risk_Report_v4.pdf",
                                   mime="application/pdf",
                                   use_container_width=True)
        else:
            st.warning("reportlab non installé : pip install reportlab")

    with st.expander("📦 Déploiement Streamlit Cloud"):
        st.markdown("**`requirements.txt`** à placer à la racine du dépôt GitHub :")
        st.code("""streamlit>=1.32
yfinance>=0.2.36
pandas>=2.0
numpy>=1.26
scipy>=1.12
matplotlib>=3.8
openpyxl>=3.1
reportlab>=4.1""", language="text")
        st.markdown("""
        **Étapes :**
        1. Pousser `app.py` + `requirements.txt` sur GitHub
        2. Aller sur [share.streamlit.io](https://share.streamlit.io)
        3. **New app** → sélectionner le repo → `app.py` → **Deploy**
        """)
