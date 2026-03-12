"""
VaR ANALYTICS SUITE — v4.2
Lancement : streamlit run app.py
"""

# ── Imports ───────────────────────────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
from matplotlib.colors import LinearSegmentedColormap
from typing import Optional, Tuple
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
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak, Image)
    from reportlab.lib.colors import HexColor
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="VaR Analytics Suite",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');

/* ── FOND SOMBRE ─────────────────────────────────────── */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background-color: #07090F !important;
}

/* ── TEXTE GLOBAL — couleur claire sur fond sombre ───── */
/* On cible le contenu principal uniquement, pas les widgets */
[data-testid="stMain"],
[data-testid="stMainBlockContainer"] {
    color: #E8EDF5 !important;
}

/* ── SIDEBAR ──────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background-color: #0C0F1A !important;
    border-right: 1px solid rgba(0,212,255,0.15) !important;
}
/* Textes dans la sidebar */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p {
    color: #7A8BA8 !important;
}
/* Bouton ouverture sidebar */
[data-testid="stSidebarCollapsedControl"] {
    display: flex !important;
    opacity: 1 !important;
    visibility: visible !important;
    background: #111827 !important;
    border: 1px solid rgba(0,212,255,0.3) !important;
    border-left: none !important;
    border-radius: 0 6px 6px 0 !important;
}
[data-testid="stSidebarCollapsedControl"] svg { color: #00D4FF !important; }

/* ── TITRES ───────────────────────────────────────────── */
/* h1 : couleur simple et fiable (pas de gradient — trop risqué) */
h1 { color: #00D4FF !important; font-weight: 800 !important; font-size: 1.9rem !important; }
h2 { color: #E8EDF5 !important; font-weight: 600 !important; }
h3 { color: #F0B429 !important; font-weight: 600 !important; }

/* Textes markdown natifs */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong,
[data-testid="stMarkdownContainer"] em {
    color: #E8EDF5 !important;
}

/* Metrics */
[data-testid="metric-container"] {
    background: rgba(17,24,37,0.9) !important;
    border: 1px solid rgba(0,212,255,0.2) !important;
    border-radius: 10px !important;
    padding: 14px !important;
}
[data-testid="stMetricLabel"] {
    color: #7A8BA8 !important;
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 1px !important;
    font-family: 'DM Mono', monospace !important;
}
[data-testid="stMetricValue"] {
    color: #00D4FF !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 1.5rem !important;
}

/* Boutons */
.stButton > button {
    background: #1a2235 !important;
    color: #E8EDF5 !important;
    border: 1px solid rgba(0,212,255,0.25) !important;
    border-radius: 6px !important;
    font-family: 'Outfit', sans-serif !important;
    font-weight: 500 !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    border-color: #00D4FF !important;
    color: #00D4FF !important;
}
.stButton > button[kind="primary"] {
    background: #00D4FF !important;
    color: #07090F !important;
    border: none !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"]:hover {
    background: #00aacc !important;
    color: #07090F !important;
}

/* Inputs */
.stSelectbox > div > div,
.stMultiSelect > div > div,
.stNumberInput input,
.stDateInput input {
    background: #1a2235 !important;
    border: 1px solid rgba(0,212,255,0.2) !important;
    border-radius: 6px !important;
    color: #E8EDF5 !important;
}
[data-testid="stRadio"] label {
    background: #1a2235 !important;
    border: 1px solid rgba(0,212,255,0.15) !important;
    border-radius: 6px !important;
    padding: 5px 12px !important;
    color: #7A8BA8 !important;
}

/* DataFrames */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(0,212,255,0.15) !important;
    border-radius: 8px !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: #111827 !important;
    border: 1px solid rgba(0,212,255,0.1) !important;
    border-radius: 6px !important;
    color: #7A8BA8 !important;
}

/* Alerts */
.stAlert { border-radius: 6px !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #07090F; }
::-webkit-scrollbar-thumb { background: rgba(0,212,255,0.2); border-radius: 10px; }

/* Cacher éléments inutiles */
#MainMenu, footer { display: none !important; }

/* Composants custom */
.var-card {
    background: rgba(17,24,37,0.9);
    border: 1px solid rgba(0,212,255,0.18);
    border-radius: 10px;
    padding: 16px 18px;
    margin-bottom: 12px;
}
.var-card-title {
    font-size: 10px; color: #3D4F6B;
    text-transform: uppercase; letter-spacing: 1.5px;
    font-family: 'DM Mono', monospace; margin-bottom: 6px;
}
.var-card-value {
    font-size: 24px; color: #FF4D6D;
    font-family: 'DM Mono', monospace; font-weight: 300;
}
.var-card-pct  { font-size: 11px; color: #7A8BA8; font-family: 'DM Mono', monospace; margin-top: 3px; }
.var-card-es   { font-size: 12px; color: rgba(255,77,109,0.6); font-family: 'DM Mono', monospace; margin-top: 5px; }
.badge-rec {
    background: #00D4FF; color: #07090F;
    font-size: 8px; font-weight: 700; padding: 2px 7px;
    border-radius: 20px; letter-spacing: 0.8px; margin-left: 8px;
    font-family: 'DM Mono', monospace;
}
.section-sep {
    border-left: 3px solid #00D4FF;
    padding-left: 10px; margin: 22px 0 12px;
    font-size: 12px; font-weight: 600;
    color: #7A8BA8; text-transform: uppercase;
    letter-spacing: 1.2px; font-family: 'DM Mono', monospace;
}
.info-box {
    background: rgba(0,212,255,0.05);
    border-left: 3px solid #00D4FF;
    border-radius: 0 6px 6px 0;
    padding: 12px 16px; margin: 12px 0;
    font-size: 13px; color: #7A8BA8; line-height: 1.7;
}
.stress-card {
    background: rgba(255,77,109,0.06);
    border: 1px solid rgba(255,77,109,0.2);
    border-radius: 10px; padding: 14px 16px; margin-bottom: 10px;
}
.mko-card {
    background: rgba(0,200,150,0.05);
    border: 1px solid rgba(0,200,150,0.2);
    border-radius: 10px; padding: 14px 16px; margin-bottom: 10px;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ACTIFS = {
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
    "AAPL":"Tech","MSFT":"Tech","MC.PA":"Luxe","TTE.PA":"Energie",
    "BNP.PA":"Finance","NESN.SW":"Conso","SAP":"Tech","AIR.PA":"Aero",
    "TSLA":"Auto","AMZN":"Commerce","NVDA":"Tech","SAF.PA":"Aero",
    "OR.PA":"Beauté","ASML.AS":"Tech","RMS.PA":"Luxe",
}

STRESS_SCENARIOS = {
    "Lehman 2008":   {"choc": -0.0469, "vol_mult": 3.8, "date": "15/09/2008"},
    "Flash Crash 2010": {"choc": -0.0340, "vol_mult": 2.5, "date": "06/05/2010"},
    "Brexit 2016":   {"choc": -0.0281, "vol_mult": 2.1, "date": "24/06/2016"},
    "COVID 2020":    {"choc": -0.0598, "vol_mult": 4.2, "date": "16/03/2020"},
    "Krach Taux 2022": {"choc": -0.0312, "vol_mult": 2.0, "date": "T1 2022"},
    "Choc −3σ":      {"choc": None,    "vol_mult": 3.0, "date": "Hypothétique"},
}

PLT_DARK = {
    "figure.facecolor": "#07090F",
    "axes.facecolor":   "#0C0F1A",
    "axes.edgecolor":   "#1a2235",
    "axes.labelcolor":  "#7A8BA8",
    "xtick.color":      "#7A8BA8",
    "ytick.color":      "#7A8BA8",
    "grid.color":       "#111827",
    "text.color":       "#E8EDF5",
    "legend.facecolor": "#111827",
    "legend.edgecolor": "#1a2235",
}

PALETTE = ["#00D4FF","#F0B429","#00C896","#FF4D6D","#9D7FEA",
           "#FF7849","#4CC9F0","#F72585","#06D6A0","#FFD166",
           "#B5E48C","#C77DFF","#00B4D8","#FFAA5B","#FF9B54"]


def sep(label: str):
    st.markdown(f"<div class='section-sep'>{label}</div>", unsafe_allow_html=True)


def info(text: str):
    st.markdown(f"<div class='info-box'>{text}</div>", unsafe_allow_html=True)


def spine(ax):
    for s in ax.spines.values():
        s.set_edgecolor("#1a2235")


# ══════════════════════════════════════════════════════════════════════════════
# DONNÉES
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def telecharger(tickers: list, debut, fin) -> pd.DataFrame:
    if not HAS_YF:
        return pd.DataFrame()
    data = yf.download(tickers, start=str(debut), end=str(fin),
                       auto_adjust=True, progress=False)
    if isinstance(data.columns, pd.MultiIndex):
        prix = data["Close"].copy() if "Close" in data.columns.get_level_values(0) \
               else data.xs(data.columns.get_level_values(0)[0], axis=1, level=0)
    else:
        prix = data[["Close"]].copy() if "Close" in data.columns else data.copy()
        if len(tickers) == 1 and isinstance(prix, pd.DataFrame):
            prix.columns = tickers
    return (prix if not isinstance(prix, pd.Series) else prix.to_frame()).dropna(axis=1, how="all")


def simuler(tickers: list, n: int = 1500) -> pd.DataFrame:
    np.random.seed(42)
    k    = len(tickers)
    corr = np.eye(k) + 0.4 * (np.ones((k, k)) - np.eye(k))
    L    = np.linalg.cholesky(corr)
    z    = np.random.randn(n, k) @ L.T
    lr   = 0.0004 + 0.012 * z
    lr[n // 2: n // 2 + 20] *= 3.5
    prices = 100 * np.exp(np.cumsum(lr, axis=0))
    return pd.DataFrame(prices, index=pd.bdate_range(end="2024-12-31", periods=n), columns=tickers)


# ══════════════════════════════════════════════════════════════════════════════
# MOTEUR VaR
# ══════════════════════════════════════════════════════════════════════════════

class VaREngine:
    def __init__(self, r: np.ndarray, pv: float = 10_000_000, h: int = 1):
        self.r = r; self.pv = pv; self.h = h

    def _scale(self, v): return v * np.sqrt(self.h)

    def historique(self, a):
        q  = np.percentile(self.r, (1 - a) * 100)
        es = self.r[self.r <= q].mean() if (self.r <= q).any() else q
        return {"VaR": -q*self.pv*self._scale(1), "ES": -es*self.pv*self._scale(1),
                "VaR_pct": float(-q), "params": "Quantile empirique"}

    def vcv(self, a):
        mu, sig = self.r.mean(), self.r.std()
        z = stats.norm.ppf(1 - a)
        v = -(mu + z * sig) * self._scale(1)
        e = (sig * stats.norm.pdf(z) / (1 - a) - mu) * self._scale(1)
        return {"VaR": v*self.pv, "ES": e*self.pv, "VaR_pct": float(v),
                "params": f"μ={mu*100:.4f}% σ={sig*100:.4f}%"}

    def riskmetrics(self, a, lam=0.94):
        s2 = np.zeros(len(self.r)); s2[0] = self.r[0]**2
        for t in range(1, len(self.r)):
            s2[t] = lam * s2[t-1] + (1 - lam) * self.r[t-1]**2
        st = np.sqrt(s2[-1]); z = stats.norm.ppf(1 - a)
        v = -z * st * self._scale(1); e = st * stats.norm.pdf(z) / (1 - a) * self._scale(1)
        return {"VaR": v*self.pv, "ES": e*self.pv, "VaR_pct": float(v),
                "params": f"λ=0.94 σ_t={st*100:.4f}%"}

    def cornish_fisher(self, a):
        mu, sig = self.r.mean(), self.r.std()
        s, k = float(stats.skew(self.r)), float(stats.kurtosis(self.r))
        z = stats.norm.ppf(1 - a)
        z_cf = z + (z**2-1)*s/6 + (z**3-3*z)*k/24 - (2*z**3-5*z)*s**2/36
        v = -(mu + z_cf * sig) * self._scale(1)
        e = (sig * stats.norm.pdf(z_cf) / (1 - a) - mu) * self._scale(1)
        return {"VaR": v*self.pv, "ES": e*self.pv, "VaR_pct": float(v),
                "params": f"z_CF={z_cf:.4f} skew={s:.3f} kurt={k:.3f}"}

    def _garch_params(self) -> Tuple[float, float, float, np.ndarray]:
        r = self.r
        def neg_ll(p):
            w, a, b = p
            if w<=0 or a<=0 or b<=0 or a+b>=1: return 1e10
            s2 = np.zeros(len(r)); s2[0] = np.var(r)
            for t in range(1, len(r)):
                s2[t] = w + a*r[t-1]**2 + b*s2[t-1]
            return 0.5 * np.sum(np.log(2*np.pi*s2) + r**2/s2)
        try:
            res = minimize(neg_ll, [1e-6, 0.08, 0.89], method="L-BFGS-B",
                           bounds=[(1e-8,None),(0.001,0.3),(0.5,0.999)])
            w, ag, bg = res.x
        except Exception:
            w, ag, bg = 5e-7, 0.09, 0.90
        s2 = np.zeros(len(r)); s2[0] = np.var(r)
        for t in range(1, len(r)):
            s2[t] = w + ag*r[t-1]**2 + bg*s2[t-1]
        return w, ag, bg, s2

    def garch(self, a):
        w, ag, bg, s2 = self._garch_params()
        sf = np.sqrt(w + ag*self.r[-1]**2 + bg*s2[-1])
        z  = stats.norm.ppf(1 - a)
        v  = -z * sf * self._scale(1); e = sf * stats.norm.pdf(z) / (1 - a) * self._scale(1)
        return {"VaR": v*self.pv, "ES": e*self.pv, "VaR_pct": float(v),
                "params": f"ω={w:.2e} α={ag:.3f} β={bg:.3f}"}

    def _gpd(self, exc: np.ndarray) -> Tuple[float, float]:
        def neg_ll(p):
            xi, beta = p
            if beta <= 0: return 1e10
            u_ = exc / beta
            if xi != 0:
                if np.any(1 + xi*u_ <= 0): return 1e10
                return len(u_)*np.log(beta) + (1+1/xi)*np.sum(np.log(1+xi*u_))
            return len(u_)*np.log(beta) + np.sum(u_)
        try:
            res = minimize(neg_ll, [0.1, float(np.std(exc))], method="L-BFGS-B",
                           bounds=[(-0.5, 0.5),(1e-6, None)])
            return float(res.x[0]), float(res.x[1]) if res.success \
                   else (0.1, float(np.std(exc)))
        except Exception:
            return 0.1, float(np.std(exc))

    def tve(self, a):
        losses = -self.r; u = float(np.percentile(losses, 90))
        exc = losses[losses > u] - u; xi, beta = self._gpd(exc)
        nu, n = len(exc), len(losses); p = 1 - a
        var = u + (beta/xi)*((n/nu*p)**(-xi)-1) if xi != 0 else u + beta*np.log(n/nu*p)
        var = max(float(var), 0.0)
        es  = (var + beta - xi*u) / (1 - xi) if xi < 1 else var * 2.0
        return {"VaR": var*self.pv*self._scale(1), "ES": es*self.pv*self._scale(1),
                "VaR_pct": float(var), "params": f"ξ={xi:.3f} β={beta:.4f} u={u:.4f}"}

    def tve_garch(self, a):
        w, ag, bg, s2 = self._garch_params()
        resid = self.r / np.sqrt(s2)
        losses_r = -resid; ur = float(np.percentile(losses_r, 90))
        exc_r = losses_r[losses_r > ur] - ur; xi, beta_r = self._gpd(exc_r)
        nu, n = len(exc_r), len(losses_r); p = 1 - a
        vr = ur + (beta_r/xi)*((n/nu*p)**(-xi)-1) if xi != 0 else ur + beta_r*np.log(n/nu*p)
        vr = max(float(vr), 0.0)
        er = (vr + beta_r - xi*ur) / (1 - xi) if xi < 1 else vr * 2.0
        sf = np.sqrt(w + ag*self.r[-1]**2 + bg*s2[-1])
        var = vr * sf * self._scale(1); es = er * sf * self._scale(1)
        return {"VaR": var*self.pv, "ES": es*self.pv, "VaR_pct": float(var),
                "params": f"GARCH+GPD ξ={xi:.3f} σ_f={sf*100:.4f}%"}

    def compute_all(self, alphas=(0.95, 0.99)) -> dict:
        fns = {
            "Historique":          self.historique,
            "Variance-Covariance": self.vcv,
            "RiskMetrics":         self.riskmetrics,
            "Cornish-Fisher":      self.cornish_fisher,
            "GARCH(1,1)":          self.garch,
            "TVE (POT)":           self.tve,
            "TVE-GARCH":           self.tve_garch,
        }
        out = {}
        for name, fn in fns.items():
            out[name] = {}
            for a in alphas:
                try:
                    out[name][a] = fn(a)
                except Exception as e:
                    out[name][a] = {"VaR": np.nan, "ES": np.nan, "VaR_pct": np.nan, "params": str(e)}
        return out


# ══════════════════════════════════════════════════════════════════════════════
# BACKTESTING
# ══════════════════════════════════════════════════════════════════════════════

def kupiec(r: np.ndarray, var_pct: float, a: float) -> dict:
    exc = (r < -var_pct).astype(int)
    N, T = int(exc.sum()), len(exc)
    if T == 0:
        return {"LR": np.nan, "p_value": np.nan, "valid": False, "N": 0, "T": 0, "rate": 0.0}
    p0 = 1 - a; ph = N / T
    if ph == 0:   lr = -2 * T * np.log(1 - p0)
    elif ph == 1: lr = -2 * N * np.log(p0)
    else:
        lr = -2 * (T*np.log(1-p0) + N*np.log(p0) - N*np.log(ph) - (T-N)*np.log(1-ph))
    pv = float(1 - stats.chi2.cdf(max(lr, 0.0), df=1))
    return {"LR": float(lr), "p_value": pv, "valid": pv > 0.05,
            "N": N, "T": T, "rate": float(ph), "expected": float(p0)}


def christoffersen(r: np.ndarray, var_pct: float) -> dict:
    exc = (r < -var_pct).astype(int)
    n00 = int(np.sum((exc[:-1]==0)&(exc[1:]==0)))
    n01 = int(np.sum((exc[:-1]==0)&(exc[1:]==1)))
    n10 = int(np.sum((exc[:-1]==1)&(exc[1:]==0)))
    n11 = int(np.sum((exc[:-1]==1)&(exc[1:]==1)))
    pi01 = n01/(n00+n01+1e-10); pi11 = n11/(n10+n11+1e-10)
    pi   = (n01+n11)/(n00+n01+n10+n11+1e-10)
    try:
        lr = -2*(
            (n00+n10)*np.log(max(1-pi,1e-15)) + (n01+n11)*np.log(max(pi,1e-15))
            - n00*np.log(max(1-pi01,1e-15)) - n01*np.log(max(pi01,1e-15))
            - n10*np.log(max(1-pi11,1e-15)) - n11*np.log(max(pi11,1e-15))
        )
    except Exception:
        lr = np.nan
    pv = float(1 - stats.chi2.cdf(max(float(lr), 0.0), df=1)) if not np.isnan(lr) else np.nan
    return {"LR_ind": float(lr) if not np.isnan(lr) else np.nan,
            "p_value_ind": pv,
            "valid": bool(pv > 0.05) if pv and not np.isnan(pv) else False}


# ══════════════════════════════════════════════════════════════════════════════
# MARKOWITZ
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def markowitz(rend_arr: np.ndarray, cols: tuple, rf: float = 0.03) -> dict:
    df  = pd.DataFrame(rend_arr, columns=list(cols))
    mu  = df.mean().values * 252
    cov = df.cov().values  * 252
    n   = len(mu)

    def stats_port(w):
        r_ = float(w @ mu); v_ = float(np.sqrt(w @ cov @ w))
        return r_, v_, (r_-rf)/v_ if v_ > 0 else 0.0

    cons   = [{"type":"eq","fun":lambda w: w.sum()-1}]
    bounds = [(0.0,1.0)]*n; w0 = np.ones(n)/n

    res_s = minimize(lambda w: -(w@mu-rf)/max(np.sqrt(w@cov@w),1e-10),
                     w0, method="SLSQP", bounds=bounds, constraints=cons)
    ws = res_s.x if res_s.success else w0

    res_v = minimize(lambda w: w@cov@w,
                     w0, method="SLSQP", bounds=bounds, constraints=cons)
    wv = res_v.x if res_v.success else w0

    # Frontière efficiente
    fv, fr = [], []
    for target in np.linspace(mu.min(), mu.max(), 50):
        c2 = cons + [{"type":"eq","fun":lambda w,t=target: w@mu-t}]
        r2 = minimize(lambda w: w@cov@w, w0, method="SLSQP", bounds=bounds, constraints=c2)
        if r2.success:
            fv.append(float(np.sqrt(r2.fun))); fr.append(float(target))

    tickers = list(cols)
    return {
        "tickers":  tickers, "mu": mu, "cov": cov,
        "sharpe":   {"weights": ws,    "stats": stats_port(ws),    "label": "Sharpe Max."},
        "minvar":   {"weights": wv,    "stats": stats_port(wv),    "label": "Variance Min."},
        "equi":     {"weights": w0,    "stats": stats_port(w0),    "label": "Équipondéré"},
        "frontier": (fv, fr),
    }


# ══════════════════════════════════════════════════════════════════════════════
# GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

def fig_perf(rend: pd.DataFrame) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        port_r = rend.mean(axis=1)
        cumul  = (1 + port_r).cumprod()
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 5.5),
                                        gridspec_kw={"height_ratios": [3, 1]})
        fig.patch.set_alpha(0.0)
        ax1.fill_between(range(len(cumul)), cumul.values, 1,
                         where=(cumul.values >= 1), alpha=0.12, color="#00D4FF", interpolate=True)
        ax1.fill_between(range(len(cumul)), cumul.values, 1,
                         where=(cumul.values < 1),  alpha=0.12, color="#FF4D6D", interpolate=True)
        ax1.plot(cumul.values, color="#00D4FF", lw=2, label="Portefeuille")
        for i, t in enumerate(list(rend.columns)[:5]):
            ax1.plot((1+rend[t]).cumprod().values,
                     color=PALETTE[(i+1)%len(PALETTE)], lw=0.8, alpha=0.5, label=t)
        ax1.axhline(1, color="#F0B429", lw=0.8, ls="--", alpha=0.5)
        ax1.set_title("Performance cumulée", fontsize=11, color="#F0B429", pad=8, fontweight="bold")
        ax1.legend(fontsize=7.5, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax1.grid(True, alpha=0.15, ls="--"); ax1.tick_params(labelsize=8.5); spine(ax1)

        cols_bar = ["#1db87a" if v >= 0 else "#e05252" for v in port_r]
        ax2.bar(range(len(port_r)), port_r*100, color=cols_bar, alpha=0.75, width=1, lw=0)
        ax2.axhline(0, color="#F0B429", lw=0.6)
        ax2.set_title("Rendements journaliers (%)", fontsize=9, color="#8aa0bc", pad=4)
        ax2.tick_params(labelsize=7.5); ax2.grid(True, alpha=0.15, ls="--"); spine(ax2)
        plt.tight_layout(); return fig


def fig_corr(rend: pd.DataFrame) -> plt.Figure:
    corr = rend.corr(); n = len(corr.columns)
    sz   = max(5.5, n * 1.05 + 1.5)
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(sz, sz * 0.82))
        fig.patch.set_alpha(0.0)
        cmap = LinearSegmentedColormap.from_list("vc", ["#e05252","#111827","#1db87a"], N=256)
        im   = ax.imshow(corr.values, cmap=cmap, vmin=-1, vmax=1, aspect="auto")
        cbar = plt.colorbar(im, ax=ax, shrink=0.82, pad=0.02)
        cbar.set_label("Corrélation", fontsize=9, color="#8aa0bc")
        plt.setp(cbar.ax.yaxis.get_ticklabels(), color="#8aa0bc")
        ax.set_xticks(range(n)); ax.set_yticks(range(n))
        ax.set_xticklabels(corr.columns, rotation=40, ha="right", fontsize=8.5)
        ax.set_yticklabels(corr.columns, fontsize=8.5)
        for i in range(n):
            for j in range(n):
                v = corr.values[i, j]
                ax.text(j, i, f"{v:.2f}", ha="center", va="center",
                        fontsize=7.5,
                        color="white" if abs(v) > 0.45 else "#d0daea",
                        fontweight="bold")
        ax.set_title("Matrice de Corrélation", fontsize=12, color="#F0B429", pad=10, fontweight="bold")
        spine(ax); plt.tight_layout(); return fig


def fig_var_bar(var_res: dict, conf: float, pv: float) -> plt.Figure:
    methods = list(var_res.keys())
    vars_   = [var_res[m][conf]["VaR"]/1000 for m in methods]
    ess_    = [var_res[m][conf]["ES"]/1000  for m in methods]
    short   = [m.replace("Variance-Covariance","VCV")
                .replace("Cornish-Fisher","C-F")
                .replace("RiskMetrics","RiskM") for m in methods]
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(12, 4.5))
        fig.patch.set_alpha(0.0)
        x = np.arange(len(methods)); w = 0.38
        b1 = ax.bar(x-w/2, vars_, w, color=PALETTE[:len(methods)], alpha=0.9,
                    label="VaR", edgecolor="#07090F", lw=0.5)
        ax.bar(x+w/2, ess_, w, color=PALETTE[:len(methods)], alpha=0.4,
               label="ES (CVaR)", edgecolor="#07090F", lw=0.5, hatch="//")
        for bar, col in zip(b1, PALETTE[:len(methods)]):
            h = bar.get_height()
            ax.text(bar.get_x()+bar.get_width()/2, h+max(vars_)*0.015,
                    f"{h:.0f}k", ha="center", va="bottom", fontsize=7.5,
                    color=col, fontweight="bold")
        ax.set_xticks(x); ax.set_xticklabels(short, rotation=20, ha="right", fontsize=9)
        ax.set_ylabel("k€", fontsize=10)
        ax.set_title(f"VaR & ES — Niveau {conf*100:.0f}%  ·  Portefeuille {pv/1e6:.0f}M€",
                     fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax.legend(fontsize=9, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax.grid(axis="y", alpha=0.18, ls="--"); spine(ax); ax.tick_params(labelsize=8.5)
        plt.tight_layout(); return fig


def fig_dist(r: np.ndarray, var_res: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(12, 4.5))
        fig.patch.set_alpha(0.0)
        ax.hist(r[r>=0]*100, bins=55, density=True, color="#00C896", alpha=0.45,
                edgecolor="#07090F", lw=0.2, label="Rdt ≥ 0")
        ax.hist(r[r<0]*100,  bins=55, density=True, color="#00D4FF", alpha=0.55,
                edgecolor="#07090F", lw=0.2, label="Rdt < 0")
        mu_, sig_ = r.mean(), r.std()
        xr = np.linspace(r.min(), r.max(), 400)
        ax.plot(xr*100, stats.norm.pdf(xr, mu_, sig_)/100,
                color="#F0B429", lw=2.2, ls="--", label="N(μ,σ)", zorder=5)
        for meth, col, ls in [("Historique","#e05252","--"),
                               ("GARCH(1,1)","#a855f7","-."),
                               ("TVE-GARCH","#f97316",":")]:
            if meth in var_res:
                p = var_res[meth].get(0.99, {}).get("VaR_pct")
                if p and not np.isnan(p):
                    ax.axvline(-p*100, color=col, lw=1.8, ls=ls,
                               label=f"VaR 99% {meth}", zorder=6)
        ax.set_xlabel("Rendement (%)", fontsize=10); ax.set_ylabel("Densité", fontsize=10)
        ax.set_title("Distribution des rendements · VaR 99% superposée",
                     fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax.legend(fontsize=8, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax.grid(True, alpha=0.15, ls="--"); spine(ax); ax.tick_params(labelsize=8.5)
        plt.tight_layout(); return fig


def fig_kupiec(bt: dict) -> plt.Figure:
    methods = list(bt.keys())
    short   = [m.replace("Variance-Covariance","VCV").replace("Cornish-Fisher","C-F")
                .replace("RiskMetrics","RiskM") for m in methods]
    p95 = [bt[m].get(0.95, {}).get("p_value", 0) for m in methods]
    p99 = [bt[m].get(0.99, {}).get("p_value", 0) for m in methods]
    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1, 2, figsize=(12, 4))
        fig.patch.set_alpha(0.0)
        for ax, pvals, title in zip(axes, [p95, p99], ["Kupiec 95%","Kupiec 99%"]):
            for xi, pv_ in enumerate(pvals):
                col = "#1db87a" if pv_ > 0.05 else "#e05252"
                ax.bar(xi, pv_, 0.6, color=col, alpha=0.85, edgecolor="#07090F", lw=0.6)
                ax.text(xi, pv_+0.006, f"{pv_:.3f}", ha="center", va="bottom",
                        fontsize=7.5, color=col, fontweight="bold")
            ax.axhline(0.05, color="#F0B429", lw=2, ls="--", label="Seuil α=5%")
            ax.axhspan(0, 0.05, alpha=0.06, color="#FF4D6D")
            ax.set_xticks(range(len(short)))
            ax.set_xticklabels(short, rotation=28, ha="right", fontsize=8.5)
            ax.set_ylabel("p-value", fontsize=10)
            ax.set_ylim(0, max(max(pvals)*1.2, 0.15))
            ax.set_title(title, fontsize=11, color="#F0B429", pad=8, fontweight="bold")
            ax.legend(fontsize=9, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
            ax.grid(axis="y", alpha=0.15, ls="--"); spine(ax); ax.tick_params(labelsize=8.5)
        plt.tight_layout(); return fig


def fig_frontier(opt: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, ax = plt.subplots(figsize=(11, 5.5))
        fig.patch.set_alpha(0.0)
        fv, fr = opt["frontier"]
        if len(fv) > 1:
            vf = [v*100 for v in fv]; rf_ = [r*100 for r in fr]
            pts  = np.array([vf, rf_]).T.reshape(-1, 1, 2)
            segs = np.concatenate([pts[:-1], pts[1:]], axis=1)
            lc   = LineCollection(segs, cmap="cool", linewidth=2.8, zorder=3)
            lc.set_array(np.linspace(0, 1, len(segs)))
            ax.add_collection(lc)
            ax.fill_between(vf, rf_, min(rf_), alpha=0.07, color="#00D4FF", zorder=1)
        mu = opt["mu"]; cov = opt["cov"]; tickers = opt["tickers"]
        for i, t in enumerate(tickers):
            vi = np.sqrt(cov[i,i])*100; mi = mu[i]*100
            col = PALETTE[i % len(PALETTE)]
            ax.scatter(vi, mi, s=55, color=col, zorder=6, alpha=0.85,
                       edgecolors="#07090F", lw=0.8)
            ax.annotate(t, (vi, mi), fontsize=7.5, color=col, fontweight="bold",
                        xytext=(5,4), textcoords="offset points")
        for key, col, mkr, sz, lbl in [
            ("sharpe","#F0B429","*",260,"Sharpe Max."),
            ("minvar","#00C896","v",130,"Variance Min."),
            ("equi",  "#9D7FEA","D",110,"Équipondéré"),
        ]:
            r_, v_, s_ = opt[key]["stats"]
            ax.scatter(v_*100, r_*100, s=sz*2.5, color=col, alpha=0.18, zorder=7, marker=mkr)
            ax.scatter(v_*100, r_*100, s=sz, color=col, zorder=8, marker=mkr,
                       edgecolors="white", lw=0.9, label=f"{lbl}  (Sharpe {s_:.2f})")
            ax.annotate(lbl, (v_*100, r_*100), fontsize=8.5, color=col, fontweight="bold",
                        xytext=(9,5), textcoords="offset points",
                        bbox=dict(boxstyle="round,pad=0.25", fc="#07090F", ec=col, alpha=0.75))
        ax.set_xlabel("Volatilité ann. (%)", fontsize=10)
        ax.set_ylabel("Rendement ann. (%)", fontsize=10)
        ax.set_title("Frontière Efficiente de Markowitz", fontsize=13,
                     color="#F0B429", pad=12, fontweight="bold")
        ax.legend(fontsize=8.5, loc="lower right", facecolor="#111827",
                  edgecolor="#1a2235", labelcolor="#d0daea")
        ax.tick_params(labelsize=8.5); ax.grid(True, alpha=0.18, ls="--"); spine(ax)
        plt.tight_layout(); return fig


def fig_poids(opt: dict) -> plt.Figure:
    with plt.rc_context(PLT_DARK):
        fig, axes = plt.subplots(1, 3, figsize=(13, 4.5))
        fig.patch.set_alpha(0.0)
        for ax, key, col in zip(axes,
                                 ["sharpe","minvar","equi"],
                                 ["#F0B429","#00C896","#9D7FEA"]):
            w    = opt[key]["weights"]
            mask = w > 0.005
            w_f  = w[mask]; t_f = [opt["tickers"][i] for i,m in enumerate(mask) if m]
            idx  = [i for i,m in enumerate(mask) if m]
            pie_cols = [PALETTE[i%len(PALETTE)] for i in idx]
            wedges, texts, autotexts = ax.pie(
                w_f, labels=t_f,
                autopct=lambda p: f"{p:.1f}%" if p > 3 else "",
                colors=pie_cols, startangle=90, pctdistance=0.72,
                wedgeprops={"edgecolor":"#07090F","linewidth":1.8},
                textprops={"fontsize":7.5})
            for txt in texts: txt.set_color("#d0daea"); txt.set_fontsize(7.5)
            for at in autotexts:
                at.set_fontsize(7); at.set_color("#07090F"); at.set_fontweight("bold")
            circle = plt.Circle((0,0), 0.42, color="#07090F", zorder=10)
            ring   = plt.Circle((0,0), 0.44, color=col, alpha=0.18, zorder=9)
            ax.add_patch(circle); ax.add_patch(ring)
            r_, v_, s_ = opt[key]["stats"]
            ax.text(0, 0.08, f"{s_:.2f}", ha="center", va="center",
                    fontsize=14, fontweight="bold", color=col, zorder=11)
            ax.text(0,-0.14, "Sharpe", ha="center", va="center",
                    fontsize=7, color="#8aa0bc", zorder=11)
            ax.set_title(f"{opt[key]['label']}\nRdt {r_*100:.1f}%  ·  Vol {v_*100:.1f}%",
                         fontsize=9, color=col, pad=8, fontweight="bold")
        plt.tight_layout(); return fig


def fig_stress(sr: dict, pv: float) -> plt.Figure:
    sc = list(sr.keys())
    pnls = [sr[s]["pnl_stress"]/1000 for s in sc]
    vnms = [sr[s]["var_normal"]/1000  for s in sc]
    rats = [sr[s]["ratio"] for s in sc]
    lbl  = [s[:18] for s in sc]
    x    = np.arange(len(sc))
    with plt.rc_context(PLT_DARK):
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 4.5))
        fig.patch.set_alpha(0.0)
        ax1.barh(x-0.2, [abs(p) for p in pnls], 0.38, color="#FF4D6D", alpha=0.85, label="P&L stressé")
        ax1.barh(x+0.2, vnms, 0.38, color="#00D4FF", alpha=0.65, label="VaR normale 99%")
        ax1.set_yticks(x); ax1.set_yticklabels(lbl, fontsize=8.5)
        ax1.set_xlabel("k€", fontsize=9)
        ax1.set_title("Impact P&L vs VaR normale", fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax1.legend(fontsize=8.5, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax1.grid(axis="x", alpha=0.18, ls="--"); spine(ax1)

        cols_r = ["#e05252" if r>2 else "#F0B429" if r>1.5 else "#1db87a" for r in rats]
        for xi, (r_, col) in enumerate(zip(rats, cols_r)):
            ax2.bar(xi, r_, 0.6, color=col, alpha=0.85, edgecolor="#07090F", lw=0.7)
            ax2.text(xi, r_+0.04, f"×{r_:.2f}", ha="center", va="bottom",
                     fontsize=8, fontweight="bold", color=col)
        ax2.axhline(1.0, color="#00C896", lw=1.8, ls="--", label="VaR normale (×1)")
        ax2.axhline(2.0, color="#FF4D6D", lw=1.4, ls=":", label="Seuil alerte (×2)")
        ax2.set_xticks(x); ax2.set_xticklabels(lbl, rotation=28, ha="right", fontsize=8)
        ax2.set_ylabel("Ratio ×", fontsize=9)
        ax2.set_title("Multiplicateurs de stress", fontsize=11, color="#F0B429", pad=10, fontweight="bold")
        ax2.legend(fontsize=8.5, facecolor="#111827", edgecolor="#1a2235", labelcolor="#d0daea")
        ax2.grid(axis="y", alpha=0.18, ls="--"); spine(ax2)
        plt.tight_layout(); return fig


# ══════════════════════════════════════════════════════════════════════════════
# EXPORTS
# ══════════════════════════════════════════════════════════════════════════════

def export_excel(rend: pd.DataFrame, var_res: dict, bt_res: dict,
                 pv: float, opt: Optional[dict] = None) -> Optional[bytes]:
    if not HAS_XLSX:
        return None
    wb = Workbook()
    N, B, W, L = "1B2A4A", "2E5FA3", "FFFFFF", "F0F4FA"

    def th(ws, r, c, v, bg=N):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Calibri", bold=True, color=W, size=10)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="CCCCCC")
        cell.border = Border(left=s, right=s, top=s, bottom=s)

    def td(ws, r, c, v, bg=W):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Calibri", size=9)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        s = Side(border_style="thin", color="DDDDDD")
        cell.border = Border(left=s, right=s, top=s, bottom=s)

    # Feuille 1 — VaR
    ws1 = wb.active; ws1.title = "Résumé VaR"
    ws1.merge_cells("A1:F1")
    c = ws1["A1"]; c.value = "RAPPORT VaR — SYNTHÈSE"
    c.font = Font(name="Calibri", bold=True, color=W, size=14)
    c.fill = PatternFill("solid", fgColor=N)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30
    for j, h in enumerate(["Méthode","VaR 95%(€)","VaR 95%(%)","VaR 99%(€)","VaR 99%(%)","ES 99%(€)"], 1):
        th(ws1, 2, j, h, bg=B)
    for i, (m, res) in enumerate(var_res.items(), 3):
        r95 = res.get(0.95,{}); r99 = res.get(0.99,{}); bg_ = L if i%2==0 else W
        td(ws1,i,1,m,bg=bg_); td(ws1,i,2,round(r95.get("VaR",0),0),bg=bg_)
        td(ws1,i,3,f"{r95.get('VaR_pct',0)*100:.3f}%",bg=bg_)
        td(ws1,i,4,round(r99.get("VaR",0),0),bg=bg_)
        td(ws1,i,5,f"{r99.get('VaR_pct',0)*100:.3f}%",bg=bg_)
        td(ws1,i,6,round(r99.get("ES",0),0),bg=bg_)
    for w_, col in zip([26,16,12,16,12,16],["A","B","C","D","E","F"]):
        ws1.column_dimensions[col].width = w_

    # Feuille 2 — Backtesting
    ws2 = wb.create_sheet("Backtesting")
    for j, h in enumerate(["Méthode","CL","Exc.","T","Taux obs.","Taux att.",
                             "Kupiec LR","Kupiec p","Kupiec OK","CC LR","CC p","CC OK"], 1):
        th(ws2, 1, j, h, bg=B)
    row = 2
    for m, alphas in bt_res.items():
        for a, res in alphas.items():
            k = res.get("kupiec",{}); cc = res.get("cc",{}); bg_ = L if row%2==0 else W
            td(ws2,row,1,m,bg=bg_); td(ws2,row,2,f"{a*100:.0f}%",bg=bg_)
            td(ws2,row,3,k.get("N",""),bg=bg_); td(ws2,row,4,k.get("T",""),bg=bg_)
            td(ws2,row,5,f"{k.get('rate',0)*100:.2f}%",bg=bg_)
            td(ws2,row,6,f"{(1-a)*100:.2f}%",bg=bg_)
            td(ws2,row,7,round(k.get("LR",0),3),bg=bg_)
            td(ws2,row,8,round(k.get("p_value",0),4),bg=bg_)
            td(ws2,row,9,"OUI" if k.get("valid") else "NON",bg=bg_)
            td(ws2,row,10,round(cc.get("LR_ind",0) if cc.get("LR_ind") else 0,3),bg=bg_)
            td(ws2,row,11,round(cc.get("p_value_ind",0) if cc.get("p_value_ind") else 0,4),bg=bg_)
            td(ws2,row,12,"OUI" if cc.get("valid") else "NON",bg=bg_)
            row += 1
    for w_, col in zip([26,6,8,8,10,10,10,10,10,10,10,10], list("ABCDEFGHIJKL")):
        ws2.column_dimensions[col].width = w_

    # Feuille 3 — Markowitz
    if opt:
        ws3 = wb.create_sheet("Markowitz")
        for j, h in enumerate(["Actif","Sharpe Max.(%)","Variance Min.(%)","Équipondéré(%)"], 1):
            th(ws3, 1, j, h, bg=B)
        for i, t in enumerate(opt["tickers"], 2):
            bg_ = L if i%2==0 else W
            td(ws3,i,1,t,bg=bg_)
            td(ws3,i,2,f"{opt['sharpe']['weights'][i-2]*100:.1f}%",bg=bg_)
            td(ws3,i,3,f"{opt['minvar']['weights'][i-2]*100:.1f}%",bg=bg_)
            td(ws3,i,4,f"{opt['equi']['weights'][i-2]*100:.1f}%",bg=bg_)
        rs = len(opt["tickers"]) + 3
        for j, h in enumerate(["Statistiques","Sharpe Max.","Variance Min.","Équipondéré"], 1):
            th(ws3, rs, j, h, bg=N)
        for j, key in enumerate(["sharpe","minvar","equi"], 2):
            r_, v_, s_ = opt[key]["stats"]
            td(ws3,rs+1,j,f"{r_*100:.2f}%"); td(ws3,rs+2,j,f"{v_*100:.2f}%"); td(ws3,rs+3,j,f"{s_:.3f}")
        for r_, lbl in zip([rs+1,rs+2,rs+3],["Rdt Ann.","Vol. Ann.","Sharpe"]):
            td(ws3,r_,1,lbl)
        for col in ["A","B","C","D"]: ws3.column_dimensions[col].width = 18

    # Feuille 4 — Données
    ws4 = wb.create_sheet("Données")
    th(ws4,1,1,"Date",bg=B); th(ws4,1,2,"Rdt Portfolio (%)",bg=B)
    r_port = rend.mean(axis=1).tail(250)
    for i, (d, v) in enumerate(r_port.items(), 2):
        bg_ = L if i%2==0 else W
        td(ws4,i,1,d.strftime("%d/%m/%Y"),bg=bg_); td(ws4,i,2,round(v*100,4),bg=bg_)
    ws4.column_dimensions["A"].width = 14; ws4.column_dimensions["B"].width = 18

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


def export_pdf(var_res: dict, bt_res: dict, pv: float,
               fv=None, fd=None) -> Optional[bytes]:
    if not HAS_PDF:
        return None
    from datetime import date
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             leftMargin=2*cm, rightMargin=2*cm,
                             topMargin=1.5*cm, bottomMargin=2*cm)
    N_c = HexColor("#1B2A4A"); B_c = HexColor("#2E5FA3")
    G_c = HexColor("#C9A84C"); LG  = HexColor("#F0F4FA")

    def ps(name, **kw): return ParagraphStyle(name, **kw)
    s_t = ps("t",fontName="Helvetica-Bold",fontSize=22,textColor=HexColor("#FFFFFF"),alignment=TA_CENTER,spaceAfter=4)
    s_s = ps("s",fontName="Helvetica-Oblique",fontSize=11,textColor=G_c,alignment=TA_CENTER,spaceAfter=6)
    s_h = ps("h",fontName="Helvetica-Bold",fontSize=12,textColor=N_c,spaceBefore=12,spaceAfter=6)
    s_b = ps("b",fontName="Helvetica",fontSize=9,textColor=HexColor("#3A3A3A"),
             alignment=TA_JUSTIFY,spaceAfter=6,leading=14)

    def tbl_s():
        return TableStyle([
            ("BACKGROUND",(0,0),(-1,0),B_c),("TEXTCOLOR",(0,0),(-1,0),HexColor("#FFFFFF")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8.5),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[HexColor("#FFFFFF"),LG]),
            ("GRID",(0,0),(-1,-1),0.3,HexColor("#CCCCCC")),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ])

    story = [Spacer(1,2*cm)]
    cov1 = Table([[Paragraph("RAPPORT DE GESTION DES RISQUES",s_t)]],[15*cm])
    cov1.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),N_c),
                               ("TOPPADDING",(0,0),(-1,-1),18),("BOTTOMPADDING",(0,0),(-1,-1),18)]))
    story.append(cov1); story.append(Spacer(1,0.3*cm))
    cov2 = Table([[Paragraph("Value at Risk · 7 méthodes · Backtesting · Stress-Testing",s_s)]],[15*cm])
    cov2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),HexColor("#243B60")),
                               ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(cov2); story.append(Spacer(1,1.5*cm))
    info_ = [["Valeur portefeuille",f"{pv:,.0f} €"],["Horizon","1 jour ouvré"],
             ["Niveaux de confiance","95% et 99%"],
             ["Date",date.today().strftime("%d/%m/%Y")]]
    t_i = Table([[Paragraph(k,s_b),Paragraph(v,s_b)] for k,v in info_],[5*cm,10*cm])
    t_i.setStyle(TableStyle([("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
                               ("ROWBACKGROUNDS",(0,0),(-1,-1),[HexColor("#FFFFFF"),LG]),
                               ("GRID",(0,0),(-1,-1),0.3,HexColor("#DDDDDD")),
                               ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
    story.append(t_i); story.append(PageBreak())

    story.append(Paragraph("1. RÉSULTATS DE LA VALUE AT RISK",s_h))
    story.append(HRFlowable(width="100%",thickness=1,color=G_c,spaceAfter=8))
    vd = [["Méthode","VaR 95%(€)","VaR 95%(%)","VaR 99%(€)","VaR 99%(%)","ES 99%(€)"]]
    for m, res in var_res.items():
        r95 = res.get(0.95,{}); r99 = res.get(0.99,{})
        vd.append([m,f"{r95.get('VaR',0):,.0f}€",f"{r95.get('VaR_pct',0)*100:.3f}%",
                   f"{r99.get('VaR',0):,.0f}€",f"{r99.get('VaR_pct',0)*100:.3f}%",
                   f"{r99.get('ES',0):,.0f}€"])
    t_v = Table(vd,[3.5*cm,2.5*cm,2*cm,2.5*cm,2*cm,2.5*cm])
    t_v.setStyle(tbl_s()); story.append(t_v); story.append(Spacer(1,0.5*cm))
    if fv:
        b2 = io.BytesIO(); fv.savefig(b2,format="png",dpi=120,bbox_inches="tight")
        b2.seek(0); story.append(Image(b2,width=15*cm,height=5*cm)); plt.close(fv)
    story.append(PageBreak())

    story.append(Paragraph("2. BACKTESTING",s_h))
    story.append(HRFlowable(width="100%",thickness=1,color=G_c,spaceAfter=8))
    story.append(Paragraph("Test de Kupiec : H₀ → fréquence observée = 1−α. "
                            "Test de Christoffersen : indépendance temporelle. "
                            "p > 0.05 → modèle non rejeté.",s_b))
    bd = [["Méthode","CL","Exc.","Taux obs.","Kupiec p","Kupiec","CC p","CC"]]
    for m, alphas in bt_res.items():
        for a, res in alphas.items():
            k = res.get("kupiec",{}); cc = res.get("cc",{})
            bd.append([m,f"{a*100:.0f}%",str(k.get("N","")),
                       f"{k.get('rate',0)*100:.2f}%",
                       f"{k.get('p_value',0):.4f}","OUI" if k.get("valid") else "NON",
                       f"{cc.get('p_value_ind',0) if cc.get('p_value_ind') else 0:.4f}",
                       "OUI" if cc.get("valid") else "NON"])
    t_b = Table(bd,[3.5*cm,1.2*cm,1.3*cm,1.8*cm,1.8*cm,1.5*cm,1.8*cm,1.5*cm])
    t_b.setStyle(tbl_s()); story.append(t_b); story.append(Spacer(1,0.5*cm))
    if fd:
        b3 = io.BytesIO(); fd.savefig(b3,format="png",dpi=120,bbox_inches="tight")
        b3.seek(0); story.append(Image(b3,width=15*cm,height=5*cm)); plt.close(fd)

    doc.build(story); buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

for k in ["prix","rendements","var_results","bt_results","pv","opt_results","stress_results"]:
    if k not in st.session_state:
        st.session_state[k] = None

if "actifs_sel" not in st.session_state:
    st.session_state["actifs_sel"] = [
        "Apple (AAPL)","Microsoft (MSFT)","LVMH (MC.PA)",
        "TotalEnergies (TTE)","BNP Paribas (BNP)"
    ]


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div style='padding:12px 4px 8px'>
      <div style='font-size:16px;font-weight:800;color:#00D4FF;letter-spacing:1px'>
        📉 VaR Analytics Suite
      </div>
      <div style='font-size:10px;color:#3D4F6B;font-family:DM Mono,monospace;margin-top:2px'>
        v4.2 · Département Risque
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    PAGE_LABELS = [
        "🏠  Accueil",
        "🏦  Portefeuille",
        "📐  Optimisation",
        "📉  Calcul VaR",
        "🧪  Backtesting",
        "🔥  Stress-Testing",
        "📊  Reporting",
    ]
    PAGE_KEYS = ["accueil","portfolio","optim","var","backtest","stress","reporting"]

    page_label = st.radio("", PAGE_LABELS, label_visibility="collapsed")
    menu = PAGE_KEYS[PAGE_LABELS.index(page_label)]

    st.divider()

    if st.session_state["rendements"] is not None:
        r_ = st.session_state["rendements"].mean(axis=1)
        ann_r_ = r_.mean() * 252
        ann_v_ = r_.std()  * np.sqrt(252)
        sh_    = (r_.mean() - 0.03/252) / r_.std() * np.sqrt(252)
        src_tag = "" if HAS_YF else " (sim.)"
        st.markdown(f"""
        <div style='background:rgba(0,212,255,0.06);border:1px solid rgba(0,212,255,0.15);
             border-radius:8px;padding:10px 12px;margin-bottom:8px'>
          <div style='font-size:9px;color:#3D4F6B;font-family:DM Mono,monospace;
               text-transform:uppercase;letter-spacing:1px;margin-bottom:6px'>
            Portefeuille{src_tag}
          </div>
          <div style='font-size:11px;font-family:DM Mono,monospace;line-height:2;color:#7A8BA8'>
            Rdt ann. :
            <span style='color:#00C896;font-weight:600'>+{ann_r_*100:.2f}%</span><br>
            Vol. ann. :
            <span style='color:#E8EDF5'>{ann_v_*100:.2f}%</span><br>
            Sharpe&nbsp;&nbsp; :
            <span style='color:#F0B429;font-weight:600'>{sh_:.3f}</span>
          </div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='font-size:10px;color:#3D4F6B;line-height:2;font-family:DM Mono,monospace'>
    <span style='color:#F0B429'>■</span> 7 méthodes VaR<br>
    <span style='color:#00D4FF'>■</span> Tests Kupiec + CC<br>
    <span style='color:#00C896'>■</span> Frontière Markowitz<br>
    <span style='color:#FF4D6D'>■</span> 6 scénarios stress<br>
    <span style='color:#9D7FEA'>■</span> Export Excel + PDF
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════

if menu == "accueil":
    st.title("VaR Analytics Suite")
    info("Progiciel professionnel de calcul, comparaison et validation de la "
         "<b>Value at Risk</b>. Standards <b>Bâle III/IV</b>. "
         "Optimisation Markowitz · Stress-Testing · ES analytique.")

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.metric("Méthodes VaR",    "7",           "Complètes")
    with c2: st.metric("Tests backtest",  "2",           "Kupiec + CC")
    with c3: st.metric("Scénarios Stress","6",           "Historiques")
    with c4: st.metric("Export",          "Excel + PDF", "4 feuilles")

    sep("Fonctionnalités")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**📥 Données de marché**
- Yahoo Finance (15 actifs) ou simulation
- Matrice de corrélation interactive

**📐 Optimisation Markowitz**
- Frontière efficiente · Sharpe Max · Variance Min
        """)
    with c2:
        st.markdown("""
**🔥 Stress-Testing**
- 6 scénarios historiques (Lehman, COVID…)
- Multiplicateurs de risque · Seuils Bâle III

**📊 Reporting**
- Excel 4 feuilles + PDF exécutif avec graphiques
        """)

    sep("Équipe Projet")
    membres = ["Anta Mbaye","Harlem D. Adjagba","Ecclésiaste Gnargo","Wariol G. Kopangoye"]
    cols = st.columns(4)
    for col, m in zip(cols, membres):
        with col:
            st.markdown(f"""
            <div class='var-card' style='text-align:center;padding:14px'>
              <div style='font-size:22px;margin-bottom:6px'>👤</div>
              <div style='font-size:11px;font-weight:600;color:#E8EDF5'>{m}</div>
            </div>""", unsafe_allow_html=True)
    st.markdown("<div style='text-align:center;font-size:11px;color:#3D4F6B;margin-top:10px'>"
                "Double diplôme M2 IFIM · Ing 3 MACS</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE PORTEFEUILLE
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "portfolio":
    st.title("Construction du Portefeuille")
    sep("Sélection des actifs")

    c1, c2 = st.columns([2,1])
    with c1:
        actifs_choisis = st.multiselect(
            "Actifs financiers", list(ACTIFS.keys()),
            default=st.session_state["actifs_sel"], key="actifs_sel")
    with c2:
        pv_m = st.number_input("Valeur (M€)", 0.1, 1000.0, 10.0, 0.5)
        pv   = pv_m * 1_000_000

    c3,c4,c5 = st.columns(3)
    with c3: debut = st.date_input("Date début", value=pd.to_datetime("2019-01-01"))
    with c4: fin   = st.date_input("Date fin",   value=pd.to_datetime("today"))
    with c5: src   = st.radio("Source", ["Yahoo Finance","Simulation"], horizontal=True)

    bc, _ = st.columns([1,3])
    with bc:
        btn = st.button("▶  Charger", type="primary", use_container_width=True)

    if btn:
        if len(actifs_choisis) < 2:
            st.warning("Sélectionnez au moins 2 actifs.")
        elif debut >= fin:
            st.warning("La date de début doit précéder la date de fin.")
        else:
            tickers = [ACTIFS[a] for a in actifs_choisis]
            with st.spinner("Chargement…"):
                if src == "Yahoo Finance" and HAS_YF:
                    prix = telecharger(tickers, debut, fin)
                    if prix.empty:
                        st.warning("Aucune donnée, simulation activée.")
                        prix = simuler(tickers)
                else:
                    prix = simuler(tickers)
                rend = prix.pct_change(fill_method=None).dropna(how="all").dropna(axis=1, how="any")
                prix = prix[rend.columns]
            st.session_state.update({
                "prix":prix,"rendements":rend,"pv":pv,
                "var_results":None,"bt_results":None,"opt_results":None,"stress_results":None
            })
            st.success(f"✅ {len(rend.columns)} actif(s) — {len(rend)} jours")

    if st.session_state["rendements"] is not None:
        rend = st.session_state["rendements"]
        port = rend.mean(axis=1)
        ann_r = port.mean()*252; ann_v = port.std()*np.sqrt(252)
        sh_   = (port.mean()-0.03/252)/port.std()*np.sqrt(252)
        mdd_  = float(((1+port).cumprod()/(1+port).cumprod().cummax()-1).min())

        sep("Statistiques du portefeuille")
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Rdt Ann.",  f"+{ann_r*100:.2f}%")
        c2.metric("Vol. Ann.", f"{ann_v*100:.2f}%")
        c3.metric("Sharpe",    f"{sh_:.3f}")
        c4.metric("Max DD",    f"{mdd_*100:.2f}%")
        c5.metric("Skewness",  f"{stats.skew(port):.4f}")
        c6.metric("Kurtosis",  f"{stats.kurtosis(port):.4f}")

        sep("Performance & Rendements")
        f1 = fig_perf(rend); st.pyplot(f1, use_container_width=True); plt.close(f1)

        sep("Matrice de Corrélation")
        f2 = fig_corr(rend); st.pyplot(f2, use_container_width=True); plt.close(f2)

        sep("Statistiques individuelles")
        df_s = pd.DataFrame({
            "Secteur":    [SECTEURS.get(t,"—") for t in rend.columns],
            "Rdt (%)":    (rend.mean()*252*100).round(2),
            "Vol. (%)":   (rend.std()*np.sqrt(252)*100).round(2),
            "Skewness":   rend.apply(lambda c: round(float(stats.skew(c)),4)),
            "Kurtosis":   rend.apply(lambda c: round(float(stats.kurtosis(c)),4)),
            "Min (%)":    (rend.min()*100).round(3),
            "Max (%)":    (rend.max()*100).round(3),
        }, index=rend.columns)
        df_s.index.name = "Ticker"
        st.dataframe(df_s, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE OPTIMISATION
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "optim":
    st.title("Optimisation — Markowitz")
    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille."); st.stop()
    rend = st.session_state["rendements"]
    info("<b>Théorie Moderne du Portefeuille (Markowitz, 1952)</b> — Frontière efficiente : "
         "meilleur rendement pour un risque donné. "
         "3 allocations : Sharpe Max · Variance Min · Équipondéré.")

    crf, cb, _ = st.columns([1,1,2])
    with crf: rf = st.number_input("Taux sans risque (%/an)", 0.0, 10.0, 3.0, 0.1) / 100
    with cb:
        st.markdown("<br>", unsafe_allow_html=True)
        ob = st.button("▶  Optimiser", type="primary", use_container_width=True)

    if ob:
        with st.spinner("Calcul de la frontière efficiente…"):
            opt = markowitz(rend.values, tuple(rend.columns), rf)
            st.session_state["opt_results"] = opt
        st.success("✅ Optimisation terminée.")

    if st.session_state["opt_results"]:
        opt = st.session_state["opt_results"]
        sep("Frontière Efficiente")
        f3 = fig_frontier(opt); st.pyplot(f3, use_container_width=True); plt.close(f3)

        sep("Allocations optimales")
        cols = st.columns(3)
        for col, (key, color) in zip(cols, [("sharpe","#F0B429"),("minvar","#00C896"),("equi","#9D7FEA")]):
            p = opt[key]; r_,v_,s_ = p["stats"]
            with col:
                st.markdown(f"""
                <div class='mko-card'>
                  <div style='font-size:11px;color:{color};font-family:DM Mono,monospace;
                       text-transform:uppercase;letter-spacing:1px;margin-bottom:8px'>{p['label']}</div>
                  <div style='font-size:13px;color:#E8EDF5;line-height:2;font-family:DM Mono,monospace'>
                  📈 Rdt : <b style='color:{color}'>{r_*100:.2f}%</b><br>
                  📊 Vol : <b style='color:#E8EDF5'>{v_*100:.2f}%</b><br>
                  ⭐ Sharpe : <b style='color:{color}'>{s_:.3f}</b>
                  </div>
                </div>""", unsafe_allow_html=True)

        sep("Répartition des poids")
        f4 = fig_poids(opt); st.pyplot(f4, use_container_width=True); plt.close(f4)

        sep("Tableau des poids")
        df_w = pd.DataFrame({
            "Actif":      opt["tickers"],
            "Sharpe Max.": [f"{w*100:.1f}%" for w in opt["sharpe"]["weights"]],
            "Var. Min.":   [f"{w*100:.1f}%" for w in opt["minvar"]["weights"]],
            "Équipondéré": [f"{w*100:.1f}%" for w in opt["equi"]["weights"]],
        }).set_index("Actif")
        st.dataframe(df_w, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE CALCUL VaR
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "var":
    st.title("Calcul de la Value at Risk")
    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille."); st.stop()

    rend   = st.session_state["rendements"]
    pv     = st.session_state["pv"] or 10_000_000
    port_r = rend.mean(axis=1).values

    sep("Paramètres")
    c1,c2,c3 = st.columns(3)
    with c1: horizon = st.slider("Horizon (jours)", 1, 10, 1)
    with c2:
        conf_opts = st.multiselect("Niveaux de confiance",
                                    [0.90,0.95,0.975,0.99], default=[0.95,0.99],
                                    format_func=lambda x: f"{x*100:.1f}%")
    with c3:
        methods_sel = st.multiselect("Méthodes",
            ["Historique","Variance-Covariance","RiskMetrics",
             "Cornish-Fisher","GARCH(1,1)","TVE (POT)","TVE-GARCH"],
            default=["Historique","Variance-Covariance","RiskMetrics",
                     "Cornish-Fisher","GARCH(1,1)","TVE (POT)","TVE-GARCH"])

    bc, _ = st.columns([1,3])
    with bc:
        cb = st.button("▶  Calculer les 7 VaR", type="primary", use_container_width=True)

    if cb:
        if not conf_opts:
            st.warning("Sélectionnez au moins un niveau de confiance.")
        else:
            with st.spinner("Calcul en cours…"):
                engine  = VaREngine(port_r, pv, horizon)
                var_res = engine.compute_all(tuple(sorted(conf_opts)))
                var_res = {k:v for k,v in var_res.items() if k in methods_sel}
                st.session_state["var_results"] = var_res
            st.success(f"✅ {len(var_res)} méthodes × {len(conf_opts)} niveaux")

    if st.session_state["var_results"]:
        var_res  = st.session_state["var_results"]
        alphas_  = sorted(list(list(var_res.values())[0].keys()))
        a_disp   = st.select_slider("Afficher pour :", options=alphas_,
                                     format_func=lambda x:f"{x*100:.0f}%", value=alphas_[-1])

        sep("Résultats par méthode")
        cols = st.columns(min(len(var_res), 4))
        for i, (method, res) in enumerate(var_res.items()):
            r = res.get(a_disp, {})
            vv = r.get("VaR",np.nan); pv_ = r.get("VaR_pct",np.nan); ev = r.get("ES",np.nan)
            rec = '<span class="badge-rec">★ Recommandé</span>' if method=="TVE-GARCH" else ""
            with cols[i % len(cols)]:
                st.markdown(f"""
                <div class='var-card'>
                  <div class='var-card-title'>{method}{rec}</div>
                  <div class='var-card-value'>{vv/1000:.1f} k€</div>
                  <div class='var-card-pct'>VaR {a_disp*100:.0f}% · {pv_*100:.3f}%</div>
                  <div class='var-card-es'>ES : {ev/1000:.1f} k€</div>
                </div>""", unsafe_allow_html=True)

        sep("Tableau comparatif")
        rows = []
        for m, res in var_res.items():
            row = {"Méthode": m}
            for a in alphas_:
                r = res.get(a,{})
                row[f"VaR {a*100:.0f}%(€)"] = f"{r.get('VaR',0):,.0f}"
                row[f"VaR {a*100:.0f}%(%)"] = f"{r.get('VaR_pct',0)*100:.3f}%"
                row[f"ES {a*100:.0f}%(€)"]  = f"{r.get('ES',0):,.0f}"
            row["Paramètres"] = list(var_res[m].values())[0].get("params","")
            rows.append(row)
        st.dataframe(pd.DataFrame(rows).set_index("Méthode"), use_container_width=True)

        sep("Graphique comparatif")
        f5 = fig_var_bar(var_res, a_disp, pv)
        st.pyplot(f5, use_container_width=True); plt.close(f5)

        sep("Distribution des rendements")
        f6 = fig_dist(port_r, var_res)
        st.pyplot(f6, use_container_width=True); plt.close(f6)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE BACKTESTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "backtest":
    st.title("Backtesting des modèles de VaR")
    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR."); st.stop()

    rend    = st.session_state["rendements"]
    var_res = st.session_state["var_results"]
    port_r  = rend.mean(axis=1).values
    alphas_ = sorted(list(list(var_res.values())[0].keys()))

    info("<b>Test de Kupiec (POF)</b> — fréquence d'exceptions vs niveau de confiance. LR ~ χ²(1).<br>"
         "<b>Test de Christoffersen (CC)</b> — indépendance temporelle des exceptions.<br>"
         "<span style='color:#00C896'>✅ p &gt; 5%</span> → validé &nbsp;"
         "<span style='color:#FF4D6D'>❌ p ≤ 5%</span> → rejeté")

    bc, _ = st.columns([1,3])
    with bc:
        bb = st.button("▶  Lancer le backtesting", type="primary", use_container_width=True)

    if bb:
        with st.spinner("Backtesting en cours…"):
            bt_res = {}
            for method, res in var_res.items():
                bt_res[method] = {}
                for a in alphas_:
                    vp = res.get(a,{}).get("VaR_pct", np.nan)
                    if np.isnan(vp): continue
                    bt_res[method][a] = {
                        "kupiec": kupiec(port_r, vp, a),
                        "cc":     christoffersen(port_r, vp),
                    }
            st.session_state["bt_results"] = bt_res
        st.success("✅ Backtesting terminé.")

    if st.session_state["bt_results"]:
        bt_res = st.session_state["bt_results"]
        rows = []
        for m, alphas in bt_res.items():
            for a, res in alphas.items():
                k = res["kupiec"]; cc = res["cc"]
                rows.append({
                    "Méthode":m,"CL":f"{a*100:.0f}%",
                    "Exc.":k["N"],"T":k["T"],
                    "Taux obs.":f"{k['rate']*100:.2f}%",
                    "Taux att.":f"{(1-a)*100:.2f}%",
                    "Kupiec LR":round(k["LR"],3),
                    "Kupiec p": round(k["p_value"],4),
                    "Kupiec ✓":"✅" if k["valid"] else "❌",
                    "CC p":     round(cc.get("p_value_ind") or 0, 4),
                    "CC ✓":     "✅" if cc.get("valid") else "❌",
                })
        st.dataframe(pd.DataFrame(rows).set_index("Méthode"), use_container_width=True)

        sep("p-values Kupiec")
        f7 = fig_kupiec({m:{a:{"p_value":bt_res[m][a]["kupiec"]["p_value"]}
                            for a in bt_res[m]} for m in bt_res})
        st.pyplot(f7, use_container_width=True); plt.close(f7)

        sep("Verdict synthétique")
        av = alphas_[-1]; cols_v = st.columns(len(bt_res))
        for col, (m, alphas) in zip(cols_v, bt_res.items()):
            r_ = alphas.get(av,{}); k=r_.get("kupiec",{}); cc=r_.get("cc",{})
            kok, cok = k.get("valid",False), cc.get("valid",False)
            score = "✅ Validé" if kok and cok else ("⚠️ Partiel" if kok or cok else "❌ Rejeté")
            color = "#00C896" if kok and cok else ("#F0B429" if kok or cok else "#FF4D6D")
            with col:
                st.markdown(f"""
                <div class='var-card' style='text-align:center'>
                  <div class='var-card-title'>{m}</div>
                  <div style='font-size:15px;font-weight:700;color:{color};margin:6px 0'>{score}</div>
                  <div style='font-size:10px;color:#7A8BA8;font-family:DM Mono,monospace'>
                  p_K={k.get('p_value',0):.4f}<br>p_CC={cc.get('p_value_ind') or 0:.4f}
                  </div>
                </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE STRESS-TESTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "stress":
    st.title("Stress-Testing — Scénarios Historiques")
    if st.session_state["rendements"] is None:
        st.info("💡 Chargez d'abord un portefeuille."); st.stop()

    rend    = st.session_state["rendements"]
    var_res = st.session_state["var_results"] or {}
    pv      = st.session_state["pv"] or 10_000_000
    port_r  = rend.mean(axis=1).values

    info("Le <b>stress-testing</b> applique des chocs déterministes calibrés sur des crises réelles. "
         "Exigé par <b>Bâle III/IV</b> (EBA, BCE).")

    sep("Sélection")
    sc_choisis = st.multiselect("Scénarios", list(STRESS_SCENARIOS.keys()),
                                 default=list(STRESS_SCENARIOS.keys()))
    ca, cb_, _ = st.columns([1,1,2])
    with ca: ast_ = st.selectbox("Niveau VaR", [0.95,0.99], format_func=lambda x:f"{x*100:.0f}%", index=1)
    with cb_:
        st.markdown("<br>", unsafe_allow_html=True)
        stb = st.button("▶  Lancer le stress-test", type="primary", use_container_width=True)

    if stb and sc_choisis:
        with st.spinner("Calcul…"):
            mu_r = port_r.mean(); sig_r = port_r.std()
            var_normal = var_res.get("Variance-Covariance",{}).get(ast_,{}).get("VaR",0)
            sr = {}
            for s in sc_choisis:
                sc = STRESS_SCENARIOS[s]
                choc = sc["choc"] if sc["choc"] is not None else mu_r - 3*sig_r
                ss = sig_r * sc["vol_mult"]; z = stats.norm.ppf(1 - ast_)
                sr[s] = {
                    "pnl_stress":  choc * pv,
                    "var_stress":  -z * ss * pv,
                    "var_normal":  var_normal,
                    "ratio":       (-z*ss*pv) / var_normal if var_normal > 0 else np.nan,
                    "choc":        choc,
                }
            st.session_state["stress_results"] = sr
        st.success(f"✅ {len(sr)} scénario(s) calculé(s).")

    if st.session_state["stress_results"]:
        sr = st.session_state["stress_results"]
        sep("Résultats")
        sc_list = list(sr.items())
        for i in range(0, len(sc_list), 3):
            cols = st.columns(3)
            for col, (sc_name, s) in zip(cols, sc_list[i:i+3]):
                ratio = s["ratio"]
                alert = "🔴" if (not np.isnan(ratio) and ratio>2) else \
                        ("🟡" if (not np.isnan(ratio) and ratio>1.5) else "🟢")
                rtxt  = f"×{ratio:.2f}" if not np.isnan(ratio) else "—"
                with col:
                    st.markdown(f"""
                    <div class='stress-card'>
                      <div style='font-size:9.5px;color:rgba(255,77,109,0.8);
                           font-family:DM Mono,monospace;text-transform:uppercase;
                           letter-spacing:1.5px;margin-bottom:6px'>{alert} {sc_name}</div>
                      <div style='font-size:10px;color:#7A8BA8;margin-bottom:6px'>
                        {STRESS_SCENARIOS.get(sc_name,{}).get('date','')}
                      </div>
                      <div style='font-size:22px;font-weight:300;color:#FF4D6D;font-family:DM Mono,monospace'>
                        {s['pnl_stress']/1000:+.0f} k€
                      </div>
                      <div style='font-size:11px;color:#7A8BA8;margin-top:4px;font-family:DM Mono,monospace'>
                        VaR stressée : <b style='color:#e87373'>{s['var_stress']/1000:.0f} k€</b><br>
                        Ratio : <b style='color:#F0B429'>{rtxt}</b><br>
                        Choc : <b style='color:#FF4D6D'>{s['choc']*100:.2f}%</b>
                      </div>
                    </div>""", unsafe_allow_html=True)

        sep("Analyse comparative")
        f8 = fig_stress(sr, pv); st.pyplot(f8, use_container_width=True); plt.close(f8)

        sep("Tableau")
        rows_st = []
        for sc_name, s in sr.items():
            ratio = s["ratio"]
            rows_st.append({
                "Scénario":         sc_name,
                "Choc (%)":         f"{s['choc']*100:.2f}%",
                "P&L (k€)":         f"{s['pnl_stress']/1000:+.1f}",
                "VaR stressée (k€)":f"{s['var_stress']/1000:.1f}",
                "VaR normale (k€)": f"{s['var_normal']/1000:.1f}",
                "Ratio ×":          f"{ratio:.2f}" if not np.isnan(ratio) else "—",
                "Statut":           "🔴 ALERTE" if (not np.isnan(ratio) and ratio>2)
                                    else ("🟡 VIGILANCE" if (not np.isnan(ratio) and ratio>1.5)
                                    else "🟢 OK"),
            })
        st.dataframe(pd.DataFrame(rows_st).set_index("Scénario"), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE REPORTING
# ══════════════════════════════════════════════════════════════════════════════

elif menu == "reporting":
    st.title("Génération des Rapports")
    if st.session_state["var_results"] is None:
        st.info("💡 Calculez d'abord les VaR."); st.stop()

    rend    = st.session_state["rendements"]
    var_res = st.session_state["var_results"]
    bt_res  = st.session_state["bt_results"] or {}
    pv      = st.session_state["pv"] or 10_000_000
    opt     = st.session_state.get("opt_results")
    port_r  = rend.mean(axis=1).values
    alphas_ = sorted(list(list(var_res.values())[0].keys()))
    a99     = max(alphas_)
    best    = "TVE-GARCH" if "TVE-GARCH" in var_res else list(var_res.keys())[-1]

    info("Générez un <b>rapport Excel</b> (4 feuilles) et un <b>rapport PDF exécutif</b> "
         "conformes aux exigences Bâle III/IV.")

    sep("Aperçu")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Méthode recommandée", best)
    c2.metric("VaR 99%", f"{var_res[best][a99]['VaR']/1000:.1f} k€")
    c3.metric("ES 99%",  f"{var_res[best][a99]['ES']/1000:.1f} k€")
    n_ok    = sum(1 for m in bt_res for a in bt_res[m] if bt_res[m][a]["kupiec"]["valid"]) if bt_res else 0
    n_total = sum(len(bt_res[m]) for m in bt_res) if bt_res else 0
    c4.metric("Backtests validés", f"{n_ok}/{n_total}" if n_total else "—")

    bt_fmt = {m:{a:res for a,res in alphas.items()} for m,alphas in bt_res.items()}

    sep("Téléchargements")
    cx, cp = st.columns(2)

    with cx:
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
                xlsx = export_excel(rend, var_res, bt_fmt, pv, opt)
            if xlsx:
                st.download_button("⬇  Télécharger Excel", data=xlsx,
                                   file_name="VaR_Report.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            st.warning("pip install openpyxl")

    with cp:
        st.markdown("""
        <div class='var-card'>
          <div class='var-card-title'>📄 Rapport PDF</div>
          <div style='font-size:12px;color:#E8EDF5;margin:8px 0;line-height:1.6'>
          Couverture · VaR · Backtesting · Graphiques intégrés<br>
          Mise en page exécutive, prêt à envoyer.
          </div>
        </div>""", unsafe_allow_html=True)
        if HAS_PDF:
            with st.spinner("Génération PDF…"):
                fv_ = fig_var_bar(var_res, a99, pv)
                fd_ = fig_dist(port_r, var_res)
                pdf = export_pdf(var_res, bt_fmt, pv, fv_, fd_)
            if pdf:
                st.download_button("⬇  Télécharger PDF", data=pdf,
                                   file_name="VaR_Risk_Report.pdf",
                                   mime="application/pdf",
                                   use_container_width=True)
        else:
            st.warning("pip install reportlab")

    with st.expander("📦 requirements.txt"):
        st.code("""streamlit>=1.32
yfinance>=0.2.36
pandas>=2.0
numpy>=1.26
scipy>=1.12
matplotlib>=3.8
openpyxl>=3.1
reportlab>=4.1""", language="text")
