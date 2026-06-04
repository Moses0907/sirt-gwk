# ============================================================
# SIRT GWK Prototype - Streamlit App
# Scope: Order-Kitchen/Bar Integration, Display Monitoring,
#        Financial Reporting, Dashboard Monitoring, User Access
# Notes: No inventory/stock management module.
# ============================================================

from __future__ import annotations

import base64
import hashlib
import io
import json
import secrets
import sqlite3
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# ============================================================
# App constants and paths
# ============================================================

APP_NAME = "SIRT GWK Prototype"
APP_SUBTITLE = "Sistem Informasi Restoran Terintegrasi untuk order, kitchen/bar display, laporan, dan dashboard monitoring."
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "sirt_gwk.db"

LOGO_FILES = ["logo.jpg", "logo.png", "logo.jpeg"]
LOGIN_BG_FILES = ["backgroundlogin.jpg", "backgroundlogin.jpeg", "backgroundlogin.png"]

COLOR = {
    # Palette disamakan dengan hasil desain Google AI Studio: forest green, sage, gold, surface, coffee.
    "forest": "#1B3022",
    "sage": "#889E81",
    "cream": "#FDFCF0",
    "cream2": "#FCFAF8",
    "terracotta": "#B86B4B",
    "gold": "#D4AF37",
    "coffee": "#2D3436",
    "dark": "#111827",
    "muted": "#6B7280",
    "danger": "#DC2626",
    "success": "#16A34A",
}

st.set_page_config(page_title="SIRT GWK", page_icon="🍃", layout="wide", initial_sidebar_state="expanded")

# ============================================================
# Utility functions
# ============================================================


def app_rerun() -> None:
    try:
        st.rerun()
    except Exception:
        try:
            st.experimental_rerun()
        except Exception:
            pass


def rupiah(value: Any) -> str:
    try:
        value = float(value or 0)
    except Exception:
        value = 0
    return "Rp{:,.0f}".format(value).replace(",", ".")


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def as_datetime_range(start_date: date, end_date: date) -> Tuple[str, str]:
    start = datetime.combine(start_date, time.min).strftime("%Y-%m-%d %H:%M:%S")
    end = datetime.combine(end_date, time.max).strftime("%Y-%m-%d %H:%M:%S")
    return start, end


def fmt_date(value: Any) -> str:
    """Display date as DD-MM-YYYY for all user-facing date labels."""
    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")
    try:
        return pd.to_datetime(value).strftime("%d-%m-%Y")
    except Exception:
        return str(value)


def date_input_id(label: str, value: date, key: Optional[str] = None) -> date:
    return st.date_input(label, value=value, key=key, format="DD-MM-YYYY")


def date_range_label(start: date, end: date) -> str:
    return f"{fmt_date(start)} - {fmt_date(end)}"


def find_asset(filenames: List[str]) -> Optional[Path]:
    candidates = []
    for filename in filenames:
        candidates.append(BASE_DIR / filename)
        candidates.append(BASE_DIR / "assets" / "images" / filename)
    for path in candidates:
        if path.exists():
            return path
    return None


def image_to_base64_src(path: Optional[Path]) -> Optional[str]:
    if path is None or not path.exists():
        return None
    suffix = path.suffix.lower().replace(".", "")
    mime = "jpeg" if suffix in ["jpg", "jpeg"] else "png"
    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:image/{mime};base64,{encoded}"


def safe_df(rows: List[sqlite3.Row]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame([dict(r) for r in rows])


def get_excel_engine() -> str:
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        return "openpyxl"

# ============================================================
# Styling
# ============================================================


def inject_global_css(login_mode: bool = False) -> None:
    """Global visual layer.

    This keeps the app in Streamlit/Python + SQLite, but makes the UI closer to
    the Google AI Studio mockup: Plus Jakarta Sans, dark executive sidebar,
    glass cards, green/gold palette, visible dark controls, and white receipt paper.
    """
    bg_src = image_to_base64_src(find_asset(LOGIN_BG_FILES))
    if bg_src:
        bg_css = f"""
            background-image:
                linear-gradient(135deg, rgba(253,252,240,0.76) 0%, rgba(255,254,250,0.82) 50%, rgba(245,242,237,0.86) 100%),
                url("{bg_src}") !important;
            background-size: cover !important;
            background-position: center center !important;
            background-attachment: fixed !important;
            background-repeat: no-repeat !important;
        """
    else:
        bg_css = "background: linear-gradient(135deg, #FDFCF0 0%, #FFFEFA 50%, #F5F2ED 100%) !important;"

    login_layout_css = """
        .block-container {
            max-width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        div[data-testid="stHorizontalBlock"] {
            gap: 0 !important;
        }
    """ if login_mode else ""

    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,0,0&display=swap');

        :root {{
            --gwk-primary: {COLOR['forest']};
            --gwk-sage: {COLOR['sage']};
            --gwk-gold: {COLOR['gold']};
            --gwk-surface: {COLOR['cream2']};
            --gwk-bg: {COLOR['cream']};
            --gwk-coffee: {COLOR['coffee']};
            --gwk-dark: {COLOR['dark']};
        }}

        html, body, [class*="css"] {{
            font-family: 'Plus Jakarta Sans', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
            color: var(--gwk-coffee) !important;
            scroll-behavior: smooth;
        }}

        .stApp {{ {bg_css} }}

        h1, h2, h3, h4, h5, h6 {{
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
            font-weight: 800 !important;
            letter-spacing: -0.035em !important;
            color: var(--gwk-coffee) !important;
        }}
        p, span, label, div {{
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
            color: var(--gwk-coffee);
        }}

        .block-container {{
            padding-top: 2.1rem !important;
            padding-bottom: 3rem !important;
        }}

        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #1B3022 0%, #122218 55%, #0B1610 100%) !important;
            border-right: 1px solid rgba(255,255,255,0.08) !important;
            box-shadow: 10px 0 28px rgba(0,0,0,0.20) !important;
        }}
        section[data-testid="stSidebar"] *,
        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] div {{
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
        }}
        section[data-testid="stSidebar"] [role="radiogroup"] label {{
            padding: 0.55rem 0.65rem !important;
            border-radius: 0.85rem !important;
            transition: all 0.18s ease !important;
            border-left: 4px solid transparent !important;
        }}
        section[data-testid="stSidebar"] [role="radiogroup"] label:hover {{
            background: rgba(255,255,255,0.06) !important;
            border-left-color: rgba(212,175,55,0.65) !important;
        }}
        section[data-testid="stSidebar"] [role="radiogroup"] label[data-checked="true"] {{
            background: rgba(255,255,255,0.11) !important;
            border-left-color: var(--gwk-gold) !important;
        }}
        section[data-testid="stSidebar"] hr {{
            border-color: rgba(255,255,255,0.10) !important;
        }}

        .hero {{
            background: linear-gradient(135deg, rgba(27,48,34,0.98) 0%, rgba(38,66,46,0.94) 54%, rgba(212,175,55,0.88) 100%);
            color: #FCFAF8 !important;
            padding: 28px 32px;
            border-radius: 28px;
            margin: 12px 0 20px 0;
            box-shadow: 0 24px 56px rgba(27, 48, 34, 0.22);
            border: 1px solid rgba(255,255,255,0.10);
        }}
        .hero h1, .hero p {{ color: #FCFAF8 !important; -webkit-text-fill-color: #FCFAF8 !important; }}
        .hero h1 {{ margin: 0 0 8px 0; font-size: 36px; }}
        .hero p {{ margin: 0; font-size: 14px; opacity: 0.92; font-weight: 600; }}

        .login-hero {{
            background: linear-gradient(135deg, rgba(27,48,34,0.96) 0%, rgba(38,66,46,0.90) 55%, rgba(212,175,55,0.88) 100%);
            color: #FCFAF8 !important;
            padding: 30px 34px;
            border-radius: 30px;
            margin: 20px 0 18px 0;
            box-shadow: 0 24px 58px rgba(27,48,34,0.24);
            display: flex;
            align-items: center;
            gap: 18px;
            border: 1px solid rgba(255,255,255,0.10);
        }}
        .login-hero img {{
            width: 92px; height: 92px; object-fit: contain;
            background: rgba(252,250,248,0.94); border-radius: 22px; padding: 8px;
            box-shadow: 0 10px 24px rgba(0,0,0,0.18);
        }}
        .login-hero h1 {{ font-size: 36px; margin: 0 0 8px 0; color: #FCFAF8 !important; -webkit-text-fill-color: #FCFAF8 !important; font-weight: 800 !important; }}
        .login-hero p {{ margin: 0; color: #FCFAF8 !important; -webkit-text-fill-color: #FCFAF8 !important; font-size: 14px; opacity: 0.92; font-weight: 600; }}

        .soft-card {{
            background: rgba(252,250,248,0.78);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid rgba(27,48,34,0.10);
            box-shadow: 0 18px 44px rgba(45,52,54,0.08);
            border-radius: 22px;
            padding: 18px 20px; margin: 12px 0;
        }}
        .mini-card {{
            background: rgba(252,250,248,0.82);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid rgba(27,48,34,0.10);
            border-left: 5px solid var(--gwk-primary);
            border-radius: 18px; padding: 16px 18px; margin: 10px 0;
            box-shadow: 0 10px 30px rgba(45,52,54,0.06);
        }}

        .receipt-box {{
            background: #FFFFFF !important;
            border: 1px dashed #94A3B8 !important;
            border-radius: 18px !important;
            padding: 20px !important;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace !important;
            white-space: pre-wrap;
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
            box-shadow: 0 12px 28px rgba(17,24,39,0.08);
        }}
        .receipt-box * {{ color: #111827 !important; -webkit-text-fill-color: #111827 !important; }}

        div[data-testid="stMetric"] {{
            background: rgba(252,250,248,0.82) !important;
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-radius: 18px !important;
            padding: 16px 18px !important;
            border: 1px solid rgba(27,48,34,0.10) !important;
            box-shadow: 0 14px 34px rgba(45,52,54,0.06) !important;
        }}
        div[data-testid="stMetric"] * {{ color: var(--gwk-coffee) !important; -webkit-text-fill-color: var(--gwk-coffee) !important; }}
        div[data-testid="stMetricDelta"] svg {{ fill: currentColor !important; }}

        div[data-testid="stDataFrame"] * {{ color: var(--gwk-coffee) !important; -webkit-text-fill-color: var(--gwk-coffee) !important; }}
        .stAlert * {{ color: var(--gwk-coffee) !important; -webkit-text-fill-color: var(--gwk-coffee) !important; }}

        input, textarea {{
            color: var(--gwk-coffee) !important;
            -webkit-text-fill-color: var(--gwk-coffee) !important;
            background-color: #FFFDF8 !important;
        }}
        input:focus, textarea:focus {{
            border-color: var(--gwk-primary) !important;
            box-shadow: 0 0 0 3px rgba(27,48,34,0.12) !important;
        }}

        .stButton > button,
        .stDownloadButton > button,
        button[kind="primary"],
        button[kind="secondary"],
        button[data-testid="baseButton-primary"],
        button[data-testid="baseButton-secondary"] {{
            background: #111827 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1px solid rgba(255,255,255,0.15) !important;
            border-radius: 12px !important;
            font-weight: 800 !important;
            box-shadow: 0 8px 18px rgba(17,24,39,0.14) !important;
        }}
        .stButton > button *, .stDownloadButton > button *,
        button[kind="primary"] *, button[kind="secondary"] *,
        button[data-testid="baseButton-primary"] *, button[data-testid="baseButton-secondary"] * {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            opacity: 1 !important;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover,
        button[kind="primary"]:hover, button[kind="secondary"]:hover {{
            background: var(--gwk-primary) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border-color: var(--gwk-primary) !important;
        }}
        .stButton > button:disabled, .stDownloadButton > button:disabled {{
            background-color: #6B7280 !important;
            color: #F9FAFB !important;
            -webkit-text-fill-color: #F9FAFB !important;
            opacity: 0.75 !important;
        }}
        .stButton > button:disabled *, .stDownloadButton > button:disabled * {{ color: #F9FAFB !important; -webkit-text-fill-color: #F9FAFB !important; }}

        div[data-baseweb="select"],
        div[data-baseweb="select"] > div {{
            background-color: #111827 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border-color: rgba(255,255,255,0.32) !important;
            border-radius: 12px !important;
        }}
        div[data-baseweb="select"] *,
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div,
        div[data-baseweb="select"] input {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}
        div[data-baseweb="select"] svg {{ color: #FFFFFF !important; fill: #FFFFFF !important; }}

        div[data-baseweb="popover"], div[data-baseweb="popover"] *,
        div[data-baseweb="menu"], div[data-baseweb="menu"] *,
        ul[role="listbox"], ul[role="listbox"] *,
        div[role="listbox"], div[role="listbox"] * {{
            background-color: #111827 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}
        li[role="option"], li[role="option"] *,
        div[role="option"], div[role="option"] * {{
            background-color: #111827 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}
        li[role="option"]:hover, li[role="option"]:hover *,
        div[role="option"]:hover, div[role="option"]:hover *,
        li[aria-selected="true"], li[aria-selected="true"] *,
        div[aria-selected="true"], div[aria-selected="true"] * {{
            background-color: var(--gwk-primary) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}

        div[data-baseweb="input"] {{ border-radius: 12px !important; }}
        div[data-baseweb="input"] input {{ color: var(--gwk-coffee) !important; -webkit-text-fill-color: var(--gwk-coffee) !important; }}


        {login_layout_css}

        /* Keep Streamlit Material icons rendered as icons, not literal text such as keyboard_double_arrow */
        .material-symbols-rounded,
        .material-symbols-outlined,
        .material-icons,
        span[class*="material-symbols"],
        span[class*="material-icons"] {{
            font-family: 'Material Symbols Rounded', 'Material Icons' !important;
            font-weight: normal !important;
            font-style: normal !important;
            line-height: 1 !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            direction: ltr !important;
            -webkit-font-feature-settings: 'liga' !important;
            font-feature-settings: 'liga' !important;
            -webkit-font-smoothing: antialiased !important;
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
        }}

        /* ===== Targeted updates: AI Studio-like login & POS ===== */
        .login-split-card {{
            min-height: 82vh;
            border-radius: 0px;
        }}
        .login-left-panel {{
            min-height: 100vh;
            background: radial-gradient(circle at 82% 10%, rgba(212,175,55,0.12) 0, rgba(212,175,55,0.04) 24%, transparent 25%),
                        radial-gradient(circle at 10% 92%, rgba(136,158,129,0.16) 0, rgba(136,158,129,0.06) 27%, transparent 28%),
                        linear-gradient(180deg, #1B3022 0%, #13271B 62%, #0C1911 100%);
            border-radius: 0px 0px 0px 0px;
            padding: 54px 44px;
            color: #FCFAF8 !important;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 24px 56px rgba(27,48,34,0.20);
        }}
        .login-left-panel * {{
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
        }}
        .login-brand-row {{
            display: flex;
            align-items: center;
            gap: 16px;
            margin-bottom: 30vh;
        }}
        .login-brand-row img {{
            width: 98px;
            height: 98px;
            object-fit: contain;
        }}
        .login-kicker {{
            font-size: 12px;
            letter-spacing: .12em;
            color: #D4AF37 !important;
            -webkit-text-fill-color: #D4AF37 !important;
            font-weight: 900;
            margin-bottom: 18px;
        }}
        .login-left-panel h1 {{
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            font-size: 42px;
            line-height: 1.02;
            max-width: 480px;
            margin: 0 0 22px 0;
            letter-spacing: -0.055em !important;
        }}
        .login-left-panel p {{
            color: rgba(252,250,248,0.84) !important;
            -webkit-text-fill-color: rgba(252,250,248,0.84) !important;
            font-size: 15px;
            line-height: 1.65;
            max-width: 520px;
        }}
        .login-proof {{
            border-top: 1px solid rgba(252,250,248,0.15);
            padding-top: 18px;
            margin-top: 26px;
            display: grid;
            gap: 12px;
            font-size: 13px;
            font-weight: 800;
        }}
        .login-proof span {{
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
        }}
        .login-footnote {{
            font-size: 11px;
            letter-spacing: .08em;
            color: rgba(252,250,248,0.56) !important;
            -webkit-text-fill-color: rgba(252,250,248,0.56) !important;
            margin-top: 20vh;
        }}
        .auth-card {{
            max-width: 520px;
            margin: 18vh auto 0 auto;
            background: rgba(255,255,255,0.76);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border-radius: 26px;
            border: 1px solid rgba(27,48,34,0.08);
            box-shadow: 0 28px 68px rgba(45,52,54,0.09);
            padding: 30px 32px 28px 32px;
        }}
        .auth-card h2 {{
            font-size: 28px;
            margin: 4px 0 8px 0;
        }}
        .auth-card .auth-subtitle {{
            color: #8A94A6 !important;
            -webkit-text-fill-color: #8A94A6 !important;
            font-size: 14px;
            line-height: 1.45;
            margin-bottom: 16px;
        }}

        /* Tabs: remove the white rounded pill/card behind tab labels across pages */
        div[data-testid="stTabs"] {{
            max-width: 560px;
            margin: 18vh auto 0 auto;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
            background: transparent !important;
            border: none !important;
            border-radius: 0 !important;
            padding: 0 !important;
            box-shadow: none !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
            gap: 12px !important;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab"] {{
            background: transparent !important;
            border-radius: 0 !important;
            box-shadow: none !important;
            font-weight: 750 !important;
            padding-left: 0 !important;
            padding-right: 0 !important;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab"] * {{
            color: var(--gwk-coffee) !important;
            -webkit-text-fill-color: var(--gwk-coffee) !important;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab-panel"] {{
            padding-top: 24px !important;
        }}
        .demo-chip {{
            display: inline-block;
            background: #F6F8FB;
            border: 1px solid #E6EAF0;
            color: #475569 !important;
            -webkit-text-fill-color: #475569 !important;
            font-size: 11px;
            font-weight: 800;
            padding: 7px 11px;
            border-radius: 999px;
            margin: 4px 6px 4px 0;
        }}
        .pos-top-filter {{
            background: rgba(255,255,255,0.78);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border: 1px solid rgba(27,48,34,0.09);
            box-shadow: 0 18px 48px rgba(45,52,54,0.07);
            border-radius: 18px;
            padding: 18px 20px;
            margin: 16px 0 22px 0;
        }}
        .pos-search-wrap {{
            background: rgba(255,255,255,0.82);
            border: 1px solid rgba(27,48,34,0.08);
            border-radius: 16px;
            padding: 14px 16px;
            margin-bottom: 14px;
        }}
        .pos-menu-card {{
            height: 170px;
            background: rgba(255,255,255,0.82);
            border: 1px solid rgba(27,48,34,0.08);
            box-shadow: 0 14px 30px rgba(45,52,54,0.055);
            border-radius: 16px;
            padding: 16px 16px 12px 16px;
            margin-bottom: 10px;
            overflow: hidden;
        }}
        .pos-menu-badge {{
            display: inline-block;
            padding: 4px 9px;
            border-radius: 999px;
            background: #F1F3F0;
            font-size: 10px;
            font-weight: 900;
            color: #6B7280 !important;
            -webkit-text-fill-color: #6B7280 !important;
            letter-spacing: .06em;
            text-transform: uppercase;
            margin-bottom: 9px;
        }}
        .pos-menu-title {{
            font-size: 14px;
            line-height: 1.2;
            font-weight: 900;
            color: #1F2937 !important;
            -webkit-text-fill-color: #1F2937 !important;
            margin-bottom: 6px;
        }}
        .pos-menu-desc {{
            font-size: 11px;
            line-height: 1.35;
            color: #94A3B8 !important;
            -webkit-text-fill-color: #94A3B8 !important;
            height: 42px;
            overflow: hidden;
        }}
        .pos-menu-price {{
            font-size: 14px;
            font-weight: 900;
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
            margin-top: 18px;
        }}
        .cart-panel {{
            min-height: 470px;
            background: rgba(255,255,255,0.84);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
            border: 1px solid rgba(27,48,34,0.08);
            border-radius: 18px;
            box-shadow: 0 22px 54px rgba(45,52,54,0.08);
            padding: 18px 18px 20px 18px;
        }}
        .cart-empty {{
            text-align: center;
            padding: 110px 10px;
            color: #94A3B8 !important;
            -webkit-text-fill-color: #94A3B8 !important;
            font-size: 13px;
        }}
        .cart-line {{
            border-bottom: 1px solid rgba(27,48,34,0.08);
            padding: 10px 0;
            font-size: 13px;
        }}
        .cart-subtotal {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 18px;
            padding-top: 14px;
            border-top: 1px solid rgba(27,48,34,0.10);
            font-weight: 900;
        }}
        /* Make dashboard plot panels look closer to app studio cards */
        div[data-testid="stPlotlyChart"] {{
            background: rgba(255,255,255,0.78);
            border: 1px solid rgba(27,48,34,0.08);
            border-radius: 18px;
            box-shadow: 0 16px 42px rgba(45,52,54,0.06);
            padding: 10px;
        }}

        /* Streamlit sidebar collapse / double-arrow control: prevent Material icon names from appearing as literal text */
        button[data-testid="stSidebarCollapseButton"],
        button[title="Collapse sidebar"],
        button[title="Close sidebar"],
        button[title="Hide sidebar"],
        button[aria-label="Collapse sidebar"],
        button[aria-label="Close sidebar"],
        button[aria-label="Hide sidebar"] {{
            width: 34px !important;
            height: 34px !important;
            min-width: 34px !important;
            overflow: hidden !important;
            border-radius: 10px !important;
            background: rgba(255,255,255,0.08) !important;
        }}
        button[data-testid="stSidebarCollapseButton"] *,
        button[title="Collapse sidebar"] *,
        button[title="Close sidebar"] *,
        button[title="Hide sidebar"] *,
        button[aria-label="Collapse sidebar"] *,
        button[aria-label="Close sidebar"] *,
        button[aria-label="Hide sidebar"] * {{
            font-size: 0 !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
        }}
        button[data-testid="stSidebarCollapseButton"]::before,
        button[title="Collapse sidebar"]::before,
        button[title="Close sidebar"]::before,
        button[title="Hide sidebar"]::before,
        button[aria-label="Collapse sidebar"]::before,
        button[aria-label="Close sidebar"]::before,
        button[aria-label="Hide sidebar"]::before {{
            content: "‹‹";
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            height: 100%;
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
            font-size: 16px !important;
            font-weight: 900 !important;
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            line-height: 1 !important;
        }}

        /* Give bordered operational cards a real white surface, including POS cart and kitchen cards */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: rgba(255,255,255,0.88) !important;
            border: 1px solid rgba(27,48,34,0.16) !important;
            border-radius: 18px !important;
            box-shadow: 0 18px 44px rgba(45,52,54,0.08) !important;
            padding: 2px !important;
        }}
        div[data-testid="stVerticalBlockBorderWrapper"] > div {{
            background: transparent !important;
        }}

        /* POS cart spacing: avoid cramped/stacked buttons */
        .cart-line {{
            padding: 12px 0 14px 0 !important;
            margin-bottom: 4px !important;
        }}
        .cart-subtotal {{
            margin-top: 20px !important;
            margin-bottom: 14px !important;
            padding-top: 18px !important;
        }}
        .cart-subtotal span:last-child {{
            font-size: 20px !important;
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
        }}



        /* ===== FINAL FIX: replace sidebar collapse/expand Material text with << / >> ===== */
        section[data-testid="stSidebar"] button:has(span[class*="material-symbols"]),
        section[data-testid="stSidebar"] button:has(span[class*="material-icons"]),
        section[data-testid="stSidebar"] button:has([class*="material-symbols"]),
        section[data-testid="stSidebar"] button:has([class*="material-icons"]) {{
            position: relative !important;
            width: 34px !important;
            height: 34px !important;
            min-width: 34px !important;
            overflow: hidden !important;
            border-radius: 10px !important;
            background: rgba(255,255,255,0.08) !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
        }}
        section[data-testid="stSidebar"] button:has(span[class*="material-symbols"]) *,
        section[data-testid="stSidebar"] button:has(span[class*="material-icons"]) *,
        section[data-testid="stSidebar"] button:has([class*="material-symbols"]) *,
        section[data-testid="stSidebar"] button:has([class*="material-icons"]) * {{
            font-size: 0 !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
            opacity: 0 !important;
        }}
        section[data-testid="stSidebar"] button:has(span[class*="material-symbols"])::after,
        section[data-testid="stSidebar"] button:has(span[class*="material-icons"])::after,
        section[data-testid="stSidebar"] button:has([class*="material-symbols"])::after,
        section[data-testid="stSidebar"] button:has([class*="material-icons"])::after {{
            content: "<<" !important;
            position: absolute !important;
            inset: 0 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
            font-size: 16px !important;
            font-weight: 900 !important;
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            opacity: 1 !important;
            line-height: 1 !important;
        }}

        button[title*="Expand"],
        button[aria-label*="Expand"],
        button[data-testid*="collapsed"],
        button[data-testid*="Collapsed"] {{
            position: relative !important;
            width: 34px !important;
            height: 34px !important;
            min-width: 34px !important;
            overflow: hidden !important;
            border-radius: 10px !important;
            background: rgba(27,48,34,0.92) !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
        }}
        button[title*="Expand"] *,
        button[aria-label*="Expand"] *,
        button[data-testid*="collapsed"] *,
        button[data-testid*="Collapsed"] * {{
            font-size: 0 !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
            opacity: 0 !important;
        }}
        button[title*="Expand"]::after,
        button[aria-label*="Expand"]::after,
        button[data-testid*="collapsed"]::after,
        button[data-testid*="Collapsed"]::after {{
            content: ">>" !important;
            position: absolute !important;
            inset: 0 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
            font-size: 16px !important;
            font-weight: 900 !important;
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            opacity: 1 !important;
            line-height: 1 !important;
        }}


        /* ===== ULTRA FINAL SIDEBAR ARROW FIX: use plain << / >> and hide Material ligature text ===== */
        span[data-testid="stIconMaterial"] {{
            font-family: "Material Symbols Rounded", "Material Symbols Outlined", "Material Icons" !important;
            font-weight: normal !important;
            font-style: normal !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            direction: ltr !important;
            -webkit-font-feature-settings: "liga" !important;
            -webkit-font-smoothing: antialiased !important;
        }}

        /* Collapse button inside the expanded sidebar header */
        section[data-testid="stSidebar"] div[data-testid="stSidebarHeader"] button,
        section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"],
        section[data-testid="stSidebar"] button[aria-label*="Collapse"],
        section[data-testid="stSidebar"] button[title*="Collapse"] {{
            position: relative !important;
            width: 38px !important;
            height: 38px !important;
            min-width: 38px !important;
            max-width: 38px !important;
            overflow: hidden !important;
            border-radius: 10px !important;
            background: rgba(255,255,255,0.10) !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
        }}

        section[data-testid="stSidebar"] div[data-testid="stSidebarHeader"] button *,
        section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"] *,
        section[data-testid="stSidebar"] button[aria-label*="Collapse"] *,
        section[data-testid="stSidebar"] button[title*="Collapse"] * {{
            font-size: 0 !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
            opacity: 0 !important;
            max-width: 0 !important;
            overflow: hidden !important;
        }}

        section[data-testid="stSidebar"] div[data-testid="stSidebarHeader"] button::after,
        section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"]::after,
        section[data-testid="stSidebar"] button[aria-label*="Collapse"]::after,
        section[data-testid="stSidebar"] button[title*="Collapse"]::after {{
            content: "<<" !important;
            position: absolute !important;
            inset: 0 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-family: 'Plus Jakarta Sans', Arial, sans-serif !important;
            font-size: 15px !important;
            font-weight: 900 !important;
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            opacity: 1 !important;
            line-height: 1 !important;
        }}

        /* Expand button when sidebar is collapsed */
        button[data-testid="collapsedControl"],
        button[data-testid="stSidebarCollapsedControl"],
        button[aria-label*="Expand"],
        button[title*="Expand"] {{
            position: relative !important;
            width: 38px !important;
            height: 38px !important;
            min-width: 38px !important;
            max-width: 38px !important;
            overflow: hidden !important;
            border-radius: 10px !important;
            background: rgba(27,48,34,0.94) !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
        }}

        button[data-testid="collapsedControl"] *,
        button[data-testid="stSidebarCollapsedControl"] *,
        button[aria-label*="Expand"] *,
        button[title*="Expand"] * {{
            font-size: 0 !important;
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
            opacity: 0 !important;
            max-width: 0 !important;
            overflow: hidden !important;
        }}

        button[data-testid="collapsedControl"]::after,
        button[data-testid="stSidebarCollapsedControl"]::after,
        button[aria-label*="Expand"]::after,
        button[title*="Expand"]::after {{
            content: ">>" !important;
            position: absolute !important;
            inset: 0 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-family: 'Plus Jakarta Sans', Arial, sans-serif !important;
            font-size: 15px !important;
            font-weight: 900 !important;
            color: #FCFAF8 !important;
            -webkit-text-fill-color: #FCFAF8 !important;
            opacity: 1 !important;
            line-height: 1 !important;
        }}

        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="hero">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def login_hero(title: str, subtitle: str) -> None:
    logo_src = image_to_base64_src(find_asset(LOGO_FILES))
    logo_html = f'<img src="{logo_src}" alt="Logo GWK">' if logo_src else ""
    st.markdown(
        f"""
        <div class="login-hero">
            {logo_html}
            <div><h1>{title}</h1><p>{subtitle}</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# Database layer
# ============================================================


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def exec_sql(sql: str, params: Tuple[Any, ...] = ()) -> None:
    with get_conn() as conn:
        conn.execute(sql, params)
        conn.commit()


def query_rows(sql: str, params: Tuple[Any, ...] = ()) -> List[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(sql, params).fetchall()


def query_one(sql: str, params: Tuple[Any, ...] = ()) -> Optional[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(sql, params).fetchone()


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    return f"sha256${salt}${digest}"


def verify_password(password: str, stored: str) -> bool:
    if stored.startswith("sha256$"):
        try:
            _, salt, digest = stored.split("$", 2)
            candidate = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
            return secrets.compare_digest(candidate, digest)
        except Exception:
            return False
    return stored == password


def init_db() -> None:
    with get_conn() as conn:
        cur = conn.cursor()
        cur.executescript(
            """
            CREATE TABLE IF NOT EXISTS roles (role_id INTEGER PRIMARY KEY AUTOINCREMENT, nama_role TEXT UNIQUE NOT NULL, deskripsi_role TEXT);
            CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY AUTOINCREMENT, nama_user TEXT NOT NULL, username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, role TEXT NOT NULL, status_aktif TEXT NOT NULL DEFAULT 'Aktif', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
            CREATE TABLE IF NOT EXISTS menu_categories (category_id INTEGER PRIMARY KEY AUTOINCREMENT, nama_kategori TEXT UNIQUE NOT NULL, deskripsi TEXT);
            CREATE TABLE IF NOT EXISTS menus (menu_id INTEGER PRIMARY KEY AUTOINCREMENT, category_id INTEGER, nama_menu TEXT UNIQUE NOT NULL, harga INTEGER NOT NULL, tipe_menu TEXT NOT NULL, bagian_produksi TEXT NOT NULL, status_menu TEXT NOT NULL DEFAULT 'Tersedia', deskripsi TEXT, FOREIGN KEY(category_id) REFERENCES menu_categories(category_id));
            CREATE TABLE IF NOT EXISTS tables (table_id INTEGER PRIMARY KEY AUTOINCREMENT, nomor_meja TEXT UNIQUE NOT NULL, area TEXT, status_meja TEXT NOT NULL DEFAULT 'Tersedia');
            CREATE TABLE IF NOT EXISTS orders (order_id INTEGER PRIMARY KEY AUTOINCREMENT, table_id INTEGER, user_id INTEGER, tanggal_order TEXT NOT NULL, tipe_layanan TEXT NOT NULL, status_order TEXT NOT NULL DEFAULT 'Submitted', customer_note TEXT, FOREIGN KEY(table_id) REFERENCES tables(table_id), FOREIGN KEY(user_id) REFERENCES users(user_id));
            CREATE TABLE IF NOT EXISTS order_details (detail_id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL, menu_id INTEGER NOT NULL, jumlah INTEGER NOT NULL, harga_satuan INTEGER NOT NULL, subtotal INTEGER NOT NULL, catatan TEXT, FOREIGN KEY(order_id) REFERENCES orders(order_id), FOREIGN KEY(menu_id) REFERENCES menus(menu_id));
            CREATE TABLE IF NOT EXISTS production_queue (queue_id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL, detail_id INTEGER NOT NULL, bagian_tujuan TEXT NOT NULL, status_produksi TEXT NOT NULL DEFAULT 'Menunggu', waktu_mulai TEXT, waktu_selesai TEXT, FOREIGN KEY(order_id) REFERENCES orders(order_id), FOREIGN KEY(detail_id) REFERENCES order_details(detail_id));
            CREATE TABLE IF NOT EXISTS payments (payment_id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL, tanggal_bayar TEXT NOT NULL, total_tagihan INTEGER NOT NULL, diskon INTEGER NOT NULL DEFAULT 0, metode_pembayaran TEXT NOT NULL, jumlah_bayar INTEGER NOT NULL, kembalian INTEGER NOT NULL DEFAULT 0, status_bayar TEXT NOT NULL DEFAULT 'Lunas', cashier_user_id INTEGER, FOREIGN KEY(order_id) REFERENCES orders(order_id), FOREIGN KEY(cashier_user_id) REFERENCES users(user_id));
            CREATE TABLE IF NOT EXISTS expenses (expense_id INTEGER PRIMARY KEY AUTOINCREMENT, tanggal TEXT NOT NULL, kategori_biaya TEXT NOT NULL, nominal INTEGER NOT NULL, metode_pembayaran TEXT NOT NULL, keterangan TEXT, user_id INTEGER, FOREIGN KEY(user_id) REFERENCES users(user_id));
            CREATE TABLE IF NOT EXISTS daily_closing (closing_id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, tanggal TEXT NOT NULL, shift TEXT NOT NULL, total_penjualan INTEGER NOT NULL, total_cash_sistem INTEGER NOT NULL, total_non_cash_sistem INTEGER NOT NULL, total_cash_fisik INTEGER NOT NULL, total_non_cash_fisik INTEGER NOT NULL, total_pengeluaran INTEGER NOT NULL, selisih INTEGER NOT NULL, catatan TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(user_id) REFERENCES users(user_id));
            CREATE TABLE IF NOT EXISTS activity_logs (log_id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, aktivitas TEXT NOT NULL, waktu_aktivitas TEXT NOT NULL, detail_perubahan TEXT, FOREIGN KEY(user_id) REFERENCES users(user_id));
            CREATE TABLE IF NOT EXISTS analytics_snapshots (snapshot_id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, nama_snapshot TEXT NOT NULL, tanggal_simpan TEXT NOT NULL, periode_awal TEXT, periode_akhir TEXT, ringkasan_json TEXT, FOREIGN KEY(user_id) REFERENCES users(user_id));
            """
        )
        conn.commit()
    seed_data()


def log_activity(user_id: Optional[int], aktivitas: str, detail: str = "") -> None:
    exec_sql("INSERT INTO activity_logs(user_id, aktivitas, waktu_aktivitas, detail_perubahan) VALUES (?, ?, ?, ?)", (user_id, aktivitas, now_str(), detail))


def seed_data() -> None:
    roles = [
        ("Owner", "Mengakses dashboard, laporan, comparison, backup, dan manajemen user."),
        ("Manager", "Mengakses dashboard, laporan, comparison, konfigurasi menu, dan user."),
        ("Kasir", "Mengelola pembayaran, pengeluaran, closing, dan laporan transaksi."),
        ("Waiter", "Mencatat pesanan pelanggan."),
        ("Kitchen/Bar", "Melihat dan memperbarui status pesanan kitchen/bar."),
    ]
    for role, desc in roles:
        exec_sql("INSERT OR IGNORE INTO roles(nama_role, deskripsi_role) VALUES (?, ?)", (role, desc))
    for i in range(1, 21):
        area = "Gazebo" if i <= 8 else ("Indoor" if i <= 14 else "Outdoor")
        exec_sql("INSERT OR IGNORE INTO tables(nomor_meja, area, status_meja) VALUES (?, ?, 'Tersedia')", (f"Meja {i}", area))
    categories = {
        "Makanan Utama": "Menu makanan utama ala carte GWK.", "Cemilan": "Menu cemilan dan snack ala carte.",
        "Minuman": "Minuman non-kopi dan minuman segar.", "Wedangan": "Minuman tradisional hangat.",
        "Kopi": "Minuman kopi dan latte.", "Juice": "Menu jus buah.", "Paket Prasmanan": "Paket prasmanan per pax.",
        "Paket Berkat": "Paket berkat untuk 5 orang.", "Snack Prasmanan": "Snack menu prasmanan.",
    }
    for cat, desc in categories.items():
        exec_sql("INSERT OR IGNORE INTO menu_categories(nama_kategori, deskripsi) VALUES (?, ?)", (cat, desc))
    cat_map = {r["nama_kategori"]: r["category_id"] for r in query_rows("SELECT * FROM menu_categories")}
    menus = [
        ("Sate Gebreg GWK", "Makanan Utama", 29000, "Ala Carte", "Kitchen", "Sate daging sapi jadul dengan kuah rempah khas."),
        ("Sop Iga Sapi", "Makanan Utama", 35000, "Ala Carte", "Kitchen", "Sop iga sapi dengan kuah rempah, kentang, wortel."),
        ("Soto Daging Sapi", "Makanan Utama", 18000, "Ala Carte", "Kitchen", "Soto daging sapi kuah bening rempah."),
        ("Pepes Ikan Patin", "Makanan Utama", 22000, "Ala Carte", "Kitchen", "Pepes ikan dengan bumbu kuning."),
        ("Pepes Ikan Nila", "Makanan Utama", 22000, "Ala Carte", "Kitchen", "Pepes ikan nila bumbu kuning."),
        ("Sop Ayam", "Makanan Utama", 23000, "Ala Carte", "Kitchen", "Sop ayam kampung dengan sayuran."),
        ("Kupat Tahu", "Makanan Utama", 17000, "Ala Carte", "Kitchen", "Ketupat, tahu, kol, tauge, kuah kacang."),
        ("Kerang Hijau", "Makanan Utama", 23000, "Ala Carte", "Kitchen", "Kerang hijau dengan varian saus."),
        ("Nasi Goreng Teri", "Makanan Utama", 24000, "Ala Carte", "Kitchen", "Nasi goreng ikan teri."),
        ("Nasi Goreng Ayam", "Makanan Utama", 24000, "Ala Carte", "Kitchen", "Nasi goreng ayam suwir."),
        ("Nasi Goreng Sapi GWK", "Makanan Utama", 29000, "Ala Carte", "Kitchen", "Nasi goreng daging sapi dengan sate gebreg."),
        ("Bakmi Godhog", "Makanan Utama", 25000, "Ala Carte", "Kitchen", "Bakmi Jawa rebus."),
        ("Bakmi Goreng", "Makanan Utama", 25000, "Ala Carte", "Kitchen", "Bakmi Jawa goreng."),
        ("Iga Bakar", "Makanan Utama", 37000, "Ala Carte", "Kitchen", "Iga sapi bakar bumbu khas."),
        ("Ayam Kampung Goreng Kremes", "Makanan Utama", 25000, "Ala Carte", "Kitchen", "Ayam kampung goreng kremes."),
        ("Ayam Kampung Bakar", "Makanan Utama", 25000, "Ala Carte", "Kitchen", "Ayam kampung bakar."),
        ("Nila Goreng Kremes", "Makanan Utama", 21000, "Ala Carte", "Kitchen", "Nila goreng kremes."),
        ("Nila Bakar", "Makanan Utama", 21000, "Ala Carte", "Kitchen", "Nila bakar."),
        ("Lele Goreng Kremes", "Makanan Utama", 17000, "Ala Carte", "Kitchen", "Lele goreng kremes."),
        ("Lele Bakar", "Makanan Utama", 17000, "Ala Carte", "Kitchen", "Lele bakar."),
        ("Nasi Putih", "Makanan Utama", 6000, "Ala Carte", "Kitchen", "Nasi putih."), ("Lontong", "Makanan Utama", 6000, "Ala Carte", "Kitchen", "Lontong."),
        ("Bola Tape Goreng", "Cemilan", 15000, "Ala Carte", "Kitchen", "Bola tape goreng."), ("Tahu Cabe Garam", "Cemilan", 17500, "Ala Carte", "Kitchen", "Tahu cabe garam."),
        ("Jasuke", "Cemilan", 13500, "Ala Carte", "Kitchen", "Jagung susu keju."), ("Tahu Tuna", "Cemilan", 18000, "Ala Carte", "Kitchen", "Tahu tuna."),
        ("Jamur Crispy", "Cemilan", 17000, "Ala Carte", "Kitchen", "Jamur crispy."), ("Pisang Goreng Coklat/Keju", "Cemilan", 17500, "Ala Carte", "Kitchen", "Pisang goreng topping."),
        ("Tempe Mendoan", "Cemilan", 14500, "Ala Carte", "Kitchen", "Tempe mendoan."), ("Tempe Garit", "Cemilan", 12000, "Ala Carte", "Kitchen", "Tempe garit."),
        ("Tahu Kemul", "Cemilan", 15000, "Ala Carte", "Kitchen", "Tahu kemul."), ("Pisang Goreng", "Cemilan", 16000, "Ala Carte", "Kitchen", "Pisang goreng."),
        ("Pisang Bakar Coklat/Keju", "Cemilan", 18000, "Ala Carte", "Kitchen", "Pisang bakar topping."), ("Singkong Goreng/Kukus", "Cemilan", 17000, "Ala Carte", "Kitchen", "Singkong goreng/kukus."),
        ("Singkong Keju", "Cemilan", 18000, "Ala Carte", "Kitchen", "Singkong keju."), ("French Fries", "Cemilan", 17000, "Ala Carte", "Kitchen", "Kentang goreng."), ("Onion Ring", "Cemilan", 18000, "Ala Carte", "Kitchen", "Onion ring."),
        ("Ice/Hot Lychee Tea", "Minuman", 19000, "Ala Carte", "Bar", "Lychee tea panas/dingin."), ("Es Timun Serut", "Minuman", 16000, "Ala Carte", "Bar", "Timun serut."),
        ("Es Cendol", "Minuman", 16000, "Ala Carte", "Bar", "Cendol."), ("Es Kelapa Jeruk", "Minuman", 18000, "Ala Carte", "Bar", "Kelapa jeruk."),
        ("Es Alpukat Coklat", "Minuman", 18000, "Ala Carte", "Bar", "Alpukat coklat."), ("Milky Choco Dino", "Minuman", 18000, "Ala Carte", "Bar", "Milky choco."),
        ("Es Teler", "Minuman", 23000, "Ala Carte", "Bar", "Es teler."), ("Es Kelapa Batok", "Minuman", 22000, "Ala Carte", "Bar", "Es kelapa batok."),
        ("Jeniper Es/Panas", "Minuman", 8500, "Ala Carte", "Bar", "Jeniper."), ("Jeruk Es/Panas", "Minuman", 8500, "Ala Carte", "Bar", "Jeruk."), ("Teh Es/Panas", "Minuman", 7000, "Ala Carte", "Bar", "Teh."),
        ("Lemon Tea", "Minuman", 13000, "Ala Carte", "Bar", "Lemon tea."), ("Teh Tarik Susu", "Minuman", 18000, "Ala Carte", "Bar", "Teh tarik."), ("Fresh Milk", "Minuman", 12000, "Ala Carte", "Bar", "Fresh milk."),
        ("Hot/Ice Chocolate", "Minuman", 19000, "Ala Carte", "Bar", "Chocolate."), ("Matcha Latte", "Minuman", 18000, "Ala Carte", "Bar", "Matcha."), ("Red Velvet Latte", "Minuman", 17000, "Ala Carte", "Bar", "Red velvet."),
        ("Milky Taro", "Minuman", 17000, "Ala Carte", "Bar", "Taro."), ("Es Cappucino Cincau", "Minuman", 17500, "Ala Carte", "Bar", "Cappucino cincau."), ("Es Kelapa Gelas", "Minuman", 15000, "Ala Carte", "Bar", "Kelapa gelas."),
        ("Wedang Jahe", "Wedangan", 13000, "Ala Carte", "Bar", "Wedang jahe."), ("Wedang Jahe Sereh", "Wedangan", 16000, "Ala Carte", "Bar", "Wedang jahe sereh."), ("Teh Jahe", "Wedangan", 14000, "Ala Carte", "Bar", "Teh jahe."),
        ("Wedang Uwuh Susu", "Wedangan", 18000, "Ala Carte", "Bar", "Wedang uwuh susu."), ("Wedang Seruni", "Wedangan", 15000, "Ala Carte", "Bar", "Wedang seruni."), ("Wedang Uwuh", "Wedangan", 16000, "Ala Carte", "Bar", "Wedang uwuh."),
        ("Wedang Jahe Sereh Susu", "Wedangan", 18000, "Ala Carte", "Bar", "Wedang jahe sereh susu."), ("Teh Poci", "Wedangan", 15000, "Ala Carte", "Bar", "Teh poci."),
        ("Hazelnut Latte", "Kopi", 19000, "Ala Carte", "Bar", "Hazelnut latte."), ("Pandan Latte", "Kopi", 18000, "Ala Carte", "Bar", "Pandan latte."), ("Kopi Susu Aren", "Kopi", 17000, "Ala Carte", "Bar", "Kopi susu aren."),
        ("Kopi Keraton", "Kopi", 16000, "Ala Carte", "Bar", "Kopi keraton."), ("Vietnam Drip", "Kopi", 16000, "Ala Carte", "Bar", "Vietnam drip."), ("Kopi Tubruk", "Kopi", 10000, "Ala Carte", "Bar", "Kopi tubruk."),
        ("Kopi Legend", "Kopi", 20000, "Ala Carte", "Bar", "Kopi legend."), ("Kopi Choco Energy", "Kopi", 18000, "Ala Carte", "Bar", "Kopi choco."), ("Es Kopi Susu Alpukat", "Kopi", 18000, "Ala Carte", "Bar", "Kopi alpukat."),
        ("Americano", "Kopi", 14000, "Ala Carte", "Bar", "Americano."), ("Kopi Jahe", "Kopi", 17000, "Ala Carte", "Bar", "Kopi jahe."), ("Kopi Jahe Susu", "Kopi", 19000, "Ala Carte", "Bar", "Kopi jahe susu."),
        ("Kopi Keraton Susu", "Kopi", 17000, "Ala Carte", "Bar", "Kopi keraton susu."), ("Choco Coffee", "Kopi", 18000, "Ala Carte", "Bar", "Choco coffee."), ("Coffee Latte", "Kopi", 18000, "Ala Carte", "Bar", "Coffee latte."),
        ("Jus Jambu", "Juice", 15000, "Ala Carte", "Bar", "Jus jambu."), ("Jeruk Sirsak", "Juice", 15000, "Ala Carte", "Bar", "Jeruk sirsak."), ("Jus Jeruk", "Juice", 15000, "Ala Carte", "Bar", "Jus jeruk."),
        ("Jus Semangka", "Juice", 15000, "Ala Carte", "Bar", "Jus semangka."), ("Jus Fiber", "Juice", 15000, "Ala Carte", "Bar", "Jus fiber."), ("Jus Buah Naga", "Juice", 16000, "Ala Carte", "Bar", "Jus buah naga."),
        ("Jus Mangga", "Juice", 18000, "Ala Carte", "Bar", "Jus mangga."), ("Jus Alpukat", "Juice", 18000, "Ala Carte", "Bar", "Jus alpukat."),
        ("Paket Prasmanan 35", "Paket Prasmanan", 35000, "Prasmanan", "Kitchen", "Paket prasmanan 35 per pax."), ("Paket Prasmanan 40", "Paket Prasmanan", 40000, "Prasmanan", "Kitchen", "Paket prasmanan 40 per pax."),
        ("Paket Prasmanan 45", "Paket Prasmanan", 45000, "Prasmanan", "Kitchen", "Paket prasmanan 45 per pax."), ("Paket Prasmanan 50", "Paket Prasmanan", 50000, "Prasmanan", "Kitchen", "Paket prasmanan 50 per pax."),
        ("Paket Prasmanan 55", "Paket Prasmanan", 55000, "Prasmanan", "Kitchen", "Paket prasmanan 55 per pax."), ("Paket Prasmanan 60", "Paket Prasmanan", 60000, "Prasmanan", "Kitchen", "Paket prasmanan 60 per pax."),
        ("Paket Berkat A (5 orang)", "Paket Berkat", 255000, "Prasmanan", "Kitchen", "Paket berkat A untuk 5 orang."), ("Paket Berkat B (5 orang)", "Paket Berkat", 276000, "Prasmanan", "Kitchen", "Paket berkat B untuk 5 orang."),
        ("Paket Liwet Special (5 orang)", "Paket Berkat", 297000, "Prasmanan", "Kitchen", "Paket liwet special."),
        ("Klepon", "Snack Prasmanan", 3000, "Prasmanan", "Kitchen", "Snack."), ("Kue Putu Ayu", "Snack Prasmanan", 2500, "Prasmanan", "Kitchen", "Snack."), ("Kue Sus Vla", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."),
        ("Dadar Gulung Enten", "Snack Prasmanan", 2500, "Prasmanan", "Kitchen", "Snack."), ("Nagasari", "Snack Prasmanan", 3000, "Prasmanan", "Kitchen", "Snack."), ("Cake Mandarin", "Snack Prasmanan", 4000, "Prasmanan", "Kitchen", "Snack."),
        ("Pie Buah", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."), ("Lumpia", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."), ("Martabak Mini", "Snack Prasmanan", 4000, "Prasmanan", "Kitchen", "Snack."),
        ("Pastel Ayam", "Snack Prasmanan", 4000, "Prasmanan", "Kitchen", "Snack."), ("Pastel Sayur Telur", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."), ("Arem-arem", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."),
        ("Lemper Isi Ayam", "Snack Prasmanan", 4000, "Prasmanan", "Kitchen", "Snack."), ("Tahu Bakso", "Snack Prasmanan", 3500, "Prasmanan", "Kitchen", "Snack."), ("Risoles Sayur", "Snack Prasmanan", 3000, "Prasmanan", "Kitchen", "Snack."),
        ("Risoles Mayo", "Snack Prasmanan", 4000, "Prasmanan", "Kitchen", "Snack."), ("Macaroni Gurih", "Snack Prasmanan", 1500, "Prasmanan", "Kitchen", "Kletikan."), ("Cheese Stick", "Snack Prasmanan", 2500, "Prasmanan", "Kitchen", "Kletikan."),
    ]
    for nama, cat, harga, tipe, bagian, desc in menus:
        exec_sql("""
            INSERT OR IGNORE INTO menus(category_id, nama_menu, harga, tipe_menu, bagian_produksi, status_menu, deskripsi)
            VALUES (?, ?, ?, ?, ?, 'Tersedia', ?)
        """, (cat_map[cat], nama, harga, tipe, bagian, desc))

# ============================================================
# Auth and permissions
# ============================================================


def get_current_user() -> Optional[Dict[str, Any]]:
    return st.session_state.get("user")


def has_role(*roles: str) -> bool:
    user = get_current_user()
    return bool(user and user.get("role") in roles)


def role_pages(role: str) -> List[str]:
    # Payment Page intentionally removed from navigation.
    # Current scope focuses on replacing manual carbon-copy order flow with Kitchen/Bar Display.
    all_pages = ["Dashboard Monitoring", "Order Page", "Kitchen/Bar Display", "Pengeluaran & Closing", "Reports", "Perbandingan Performa", "Konfigurasi Menu & Hak Akses", "Data & Backup"]
    allowed = {
        "Owner": all_pages,
        "Manager": all_pages,
        "Kasir": ["Order Page", "Kitchen/Bar Display", "Pengeluaran & Closing", "Reports", "Dashboard Monitoring"],
        "Waiter": ["Order Page", "Kitchen/Bar Display"],
        "Kitchen/Bar": ["Kitchen/Bar Display"],
    }
    return allowed.get(role, [])

# ============================================================
# Data access helpers
# ============================================================


def get_menu_df(active_only: bool = True) -> pd.DataFrame:
    sql = """
        SELECT m.menu_id, c.nama_kategori, m.nama_menu, m.harga, m.tipe_menu,
               m.bagian_produksi, m.status_menu, m.deskripsi
        FROM menus m LEFT JOIN menu_categories c ON c.category_id = m.category_id
    """
    if active_only:
        sql += " WHERE m.status_menu='Tersedia' "
    sql += " ORDER BY c.nama_kategori, m.nama_menu"
    return safe_df(query_rows(sql))


def get_unpaid_orders() -> pd.DataFrame:
    rows = query_rows("""
        SELECT o.order_id, o.tanggal_order, o.tipe_layanan, o.status_order, t.nomor_meja,
               u.nama_user AS staff, COALESCE(SUM(od.subtotal),0) AS total
        FROM orders o
        LEFT JOIN tables t ON t.table_id=o.table_id
        LEFT JOIN users u ON u.user_id=o.user_id
        LEFT JOIN order_details od ON od.order_id=o.order_id
        LEFT JOIN payments p ON p.order_id=o.order_id
        WHERE p.payment_id IS NULL
        GROUP BY o.order_id
        ORDER BY o.order_id DESC
    """)
    return safe_df(rows)


def get_order_operational_df(start: date, end: date) -> pd.DataFrame:
    start_dt, end_dt = as_datetime_range(start, end)
    rows = query_rows("""
        SELECT
            o.order_id, o.tanggal_order, o.tipe_layanan, o.status_order,
            t.nomor_meja, u.nama_user AS staff,
            COALESCE(SUM(od.subtotal),0) AS total_order_value,
            COUNT(q.queue_id) AS total_items,
            SUM(CASE WHEN q.status_produksi='Menunggu' THEN 1 ELSE 0 END) AS menunggu_items,
            SUM(CASE WHEN q.status_produksi='Diproses' THEN 1 ELSE 0 END) AS diproses_items,
            SUM(CASE WHEN q.status_produksi='Selesai' THEN 1 ELSE 0 END) AS selesai_items,
            CASE
                WHEN COUNT(q.queue_id) > 0 AND SUM(CASE WHEN q.status_produksi='Selesai' THEN 1 ELSE 0 END) = COUNT(q.queue_id) THEN 'Selesai'
                WHEN SUM(CASE WHEN q.status_produksi='Diproses' THEN 1 ELSE 0 END) > 0 OR SUM(CASE WHEN q.status_produksi='Selesai' THEN 1 ELSE 0 END) > 0 THEN 'Diproses'
                ELSE 'Menunggu'
            END AS operational_status
        FROM orders o
        LEFT JOIN tables t ON t.table_id=o.table_id
        LEFT JOIN users u ON u.user_id=o.user_id
        LEFT JOIN order_details od ON od.order_id=o.order_id
        LEFT JOIN production_queue q ON q.detail_id=od.detail_id
        WHERE o.tanggal_order BETWEEN ? AND ?
        GROUP BY o.order_id
        ORDER BY o.tanggal_order DESC
    """, (start_dt, end_dt))
    return safe_df(rows)


def get_menu_order_df(start: date, end: date) -> pd.DataFrame:
    start_dt, end_dt = as_datetime_range(start, end)
    rows = query_rows("""
        SELECT m.nama_menu, c.nama_kategori, SUM(od.jumlah) AS qty, SUM(od.subtotal) AS order_value
        FROM orders o
        JOIN order_details od ON od.order_id=o.order_id
        JOIN menus m ON m.menu_id=od.menu_id
        LEFT JOIN menu_categories c ON c.category_id=m.category_id
        WHERE o.tanggal_order BETWEEN ? AND ?
        GROUP BY m.menu_id
        ORDER BY qty DESC, order_value DESC
    """, (start_dt, end_dt))
    return safe_df(rows)


def get_queue_df(start: date, end: date) -> pd.DataFrame:
    start_dt, end_dt = as_datetime_range(start, end)
    rows = query_rows("""
        SELECT q.queue_id, q.order_id, q.bagian_tujuan, q.status_produksi, q.waktu_mulai, q.waktu_selesai,
               o.tanggal_order, t.nomor_meja, m.nama_menu, od.jumlah, od.catatan
        FROM production_queue q
        JOIN order_details od ON od.detail_id=q.detail_id
        JOIN menus m ON m.menu_id=od.menu_id
        JOIN orders o ON o.order_id=q.order_id
        LEFT JOIN tables t ON t.table_id=o.table_id
        WHERE o.tanggal_order BETWEEN ? AND ?
        ORDER BY q.queue_id DESC
    """, (start_dt, end_dt))
    return safe_df(rows)


def order_period_summary(start: date, end: date) -> Dict[str, Any]:
    orders = get_order_operational_df(start, end)
    detail = get_menu_order_df(start, end)
    total_value = int(orders["total_order_value"].sum()) if not orders.empty else 0
    total_orders = int(len(orders)) if not orders.empty else 0
    completed_orders = int((orders["operational_status"] == "Selesai").sum()) if not orders.empty else 0
    processing_orders = int((orders["operational_status"] == "Diproses").sum()) if not orders.empty else 0
    waiting_orders = int((orders["operational_status"] == "Menunggu").sum()) if not orders.empty else 0
    completed_value = int(orders.loc[orders["operational_status"] == "Selesai", "total_order_value"].sum()) if not orders.empty else 0
    completion_rate = (completed_orders / total_orders * 100) if total_orders else 0.0
    avg_order_value = int(total_value / total_orders) if total_orders else 0
    top_menu = "-"
    if not detail.empty:
        top_menu = str(detail.iloc[0]["nama_menu"])
    return {
        "start": start.isoformat(), "end": end.isoformat(),
        "total_order_value": total_value, "completed_order_value": completed_value,
        "total_orders": total_orders, "completed_orders": completed_orders,
        "processing_orders": processing_orders, "waiting_orders": waiting_orders,
        "completion_rate": completion_rate, "avg_order_value": avg_order_value,
        "top_menu": top_menu,
    }


def pct_change(current: float, previous: float) -> Optional[float]:
    if previous in (0, None):
        return None
    return ((current - previous) / previous) * 100


def fmt_pct(value: Optional[float]) -> str:
    return "N/A" if value is None else f"{value:.2f}%"


def metric_delta(current: float, previous: float, suffix: str = "") -> str:
    change = pct_change(current, previous)
    return "N/A" if change is None else f"{change:.2f}%{suffix}"


def make_order_line_df(start: date, end: date, label: str = "Periode") -> pd.DataFrame:
    orders = get_order_operational_df(start, end)
    if orders.empty:
        return pd.DataFrame(columns=["tanggal", "Tanggal", "Nilai", "Jumlah Order", "Periode", "Status"] )
    orders["tanggal"] = pd.to_datetime(orders["tanggal_order"]).dt.date
    daily = orders.groupby(["tanggal", "operational_status"], as_index=False).agg(
        Nilai=("total_order_value", "sum"),
        **{"Jumlah Order": ("order_id", "count")}
    )
    daily["Tanggal"] = daily["tanggal"].apply(fmt_date)
    daily["Periode"] = label
    daily["Status"] = daily["operational_status"]
    return daily


def get_sales_df(start: date, end: date) -> pd.DataFrame:
    start_dt, end_dt = as_datetime_range(start, end)
    rows = query_rows("""
        SELECT p.payment_id, p.order_id, p.tanggal_bayar, p.total_tagihan, p.diskon,
               (p.total_tagihan - p.diskon) AS net_sales, p.metode_pembayaran,
               p.jumlah_bayar, p.kembalian, p.status_bayar, u.nama_user AS kasir
        FROM payments p LEFT JOIN users u ON u.user_id=p.cashier_user_id
        WHERE p.tanggal_bayar BETWEEN ? AND ? ORDER BY p.tanggal_bayar DESC
    """, (start_dt, end_dt))
    return safe_df(rows)


def get_expenses_df(start: date, end: date) -> pd.DataFrame:
    rows = query_rows("""
        SELECT e.expense_id, e.tanggal, e.kategori_biaya, e.nominal, e.metode_pembayaran,
               e.keterangan, u.nama_user AS staff
        FROM expenses e LEFT JOIN users u ON u.user_id=e.user_id
        WHERE date(e.tanggal) BETWEEN date(?) AND date(?) ORDER BY e.tanggal DESC
    """, (start.isoformat(), end.isoformat()))
    return safe_df(rows)


def get_closing_df(start: date, end: date) -> pd.DataFrame:
    rows = query_rows("""
        SELECT d.*, u.nama_user AS staff FROM daily_closing d
        LEFT JOIN users u ON u.user_id=d.user_id
        WHERE date(d.tanggal) BETWEEN date(?) AND date(?) ORDER BY d.tanggal DESC, d.shift DESC
    """, (start.isoformat(), end.isoformat()))
    return safe_df(rows)


def get_order_details_df(order_id: int) -> pd.DataFrame:
    rows = query_rows("""
        SELECT od.detail_id, m.nama_menu, m.bagian_produksi, od.jumlah, od.harga_satuan,
               od.subtotal, od.catatan
        FROM order_details od JOIN menus m ON m.menu_id=od.menu_id
        WHERE od.order_id=? ORDER BY od.detail_id
    """, (order_id,))
    return safe_df(rows)


def get_menu_sales_df(start: date, end: date) -> pd.DataFrame:
    start_dt, end_dt = as_datetime_range(start, end)
    rows = query_rows("""
        SELECT m.nama_menu, c.nama_kategori, SUM(od.jumlah) AS qty, SUM(od.subtotal) AS gross_sales
        FROM payments p
        JOIN orders o ON o.order_id=p.order_id
        JOIN order_details od ON od.order_id=o.order_id
        JOIN menus m ON m.menu_id=od.menu_id
        LEFT JOIN menu_categories c ON c.category_id=m.category_id
        WHERE p.tanggal_bayar BETWEEN ? AND ?
        GROUP BY m.menu_id ORDER BY qty DESC, gross_sales DESC
    """, (start_dt, end_dt))
    return safe_df(rows)


def period_summary(start: date, end: date) -> Dict[str, Any]:
    sales = get_sales_df(start, end)
    expenses = get_expenses_df(start, end)
    detail = get_menu_sales_df(start, end)
    total_sales = int(sales["net_sales"].sum()) if not sales.empty else 0
    total_exp = int(expenses["nominal"].sum()) if not expenses.empty else 0
    trx = int(len(sales))
    avg_trx = int(total_sales / trx) if trx else 0
    top_menu = "-"
    if not detail.empty:
        top_row = detail.sort_values("qty", ascending=False).head(1)
        if not top_row.empty:
            top_menu = str(top_row.iloc[0]["nama_menu"])
    return {"start": start.isoformat(), "end": end.isoformat(), "total_sales": total_sales, "total_expenses": total_exp, "net_cashflow": total_sales - total_exp, "transactions": trx, "avg_transaction": avg_trx, "top_menu": top_menu}

# ============================================================
# Excel helpers
# ============================================================


def make_interpretation(summary: Dict[str, Any]) -> str:
    total_sales = summary.get("total_sales", 0)
    total_exp = summary.get("total_expenses", 0)
    ratio = (total_exp / total_sales * 100) if total_sales else 0
    net = summary.get("net_cashflow", 0)
    if total_sales == 0:
        return "Belum terdapat transaksi penjualan pada periode ini, sehingga performa belum dapat dianalisis secara memadai."
    if ratio <= 35 and net > 0:
        return "Performa periode ini relatif sehat. Penjualan mampu menutup pengeluaran operasional dengan rasio pengeluaran yang terkendali."
    if ratio <= 60 and net > 0:
        return "Performa periode ini cukup baik, namun pengeluaran perlu dipantau agar margin kas tetap aman."
    return "Performa periode ini perlu mendapat perhatian karena pengeluaran relatif tinggi dibandingkan penjualan atau net cashflow rendah."


def make_recommendation(summary: Dict[str, Any]) -> str:
    total_sales = summary.get("total_sales", 0)
    total_exp = summary.get("total_expenses", 0)
    ratio = (total_exp / total_sales * 100) if total_sales else 0
    if total_sales == 0:
        return "Dorong pencatatan transaksi secara rutin agar dashboard dan laporan dapat digunakan untuk pengambilan keputusan."
    if ratio > 60:
        return "Manager perlu meninjau pengeluaran harian dan memastikan biaya operasional dicatat dengan kategori yang jelas."
    return "Pertahankan pola operasional saat ini dan gunakan menu terlaris sebagai dasar promosi atau rekomendasi waiter."


def make_excel_report(pay: pd.DataFrame, detail: pd.DataFrame, exp: pd.DataFrame, closing: pd.DataFrame, summary: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    engine = get_excel_engine()
    interpretation_df = pd.DataFrame([
        {"Item": "Periode Awal", "Nilai": summary.get("start")}, {"Item": "Periode Akhir", "Nilai": summary.get("end")},
        {"Item": "Total Penjualan", "Nilai": summary.get("total_sales")}, {"Item": "Total Pengeluaran", "Nilai": summary.get("total_expenses")},
        {"Item": "Net Cashflow", "Nilai": summary.get("net_cashflow")}, {"Item": "Jumlah Transaksi", "Nilai": summary.get("transactions")},
        {"Item": "Rata-rata Transaksi", "Nilai": summary.get("avg_transaction")}, {"Item": "Menu Terlaris", "Nilai": summary.get("top_menu")},
        {"Item": "Interpretasi", "Nilai": make_interpretation(summary)}, {"Item": "Rekomendasi", "Nilai": make_recommendation(summary)},
    ])
    with pd.ExcelWriter(output, engine=engine) as writer:
        interpretation_df.to_excel(writer, sheet_name="Interpretasi Report", index=False)
        (pay if not pay.empty else pd.DataFrame()).to_excel(writer, sheet_name="Penjualan", index=False)
        (detail if not detail.empty else pd.DataFrame()).to_excel(writer, sheet_name="Per Menu", index=False)
        (exp if not exp.empty else pd.DataFrame()).to_excel(writer, sheet_name="Pengeluaran", index=False)
        (closing if not closing.empty else pd.DataFrame()).to_excel(writer, sheet_name="Closing", index=False)
    return output.getvalue()


def make_comparison_excel(a: Dict[str, Any], b: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    engine = get_excel_engine()
    comp = pd.DataFrame([
        {"Metrik": "Total Penjualan", "Periode A": a["total_sales"], "Periode B": b["total_sales"]},
        {"Metrik": "Total Pengeluaran", "Periode A": a["total_expenses"], "Periode B": b["total_expenses"]},
        {"Metrik": "Net Cashflow", "Periode A": a["net_cashflow"], "Periode B": b["net_cashflow"]},
        {"Metrik": "Jumlah Transaksi", "Periode A": a["transactions"], "Periode B": b["transactions"]},
        {"Metrik": "Rata-rata Transaksi", "Periode A": a["avg_transaction"], "Periode B": b["avg_transaction"]},
        {"Metrik": "Menu Terlaris", "Periode A": a["top_menu"], "Periode B": b["top_menu"]},
    ])
    growth = ((b["total_sales"] - a["total_sales"]) / a["total_sales"] * 100) if a["total_sales"] else None
    interp = pd.DataFrame([
        {"Item": "Periode A", "Nilai": f"{a['start']} s.d. {a['end']}"}, {"Item": "Periode B", "Nilai": f"{b['start']} s.d. {b['end']}"},
        {"Item": "Growth Penjualan", "Nilai": "N/A" if growth is None else f"{growth:.2f}%"},
        {"Item": "Interpretasi", "Nilai": "Penjualan periode B meningkat dibanding periode A." if growth and growth > 0 else "Penjualan periode B belum meningkat dibanding periode A atau belum ada basis pembanding."},
    ])
    with pd.ExcelWriter(output, engine=engine) as writer:
        interp.to_excel(writer, sheet_name="Interpretasi", index=False)
        comp.to_excel(writer, sheet_name="Perbandingan", index=False)
    return output.getvalue()

# ============================================================
# Login page
# ============================================================


def login_page() -> None:
    inject_global_css(login_mode=True)
    logo_src = image_to_base64_src(find_asset(LOGO_FILES))
    logo_html = f'<img src="{logo_src}" alt="Gubug Watu Kali Logo">' if logo_src else ''

    left, right = st.columns([0.42, 0.58], gap="large")
    with left:
        st.markdown(
            f"""
            <div class="login-left-panel">
                <div>
                    <div class="login-brand-row">{logo_html}</div>
                    <div class="login-kicker">✦ REDESIGNED VISUAL ENTERPRISE</div>
                    <h1>Sistem Informasi<br>Restoran Terpadu</h1>
                    <p>Menggabungkan pencatatan order modern, antrean dapur saji cepat, dan dashboard monitor operasional dalam satu kesatuan sistem handal.</p>
                    <div class="login-proof">
                        <span>🛡️ Enforce Local Offline Relational State Data</span>
                        <span>🛡️ Structured shift reconciliation clearances</span>
                    </div>
                </div>
                <div class="login-footnote">© 2026 Restoran Gubug Watu Kali (GWK). Built for academic prototype.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        tab_login, tab_register = st.tabs(["↪ Masuk Sistem", "⚭ Daftar Akun Baru"])

        with tab_login:
            st.markdown("<h2>Selamat Datang Kembali</h2><div class='auth-subtitle'>Masuk menggunakan akun operasional SIRT yang sudah tersimpan di database lokal.</div>", unsafe_allow_html=True)
            with st.form("login_form"):
                username = st.text_input("USERNAME AKUN", placeholder="Username")
                password = st.text_input("SANDI KEAMANAN", type="password", placeholder="Password")
                submitted = st.form_submit_button("Masuk ke Console SIRT")
                if submitted:
                    user = query_one("SELECT * FROM users WHERE username=? AND status_aktif='Aktif'", (username.strip(),))
                    if user and verify_password(password, user["password"]):
                        st.session_state["user"] = dict(user)
                        st.session_state.setdefault("cart", [])
                        log_activity(user["user_id"], "Login", f"User {username} login")
                        st.success("Login berhasil.")
                        app_rerun()
                    else:
                        st.error("Username atau password salah, atau akun tidak aktif.")

        with tab_register:
            st.markdown("<h2>Pendaftaran Anggota SIRT</h2><div class='auth-subtitle'>Daftarkan akun karyawan, kasir, koki, atau direksi baru untuk mendapatkan hak akses login.</div>", unsafe_allow_html=True)
            with st.form("register_form"):
                nama = st.text_input("NAMA LENGKAP ANGGOTA", placeholder="Nama Lengkap Karyawan")
                c1, c2 = st.columns([1, 1])
                with c1:
                    new_username = st.text_input("USERNAME UNIK", placeholder="Username")
                with c2:
                    roles = [r["nama_role"] for r in query_rows("SELECT nama_role FROM roles ORDER BY role_id")]
                    role = st.selectbox("JABATAN OPERASIONAL (ROLE)", roles)
                new_password = st.text_input("SANDI KEAMANAN BARU", type="password", placeholder="Sandi baru Anda")
                confirm_password = st.text_input("KONFIRMASI SANDI KEAMANAN", type="password", placeholder="Ketik ulang sandi baru")
                submitted_reg = st.form_submit_button("Buat Akun SIRT Terdaftar")
                if submitted_reg:
                    if not nama.strip() or not new_username.strip() or not new_password.strip():
                        st.warning("Nama, username, dan password wajib diisi.")
                    elif new_password != confirm_password:
                        st.error("Konfirmasi password belum sama.")
                    elif query_one("SELECT user_id FROM users WHERE username=?", (new_username.strip(),)):
                        st.error("Username sudah digunakan. Pilih username lain.")
                    else:
                        exec_sql("INSERT INTO users(nama_user, username, password, role, status_aktif) VALUES (?, ?, ?, ?, 'Aktif')", (nama.strip(), new_username.strip(), hash_password(new_password), role))
                        new_user = query_one("SELECT user_id FROM users WHERE username=?", (new_username.strip(),))
                        log_activity(new_user["user_id"] if new_user else None, "Buat akun", f"Akun {new_username} dibuat sebagai {role}")
                        st.success("Akun berhasil dibuat. Silakan login melalui tab Masuk Sistem.")
            st.markdown("""
            <div style="margin-top:14px; font-size:12px; font-weight:800; color:#8A94A6; -webkit-text-fill-color:#8A94A6; letter-spacing:.08em;">⊙ AKUN DEMO PENGUJI (QUICK FILL)</div>
            <span class="demo-chip">Waiter: Budi</span><span class="demo-chip">Koki: Joko</span><span class="demo-chip">Kasir: Ana</span><span class="demo-chip">Manajer: Siti</span><span class="demo-chip">Owner: Kemal</span>
            <div style="margin-top:12px; font-size:12px; color:#94A3B8; -webkit-text-fill-color:#94A3B8;">Sandi penguji default adalah <b>123</b> untuk semua user demo jika dibuat melalui database seed atau konfigurasi kelompok.</div>
            """, unsafe_allow_html=True)

# ============================================================
# Pages
# ============================================================


def page_dashboard() -> None:
    hero("Executive Dashboard", "Ringkasan order value, status kitchen/bar, completion rate, dan performa operasional.")
    c1, c2 = st.columns(2)
    with c1:
        start = date_input_id("Tanggal awal", value=date.today() - timedelta(days=7), key="dash_start")
    with c2:
        end = date_input_id("Tanggal akhir", value=date.today(), key="dash_end")
    if start > end:
        st.error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.")
        return

    # Previous comparable period with the same length.
    period_days = (end - start).days + 1
    prev_end = start - timedelta(days=1)
    prev_start = prev_end - timedelta(days=period_days - 1)

    current = order_period_summary(start, end)
    previous = order_period_summary(prev_start, prev_end)
    orders = get_order_operational_df(start, end)
    queue = get_queue_df(start, end)

    st.caption(f"Periode berjalan: {date_range_label(start, end)} | Pembanding: {date_range_label(prev_start, prev_end)}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Sales Order Value", rupiah(current["total_order_value"]), delta=metric_delta(current["total_order_value"], previous["total_order_value"]))
    m2.metric("SO Value Growth", fmt_pct(pct_change(current["total_order_value"], previous["total_order_value"])))
    m3.metric("Completed Order Value", rupiah(current["completed_order_value"]), delta=metric_delta(current["completed_order_value"], previous["completed_order_value"]))
    m4.metric("Completion Rate", f"{current['completion_rate']:.2f}%", delta=metric_delta(current["completion_rate"], previous["completion_rate"]))

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("Total Submitted Order", f"{current['total_orders']}", delta=metric_delta(current["total_orders"], previous["total_orders"]))
    m6.metric("Completed Order", f"{current['completed_orders']}", delta=metric_delta(current["completed_orders"], previous["completed_orders"]))
    m7.metric("In Progress Order", f"{current['processing_orders']}")
    m8.metric("Waiting Order", f"{current['waiting_orders']}")

    st.markdown("### Order Over Time")
    daily = make_order_line_df(start, end, "Periode Berjalan")
    if not daily.empty:
        fig = px.line(
            daily,
            x="Tanggal",
            y="Nilai",
            color="Status",
            markers=True,
            title="Nilai Order Harian Berdasarkan Status Kitchen/Bar",
        )
        fig.update_layout(yaxis_title="Nilai Order", xaxis_title="Tanggal")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Belum ada order pada periode ini.")

    st.markdown("### Status Order Kitchen/Bar")
    status_daily = pd.DataFrame()
    if not orders.empty:
        status_daily = orders.copy()
        status_daily["tanggal"] = pd.to_datetime(status_daily["tanggal_order"]).dt.date
        status_line = status_daily.groupby(["tanggal", "operational_status"], as_index=False)["order_id"].count()
        status_line["Tanggal"] = status_line["tanggal"].apply(fmt_date)
        status_line = status_line.rename(columns={"order_id": "Jumlah Order", "operational_status": "Status"})
        fig2 = px.line(status_line, x="Tanggal", y="Jumlah Order", color="Status", markers=True, title="Jumlah Order Harian per Status")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Belum ada status order untuk ditampilkan.")

    st.markdown("### Antrean Kitchen/Bar")
    if not queue.empty:
        show_cols = ["order_id", "nomor_meja", "nama_menu", "jumlah", "bagian_tujuan", "status_produksi", "tanggal_order"]
        st.dataframe(queue[show_cols], use_container_width=True)
    else:
        st.info("Belum ada antrean kitchen/bar pada periode ini.")

    st.markdown("### Snapshot Analytics")
    snapshot_name = st.text_input("Nama snapshot", value=f"Dashboard {fmt_date(start)} s.d. {fmt_date(end)}")
    if st.button("Simpan Snapshot Analytics"):
        user = get_current_user()
        exec_sql(
            "INSERT INTO analytics_snapshots(user_id, nama_snapshot, tanggal_simpan, periode_awal, periode_akhir, ringkasan_json) VALUES (?, ?, ?, ?, ?, ?)",
            (user["user_id"], snapshot_name, now_str(), start.isoformat(), end.isoformat(), json.dumps(current)),
        )
        log_activity(user["user_id"], "Simpan snapshot", snapshot_name)
        st.success("Snapshot analytics berhasil disimpan.")

    st.markdown(
        f"""<div class="mini-card"><b>Menu paling sering dipesan:</b> {current['top_menu']}<br><b>Interpretasi:</b> {'Order value meningkat dibanding periode sebelumnya.' if (pct_change(current['total_order_value'], previous['total_order_value']) or 0) > 0 else 'Order value belum meningkat dibanding periode sebelumnya atau belum ada basis pembanding.'}</div>""",
        unsafe_allow_html=True,
    )

def page_order() -> None:
    st.markdown("<h1 style='margin-bottom:2px;'>Point of Sale (POS)</h1>", unsafe_allow_html=True)
    st.caption("Mencatat pembelian menu cepat. Pesanan otomatis diteruskan ke monitor antrean koki.")
    user = get_current_user()
    st.session_state.setdefault("cart", [])
    tables_df = safe_df(query_rows("SELECT * FROM tables ORDER BY table_id"))
    menu_df = get_menu_df(True)

    st.markdown('<div class="pos-top-filter">', unsafe_allow_html=True)
    top1, top2, top3 = st.columns([1.1, 1.1, 1.25])
    with top1:
        table_label = st.selectbox("PILIH MEJA / LOKASI", tables_df["nomor_meja"].tolist() if not tables_df.empty else [], format_func=lambda x: f"{x} — Area {tables_df.loc[tables_df['nomor_meja'] == x, 'area'].iloc[0]} (Tersedia)" if not tables_df.empty else x)
    with top2:
        service_type = st.radio("KATEGORI LAYANAN", ["Ala Carte", "Prasmanan"], horizontal=True)
    with top3:
        customer_note = st.text_input("MEMO ORDER UMUM", placeholder="Contoh: rombongan bupati, sajikan bersamaan...")
    st.markdown('</div>', unsafe_allow_html=True)

    filtered = menu_df[menu_df["tipe_menu"].isin(["Prasmanan"])] if service_type == "Prasmanan" else menu_df[menu_df["tipe_menu"].isin(["Ala Carte"])]

    left, right = st.columns([2.05, 0.95], gap="large")
    with left:
        st.markdown('<div class="pos-search-wrap">', unsafe_allow_html=True)
        s1, s2 = st.columns([2.4, 0.8])
        with s1:
            search_term = st.text_input("Cari menu", placeholder="🔎 Cari Menu Makanan Utama, Cemilan, Wedangan, Kopi...", label_visibility="collapsed")
        with s2:
            st.markdown("<div style='font-size:12px; font-weight:900; color:#64748B; -webkit-text-fill-color:#64748B; padding-top:10px;'>☷ Pencarian Cepat</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        categories = ["Semua"] + sorted(filtered["nama_kategori"].dropna().unique().tolist()) if not filtered.empty else ["Semua"]
        cat = st.radio("Kategori menu", categories, horizontal=True, label_visibility="collapsed")
        data_menu = filtered if cat == "Semua" else filtered[filtered["nama_kategori"] == cat]
        if search_term.strip():
            mask = data_menu["nama_menu"].str.contains(search_term.strip(), case=False, na=False) | data_menu["deskripsi"].fillna("").str.contains(search_term.strip(), case=False, na=False)
            data_menu = data_menu[mask]

        if data_menu.empty:
            st.warning("Belum ada menu sesuai filter ini.")
        else:
            rows = data_menu.to_dict("records")
            for i in range(0, len(rows), 3):
                cols = st.columns(3)
                for col, row in zip(cols, rows[i:i+3]):
                    with col:
                        desc = (row.get("deskripsi") or "").strip()
                        if len(desc) > 82:
                            desc = desc[:82] + "..."
                        st.markdown(
                            f"""
                            <div class="pos-menu-card">
                                <div class="pos-menu-badge">{row.get('nama_kategori') or '-'}</div>
                                <div class="pos-menu-title">{row['nama_menu']}</div>
                                <div class="pos-menu-desc">{desc or 'Menu GWK siap ditambahkan ke order.'}</div>
                                <div class="pos-menu-price">{rupiah(row['harga'])}</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        qty = st.number_input("Qty", min_value=1, value=1, step=1, key=f"qty_{row['menu_id']}", label_visibility="collapsed")
                        note = st.text_input("Catatan", placeholder="Catatan item", key=f"note_{row['menu_id']}", label_visibility="collapsed")
                        if st.button("＋ Tambah", key=f"add_{row['menu_id']}"):
                            st.session_state["cart"].append({
                                "menu_id": int(row["menu_id"]),
                                "nama_menu": row["nama_menu"],
                                "harga": int(row["harga"]),
                                "bagian_produksi": row["bagian_produksi"],
                                "jumlah": int(qty),
                                "subtotal": int(qty) * int(row["harga"]),
                                "catatan": note.strip(),
                            })
                            st.success("Item ditambahkan ke ringkasan.")
                            app_rerun()

    with right:
        with st.container(border=True):
            cart = st.session_state.get("cart", [])
            st.markdown(f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;'><b>🛒 Ringkasan Cart</b><span style='background:#EEF2EA; color:#1B3022; -webkit-text-fill-color:#1B3022; font-size:11px; font-weight:900; padding:4px 8px; border-radius:999px;'>{len(cart)} Item</span></div>", unsafe_allow_html=True)
            if cart:
                for idx, item in enumerate(cart):
                    st.markdown(
                        f"""
                        <div class="cart-line">
                            <b>{item['nama_menu']}</b><br>
                            <span style="font-size:12px; color:#64748B; -webkit-text-fill-color:#64748B;">{item['jumlah']} × {rupiah(item['harga'])} • {item['bagian_produksi']}</span><br>
                            <b>{rupiah(item['subtotal'])}</b>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button("Hapus", key=f"remove_cart_{idx}"):
                        st.session_state["cart"].pop(idx)
                        app_rerun()
                cart_df = pd.DataFrame(cart)
                total = int(cart_df["subtotal"].sum())
                st.markdown(f"<div class='cart-subtotal'><span>SUBTOTAL POS</span><span>{rupiah(total)}</span></div>", unsafe_allow_html=True)
                c_clear, c_submit = st.columns([0.95, 1.35], gap="medium")
                with c_clear:
                    if st.button("↺ Clear Cart"):
                        st.session_state["cart"] = []
                        app_rerun()
                with c_submit:
                    if st.button("▣ Submit Order"):
                        table_id = int(tables_df.loc[tables_df["nomor_meja"] == table_label, "table_id"].iloc[0])
                        with get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute(
                                "INSERT INTO orders(table_id, user_id, tanggal_order, tipe_layanan, status_order, customer_note) VALUES (?, ?, ?, ?, 'Submitted', ?)",
                                (table_id, user["user_id"], now_str(), service_type, customer_note),
                            )
                            order_id = cur.lastrowid
                            for item in cart:
                                cur.execute(
                                    "INSERT INTO order_details(order_id, menu_id, jumlah, harga_satuan, subtotal, catatan) VALUES (?, ?, ?, ?, ?, ?)",
                                    (order_id, item["menu_id"], item["jumlah"], item["harga"], item["subtotal"], item["catatan"]),
                                )
                                detail_id = cur.lastrowid
                                cur.execute(
                                    "INSERT INTO production_queue(order_id, detail_id, bagian_tujuan, status_produksi) VALUES (?, ?, ?, 'Menunggu')",
                                    (order_id, detail_id, item["bagian_produksi"]),
                                )
                            conn.commit()
                        log_activity(user["user_id"], "Submit order", f"Order #{order_id} dibuat")
                        st.session_state["cart"] = []
                        st.success(f"Order #{order_id} berhasil dibuat dan dikirim ke Kitchen/Bar.")
                        app_rerun()
            else:
                st.markdown('<div class="cart-empty">🛒<br><br>Keranjang masih kosong. Pilih menu di sisi kiri untuk memulai.</div>', unsafe_allow_html=True)
                st.markdown("<div class='cart-subtotal'><span>SUBTOTAL POS</span><span>Rp0</span></div>", unsafe_allow_html=True)

def page_kitchen() -> None:
    hero("Antrean Kitchen Board", "Display kitchen/bar untuk melihat order masuk dan menandai status menunggu, diproses, atau selesai.")
    status_filter = st.multiselect("Filter status", ["Menunggu", "Diproses", "Selesai"], default=["Menunggu", "Diproses"])
    placeholders = ",".join(["?"] * len(status_filter)) if status_filter else "'x'"
    rows = query_rows(f"""
        SELECT q.queue_id, q.order_id, q.detail_id, q.bagian_tujuan, q.status_produksi, q.waktu_mulai, q.waktu_selesai,
               t.nomor_meja, m.nama_menu, od.jumlah, od.catatan, o.tanggal_order
        FROM production_queue q JOIN order_details od ON od.detail_id=q.detail_id JOIN menus m ON m.menu_id=od.menu_id
        JOIN orders o ON o.order_id=q.order_id LEFT JOIN tables t ON t.table_id=o.table_id
        WHERE q.status_produksi IN ({placeholders}) ORDER BY q.queue_id DESC
    """, tuple(status_filter))
    df = safe_df(rows)
    if df.empty: st.info("Tidak ada antrean sesuai filter."); return
    for row in df.to_dict("records"):
        with st.container(border=True):
            cols = st.columns([1.2,2.4,1.4,1.2,2.6])
            cols[0].markdown(f"**Order #{row['order_id']}**  \n{row['nomor_meja']}"); cols[1].markdown(f"**{row['nama_menu']}**  \nJumlah: {row['jumlah']}  \nCatatan: {row['catatan'] or '-'}")
            cols[2].markdown(f"**Tujuan**  \n{row['bagian_tujuan']}"); cols[3].markdown(f"**Status**  \n{row['status_produksi']}")
            with cols[4]:
                b1,b2,b3=st.columns(3)
                if b1.button("Menunggu", key=f"wait_{row['queue_id']}"): exec_sql("UPDATE production_queue SET status_produksi='Menunggu' WHERE queue_id=?", (row["queue_id"],)); app_rerun()
                if b2.button("Diproses", key=f"proc_{row['queue_id']}"): exec_sql("UPDATE production_queue SET status_produksi='Diproses', waktu_mulai=COALESCE(waktu_mulai, ?) WHERE queue_id=?", (now_str(), row["queue_id"])); app_rerun()
                if b3.button("Selesai", key=f"done_{row['queue_id']}"):
                    exec_sql("UPDATE production_queue SET status_produksi='Selesai', waktu_selesai=? WHERE queue_id=?", (now_str(), row["queue_id"]))
                    pending = query_one("SELECT COUNT(*) AS n FROM production_queue WHERE order_id=? AND status_produksi!='Selesai'", (row["order_id"],))["n"]
                    exec_sql("UPDATE orders SET status_order=? WHERE order_id=?", ("Selesai" if pending == 0 else "Diproses", row["order_id"])); app_rerun()


def page_payment() -> None:
    hero("Payment Page", "Proses pembayaran kasir dengan diskon berbasis persen dan struk otomatis.")
    user = get_current_user(); unpaid = get_unpaid_orders()
    if unpaid.empty: st.info("Tidak ada order yang belum dibayar."); return
    options = [f"Order #{int(r.order_id)} — {r.nomor_meja} — {rupiah(r.total)}" for r in unpaid.itertuples()]
    selected = st.selectbox("Pilih order", options); order = unpaid.iloc[options.index(selected)]; order_id = int(order["order_id"]); details = get_order_details_df(order_id)
    st.dataframe(details, use_container_width=True); gross_total = int(details["subtotal"].sum()) if not details.empty else 0; st.metric("Total tagihan bruto", rupiah(gross_total))
    c1,c2,c3 = st.columns(3)
    with c1: discount_pct = st.number_input("Diskon (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
    discount_amount = int(round(gross_total * (discount_pct / 100))); net_total = max(gross_total - discount_amount, 0)
    with c2: method = st.selectbox("Metode pembayaran", ["Cash", "QRIS", "Debit", "Transfer", "Lainnya"])
    with c3: amount_paid = st.number_input("Jumlah bayar", min_value=0, value=net_total if method != "Cash" else 0, step=1000)
    change = max(int(amount_paid) - net_total, 0)
    st.markdown(f"""<div class="mini-card"><b>Diskon:</b> {discount_pct:.1f}% ({rupiah(discount_amount)})<br><b>Total bersih:</b> {rupiah(net_total)}<br><b>Kembalian:</b> {rupiah(change)}</div>""", unsafe_allow_html=True)
    if st.button("Simpan Pembayaran & Generate Struk"):
        if amount_paid < net_total: st.error("Jumlah bayar belum mencukupi total bersih.")
        else:
            exec_sql("""INSERT INTO payments(order_id, tanggal_bayar, total_tagihan, diskon, metode_pembayaran, jumlah_bayar, kembalian, status_bayar, cashier_user_id) VALUES (?, ?, ?, ?, ?, ?, ?, 'Lunas', ?)""", (order_id, now_str(), gross_total, discount_amount, method, int(amount_paid), change, user["user_id"]))
            exec_sql("UPDATE orders SET status_order='Lunas' WHERE order_id=?", (order_id,))
            receipt = build_receipt(order_id, method, gross_total, discount_pct, discount_amount, net_total, int(amount_paid), change, user["nama_user"])
            st.session_state["last_receipt"] = receipt; log_activity(user["user_id"], "Pembayaran", f"Order #{order_id} dibayar {method} {rupiah(net_total)}"); st.success("Pembayaran berhasil disimpan."); app_rerun()
    if st.session_state.get("last_receipt"):
        st.markdown("### Struk Terakhir"); st.markdown(f"<div class='receipt-box'>{st.session_state['last_receipt']}</div>", unsafe_allow_html=True)
        st.download_button("Download Struk .txt", st.session_state["last_receipt"].encode("utf-8"), file_name=f"struk_order_{order_id}.txt")


def build_receipt(order_id: int, method: str, gross_total: int, pct: float, discount: int, net_total: int, paid: int, change: int, cashier: str) -> str:
    order = query_one("SELECT o.*, t.nomor_meja FROM orders o LEFT JOIN tables t ON t.table_id=o.table_id WHERE o.order_id=?", (order_id,))
    details = get_order_details_df(order_id)
    lines = ["GUBUG WATU KALI", "SIRT GWK - STRUK PEMBAYARAN", "="*36, f"No Order  : #{order_id}", f"Meja      : {order['nomor_meja'] if order else '-'}", f"Tanggal   : {now_str()}", f"Kasir     : {cashier}", "-"*36]
    for r in details.to_dict("records"):
        lines.append(f"{r['nama_menu']} x{r['jumlah']} @ {rupiah(r['harga_satuan'])}"); lines.append(f"  Subtotal: {rupiah(r['subtotal'])}")
    lines += ["-"*36, f"Total Bruto : {rupiah(gross_total)}", f"Diskon      : {pct:.1f}% ({rupiah(discount)})", f"Total Bayar : {rupiah(net_total)}", f"Metode      : {method}", f"Dibayar     : {rupiah(paid)}", f"Kembalian   : {rupiah(change)}", "="*36, "Terima kasih."]
    return "\n".join(lines)


def page_expense_closing() -> None:
    hero("Pengeluaran & Closing", "Catat pengeluaran operasional dan lakukan closing shift harian.")
    user = get_current_user(); tab1, tab2 = st.tabs(["Input Pengeluaran", "Closing Shift"])
    with tab1:
        with st.form("expense_form"):
            tanggal=date_input_id("Tanggal pengeluaran", value=date.today()); kategori=st.selectbox("Kategori biaya", ["Bahan baku", "Perlengkapan", "Listrik", "Transportasi", "Kebersihan", "Operasional Lainnya"])
            nominal=st.number_input("Nominal", min_value=0, step=1000); metode=st.selectbox("Metode pembayaran", ["Cash", "Transfer", "QRIS", "Lainnya"]); ket=st.text_area("Keterangan")
            if st.form_submit_button("Simpan Pengeluaran"):
                if nominal <= 0: st.warning("Nominal harus lebih dari 0.")
                else:
                    exec_sql("INSERT INTO expenses(tanggal, kategori_biaya, nominal, metode_pembayaran, keterangan, user_id) VALUES (?, ?, ?, ?, ?, ?)", (tanggal.isoformat(), kategori, int(nominal), metode, ket, user["user_id"]))
                    log_activity(user["user_id"], "Input pengeluaran", f"{kategori} {rupiah(nominal)}"); st.success("Pengeluaran berhasil disimpan.")
    with tab2:
        closing_date=date_input_id("Tanggal closing", value=date.today(), key="closing_date"); shift=st.selectbox("Shift", ["Pagi", "Siang", "Malam", "Full Day"])
        sales=get_sales_df(closing_date, closing_date); expenses=get_expenses_df(closing_date, closing_date)
        total_sales=int(sales["net_sales"].sum()) if not sales.empty else 0; cash_system=int(sales.loc[sales["metode_pembayaran"]=="Cash", "net_sales"].sum()) if not sales.empty else 0; non_cash_system=total_sales-cash_system; total_expenses=int(expenses["nominal"].sum()) if not expenses.empty else 0
        c1,c2,c3,c4=st.columns(4); c1.metric("Penjualan Sistem", rupiah(total_sales)); c2.metric("Cash Sistem", rupiah(cash_system)); c3.metric("Non-Cash Sistem", rupiah(non_cash_system)); c4.metric("Pengeluaran", rupiah(total_expenses))
        with st.form("closing_form"):
            physical_cash=st.number_input("Total Cash Fisik", min_value=0, value=cash_system, step=1000); physical_non_cash=st.number_input("Total Non-Cash Fisik", min_value=0, value=non_cash_system, step=1000); note=st.text_area("Catatan closing")
            if st.form_submit_button("Simpan Closing"):
                diff=int(physical_cash+physical_non_cash-total_sales)
                exec_sql("""INSERT INTO daily_closing(user_id, tanggal, shift, total_penjualan, total_cash_sistem, total_non_cash_sistem, total_cash_fisik, total_non_cash_fisik, total_pengeluaran, selisih, catatan) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", (user["user_id"], closing_date.isoformat(), shift, total_sales, cash_system, non_cash_system, int(physical_cash), int(physical_non_cash), total_expenses, diff, note))
                log_activity(user["user_id"], "Closing shift", f"{closing_date} {shift}, selisih {rupiah(diff)}"); st.success(f"Closing berhasil disimpan. Selisih: {rupiah(diff)}")


def page_reports() -> None:
    hero("Reports", "Laporan penjualan, pengeluaran, closing, dan interpretasi report Excel.")
    c1,c2,c3=st.columns(3)
    with c1: start=date_input_id("Tanggal awal", value=date.today()-timedelta(days=7), key="rep_start")
    with c2: end=date_input_id("Tanggal akhir", value=date.today(), key="rep_end")
    with c3: method_filter=st.selectbox("Filter metode", ["Semua", "Cash", "QRIS", "Debit", "Transfer", "Lainnya"])
    if start>end: st.error("Tanggal awal tidak boleh lebih besar dari tanggal akhir."); return
    pay=get_sales_df(start,end); exp=get_expenses_df(start,end); closing=get_closing_df(start,end); detail=get_menu_sales_df(start,end); summary=period_summary(start,end)
    if method_filter!="Semua" and not pay.empty: pay=pay[pay["metode_pembayaran"]==method_filter]
    st.markdown(f"""<div class="mini-card"><b>Total Penjualan:</b> {rupiah(summary['total_sales'])}<br><b>Total Pengeluaran:</b> {rupiah(summary['total_expenses'])}<br><b>Net Cashflow:</b> {rupiah(summary['net_cashflow'])}<br><b>Interpretasi:</b> {make_interpretation(summary)}<br><b>Rekomendasi:</b> {make_recommendation(summary)}</div>""", unsafe_allow_html=True)
    st.download_button("Download Interpretasi Report Excel", data=make_excel_report(pay,detail,exp,closing,summary), file_name=f"SIRT_GWK_Report_{start}_{end}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    tabs=st.tabs(["Penjualan", "Per Menu", "Pengeluaran", "Closing"])
    with tabs[0]: st.dataframe(pay,use_container_width=True); st.download_button("Download penjualan.csv", pay.to_csv(index=False).encode("utf-8"), "penjualan.csv", "text/csv")
    with tabs[1]: st.dataframe(detail,use_container_width=True); st.download_button("Download menu_sales.csv", detail.to_csv(index=False).encode("utf-8"), "menu_sales.csv", "text/csv")
    with tabs[2]: st.dataframe(exp,use_container_width=True); st.download_button("Download pengeluaran.csv", exp.to_csv(index=False).encode("utf-8"), "pengeluaran.csv", "text/csv")
    with tabs[3]: st.dataframe(closing,use_container_width=True); st.download_button("Download closing.csv", closing.to_csv(index=False).encode("utf-8"), "closing.csv", "text/csv")


def page_comparison() -> None:
    hero("Perbandingan Performa", "Bandingkan nilai order, completed order, dan completion rate antarperiode.")
    if not has_role("Owner", "Manager"):
        st.warning("Halaman ini hanya untuk Owner dan Manager.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        a_start = date_input_id("Periode A - awal", value=date.today() - timedelta(days=14), key="comp_a_start")
    with c2:
        a_end = date_input_id("Periode A - akhir", value=date.today() - timedelta(days=8), key="comp_a_end")
    with c3:
        b_start = date_input_id("Periode B - awal", value=date.today() - timedelta(days=7), key="comp_b_start")
    with c4:
        b_end = date_input_id("Periode B - akhir", value=date.today(), key="comp_b_end")

    if a_start > a_end or b_start > b_end:
        st.error("Tanggal awal tidak boleh lebih besar dari tanggal akhir.")
        return

    a = order_period_summary(a_start, a_end)
    b = order_period_summary(b_start, b_end)

    order_growth = pct_change(b["total_order_value"], a["total_order_value"])
    completed_growth = pct_change(b["completed_order_value"], a["completed_order_value"])
    rate_growth = pct_change(b["completion_rate"], a["completion_rate"])
    trx_growth = pct_change(b["total_orders"], a["total_orders"])

    st.caption(f"Periode A: {date_range_label(a_start, a_end)} | Periode B: {date_range_label(b_start, b_end)}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Order Value", rupiah(b["total_order_value"]), delta=fmt_pct(order_growth))
    m2.metric("Completed Order Value", rupiah(b["completed_order_value"]), delta=fmt_pct(completed_growth))
    m3.metric("Completion Rate", f"{b['completion_rate']:.2f}%", delta=fmt_pct(rate_growth))
    m4.metric("Total Orders", f"{b['total_orders']}", delta=fmt_pct(trx_growth))

    comp_df = pd.DataFrame([
        {"Metrik": "Total Order Value", "Periode A": a["total_order_value"], "Periode B": b["total_order_value"], "Growth": fmt_pct(order_growth)},
        {"Metrik": "Completed Order Value", "Periode A": a["completed_order_value"], "Periode B": b["completed_order_value"], "Growth": fmt_pct(completed_growth)},
        {"Metrik": "Completion Rate", "Periode A": round(a["completion_rate"], 2), "Periode B": round(b["completion_rate"], 2), "Growth": fmt_pct(rate_growth)},
        {"Metrik": "Total Orders", "Periode A": a["total_orders"], "Periode B": b["total_orders"], "Growth": fmt_pct(trx_growth)},
        {"Metrik": "Average Order Value", "Periode A": a["avg_order_value"], "Periode B": b["avg_order_value"], "Growth": fmt_pct(pct_change(b["avg_order_value"], a["avg_order_value"]))},
    ])
    st.dataframe(comp_df, use_container_width=True)

    st.markdown("### Grafik Garis Perbandingan Performa")
    line_a = make_order_line_df(a_start, a_end, "Periode A")
    line_b = make_order_line_df(b_start, b_end, "Periode B")
    line_df = pd.concat([line_a, line_b], ignore_index=True)
    if not line_df.empty:
        # Align comparison by relative day so different date ranges are comparable on one line chart.
        line_df["tanggal_dt"] = pd.to_datetime(line_df["tanggal"])
        line_df["Hari ke"] = line_df.groupby("Periode")["tanggal_dt"].rank(method="dense").astype(int)
        daily_period = line_df.groupby(["Hari ke", "Periode"], as_index=False)["Nilai"].sum()
        fig = px.line(daily_period, x="Hari ke", y="Nilai", color="Periode", markers=True, title="Order Value Over Time")
        fig.update_layout(yaxis_title="Nilai Order", xaxis_title="Hari ke-")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Belum ada order pada kedua periode tersebut.")

    if order_growth is None:
        interp = "Belum ada basis pembanding pada Periode A, sehingga growth belum dapat dihitung."
    elif order_growth > 0:
        interp = "Periode B menunjukkan kenaikan order value dibanding Periode A."
    elif order_growth < 0:
        interp = "Periode B menunjukkan penurunan order value dibanding Periode A."
    else:
        interp = "Order value Periode B relatif sama dengan Periode A."

    st.markdown(
        f"""<div class="mini-card"><b>Top Menu Periode A:</b> {a['top_menu']}<br><b>Top Menu Periode B:</b> {b['top_menu']}<br><b>Interpretasi:</b> {interp}</div>""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "Download Perbandingan Excel",
        make_comparison_excel(
            {"start": fmt_date(a_start), "end": fmt_date(a_end), "total_sales": a["total_order_value"], "total_expenses": 0, "net_cashflow": a["completed_order_value"], "transactions": a["total_orders"], "avg_transaction": a["avg_order_value"], "top_menu": a["top_menu"]},
            {"start": fmt_date(b_start), "end": fmt_date(b_end), "total_sales": b["total_order_value"], "total_expenses": 0, "net_cashflow": b["completed_order_value"], "transactions": b["total_orders"], "avg_transaction": b["avg_order_value"], "top_menu": b["top_menu"]},
        ),
        file_name=f"SIRT_GWK_Perbandingan_{fmt_date(a_start)}_{fmt_date(b_end)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def page_config() -> None:
    hero("Konfigurasi System", "Kelola menu, status menu, akun pengguna, dan activity log.")
    if not has_role("Owner", "Manager"): st.warning("Halaman ini hanya untuk Owner dan Manager."); return
    tabs=st.tabs(["Menu", "User & Role", "Activity Log", "Analytics Snapshot"])
    with tabs[0]:
        categories=safe_df(query_rows("SELECT * FROM menu_categories ORDER BY nama_kategori")); st.subheader("Tambah Menu")
        with st.form("add_menu_form"):
            nama_menu=st.text_input("Nama menu"); cat=st.selectbox("Kategori", categories["nama_kategori"].tolist()); harga=st.number_input("Harga", min_value=0, step=500); tipe=st.selectbox("Tipe menu", ["Ala Carte","Prasmanan"]); bagian=st.selectbox("Bagian produksi", ["Kitchen","Bar"]); desc=st.text_area("Deskripsi")
            if st.form_submit_button("Simpan Menu"):
                cat_id=int(categories.loc[categories["nama_kategori"]==cat,"category_id"].iloc[0])
                try: exec_sql("INSERT INTO menus(category_id,nama_menu,harga,tipe_menu,bagian_produksi,status_menu,deskripsi) VALUES (?, ?, ?, ?, ?, 'Tersedia', ?)", (cat_id,nama_menu.strip(),int(harga),tipe,bagian,desc)); log_activity(get_current_user()["user_id"],"Tambah menu",nama_menu); st.success("Menu berhasil disimpan.")
                except sqlite3.IntegrityError: st.error("Nama menu sudah ada.")
        menu=get_menu_df(False); st.dataframe(menu,use_container_width=True)
        if not menu.empty:
            selected=st.selectbox("Pilih menu untuk ubah status", menu["nama_menu"].tolist()); current_status=menu.loc[menu["nama_menu"]==selected,"status_menu"].iloc[0]; new_status=st.selectbox("Status baru", ["Tersedia","Tidak Tersedia"], index=0 if current_status=="Tersedia" else 1)
            if st.button("Update Status Menu"): exec_sql("UPDATE menus SET status_menu=? WHERE nama_menu=?", (new_status,selected)); log_activity(get_current_user()["user_id"],"Update status menu",f"{selected} -> {new_status}"); st.success("Status menu diperbarui."); app_rerun()
    with tabs[1]:
        st.subheader("Tambah User")
        with st.form("admin_user_form"):
            nama=st.text_input("Nama user"); username=st.text_input("Username"); password=st.text_input("Password", type="password"); role=st.selectbox("Role user", ["Owner","Manager","Kasir","Waiter","Kitchen/Bar"])
            if st.form_submit_button("Simpan User"):
                if not nama or not username or not password: st.warning("Semua field wajib diisi.")
                else:
                    try: exec_sql("INSERT INTO users(nama_user, username, password, role, status_aktif) VALUES (?, ?, ?, ?, 'Aktif')", (nama.strip(), username.strip(), hash_password(password), role)); log_activity(get_current_user()["user_id"],"Tambah user",username); st.success("User berhasil dibuat.")
                    except sqlite3.IntegrityError: st.error("Username sudah dipakai.")
        users=safe_df(query_rows("SELECT user_id,nama_user,username,role,status_aktif,created_at FROM users ORDER BY user_id DESC")); st.dataframe(users,use_container_width=True)
        if not users.empty:
            selected_user=st.selectbox("Pilih user untuk ubah status", users["username"].tolist()); status=st.selectbox("Status akun", ["Aktif","Nonaktif"])
            if st.button("Update Status Akun"): exec_sql("UPDATE users SET status_aktif=? WHERE username=?", (status,selected_user)); log_activity(get_current_user()["user_id"],"Update status user",f"{selected_user} -> {status}"); st.success("Status akun diperbarui."); app_rerun()
    with tabs[2]:
        logs=safe_df(query_rows("SELECT l.log_id,l.waktu_aktivitas,u.username,u.role,l.aktivitas,l.detail_perubahan FROM activity_logs l LEFT JOIN users u ON u.user_id=l.user_id ORDER BY l.log_id DESC LIMIT 300")); st.dataframe(logs,use_container_width=True)
    with tabs[3]:
        snaps=safe_df(query_rows("SELECT s.snapshot_id,s.nama_snapshot,s.tanggal_simpan,s.periode_awal,s.periode_akhir,u.username FROM analytics_snapshots s LEFT JOIN users u ON u.user_id=s.user_id ORDER BY s.snapshot_id DESC")); st.dataframe(snaps,use_container_width=True); st.download_button("Download analytics_snapshots.csv", snaps.to_csv(index=False).encode("utf-8"), "analytics_snapshots.csv", "text/csv")


def page_backup_clear() -> None:
    hero("Data & Backup", "Download data, backup database, dan clear data transaksi dengan konfirmasi.")
    if not has_role("Owner", "Manager"): st.warning("Halaman ini hanya untuk Owner dan Manager."); return
    tab1,tab2=st.tabs(["Download Data","Clear Data"])
    with tab1:
        tables=["orders","order_details","production_queue","payments","expenses","daily_closing","menus","users","activity_logs"]
        selected_table=st.selectbox("Pilih tabel", tables); df=safe_df(query_rows(f"SELECT * FROM {selected_table}")); st.dataframe(df,use_container_width=True); st.download_button(f"Download {selected_table}.csv", df.to_csv(index=False).encode("utf-8"), f"{selected_table}.csv", "text/csv")
        if DB_PATH.exists(): st.download_button("Download Full SQLite Database", DB_PATH.read_bytes(), "sirt_gwk.db", "application/octet-stream")
    with tab2:
        st.error("Area ini akan menghapus data transaksi. Data master seperti user, role, menu, dan meja tetap dipertahankan.")
        confirm_check=st.checkbox("Saya paham bahwa data transaksi akan dihapus."); confirm_text=st.text_input("Ketik HAPUS untuk konfirmasi")
        if st.button("Clear Data Transaksi"):
            if not confirm_check or confirm_text!="HAPUS": st.warning("Konfirmasi belum lengkap. Centang persetujuan dan ketik HAPUS.")
            else:
                with get_conn() as conn:
                    cur=conn.cursor()
                    for table in ["order_details","production_queue","payments","expenses","daily_closing","orders","activity_logs","analytics_snapshots"]: cur.execute(f"DELETE FROM {table}")
                    conn.commit()
                st.session_state["cart"]=[]; st.session_state.pop("last_receipt", None); st.success("Data transaksi berhasil dibersihkan."); app_rerun()

# ============================================================
# Main app
# ============================================================


PAGE_LABELS = {
    "Dashboard Monitoring": "Executive Dashboard",
    "Order Page": "Pintu POS Order",
    "Kitchen/Bar Display": "Antrean Kitchen Board",
    "Pengeluaran & Closing": "Pengeluaran & Closing",
    "Reports": "Laporan Terinci",
    "Perbandingan Performa": "Perbandingan Laju",
    "Konfigurasi Menu & Hak Akses": "Konfigurasi System",
    "Data & Backup": "Data & Backup",
}


def sidebar() -> Optional[str]:
    user = get_current_user()
    if not user:
        return None

    st.sidebar.markdown("### GUBUG WATU KALI")
    st.sidebar.caption("SIRT Operational Console")
    st.sidebar.markdown(
        f"""
        <div style="background:rgba(255,255,255,0.06); border:1px solid rgba(255,255,255,0.08); border-radius:14px; padding:12px 14px; margin:12px 0 18px 0;">
            <div style="font-size:10px; font-weight:800; letter-spacing:.08em; opacity:.72;">ACTIVE USER</div>
            <div style="font-size:13px; font-weight:800; margin-top:4px;">{user['nama_user']}</div>
            <div style="display:inline-block; margin-top:6px; background:#D4AF37; color:#1B3022; -webkit-text-fill-color:#1B3022; font-size:10px; font-weight:900; padding:2px 8px; border-radius:999px;">OTORITAS: {user['role']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.sidebar.markdown("##### MODULES NAVIGATION")
    page = st.sidebar.radio(
        "Navigasi",
        role_pages(user["role"]),
        format_func=lambda x: PAGE_LABELS.get(x, x),
        label_visibility="collapsed",
    )
    st.sidebar.divider()
    if st.sidebar.button("Keluar Sesi Karyawan"):
        log_activity(user["user_id"], "Logout", f"User {user['username']} logout")
        for key in ["user", "cart", "last_receipt"]:
            st.session_state.pop(key, None)
        app_rerun()
    st.sidebar.caption("SIRT Gubug Watu Kali • SQLite Local")
    return page


def main() -> None:
    init_db()
    if not get_current_user():
        login_page(); return
    inject_global_css(login_mode=False)
    page=sidebar()
    if page=="Dashboard Monitoring": page_dashboard()
    elif page=="Order Page": page_order()
    elif page=="Kitchen/Bar Display": page_kitchen()
    elif page=="Pengeluaran & Closing": page_expense_closing()
    elif page=="Reports": page_reports()
    elif page=="Perbandingan Performa": page_comparison()
    elif page=="Konfigurasi Menu & Hak Akses": page_config()
    elif page=="Data & Backup": page_backup_clear()
    else: st.info("Pilih halaman dari sidebar.")


if __name__ == "__main__":
    main()
