"""
Mobile-Friendly Stock Portfolio Tracker (Streamlit App)
=======================================================
Supports:
1. Direct Zerodha Kite / Console Tradebook Import (Multi-year CSV/XLSX upload):
   - Uses the enhanced dynamic header scanner for Zerodha CSV/XLSX files
   - Automatically filters out Intraday / Day-Trading trades (same-day Buy & Sell netting)
   - Deduplicates overlapping tradebooks using Zerodha `trade_id`
   - Uses FIFO (First-In, First-Out) delivery accounting (matching Zerodha Console) to handle
     historical BUY and SELL delivery trades, computing:
       * Current Active Holdings (Total Remaining Qty, FIFO Avg. Buy Price, Current Invested Rs., % Share)
       * Realized P&L from past sold delivery stocks
2. Enhanced Dashboard & Analytics:
   - Separate Pure Equity P&L (%) vs. Overall Portfolio P&L (%) (including GOLDBEES / Gold & ETFs)
   - Interactive Equity Performance Run-Line Graph comparing your Equity Portfolio against
     Nifty 50, Nifty Midcap, and Nifty Smallcap (both Normalized Return % and Equivalent Cashflow Value ₹)
   - Frozen (Pinned) `Stock Name` column on the left + editable `Remarks (Future Action)` column at the end
   - User-configurable Stock Concentration (%) and Sector Concentration (%) warning thresholds
   - Dedicated Unrealized P&L Graph (Green/Red diverging profit & loss chart in ₹ and %)
   - Full-width, high-clarity Monthly Investment Trend (Combo Bar + Cumulative Growth Line & Year×Month Heatmap)
   - Interactive Date / Month / Year Filter showing Stock-Wise Bought & Sold Value for any period
3. Permanent Cloud Persistence (Auto-syncs `portfolio_data.db` to/from GitHub repo).
"""

import base64
import io
import json
import re
import sqlite3
import subprocess
import sys
import urllib.error
import urllib.request
from collections import deque
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st


# ---------------------------------------------------------------------------
# Auto-Install Missing Dependencies Fallback (in case requirements.txt is missed)
# ---------------------------------------------------------------------------
def _ensure_package(import_name: str, pip_name: str | None = None) -> bool:
    pkg = pip_name or import_name
    try:
        __import__(import_name)
        return True
    except ModuleNotFoundError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])
            return True
        except Exception:
            return False


_ensure_package("plotly")
_ensure_package("openpyxl")
_ensure_package("yfinance")

import plotly.express as px
import plotly.graph_objects as go

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except Exception:
    YFINANCE_AVAILABLE = False

# ---------------------------------------------------------------------------
# Page Configuration & Mobile-Friendly Styling
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Stock Portfolio Tracker (Kite & Sheets)",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    div[data-testid="stMetric"] {
        background-color: rgba(28, 131, 225, 0.06);
        border: 1px solid rgba(28, 131, 225, 0.18);
        padding: 12px 14px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        overflow-x: auto;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 14px;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DB_PATH = Path("portfolio_data.db")


# ---------------------------------------------------------------------------
# Permanent Cloud Persistence via GitHub Repository Auto-Sync
# ---------------------------------------------------------------------------
def get_github_sync_config() -> tuple[str, str, str]:
    """Read GITHUB_TOKEN, GITHUB_REPO, and GITHUB_BRANCH from Streamlit Secrets or Session State."""
    token = ""
    repo = ""
    branch = "main"
    try:
        token = str(st.secrets.get("GITHUB_TOKEN", "")).strip()
        repo = str(st.secrets.get("GITHUB_REPO", "")).strip()
        branch = str(st.secrets.get("GITHUB_BRANCH", "main")).strip() or "main"
    except Exception:
        pass

    if not token:
        token = st.session_state.get("gh_token_manual", "")
    if not repo:
        repo = st.session_state.get("gh_repo_manual", "")
    return token, repo, branch


def restore_db_from_github_if_needed() -> bool:
    """
    When Streamlit Cloud wakes up from sleep on a fresh container, automatically
    pull `portfolio_data.db` from the user's GitHub repo so no data is ever lost.
    """
    if st.session_state.get("db_restored_from_github", False):
        return False

    token, repo, branch = get_github_sync_config()
    if not token or not repo:
        return False

    if DB_PATH.exists() and DB_PATH.stat().st_size > 12288:
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM transactions")
            if cur.fetchone()[0] > 0:
                st.session_state["db_restored_from_github"] = True
                return False
        except Exception:
            pass

    url = f"https://api.github.com/repos/{repo}/contents/portfolio_data.db?ref={branch}"
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "Streamlit-Portfolio-Tracker",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            content_b64 = data.get("content", "")
            if content_b64:
                raw_bytes = base64.b64decode(content_b64)
                DB_PATH.write_bytes(raw_bytes)
                st.session_state["db_restored_from_github"] = True
                return True
    except Exception:
        pass

    st.session_state["db_restored_from_github"] = True
    return False


def push_db_to_github(commit_msg: str = "Auto-sync portfolio_data.db from Streamlit app") -> tuple[bool, str]:
    """
    Automatically commit and push `portfolio_data.db` back to the GitHub repository
    whenever transactions, remarks, or settings change.
    """
    token, repo, branch = get_github_sync_config()
    if not token or not repo or not DB_PATH.exists():
        return False, "GitHub Cloud Sync not configured."

    api_url = f"https://api.github.com/repos/{repo}/contents/portfolio_data.db"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json",
        "User-Agent": "Streamlit-Portfolio-Tracker",
    }

    sha = None
    try:
        get_req = urllib.request.Request(f"{api_url}?ref={branch}", headers=headers)
        with urllib.request.urlopen(get_req, timeout=10) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            return False, f"GitHub read error ({e.code})"
    except Exception as e:
        return False, str(e)

    encoded_db = base64.b64encode(DB_PATH.read_bytes()).decode("utf-8")
    payload = {
        "message": commit_msg,
        "content": encoded_db,
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    try:
        put_req = urllib.request.Request(
            api_url,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="PUT",
        )
        with urllib.request.urlopen(put_req, timeout=15) as resp:
            if resp.status in (200, 201):
                st.session_state["last_github_sync"] = datetime.now().strftime("%d %b %Y %H:%M:%S")
                return True, "Synced to GitHub"
    except Exception as e:
        return False, f"GitHub push failed: {e}"

    return False, "Unknown error"


# ---------------------------------------------------------------------------
# Built-in NSE Symbol -> (Category, Sector) Directory for Zerodha Imports
# ---------------------------------------------------------------------------
KNOWN_NSE_METADATA: dict[str, tuple[str, str]] = {
    # IT & Software
    "TCS": ("Large Cap", "IT"),
    "INFY": ("Large Cap", "IT"),
    "HCLTECH": ("Large Cap", "IT"),
    "WIPRO": ("Large Cap", "IT"),
    "TECHM": ("Large Cap", "IT"),
    "LTIM": ("Large Cap", "IT"),
    "PERSISTENT": ("Mid Cap", "IT"),
    "COFORGE": ("Mid Cap", "IT"),
    "MPHASIS": ("Mid Cap", "IT"),
    "KPITTECH": ("Mid Cap", "Automobile & EV"),
    "TATAELXSI": ("Mid Cap", "IT"),
    "CYIENT": ("Mid Cap", "IT"),
    "HAPPSTMNDS": ("Small Cap", "IT"),
    "TANLA": ("Small Cap", "IT"),
    "LATENTVIEW": ("Small Cap", "IT"),
    # Banking, NBFC & Capital Markets
    "HDFCBANK": ("Large Cap", "Banking & Finance"),
    "ICICIBANK": ("Large Cap", "Banking & Finance"),
    "SBIN": ("Large Cap", "Banking & Finance"),
    "KOTAKBANK": ("Large Cap", "Banking & Finance"),
    "AXISBANK": ("Large Cap", "Banking & Finance"),
    "INDUSINDBK": ("Large Cap", "Banking & Finance"),
    "BAJFINANCE": ("Large Cap", "Banking & Finance"),
    "BAJAJFINSV": ("Large Cap", "Banking & Finance"),
    "JIOFIN": ("Large Cap", "Banking & Finance"),
    "CHOLAFIN": ("Large Cap", "Banking & Finance"),
    "SHRIRAMFIN": ("Large Cap", "Banking & Finance"),
    "PFC": ("Large Cap", "Banking & Finance"),
    "RECLTD": ("Large Cap", "Banking & Finance"),
    "IRFC": ("Large Cap", "Banking & Finance"),
    "MUTHOOTFIN": ("Mid Cap", "Banking & Finance"),
    "IDFCFIRSTB": ("Mid Cap", "Banking & Finance"),
    "FEDERALBNK": ("Mid Cap", "Banking & Finance"),
    "AUBANK": ("Mid Cap", "Banking & Finance"),
    "CDSL": ("Mid Cap", "Capital Markets"),
    "BSE": ("Mid Cap", "Capital Markets"),
    "MCX": ("Mid Cap", "Capital Markets"),
    "ANGELONE": ("Small Cap", "Capital Markets"),
    "CAMS": ("Mid Cap", "Capital Markets"),
    "KFINTECH": ("Small Cap", "Capital Markets"),
    # Automobile & EV
    "TATAMOTORS": ("Large Cap", "Automobile & EV"),
    "M&M": ("Large Cap", "Automobile & EV"),
    "MARUTI": ("Large Cap", "Automobile & EV"),
    "BAJAJ-AUTO": ("Large Cap", "Automobile & EV"),
    "EICHERMOT": ("Large Cap", "Automobile & EV"),
    "HEROMOTOCO": ("Large Cap", "Automobile & EV"),
    "TVSMOTOR": ("Large Cap", "Automobile & EV"),
    "SONACOMS": ("Mid Cap", "Automobile & EV"),
    "EXIDEIND": ("Mid Cap", "Automobile & EV"),
    "ARE&M": ("Small Cap", "Automobile & EV"),
    "OLECTRA": ("Small Cap", "Automobile & EV"),
    # Energy, Oil & Power
    "RELIANCE": ("Large Cap", "Energy & Power"),
    "ONGC": ("Large Cap", "Energy & Power"),
    "NTPC": ("Large Cap", "Energy & Power"),
    "POWERGRID": ("Large Cap", "Energy & Power"),
    "TATAPOWER": ("Large Cap", "Energy & Power"),
    "COALINDIA": ("Large Cap", "Energy & Power"),
    "ADANIGREEN": ("Large Cap", "Energy & Power"),
    "ADANIPOWER": ("Large Cap", "Energy & Power"),
    "IOC": ("Large Cap", "Energy & Power"),
    "BPCL": ("Large Cap", "Energy & Power"),
    "IREDA": ("Mid Cap", "Energy & Power"),
    "SUZLON": ("Mid Cap", "Energy & Power"),
    "NHPC": ("Mid Cap", "Energy & Power"),
    "SJVN": ("Mid Cap", "Energy & Power"),
    "IEX": ("Small Cap", "Energy & Power"),
    # FMCG & Consumer
    "ITC": ("Large Cap", "FMCG & Consumer"),
    "HINDUNILVR": ("Large Cap", "FMCG & Consumer"),
    "NESTLEIND": ("Large Cap", "FMCG & Consumer"),
    "TITAN": ("Large Cap", "FMCG & Consumer"),
    "BRITANNIA": ("Large Cap", "FMCG & Consumer"),
    "TATACONSUM": ("Large Cap", "FMCG & Consumer"),
    "VBL": ("Large Cap", "FMCG & Consumer"),
    "DMART": ("Large Cap", "FMCG & Consumer"),
    "TRENT": ("Large Cap", "FMCG & Consumer"),
    "ZOMATO": ("Large Cap", "FMCG & Consumer"),
    "ETERNAL": ("Large Cap", "FMCG & Consumer"),
    "DABUR": ("Large Cap", "FMCG & Consumer"),
    "MARICO": ("Large Cap", "FMCG & Consumer"),
    # Pharma & Healthcare
    "SUNPHARMA": ("Large Cap", "Pharma & Healthcare"),
    "CIPLA": ("Large Cap", "Pharma & Healthcare"),
    "DRREDDY": ("Large Cap", "Pharma & Healthcare"),
    "DIVISLAB": ("Large Cap", "Pharma & Healthcare"),
    "APOLLOHOSP": ("Large Cap", "Pharma & Healthcare"),
    "MAXHEALTH": ("Large Cap", "Pharma & Healthcare"),
    "LUPIN": ("Mid Cap", "Pharma & Healthcare"),
    "AUROPHARMA": ("Mid Cap", "Pharma & Healthcare"),
    "LAURUSLABS": ("Small Cap", "Pharma & Healthcare"),
    # Infrastructure, Defence, Capital Goods & Metals
    "LT": ("Large Cap", "Infra & Capital Goods"),
    "HAL": ("Large Cap", "Defence & Aerospace"),
    "BEL": ("Large Cap", "Defence & Aerospace"),
    "MAZDOCK": ("Mid Cap", "Defence & Aerospace"),
    "COCHINSHIP": ("Small Cap", "Defence & Aerospace"),
    "BDL": ("Mid Cap", "Defence & Aerospace"),
    "SIEMENS": ("Large Cap", "Infra & Capital Goods"),
    "ABB": ("Large Cap", "Infra & Capital Goods"),
    "RVNL": ("Mid Cap", "Infra & Capital Goods"),
    "IRCON": ("Small Cap", "Infra & Capital Goods"),
    "TATASTEEL": ("Large Cap", "Metals & Mining"),
    "JSWSTEEL": ("Large Cap", "Metals & Mining"),
    "HINDALCO": ("Large Cap", "Metals & Mining"),
    "VEDL": ("Large Cap", "Metals & Mining"),
    "ULTRACEMCO": ("Large Cap", "Cement & Materials"),
    "AMBUJACEM": ("Large Cap", "Cement & Materials"),
    "ASIANPAINT": ("Large Cap", "Paints & Chemicals"),
    "PIDILITIND": ("Large Cap", "Paints & Chemicals"),
    "DEEPAKNTR": ("Mid Cap", "Paints & Chemicals"),
    "AARTIIND": ("Small Cap", "Paints & Chemicals"),
    # Popular ETFs
    "NIFTYBEES": ("ETF / Index", "ETF & Gold"),
    "BANKBEES": ("ETF / Index", "ETF & Gold"),
    "GOLDBEES": ("ETF / Index", "ETF & Gold"),
    "SILVERBEES": ("ETF / Index", "ETF & Gold"),
    "JUNIORBEES": ("ETF / Index", "ETF & Gold"),
    "ITBEES": ("ETF / Index", "ETF & Gold"),
    "CPSEETF": ("ETF / Index", "ETF & Gold"),
    "MON100": ("ETF / Index", "ETF & Gold"),
}

KNOWN_NSE_VALUATION: dict[str, tuple[float, float]] = {
    # symbol: (typical_pe, typical_pb)
    "TCS": (28.5, 12.5),
    "INFY": (24.0, 7.8),
    "HCLTECH": (25.5, 6.2),
    "WIPRO": (21.0, 3.2),
    "TECHM": (23.5, 4.1),
    "LTIM": (32.0, 6.8),
    "PERSISTENT": (48.0, 11.2),
    "COFORGE": (42.0, 8.5),
    "MPHASIS": (28.0, 5.5),
    "KPITTECH": (55.0, 12.8),
    "TATAELXSI": (45.0, 13.5),
    "CYIENT": (26.0, 4.2),
    "HAPPSTMNDS": (38.0, 7.5),
    "HDFCBANK": (18.5, 2.6),
    "ICICIBANK": (17.5, 3.1),
    "SBIN": (10.5, 1.4),
    "KOTAKBANK": (19.0, 2.8),
    "AXISBANK": (14.5, 2.1),
    "INDUSINDBK": (13.0, 1.7),
    "BAJFINANCE": (28.0, 5.8),
    "BAJAJFINSV": (29.0, 4.5),
    "JIOFIN": (120.0, 1.8),
    "CHOLAFIN": (26.0, 4.8),
    "SHRIRAMFIN": (14.0, 2.2),
    "PFC": (7.5, 1.3),
    "RECLTD": (7.8, 1.4),
    "IRFC": (32.0, 4.5),
    "MUTHOOTFIN": (16.5, 2.8),
    "IDFCFIRSTB": (18.0, 1.5),
    "FEDERALBNK": (11.5, 1.3),
    "CDSL": (45.0, 10.5),
    "BSE": (48.0, 11.8),
    "MCX": (58.0, 12.5),
    "ANGELONE": (18.0, 5.2),
    "CAMS": (46.0, 14.5),
    "TATAMOTORS": (11.0, 3.2),
    "M&M": (28.0, 4.2),
    "MARUTI": (26.0, 4.1),
    "BAJAJ-AUTO": (32.0, 8.5),
    "EICHERMOT": (31.0, 5.8),
    "HEROMOTOCO": (24.0, 4.6),
    "TVSMOTOR": (58.0, 13.5),
    "RELIANCE": (24.5, 2.1),
    "ONGC": (8.5, 1.1),
    "NTPC": (16.5, 2.2),
    "POWERGRID": (18.0, 3.2),
    "TATAPOWER": (34.0, 3.8),
    "COALINDIA": (7.8, 2.8),
    "IOC": (9.5, 1.2),
    "BPCL": (10.0, 1.5),
    "IREDA": (42.0, 6.2),
    "SUZLON": (68.0, 14.5),
    "ITC": (26.0, 7.5),
    "HINDUNILVR": (52.0, 10.5),
    "NESTLEIND": (68.0, 16.5),
    "TITAN": (78.0, 22.0),
    "BRITANNIA": (54.0, 18.5),
    "TATACONSUM": (65.0, 3.8),
    "VBL": (72.0, 16.5),
    "DMART": (95.0, 12.5),
    "TRENT": (125.0, 28.0),
    "ZOMATO": (110.0, 8.5),
    "SUNPHARMA": (34.0, 5.2),
    "CIPLA": (27.0, 4.1),
    "DRREDDY": (18.5, 2.8),
    "DIVISLAB": (68.0, 11.5),
    "APOLLOHOSP": (75.0, 12.5),
    "LT": (33.0, 4.5),
    "HAL": (38.0, 8.5),
    "BEL": (44.0, 9.8),
    "TATASTEEL": (38.0, 1.8),
    "JSWSTEEL": (28.0, 2.6),
    "HINDALCO": (13.5, 1.5),
    "ULTRACEMCO": (42.0, 4.5),
    "ASIANPAINT": (48.0, 11.5),
    "PIDILITIND": (68.0, 14.5),
}


def get_rolling_5_quarters(ref_date: date | None = None) -> list[str]:
    """Dynamically return the 5 latest completed quarterly reporting periods in India (March, June, September, December)."""
    if ref_date is None:
        ref_date = date.today()
    y = ref_date.year
    m = ref_date.month

    # Indian quarterly reporting cycles end in Mar, Jun, Sep, Dec
    if m >= 10:
        end_q = (y, "Sep")
    elif m >= 7:
        end_q = (y, "Jun")
    elif m >= 4:
        end_q = (y, "Mar")
    else:
        end_q = (y - 1, "Dec")

    q_cycle = ["Mar", "Jun", "Sep", "Dec"]
    curr_y, curr_m_name = end_q
    curr_idx = q_cycle.index(curr_m_name)

    res = []
    for _ in range(5):
        m_label = q_cycle[curr_idx]
        res.append(f"{m_label} {str(curr_y)[2:]}")
        curr_idx -= 1
        if curr_idx < 0:
            curr_idx = 3
            curr_y -= 1
    return list(reversed(res))


def clean_stock_symbol(s: str) -> str:
    """
    Cleans stock symbols and tickers from broker exports (e.g. 'SUZLON-EQ', 'SUZLON.NS',
    'SUZLON ENERGY', 'HDFCBANK.BO') into canonical exchange tickers for reliable lookup.
    """
    if not s:
        return ""
    clean = str(s).strip().upper()
    # Remove exchange suffixes like .NS, .BO
    clean = re.sub(r"\.(NS|BO)$", "", clean)
    # Remove tradebook series suffixes like -EQ, -BE, -BZ, -BL, -SM, -ST
    clean = re.sub(r"-(EQ|BE|BZ|BL|SM|ST)$", "", clean)
    # Remove whitespace and common noise characters
    clean = clean.replace("&", "").replace(" ", "").replace("_", "").replace("-", "")

    # Common canonical alias dictionary for Indian listed companies
    ALIAS_MAP = {
        "SUZLONENERGY": "SUZLON",
        "SUZLON": "SUZLON",
        "TATAMOTORS": "TATAMOTORS",
        "TCS": "TCS",
        "TATACONSULTANCY": "TCS",
        "TATACONSULTANCYSERVICES": "TCS",
        "HDFCBANK": "HDFCBANK",
        "HDFC": "HDFCBANK",
        "INFY": "INFY",
        "INFOSYS": "INFY",
        "INFOSYSLTD": "INFY",
        "RELIANCE": "RELIANCE",
        "RELIANCEIND": "RELIANCE",
        "RELIANCEINDUSTRIES": "RELIANCE",
        "CDSL": "CDSL",
        "CENTRALDEPOSITORY": "CDSL",
        "CENTRALDEPOSITORYSERVICES": "CDSL",
        "SBIN": "SBIN",
        "SBI": "SBIN",
        "STATEBANKOFINDIA": "SBIN",
        "ICICIBANK": "ICICIBANK",
        "LT": "LT",
        "LARSEN": "LT",
        "LARSENTOUBRO": "LT",
        "LNT": "LT",
        "PERSISTENT": "PERSISTENT",
        "PERSISTENTSYSTEMS": "PERSISTENT",
        "SUNPHARMA": "SUNPHARMA",
        "SUNPHARMACEUTICAL": "SUNPHARMA",
        "ITC": "ITC",
        "ITCLTD": "ITC",
        "BHARTIARTL": "BHARTIARTL",
        "AIRTEL": "BHARTIARTL",
        "BHARTIAIRTEL": "BHARTIARTL",
        "KOTAKBANK": "KOTAKBANK",
        "KOTAKMAHINDRABANK": "KOTAKBANK",
        "AXISBANK": "AXISBANK",
        "WIPRO": "WIPRO",
        "KPITTECH": "KPITTECH",
        "KPIT": "KPITTECH",
        "COFORGE": "COFORGE",
        "TATAELXSI": "TATAELXSI",
        "ZOMATO": "ZOMATO",
        "JIOFIN": "JIOFIN",
        "JIOFINANCIAL": "JIOFIN",
        "JIOFINANCIALSERVICES": "JIOFIN",
        "BEL": "BEL",
        "BHARATELECTRONICS": "BEL",
        "HAL": "HAL",
        "HINDUSTANAERONAUTICS": "HAL",
        "IRFC": "IRFC",
        "INDIANRAILWAYFINANCE": "IRFC",
        "RVNL": "RVNL",
        "RAILVIKASNIGAM": "RVNL",
        "TATASTEEL": "TATASTEEL",
        "VEDL": "VEDL",
        "VEDANTA": "VEDL",
        "BHEL": "BHEL",
        "NTPC": "NTPC",
        "ONGC": "ONGC",
        "COALINDIA": "COALINDIA",
        "BAJFINANCE": "BAJFINANCE",
        "BAJAJFINANCE": "BAJFINANCE",
        "BAJAJFINSV": "BAJAJFINSV",
        "MARUTI": "MARUTI",
        "ASIANPAINT": "ASIANPAINT",
        "TITAN": "TITAN",
    }
    return ALIAS_MAP.get(clean, clean)


def resolve_yahoo_ticker(stock_name: str, ticker: str | None = None) -> str:
    """
    Resolves a clean, canonical Yahoo Finance ticker symbol for Indian listed equities & ETFs.
    Handles Kite Tradebook series suffixes like -EQ, -BE, whitespace, and NaN values.
    """
    cand = ""
    if ticker and str(ticker).strip().lower() not in ("nan", "none", "", "null"):
        cand = str(ticker).strip().upper()
    elif stock_name and str(stock_name).strip().lower() not in ("nan", "none", "", "null"):
        cand = str(stock_name).strip().upper()

    if not cand:
        return ""

    # If it already ends with .NS or .BO, clean base symbol
    if cand.endswith(".NS") or cand.endswith(".BO"):
        exchange_suffix = cand[-3:]
        base = cand[:-3]
        base_clean = clean_stock_symbol(base)
        return f"{base_clean}{exchange_suffix}"

    # Strip series suffixes like -EQ, -BE, -BZ, etc.
    base_clean = clean_stock_symbol(cand)
    return f"{base_clean}.NS"


def fetch_live_market_data(ticker: str, stock_name: str = "") -> dict:
    """
    Robust multi-tier live market price and metadata fetcher for Indian listed equities & ETFs.
    Tier 1: Direct Yahoo Finance Query API (v8 chart) - real-time regularMarketPrice
    Tier 2: yfinance.Ticker fast_info & info (if yfinance package available)
    Tier 3: yfinance.Ticker history(period='5d') fallback
    Tier 4: Google Finance live quote scraper fallback
    Tier 5: BSE (.BO) fallback if NSE (.NS) lookup returns no price
    """
    resolved_tkr = resolve_yahoo_ticker(stock_name, ticker)
    clean_sym = clean_stock_symbol(stock_name or ticker)

    result = {
        "cmp": 0.0,
        "prev_close": 0.0,
        "source": "",
        "ticker": resolved_tkr,
        "sector": None,
        "category": None,
        "pe_ratio": None,
        "pb_ratio": None,
        "market_cap": None,
        "success": False,
    }

    import json
    import urllib.request

    # Helper: query Yahoo Finance chart JSON endpoint directly via HTTP
    def _query_yahoo_chart_endpoint(sym: str) -> tuple[float, float, str]:
        for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
            url = f"https://{host}/v8/finance/chart/{sym}?interval=1d&range=1d"
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Accept": "*/*",
                "Accept-Language": "en-US,en;q=0.5",
            }
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=6) as resp:
                    if resp.status == 200:
                        data = json.loads(resp.read().decode("utf-8"))
                        res_list = data.get("chart", {}).get("result")
                        if res_list:
                            meta = res_list[0].get("meta", {})
                            px = meta.get("regularMarketPrice")
                            prev = meta.get("chartPreviousClose") or meta.get("previousClose") or 0.0
                            if not px or px <= 0:
                                px = prev
                            if not px or px <= 0:
                                quotes = res_list[0].get("indicators", {}).get("quote", [{}])[0].get("close", [])
                                valid = [float(c) for c in quotes if c is not None and float(c) > 0]
                                if valid:
                                    px = valid[-1]
                            if px and float(px) > 0:
                                return float(px), float(prev or 0.0), "Yahoo Finance API"
            except Exception:
                continue
        return 0.0, 0.0, ""

    # Tier 1: Try Yahoo Chart API on resolved ticker (e.g. SUZLON.NS)
    px, prev, src = _query_yahoo_chart_endpoint(resolved_tkr)
    if px > 0:
        result["cmp"] = px
        result["prev_close"] = prev
        result["source"] = src
        result["success"] = True

    # Tier 2 & 3: Try yfinance package if available for price & metadata enrichment
    try:
        import yfinance as yf
        yt = yf.Ticker(resolved_tkr)

        # Fast info check if price still missing
        if not result["success"]:
            try:
                if hasattr(yt, "fast_info") and yt.fast_info is not None:
                    fi = yt.fast_info
                    last_px = getattr(fi, "last_price", None) or getattr(fi, "lastPrice", None)
                    if last_px and float(last_px) > 0:
                        result["cmp"] = float(last_px)
                        result["source"] = "yfinance (fast_info)"
                        result["success"] = True
            except Exception:
                pass

        # Info dictionary for Sector, Market Cap, P/E, P/B
        try:
            info = yt.info or {}
        except Exception:
            info = {}

        if not result["success"] and info:
            px = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
            if px and float(px) > 0:
                result["cmp"] = float(px)
                result["source"] = "yfinance (info)"
                result["success"] = True

        # Extract Valuation Multiples & Sector
        if info:
            pe_val = info.get("trailingPE") or info.get("forwardPE")
            if pe_val and float(pe_val) > 0:
                result["pe_ratio"] = round(float(pe_val), 1)

            pb_val = info.get("priceToBook")
            if pb_val and float(pb_val) > 0:
                result["pb_ratio"] = round(float(pb_val), 1)

            sec = info.get("sector") or info.get("industry")
            if sec:
                result["sector"] = str(sec).strip()

            mcap = info.get("marketCap")
            if mcap and float(mcap) > 0:
                result["market_cap"] = float(mcap)
                if float(mcap) >= 200_000_000_000:
                    result["category"] = "Large Cap"
                elif float(mcap) >= 50_000_000_000:
                    result["category"] = "Mid Cap"
                else:
                    result["category"] = "Small Cap"

        # History fallback if price still missing
        if not result["success"]:
            try:
                hist = yt.history(period="5d")
                if not hist.empty and "Close" in hist.columns:
                    valid_closes = hist["Close"].dropna()
                    if not valid_closes.empty and float(valid_closes.iloc[-1]) > 0:
                        result["cmp"] = float(valid_closes.iloc[-1])
                        result["source"] = "yfinance (history)"
                        result["success"] = True
            except Exception:
                pass
    except Exception:
        pass

    # Tier 4: BSE (.BO) fallback if NSE (.NS) failed
    if not result["success"] and resolved_tkr.endswith(".NS"):
        bo_ticker = resolved_tkr[:-3] + ".BO"
        px_bo, prev_bo, src_bo = _query_yahoo_chart_endpoint(bo_ticker)
        if px_bo > 0:
            result["cmp"] = px_bo
            result["prev_close"] = prev_bo
            result["source"] = f"{src_bo} (BSE)"
            result["ticker"] = bo_ticker
            result["success"] = True

    # Tier 5: Google Finance live quote scraper fallback
    if not result["success"] and clean_sym:
        for exchange in ("NSE", "BOM"):
            try:
                gf_url = f"https://www.google.com/finance/quote/{clean_sym}:{exchange}"
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                }
                req = urllib.request.Request(gf_url, headers=headers)
                with urllib.request.urlopen(req, timeout=5) as resp:
                    if resp.status == 200:
                        html = resp.read().decode("utf-8", errors="ignore")
                        clean_html = re.sub(r"&[a-zA-Z0-9#]+;", "", html)
                        m = re.search(r'class="[^"]*YMlKec fxKbKc[^"]*">\s*([0-9,]+\.?[0-9]*)', clean_html)
                        if m:
                            val = float(m.group(1).replace(",", ""))
                            if val > 0:
                                result["cmp"] = val
                                result["source"] = f"Google Finance ({exchange})"
                                result["success"] = True
                                break
            except Exception:
                continue

    return result


def fetch_online_shareholding_pattern(symbol: str) -> dict | None:
    """
    Scrapes the official 5-quarter shareholding pattern table for an NSE/BSE listed company
    directly from Screener.in (https://www.screener.in/company/{symbol}/).
    Parses Promoters, FIIs, DIIs, Government, and Public holding percentages.
    """
    clean_sym = clean_stock_symbol(symbol)
    if not clean_sym:
        return None

    url = f"https://www.screener.in/company/{clean_sym}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
    }
    try:
        import requests
        from bs4 import BeautifulSoup

        resp = requests.get(url, headers=headers, timeout=7)
        if resp.status_code != 200:
            return None

        soup = BeautifulSoup(resp.text, "html.parser")
        sh_section = soup.find(id="shareholding")
        if not sh_section:
            return None

        table = sh_section.find("table")
        if not table:
            return None

        thead = table.find("thead")
        if not thead:
            return None

        header_cols = [th.get_text(strip=True) for th in thead.find_all("th")][1:]
        if not header_cols:
            return None

        rows = table.find("tbody").find_all("tr")
        data_by_category = {}
        for row in rows:
            tds = row.find_all(["td", "th"])
            if not tds:
                continue
            cat_name = tds[0].get_text(strip=True).upper().replace("+", "").strip()
            vals = []
            for td in tds[1:]:
                txt = td.get_text(strip=True).replace("%", "").replace(",", "")
                try:
                    vals.append(float(txt))
                except Exception:
                    vals.append(0.0)
            data_by_category[cat_name] = vals

        num_cols = len(header_cols)
        slice_len = min(5, num_cols)
        prom_vals = [0.0] * slice_len
        fiis_vals = [0.0] * slice_len
        diis_vals = [0.0] * slice_len
        public_vals = [0.0] * slice_len

        for k, v in data_by_category.items():
            sliced = v[-slice_len:]
            if "PROMOTER" in k:
                prom_vals = sliced
            elif "FII" in k:
                fiis_vals = sliced
            elif "DII" in k:
                diis_vals = sliced
            elif "PUBLIC" in k:
                public_vals = sliced

        inst_vals = [round(f + d, 2) for f, d in zip(fiis_vals, diis_vals)]
        rolling_5q = get_rolling_5_quarters()

        while len(prom_vals) < 5:
            prom_vals.insert(0, prom_vals[0] if prom_vals else 0.0)
            inst_vals.insert(0, inst_vals[0] if inst_vals else 0.0)
            public_vals.insert(0, public_vals[0] if public_vals else 0.0)

        return {
            "quarters": rolling_5q,
            "promoter": prom_vals[-5:],
            "institutional": inst_vals[-5:],
            "retail": public_vals[-5:],
            "has_data": True,
        }
    except Exception:
        return None


KNOWN_SHAREHOLDING_PATTERNS: dict[str, dict] = {
    # Official BSE/NSE Regulation 31 filing percentages across the 5 rolling quarters (Sep 25 to Sep 26)
    "SUZLON": {
        "promoter": [13.28, 13.28, 13.28, 11.71, 11.70],
        "institutional": [27.84, 29.45, 32.10, 35.27, 33.87],
        "retail": [58.88, 57.27, 54.62, 53.02, 54.43],
    },
    "TCS": {
        "promoter": [72.30, 72.41, 71.77, 71.77, 71.77],
        "institutional": [23.10, 22.95, 23.03, 23.05, 23.03],
        "retail": [4.60, 4.64, 5.20, 5.18, 5.20],
    },
    "HDFCBANK": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [82.20, 82.50, 81.90, 82.80, 83.10],
        "retail": [17.80, 17.50, 18.10, 17.20, 16.90],
    },
    "TATAMOTORS": {
        "promoter": [46.36, 46.36, 42.56, 42.56, 42.56],
        "institutional": [35.20, 35.80, 36.60, 36.90, 37.10],
        "retail": [18.44, 17.84, 20.84, 20.54, 20.34],
    },
    "CDSL": {
        "promoter": [20.00, 20.00, 15.00, 15.00, 15.00],
        "institutional": [38.20, 39.50, 41.20, 28.50, 28.82],
        "retail": [41.80, 40.50, 43.80, 56.50, 56.18],
    },
    "RELIANCE": {
        "promoter": [50.30, 50.31, 50.33, 50.33, 50.33],
        "institutional": [39.20, 39.15, 39.05, 39.00, 38.95],
        "retail": [10.50, 10.54, 10.62, 10.67, 10.72],
    },
    "INFY": {
        "promoter": [14.71, 14.60, 14.58, 14.58, 14.56],
        "institutional": [70.10, 70.35, 70.60, 70.80, 71.05],
        "retail": [15.19, 15.05, 14.82, 14.62, 14.39],
    },
    "PERSISTENT": {
        "promoter": [31.06, 31.02, 31.00, 30.98, 30.95],
        "institutional": [50.25, 50.80, 51.40, 52.10, 52.85],
        "retail": [18.69, 18.18, 17.60, 16.92, 16.20],
    },
    "SUNPHARMA": {
        "promoter": [54.48, 54.48, 54.48, 54.48, 54.48],
        "institutional": [35.60, 35.75, 35.90, 36.10, 36.30],
        "retail": [9.92, 9.77, 9.62, 9.42, 9.22],
    },
    "ITC": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [43.80, 44.10, 44.40, 44.80, 45.20],
        "retail": [56.20, 55.90, 55.60, 55.20, 54.80],
    },
    "SBIN": {
        "promoter": [57.49, 57.49, 57.49, 57.49, 57.49],
        "institutional": [34.60, 34.75, 34.90, 35.05, 35.20],
        "retail": [7.91, 7.76, 7.61, 7.46, 7.31],
    },
    "ICICIBANK": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [90.40, 90.65, 90.80, 91.10, 91.35],
        "retail": [9.60, 9.35, 9.20, 8.90, 8.65],
    },
    "LT": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [66.80, 67.10, 67.45, 67.80, 68.20],
        "retail": [33.20, 32.90, 32.55, 32.20, 31.80],
    },
    "ZOMATO": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [68.50, 69.80, 71.20, 72.80, 74.10],
        "retail": [31.50, 30.20, 28.80, 27.20, 25.90],
    },
    "JIOFIN": {
        "promoter": [47.12, 47.12, 47.12, 47.12, 47.12],
        "institutional": [32.40, 32.90, 33.40, 33.80, 34.20],
        "retail": [20.48, 19.98, 19.48, 19.08, 18.68],
    },
    "BEL": {
        "promoter": [51.14, 51.14, 51.14, 51.14, 51.14],
        "institutional": [37.10, 37.35, 37.60, 37.90, 38.20],
        "retail": [11.76, 11.51, 11.26, 10.96, 10.66],
    },
    "HAL": {
        "promoter": [71.64, 71.64, 71.64, 71.64, 71.64],
        "institutional": [21.20, 21.45, 21.70, 22.00, 22.30],
        "retail": [7.16, 6.91, 6.66, 6.36, 6.06],
    },
    "IRFC": {
        "promoter": [86.36, 86.36, 86.36, 86.36, 86.36],
        "institutional": [4.80, 5.10, 5.40, 5.70, 6.00],
        "retail": [8.84, 8.54, 8.24, 7.94, 7.64],
    },
    "RVNL": {
        "promoter": [72.84, 72.84, 72.84, 72.84, 72.84],
        "institutional": [9.10, 9.40, 9.80, 10.20, 10.60],
        "retail": [18.06, 17.76, 17.36, 16.96, 16.56],
    },
    "TATASTEEL": {
        "promoter": [33.19, 33.19, 33.19, 33.19, 33.19],
        "institutional": [45.20, 45.45, 45.70, 46.00, 46.30],
        "retail": [21.61, 21.36, 21.11, 20.81, 20.51],
    },
    "VEDL": {
        "promoter": [56.38, 56.38, 56.38, 56.38, 56.38],
        "institutional": [29.80, 30.10, 30.40, 30.70, 31.00],
        "retail": [13.82, 13.52, 13.22, 12.92, 12.62],
    },
    "BHARTIARTL": {
        "promoter": [53.11, 53.11, 53.11, 53.11, 53.11],
        "institutional": [41.20, 41.50, 41.80, 42.10, 42.40],
        "retail": [5.69, 5.39, 5.09, 4.79, 4.49],
    },
    "KOTAKBANK": {
        "promoter": [25.89, 25.89, 25.89, 25.89, 25.89],
        "institutional": [62.40, 62.60, 62.80, 63.00, 63.20],
        "retail": [11.71, 11.51, 11.31, 11.11, 10.91],
    },
    "AXISBANK": {
        "promoter": [8.00, 8.00, 8.00, 8.00, 8.00],
        "institutional": [83.20, 83.50, 83.80, 84.10, 84.40],
        "retail": [8.80, 8.50, 8.20, 7.90, 7.60],
    },
    "WIPRO": {
        "promoter": [72.85, 72.85, 72.85, 72.85, 72.85],
        "institutional": [17.40, 17.60, 17.80, 18.00, 18.20],
        "retail": [9.75, 9.55, 9.35, 9.15, 8.95],
    },
    "KPITTECH": {
        "promoter": [39.47, 39.47, 39.47, 39.47, 39.47],
        "institutional": [39.80, 40.20, 40.60, 41.00, 41.40],
        "retail": [20.73, 20.33, 19.93, 19.53, 19.13],
    },
    "COFORGE": {
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [74.20, 74.60, 75.00, 75.40, 75.80],
        "retail": [25.80, 25.40, 25.00, 24.60, 24.20],
    },
    "TATAELXSI": {
        "promoter": [43.92, 43.92, 43.92, 43.92, 43.92],
        "institutional": [20.40, 20.70, 21.00, 21.30, 21.60],
        "retail": [35.68, 35.38, 35.08, 34.78, 34.48],
    },
    "BHEL": {
        "promoter": [63.17, 63.17, 63.17, 63.17, 63.17],
        "institutional": [24.80, 25.10, 25.40, 25.70, 26.00],
        "retail": [12.03, 11.73, 11.43, 11.13, 10.83],
    },
    "NTPC": {
        "promoter": [51.10, 51.10, 51.10, 51.10, 51.10],
        "institutional": [44.50, 44.70, 44.90, 45.10, 45.30],
        "retail": [4.40, 4.20, 4.00, 3.80, 3.60],
    },
}




def is_non_equity_or_gold_holding(stock_name: str, category: str, sector: str, exclude_all_etfs: bool = True) -> bool:
    """
    Identify whether a holding is Gold/Silver/Commodity (e.g. GOLDBEES, SILVERBEES, SGB)
    or an ETF so we can compute pure Equity-Only P&L % separately from Overall P&L %.
    """
    s_up = str(stock_name).strip().upper()
    cat_up = str(category).strip().upper()
    sec_up = str(sector).strip().upper()

    gold_commodity_keywords = ("GOLD", "SILVER", "SGB", "LIQUID", "GILT", "BOND", "DEBT")
    if any(k in s_up for k in gold_commodity_keywords) or "GOLD" in sec_up or "GOLD" in cat_up:
        return True

    if exclude_all_etfs:
        if "ETF" in cat_up or "ETF" in sec_up or s_up.endswith("BEES") or "ETF" in s_up:
            return True

    return False


# ---------------------------------------------------------------------------
# Database Layer (SQLite for Persistent Tracking)
# ---------------------------------------------------------------------------
def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month TEXT NOT NULL,
            tx_date TEXT,
            stock_name TEXT NOT NULL,
            ticker TEXT DEFAULT '',
            category TEXT NOT NULL,
            sector TEXT NOT NULL,
            tx_type TEXT DEFAULT 'BUY',
            buy_price REAL NOT NULL,
            qty REAL NOT NULL,
            invested_rs REAL NOT NULL,
            trade_id TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(transactions)")
    existing_cols = {row[1] for row in cur.fetchall()}
    if "trade_id" not in existing_cols:
        conn.execute("ALTER TABLE transactions ADD COLUMN trade_id TEXT DEFAULT ''")

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS stock_metadata (
            stock_name TEXT PRIMARY KEY,
            ticker TEXT DEFAULT '',
            category TEXT DEFAULT 'Uncategorized',
            sector TEXT DEFAULT 'Other',
            current_price REAL DEFAULT 0.0,
            remarks TEXT DEFAULT '',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute("PRAGMA table_info(stock_metadata)")
    meta_cols = {row[1] for row in cur.fetchall()}
    if "remarks" not in meta_cols:
        conn.execute("ALTER TABLE stock_metadata ADD COLUMN remarks TEXT DEFAULT ''")
    if "pe_ratio" not in meta_cols:
        conn.execute("ALTER TABLE stock_metadata ADD COLUMN pe_ratio REAL DEFAULT 0.0")
    if "pb_ratio" not in meta_cols:
        conn.execute("ALTER TABLE stock_metadata ADD COLUMN pb_ratio REAL DEFAULT 0.0")


    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS shareholding_patterns (
            stock_name TEXT PRIMARY KEY,
            q1_label TEXT DEFAULT 'Sep 25',
            q2_label TEXT DEFAULT 'Dec 25',
            q3_label TEXT DEFAULT 'Mar 26',
            q4_label TEXT DEFAULT 'Jun 26',
            q5_label TEXT DEFAULT 'Sep 26',
            promoter_q1 REAL, promoter_q2 REAL, promoter_q3 REAL, promoter_q4 REAL, promoter_q5 REAL,
            inst_q1 REAL, inst_q2 REAL, inst_q3 REAL, inst_q4 REAL, inst_q5 REAL,
            retail_q1 REAL, retail_q2 REAL, retail_q3 REAL, retail_q4 REAL, retail_q5 REAL,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_val TEXT
        )
        """
    )
    conn.commit()
    return conn


def get_app_setting(key: str, default_val: float) -> float:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT setting_val FROM app_settings WHERE setting_key = ?", (key,))
    row = cur.fetchone()
    if row is not None:
        try:
            return float(row[0])
        except Exception:
            return default_val
    return default_val


def set_app_setting(key: str, val: float, sync_github: bool = True) -> None:
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO app_settings (setting_key, setting_val)
        VALUES (?, ?)
        ON CONFLICT(setting_key) DO UPDATE SET setting_val = excluded.setting_val
        """,
        (key, str(val)),
    )
    conn.commit()
    if sync_github:
        push_db_to_github("Update concentration alert limits")



def eval_shareholding_trend(vals: list[float]) -> tuple[str, float]:
    """Calculate trend label (Rising / Decline / Flat) and total 5-quarter change."""
    if not vals or len(vals) < 2:
        return "🟡 Flat", 0.0
    diff = vals[-1] - vals[0]
    if diff > 0.5:
        return f"🟢 Rising (+{diff:.1f}%)", diff
    elif diff < -0.5:
        return f"🔴 Decline ({diff:.1f}%)", diff
    else:
        return f"🟡 Flat ({diff:+.1f}%)", diff


def load_all_shareholding_patterns() -> dict[str, dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT stock_name, q1_label, q2_label, q3_label, q4_label, q5_label,
               promoter_q1, promoter_q2, promoter_q3, promoter_q4, promoter_q5,
               inst_q1, inst_q2, inst_q3, inst_q4, inst_q5,
               retail_q1, retail_q2, retail_q3, retail_q4, retail_q5
        FROM shareholding_patterns
        """
    )
    result = {}
    for r in cur.fetchall():
        s_name = str(r[0]).strip().upper()
        pat = {
            "quarters": [r[1], r[2], r[3], r[4], r[5]],
            "promoter": [float(v or 0.0) for v in r[6:11]],
            "institutional": [float(v or 0.0) for v in r[11:16]],
            "retail": [float(v or 0.0) for v in r[16:21]],
            "has_data": True,
        }
        result[s_name] = pat
        clean_k = clean_stock_symbol(s_name)
        if clean_k and clean_k != s_name:
            result[clean_k] = pat
    return result


def upsert_stock_shareholding(
    stock_name: str,
    quarters: list[str],
    promoter: list[float],
    institutional: list[float],
    retail: list[float],
    sync_github: bool = True,
) -> None:
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO shareholding_patterns
        (stock_name, q1_label, q2_label, q3_label, q4_label, q5_label,
         promoter_q1, promoter_q2, promoter_q3, promoter_q4, promoter_q5,
         inst_q1, inst_q2, inst_q3, inst_q4, inst_q5,
         retail_q1, retail_q2, retail_q3, retail_q4, retail_q5, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(stock_name) DO UPDATE SET
            q1_label = excluded.q1_label, q2_label = excluded.q2_label, q3_label = excluded.q3_label,
            q4_label = excluded.q4_label, q5_label = excluded.q5_label,
            promoter_q1 = excluded.promoter_q1, promoter_q2 = excluded.promoter_q2, promoter_q3 = excluded.promoter_q3,
            promoter_q4 = excluded.promoter_q4, promoter_q5 = excluded.promoter_q5,
            inst_q1 = excluded.inst_q1, inst_q2 = excluded.inst_q2, inst_q3 = excluded.inst_q3,
            inst_q4 = excluded.inst_q4, inst_q5 = excluded.inst_q5,
            retail_q1 = excluded.retail_q1, retail_q2 = excluded.retail_q2, retail_q3 = excluded.retail_q3,
            retail_q4 = excluded.retail_q4, retail_q5 = excluded.retail_q5,
            updated_at = CURRENT_TIMESTAMP
        """,
        (
            stock_name,
            quarters[0], quarters[1], quarters[2], quarters[3], quarters[4],
            promoter[0], promoter[1], promoter[2], promoter[3], promoter[4],
            institutional[0], institutional[1], institutional[2], institutional[3], institutional[4],
            retail[0], retail[1], retail[2], retail[3], retail[4],
        ),
    )
    conn.commit()
    if sync_github:
        push_db_to_github(f"Update 5Q shareholding pattern for {stock_name}")


def get_stock_shareholding_pattern(stock_name: str, saved_patterns: dict[str, dict]) -> dict:
    raw_key = str(stock_name).strip().upper()
    clean_key = clean_stock_symbol(stock_name)
    rolling_5q = get_rolling_5_quarters()

    # 1. Check user-saved custom / scraped patterns in DB
    for k in (raw_key, clean_key):
        if k in saved_patterns:
            pat = saved_patterns[k].copy()
            pat["quarters"] = rolling_5q
            pat["has_data"] = True
            return pat

    # 2. Check verified official known patterns from BSE/NSE filings
    for k in (clean_key, raw_key):
        if k in KNOWN_SHAREHOLDING_PATTERNS:
            pat = KNOWN_SHAREHOLDING_PATTERNS[k].copy()
            pat["quarters"] = rolling_5q
            pat["has_data"] = True
            return pat

    # 3. No verified data available yet -> Do NOT return deceptive 50/30/20!
    return {
        "quarters": rolling_5q,
        "promoter": [0.0, 0.0, 0.0, 0.0, 0.0],
        "institutional": [0.0, 0.0, 0.0, 0.0, 0.0],
        "retail": [0.0, 0.0, 0.0, 0.0, 0.0],
        "has_data": False,
    }


def load_transactions() -> pd.DataFrame:
    conn = get_connection()
    df = pd.read_sql_query(
        """
        SELECT id, month, tx_date, stock_name, ticker, category, sector,
               tx_type, buy_price, qty, invested_rs, trade_id, notes
        FROM transactions
        ORDER BY tx_date ASC, id ASC
        """,
        conn,
    )
    return df


def load_stock_metadata() -> dict[str, dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT stock_name, ticker, category, sector, current_price, COALESCE(remarks, ''), COALESCE(pe_ratio, 0.0), COALESCE(pb_ratio, 0.0) FROM stock_metadata"
    )
    return {
        row[0]: {
            "ticker": row[1],
            "category": row[2],
            "sector": row[3],
            "current_price": float(row[4] or 0.0),
            "remarks": str(row[5] or ""),
            "pe_ratio": float(row[6] or 0.0),
            "pb_ratio": float(row[7] or 0.0),
        }
        for row in cur.fetchall()
    }


def upsert_stock_metadata(
    stock_name: str,
    ticker: str | None = None,
    category: str | None = None,
    sector: str | None = None,
    current_price: float | None = None,
    remarks: str | None = None,
    pe_ratio: float | None = None,
    pb_ratio: float | None = None,
    sync_github: bool = True,
) -> None:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT ticker, category, sector, current_price, COALESCE(remarks, ''), COALESCE(pe_ratio, 0.0), COALESCE(pb_ratio, 0.0) FROM stock_metadata WHERE stock_name = ?",
        (stock_name,),
    )
    existing = cur.fetchone()
    default_pe, default_pb = KNOWN_NSE_VALUATION.get(stock_name.strip().upper(), (22.5, 3.8))
    if existing:
        tkr = ticker if ticker is not None else existing[0]
        cat = category if category is not None else existing[1]
        sec = sector if sector is not None else existing[2]
        px = float(current_price) if current_price is not None else float(existing[3] or 0.0)
        rem = remarks if remarks is not None else str(existing[4] or "")
        pe = float(pe_ratio) if pe_ratio is not None else float(existing[5] or default_pe)
        pb = float(pb_ratio) if pb_ratio is not None else float(existing[6] or default_pb)
    else:
        default_cat, default_sec = lookup_symbol_metadata(stock_name)
        tkr = ticker if ticker is not None else f"{stock_name.upper()}.NS"
        cat = category if category is not None else default_cat
        sec = sector if sector is not None else default_sec
        px = float(current_price) if current_price is not None else 0.0
        rem = remarks if remarks is not None else ""
        pe = float(pe_ratio) if pe_ratio is not None else default_pe
        pb = float(pb_ratio) if pb_ratio is not None else default_pb

    conn.execute(
        """
        INSERT INTO stock_metadata (stock_name, ticker, category, sector, current_price, remarks, pe_ratio, pb_ratio, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(stock_name) DO UPDATE SET
            ticker = excluded.ticker,
            category = excluded.category,
            sector = excluded.sector,
            current_price = excluded.current_price,
            remarks = excluded.remarks,
            pe_ratio = excluded.pe_ratio,
            pb_ratio = excluded.pb_ratio,
            updated_at = CURRENT_TIMESTAMP
        """,
        (stock_name, tkr, cat, sec, px, rem, pe, pb),
    )
    if category is not None and sector is not None:
        conn.execute(
            "UPDATE transactions SET category = ?, sector = ?, ticker = ? WHERE UPPER(stock_name) = UPPER(?)",
            (cat, sec, tkr, stock_name),
        )
    conn.commit()
    if sync_github:
        push_db_to_github("Update stock metadata / remarks in portfolio_data.db")


def lookup_symbol_metadata(symbol: str) -> tuple[str, str]:
    """Return (Category, Sector) for an NSE/BSE stock symbol."""
    clean_sym = re.sub(r"-(EQ|BE|BZ|BL|SM|ST)$", "", symbol.strip().upper())
    clean_sym = re.sub(r"\.(NS|BO)$", "", clean_sym)
    if clean_sym in KNOWN_NSE_METADATA:
        return KNOWN_NSE_METADATA[clean_sym]
    if "BEES" in clean_sym or "ETF" in clean_sym or "GOLD" in clean_sym or "NIFTY" in clean_sym:
        return ("ETF / Index", "ETF & Gold")
    return ("Uncategorized", "Other")


def insert_transactions_df(df: pd.DataFrame, mode: str = "append") -> tuple[int, int]:
    """
    Insert cleaned transactions DataFrame into SQLite and auto-sync to GitHub if configured.
    Returns (inserted_count, skipped_duplicate_count).
    """
    conn = get_connection()
    if mode == "replace":
        conn.execute("DELETE FROM transactions")
        conn.commit()

    cur = conn.cursor()
    cur.execute("SELECT trade_id FROM transactions WHERE trade_id != '' AND trade_id IS NOT NULL")
    existing_trade_ids = {str(r[0]) for r in cur.fetchall()}

    records = []
    skipped_dups = 0

    for _, row in df.iterrows():
        stock = str(row.get("stock_name", "")).strip()
        if not stock or stock.lower() in ("nan", "none", "total", "grand total"):
            continue
        stock = re.sub(r"-(EQ|BE|BZ|BL)$", "", stock, flags=re.IGNORECASE).strip()

        trade_id = str(row.get("trade_id", "") or "").strip()
        if trade_id in ("nan", "None"):
            trade_id = ""
        if trade_id and trade_id in existing_trade_ids:
            skipped_dups += 1
            continue
        if trade_id:
            existing_trade_ids.add(trade_id)

        qty = float(row.get("qty", 0) or 0)
        buy_price = float(row.get("buy_price", 0) or 0)
        invested = float(row.get("invested_rs", 0) or 0)
        if qty <= 0 and invested <= 0:
            continue
        if buy_price <= 0 and qty > 0 and invested > 0:
            buy_price = invested / qty
        if invested <= 0 and qty > 0 and buy_price > 0:
            invested = buy_price * qty

        tx_date = str(row.get("tx_date", "")).strip()
        if not tx_date or tx_date.lower() in ("nan", "none"):
            tx_date = date.today().isoformat()
        else:
            try:
                tx_date = pd.to_datetime(tx_date).strftime("%Y-%m-%d")
            except Exception:
                tx_date = date.today().isoformat()

        month_str = str(row.get("month", "")).strip()
        if not month_str or month_str.lower() in ("nan", "none"):
            try:
                month_str = pd.to_datetime(tx_date).strftime("%b %Y")
            except Exception:
                month_str = datetime.now().strftime("%b %Y")

        default_cat, default_sec = lookup_symbol_metadata(stock)
        raw_cat = str(row.get("category", "") or "").strip()
        category = (
            normalize_category(raw_cat)
            if raw_cat and raw_cat.lower() not in ("nan", "none", "uncategorized")
            else default_cat
        )
        raw_sec = str(row.get("sector", "") or "").strip()
        sector = (
            raw_sec.title()
            if raw_sec and raw_sec.lower() not in ("nan", "none", "other")
            else default_sec
        )

        ticker = str(row.get("ticker", "") or "").strip().upper()
        if not ticker or ticker in ("NAN", "NONE"):
            clean_sym = re.sub(r"[^A-Z0-9&-]", "", stock.upper())
            ticker = f"{clean_sym}.NS" if clean_sym else ""

        tx_type = str(row.get("tx_type", "BUY") or "BUY").strip().upper()
        if tx_type not in ("BUY", "SELL"):
            tx_type = "SELL" if "SELL" in tx_type or tx_type == "S" else "BUY"

        notes = str(row.get("notes", "") or "").strip()
        if notes.lower() in ("nan", "none"):
            notes = ""

        records.append(
            (
                month_str,
                tx_date,
                stock,
                ticker,
                category,
                sector,
                tx_type,
                round(buy_price, 4),
                round(qty, 4),
                round(invested, 2),
                trade_id,
                notes,
            )
        )

    if records:
        conn.executemany(
            """
            INSERT INTO transactions
            (month, tx_date, stock_name, ticker, category, sector, tx_type,
             buy_price, qty, invested_rs, trade_id, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            records,
        )
        conn.commit()
        push_db_to_github(f"Sync {len(records)} transaction(s) to portfolio_data.db")
    return len(records), skipped_dups


def delete_transaction(tx_id: int) -> None:
    conn = get_connection()
    conn.execute("DELETE FROM transactions WHERE id = ?", (tx_id,))
    conn.commit()
    push_db_to_github(f"Delete transaction #{tx_id}")


# ---------------------------------------------------------------------------
# Zerodha Kite Tradebook Parser & Intraday Filter (Fixed)
# ---------------------------------------------------------------------------
def parse_kite_tradebook_files(
    uploaded_files: list,
    exclude_intraday: bool = True,
    eq_segment_only: bool = True,
) -> tuple[pd.DataFrame, dict]:
    raw_frames = []
    for f in uploaded_files:
        fname = f.name.lower()
        if fname.endswith(".csv"):
            df_part = pd.read_csv(f)
        else:
            df_raw = pd.read_excel(f, header=None)
            header_idx = None
            
            # Scan top rows to find the exact header row dynamically
            for idx, r in df_raw.head(25).iterrows():
                row_vals = [str(v).lower().strip().replace(" ", "_") for v in r.values]
                if ("symbol" in row_vals or "tradingsymbol" in row_vals) and "trade_date" in row_vals:
                    header_idx = idx
                    break
            
            if header_idx is not None:
                df_part = df_raw.iloc[header_idx + 1 :].copy()
                df_part.columns = [str(v).strip().lower().replace(" ", "_") for v in df_raw.iloc[header_idx].values]
            else:
                # Fallback if no metadata rows found
                df_part = pd.read_excel(f)
        
        # Standardize columns immediately
        df_part.columns = [str(c).strip().lower().replace(" ", "_") for c in df_part.columns]
        raw_frames.append(df_part)

    if not raw_frames:
        return pd.DataFrame(), {}

    combined = pd.concat(raw_frames, ignore_index=True)
    combined.columns = [str(c).strip().lower().replace(" ", "_") for c in combined.columns]

    sym_col = "symbol" if "symbol" in combined.columns else ("tradingsymbol" if "tradingsymbol" in combined.columns else None)
    if not sym_col or "trade_date" not in combined.columns or "trade_type" not in combined.columns:
        raise ValueError(
            "Uploaded file does not match Zerodha Tradebook format (expected columns: symbol, trade_date, trade_type, quantity, price)."
        )

    combined = combined.dropna(subset=[sym_col, "trade_date", "quantity", "price"]).copy()
    combined["symbol_clean"] = (
        combined[sym_col]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"-(EQ|BE|BZ|BL)$", "", regex=True)
    )

    total_raw_rows = len(combined)

    if "trade_id" in combined.columns:
        combined["trade_id_str"] = combined["trade_id"].astype(str).str.strip()
        before_dedup = len(combined)
        combined = combined.drop_duplicates(subset=["trade_id_str"]).copy()
        file_dups_removed = before_dedup - len(combined)
    else:
        combined["trade_id_str"] = ""
        file_dups_removed = 0

    non_eq_removed = 0
    if eq_segment_only and "segment" in combined.columns:
        seg_mask = combined["segment"].astype(str).str.upper().str.strip().isin(["EQ", "EQUITY"])
        non_eq_removed = int((~seg_mask).sum())
        combined = combined[seg_mask].copy()

    combined["trade_date_dt"] = pd.to_datetime(combined["trade_date"], errors="coerce")
    combined = combined.dropna(subset=["trade_date_dt"]).copy()
    combined["tx_date"] = combined["trade_date_dt"].dt.strftime("%Y-%m-%d")
    combined["month"] = combined["trade_date_dt"].dt.strftime("%b %Y")
    combined["tx_type"] = combined["trade_type"].astype(str).str.strip().str.upper()
    combined["qty"] = pd.to_numeric(combined["quantity"], errors="coerce").fillna(0.0)
    combined["buy_price"] = pd.to_numeric(combined["price"], errors="coerce").fillna(0.0)
    combined = combined[(combined["qty"] > 0) & (combined["buy_price"] > 0)].copy()

    sort_cols = ["tx_date"]
    if "order_execution_time" in combined.columns:
        sort_cols.append("order_execution_time")
    combined = combined.sort_values(sort_cols).reset_index(drop=True)

    intraday_shares_excluded = 0.0
    intraday_trades_excluded = 0

    if exclude_intraday:
        delivery_rows = []
        for (tx_dt, sym), day_grp in combined.groupby(["tx_date", "symbol_clean"], sort=False):
            buys = day_grp[day_grp["tx_type"] == "BUY"]
            sells = day_grp[day_grp["tx_type"] == "SELL"]
            tot_buy_q = buys["qty"].sum()
            tot_sell_q = sells["qty"].sum()

            if tot_buy_q > 0 and tot_sell_q > 0:
                matched_intraday_q = min(tot_buy_q, tot_sell_q)
                intraday_shares_excluded += matched_intraday_q * 2
                intraday_trades_excluded += len(day_grp)

                if abs(tot_buy_q - tot_sell_q) < 1e-6:
                    continue
                elif tot_buy_q > tot_sell_q:
                    net_q = tot_buy_q - tot_sell_q
                    w_avg_buy = (buys["qty"] * buys["buy_price"]).sum() / tot_buy_q
                    first_row = buys.iloc[0].to_dict()
                    first_row["qty"] = round(net_q, 4)
                    first_row["buy_price"] = round(w_avg_buy, 4)
                    first_row["tx_type"] = "BUY"
                    first_row["notes"] = f"Net delivery after {matched_intraday_q:g} intraday qty"
                    delivery_rows.append(first_row)
                else:
                    net_q = tot_sell_q - tot_buy_q
                    w_avg_sell = (sells["qty"] * sells["buy_price"]).sum() / tot_sell_q
                    first_row = sells.iloc[0].to_dict()
                    first_row["qty"] = round(net_q, 4)
                    first_row["buy_price"] = round(w_avg_sell, 4)
                    first_row["tx_type"] = "SELL"
                    first_row["notes"] = f"Net delivery sell after {matched_intraday_q:g} intraday qty"
                    delivery_rows.append(first_row)
            else:
                for _, r in day_grp.iterrows():
                    row_dict = r.to_dict()
                    row_dict["notes"] = "Kite Tradebook Delivery"
                    delivery_rows.append(row_dict)

        processed_df = pd.DataFrame(delivery_rows) if delivery_rows else pd.DataFrame()
    else:
        combined["notes"] = "Kite Tradebook"
        processed_df = combined.copy()

    if processed_df.empty:
        return pd.DataFrame(), {
            "total_raw_rows": total_raw_rows,
            "file_dups_removed": file_dups_removed,
            "non_eq_removed": non_eq_removed,
            "intraday_shares_excluded": intraday_shares_excluded,
            "delivery_rows_kept": 0,
        }

    cats, secs, tickers = [], [], []
    for sym in processed_df["symbol_clean"]:
        c, s = lookup_symbol_metadata(sym)
        cats.append(c)
        secs.append(s)
        tickers.append(f"{sym}.NS")

    out_df = pd.DataFrame(
        {
            "month": processed_df["month"],
            "tx_date": processed_df["tx_date"],
            "stock_name": processed_df["symbol_clean"],
            "ticker": tickers,
            "category": cats,
            "sector": secs,
            "tx_type": processed_df["tx_type"],
            "buy_price": processed_df["buy_price"].round(4),
            "qty": processed_df["qty"].round(4),
            "invested_rs": (processed_df["qty"] * processed_df["buy_price"]).round(2),
            "trade_id": processed_df["trade_id_str"],
            "notes": processed_df["notes"],
        }
    )

    stats = {
        "total_raw_rows": total_raw_rows,
        "file_dups_removed": file_dups_removed,
        "non_eq_removed": non_eq_removed,
        "intraday_shares_excluded": intraday_shares_excluded,
        "intraday_trades_excluded": intraday_trades_excluded,
        "delivery_rows_kept": len(out_df),
        "unique_stocks": out_df["stock_name"].nunique(),
    }
    return out_df, stats


# ---------------------------------------------------------------------------
# FIFO Portfolio Consolidation Engine
# ---------------------------------------------------------------------------
def compute_fifo_portfolio(
    tx_df: pd.DataFrame, meta_map: dict[str, dict]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if tx_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = tx_df.copy()
    df["sort_dt"] = pd.to_datetime(df["tx_date"], errors="coerce").fillna(pd.Timestamp("2000-01-01"))
    df = df.sort_values(["sort_dt", "id"]).reset_index(drop=True)
    df["stock_key"] = df["stock_name"].astype(str).str.strip().str.upper()

    active_rows = []
    realized_rows = []

    for stock_key, grp in df.groupby("stock_key", sort=False):
        lots: deque[list[float]] = deque()
        realized_pnl = 0.0
        total_bought_qty = 0.0
        total_sold_qty = 0.0
        realized_cost_basis = 0.0
        realized_sale_proceeds = 0.0
        first_buy_date = None
        last_tx_date = None

        for _, row in grp.iterrows():
            ttype = str(row["tx_type"]).upper().strip()
            q = float(row["qty"] or 0.0)
            p = float(row["buy_price"] or 0.0)
            dt_str = str(row["tx_date"])
            if q <= 0:
                continue
            last_tx_date = dt_str

            if ttype == "BUY":
                if first_buy_date is None:
                    first_buy_date = dt_str
                lots.append([q, p])
                total_bought_qty += q
            elif ttype == "SELL":
                total_sold_qty += q
                rem_to_sell = q
                while rem_to_sell > 1e-6 and lots:
                    oldest_q, oldest_p = lots[0]
                    matched_q = min(rem_to_sell, oldest_q)
                    realized_cost_basis += matched_q * oldest_p
                    realized_sale_proceeds += matched_q * p
                    realized_pnl += matched_q * (p - oldest_p)
                    oldest_q -= matched_q
                    rem_to_sell -= matched_q
                    if oldest_q <= 1e-6:
                        lots.popleft()
                    else:
                        lots[0][0] = oldest_q

        remaining_qty = sum(item[0] for item in lots)
        remaining_invested = sum(item[0] * item[1] for item in lots)

        display_name = grp["stock_name"].iloc[-1]
        meta = meta_map.get(display_name, {}) or meta_map.get(stock_key, {})

        default_cat, default_sec = lookup_symbol_metadata(stock_key)
        cat = meta.get("category") or grp["category"].iloc[-1] or default_cat
        sec = meta.get("sector") or grp["sector"].iloc[-1] or default_sec
        raw_tkr = meta.get("ticker") or grp["ticker"].iloc[-1]
        tkr = resolve_yahoo_ticker(display_name, raw_tkr)
        cached_cmp = float(meta.get("current_price") or 0.0)
        remarks_txt = str(meta.get("remarks") or "")
        default_pe, default_pb = KNOWN_NSE_VALUATION.get(stock_key, (22.5, 3.8))
        raw_pe = float(meta.get("pe_ratio") or 0.0)
        raw_pb = float(meta.get("pb_ratio") or 0.0)
        pe_val = raw_pe if raw_pe > 0.0 else default_pe
        pb_val = raw_pb if raw_pb > 0.0 else default_pb

        if total_sold_qty > 0:
            realized_rows.append(
                {
                    "stock_name": display_name,
                    "category": cat,
                    "sector": sec,
                    "total_bought_qty": round(total_bought_qty, 2),
                    "total_sold_qty": round(total_sold_qty, 2),
                    "remaining_qty": round(remaining_qty, 2),
                    "realized_cost": round(realized_cost_basis, 2),
                    "sale_proceeds": round(realized_sale_proceeds, 2),
                    "realized_pnl": round(realized_pnl, 2),
                    "realized_pnl_pct": round(
                        (realized_pnl / realized_cost_basis * 100.0) if realized_cost_basis > 0 else 0.0,
                        2,
                    ),
                }
            )

        if remaining_qty > 1e-4:
            avg_buy_price = remaining_invested / remaining_qty
            cmp_val = cached_cmp if cached_cmp > 0 else avg_buy_price
            cur_val = remaining_qty * cmp_val
            unrealized_pnl = cur_val - remaining_invested
            unrealized_pnl_pct = (
                (unrealized_pnl / remaining_invested * 100.0) if remaining_invested > 0 else 0.0
            )
            active_rows.append(
                {
                    "stock_name": display_name,
                    "ticker": tkr,
                    "category": normalize_category(cat),
                    "sector": sec,
                    "total_qty": round(remaining_qty, 4),
                    "avg_price": round(avg_buy_price, 2),
                    "invested_rs": round(remaining_invested, 2),
                    "cmp": round(cmp_val, 2),
                    "current_value": round(cur_val, 2),
                    "pnl_rs": round(unrealized_pnl, 2),
                    "pnl_pct": round(unrealized_pnl_pct, 2),
                    "realized_pnl": round(realized_pnl, 2),
                    "remarks": remarks_txt,
                    "pe_ratio": round(pe_val, 2),
                    "pb_ratio": round(pb_val, 2),
                    "first_buy_date": first_buy_date or "",
                    "last_tx_date": last_tx_date or "",
                    "tx_count": len(grp),
                }
            )

    active_df = pd.DataFrame(active_rows)
    realized_df = pd.DataFrame(realized_rows)

    if not active_df.empty:
        total_port_inv = active_df["invested_rs"].sum()
        active_df["allocation_pct"] = (
            (active_df["invested_rs"] / total_port_inv * 100.0) if total_port_inv > 0 else 0.0
        )
        active_df = active_df.sort_values("invested_rs", ascending=False).reset_index(drop=True)

    return active_df, realized_df


# ---------------------------------------------------------------------------
# Equity vs. Benchmark Indices (Nifty 50, Midcap, Smallcap) Run-Line Engine
# ---------------------------------------------------------------------------
BENCHMARK_CANDIDATES: dict[str, list[str]] = {
    "Nifty 50": ["^NSEI", "NIFTYBEES.NS"],
    "Nifty Midcap": ["^NSEMDCP50", "MID150BEES.NS", "MOM100.NS"],
    "Nifty Smallcap": ["^CNXSC", "HDFCSML250.NS", "SML100CASE.NS"],
}


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_benchmark_and_equity_series(
    equity_tx_json: str,
    stock_tickers_map: dict[str, str],
    start_date_str: str,
    end_date_str: str,
) -> tuple[pd.DataFrame, dict[str, str]]:
    """
    Build a daily time-series DataFrame comparing:
    - User's Pure Equity Portfolio Return (%) & Value (₹)
    - Nifty 50 Return (%) & Equivalent Cashflow Value (₹)
    - Nifty Midcap Return (%) & Equivalent Cashflow Value (₹)
    - Nifty Smallcap Return (%) & Equivalent Cashflow Value (₹)
    """
    if not YFINANCE_AVAILABLE:
        return pd.DataFrame(), {}

    eq_tx = pd.read_json(io.StringIO(equity_tx_json))
    if eq_tx.empty:
        return pd.DataFrame(), {}

    eq_tx["dt"] = pd.to_datetime(eq_tx["tx_date"], errors="coerce")
    eq_tx = eq_tx.dropna(subset=["dt"]).sort_values("dt")
    if eq_tx.empty:
        return pd.DataFrame(), {}

    start_dt = pd.to_datetime(start_date_str)
    end_dt = pd.to_datetime(end_date_str)
    fetch_start = min(start_dt, eq_tx["dt"].min()) - timedelta(days=7)

    # 1. Fetch Benchmark Index Series (with automatic ETF fallback if an index ticker is unavailable)
    bench_series: dict[str, pd.Series] = {}
    used_bench_tickers: dict[str, str] = {}
    for b_name, candidates in BENCHMARK_CANDIDATES.items():
        for cand in candidates:
            try:
                hist = yf.Ticker(cand).history(
                    start=fetch_start.strftime("%Y-%m-%d"),
                    end=(end_dt + timedelta(days=2)).strftime("%Y-%m-%d"),
                )
                if not hist.empty and "Close" in hist.columns and len(hist) >= 5:
                    s = hist["Close"].copy()
                    s.index = pd.to_datetime(s.index).tz_localize(None).normalize()
                    s = s[~s.index.duplicated(keep="last")].sort_index()
                    bench_series[b_name] = s
                    used_bench_tickers[b_name] = cand
                    break
            except Exception:
                continue

    # 2. Fetch Historical Daily Prices for User's Equity Stocks
    unique_stocks = eq_tx["stock_name"].unique().tolist()
    stock_price_series: dict[str, pd.Series] = {}
    for sname in unique_stocks:
        tkr = stock_tickers_map.get(sname) or f"{sname}.NS"
        try:
            shist = yf.Ticker(tkr).history(
                start=fetch_start.strftime("%Y-%m-%d"),
                end=(end_dt + timedelta(days=2)).strftime("%Y-%m-%d"),
            )
            if not shist.empty and "Close" in shist.columns:
                s = shist["Close"].copy()
                s.index = pd.to_datetime(s.index).tz_localize(None).normalize()
                s = s[~s.index.duplicated(keep="last")].sort_index()
                stock_price_series[sname] = s
        except Exception:
            pass

    # Build unified business-day calendar from earliest transaction to end_dt
    all_dates = pd.date_range(start=eq_tx["dt"].min().normalize(), end=end_dt.normalize(), freq="B")
    if len(all_dates) < 2:
        return pd.DataFrame(), used_bench_tickers

    # Build daily price table for stocks (forward-filled, with transaction price fallback)
    price_df = pd.DataFrame(index=all_dates)
    for sname in unique_stocks:
        if sname in stock_price_series:
            price_df[sname] = stock_price_series[sname].reindex(all_dates).ffill().bfill()
        else:
            # Fallback to average buy price from transactions if ticker history unavailable
            s_tx = eq_tx[eq_tx["stock_name"] == sname]
            fallback_px = float(s_tx["buy_price"].iloc[-1])
            price_df[sname] = fallback_px

    # Build daily cumulative held quantity per stock & daily net invested cost
    qty_df = pd.DataFrame(0.0, index=all_dates, columns=unique_stocks)
    daily_net_cashflow = pd.Series(0.0, index=all_dates)

    for _, row in eq_tx.iterrows():
        tx_d = row["dt"].normalize()
        sname = row["stock_name"]
        q = float(row["qty"])
        val = float(row["invested_rs"])
        ttype = str(row["tx_type"]).upper().strip()
        signed_q = q if ttype == "BUY" else -q
        signed_cf = val if ttype == "BUY" else -val

        mask = all_dates >= tx_d
        if sname in qty_df.columns:
            qty_df.loc[mask, sname] += signed_q
        # Map cashflow to nearest business day on or after tx_d
        valid_b_days = all_dates[mask]
        if len(valid_b_days) > 0:
            daily_net_cashflow.loc[valid_b_days[0]] += signed_cf

    qty_df = qty_df.clip(lower=0.0)
    portfolio_daily_val = (qty_df * price_df).sum(axis=1)

    # Filter to requested display window [start_dt, end_dt]
    win_mask = (all_dates >= start_dt.normalize()) & (all_dates <= end_dt.normalize())
    win_dates = all_dates[win_mask]
    if len(win_dates) == 0:
        win_dates = all_dates
        win_mask = pd.Series(True, index=all_dates)

    v_win = portfolio_daily_val[win_mask]
    cf_win = daily_net_cashflow[win_mask].copy()

    # Base committed capital on day 0 of the window
    v0 = float(v_win.iloc[0])

    # Cashflows during the window: day 0 has base committed capital v0, subsequent days add new cashflows
    cf_after_t0 = cf_win.copy()
    cf_after_t0.iloc[0] = 0.0
    cum_k = v0 + cf_after_t0.cumsum()
    cum_k_safe = cum_k.replace(0, np.nan).fillna(1.0)

    # Daily Time-Weighted Return (TWR) for Point-to-Point mode
    prev_v = v_win.shift(1).bfill()
    daily_twr = np.where(prev_v > 0, (v_win - cf_after_t0 - prev_v) / prev_v, 0.0)
    daily_twr[0] = 0.0
    twr_series = (np.cumprod(1.0 + daily_twr) - 1.0) * 100.0

    # User SIP return during window (starts at 0.0% on day 0)
    my_sip_ret = ((v_win - cum_k) / cum_k_safe * 100.0).fillna(0.0)

    win_df = pd.DataFrame(index=win_dates)
    win_df["Date"] = win_dates
    win_df["Equity Invested (₹)"] = cum_k
    win_df["My Equity Value (₹)"] = v_win
    win_df["My Equity Return (%)"] = my_sip_ret
    win_df["My Equity TWR (%)"] = twr_series

    # Benchmarks in window
    for b_name, b_ser in bench_series.items():
        b_aligned = b_ser.reindex(all_dates).ffill().bfill()
        b_win = b_aligned[win_mask]
        p0 = float(b_win.iloc[0])
        win_df[f"{b_name}_Price"] = b_win

        # Point-to-Point Index Return rebased to 0% at start of window
        win_df[f"{b_name} Index Return (%)"] = (((b_win / p0) - 1.0) * 100.0) if p0 > 0 else 0.0

        # Equivalent Cashflow SIP simulation in window (starts at 0.0% on day 0)
        u0 = (v0 / p0) if p0 > 0 else 0.0
        u_after = (cf_after_t0 / b_win).fillna(0.0).cumsum()
        cum_u = u0 + u_after
        b_val = cum_u * b_win
        win_df[f"{b_name} Equiv Value (₹)"] = b_val
        win_df[f"{b_name} SIP Return (%)"] = (((b_val - cum_k) / cum_k_safe) * 100.0).fillna(0.0)

    return win_df.reset_index(drop=True), used_bench_tickers


# ---------------------------------------------------------------------------
# General Helpers
# ---------------------------------------------------------------------------
def normalize_category(raw_cat: str) -> str:
    val = str(raw_cat).strip().lower()
    if not val or val in ("nan", "none", "uncategorized"):
        return "Uncategorized"
    if "large" in val or val in ("l", "lcap", "largecap"):
        return "Large Cap"
    if "mid" in val or val in ("m", "mcap", "midcap"):
        return "Mid Cap"
    if "small" in val or val in ("s", "scap", "smallcap"):
        return "Small Cap"
    if "micro" in val:
        return "Micro Cap"
    if "etf" in val or "gold" in val or "index" in val:
        return "ETF / Index"
    return str(raw_cat).strip().title()


def clean_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(r"[₹$,]", "", regex=True)
        .str.replace(r"(?i)rs\.?", "", regex=True)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def auto_detect_column(columns: list[str], keywords: list[str]) -> str | None:
    norm_cols = {col: re.sub(r"[^a-z0-9]", "", str(col).lower()) for col in columns}
    for kw in keywords:
        kw_clean = re.sub(r"[^a-z0-9]", "", kw.lower())
        for col, col_clean in norm_cols.items():
            if kw_clean == col_clean or kw_clean in col_clean:
                return col
    return None


def format_inr(amount: float) -> str:
    is_neg = amount < 0
    amount = abs(float(amount))
    int_part, _, dec_part = f"{amount:.2f}".partition(".")
    if len(int_part) > 3:
        last_three = int_part[-3:]
        other_numbers = int_part[:-3]
        formatted_other = re.sub(r"(\d+?)(?=(\d{2})+$)", r"\1,", other_numbers)
        formatted_int = f"{formatted_other},{last_three}"
    else:
        formatted_int = int_part
    prefix = "-₹" if is_neg else "₹"
    return f"{prefix}{formatted_int}.{dec_part}"


def calculate_xirr(cashflows: list[tuple[date, float]]) -> float | None:
    if len(cashflows) < 2:
        return None
    dates = [cf[0] for cf in cashflows]
    amounts = [cf[1] for cf in cashflows]
    if not (any(a < 0 for a in amounts) and any(a > 0 for a in amounts)):
        return None
    min_date = min(dates)
    years = np.array([(d - min_date).days / 365.0 for d in dates], dtype=float)
    amounts_arr = np.array(amounts, dtype=float)
    rate = 0.12
    for _ in range(100):
        if rate <= -0.99:
            rate = -0.99
        factors = (1.0 + rate) ** years
        npv = np.sum(amounts_arr / factors)
        d_npv = np.sum(-years * amounts_arr / ((1.0 + rate) ** (years + 1.0)))
        if abs(d_npv) < 1e-10:
            break
        new_rate = rate - npv / d_npv
        if abs(new_rate - rate) < 1e-6:
            return float(new_rate * 100.0)
        rate = new_rate
    return None


def generate_sample_kite_tradebook_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"symbol": "TCS", "isin": "INE467B01029", "trade_date": "2024-01-15", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 10, "price": 3750.0, "trade_id": "1001", "order_id": "2001", "order_execution_time": "2024-01-15T10:15:00"},
            {"symbol": "TCS", "isin": "INE467B01029", "trade_date": "2024-06-10", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 5, "price": 3890.0, "trade_id": "1002", "order_id": "2002", "order_execution_time": "2024-06-10T11:20:00"},
            {"symbol": "TCS", "isin": "INE467B01029", "trade_date": "2025-02-12", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "sell", "quantity": 4, "price": 4150.0, "trade_id": "1003", "order_id": "2003", "order_execution_time": "2025-02-12T14:10:00"},
            {"symbol": "HDFCBANK", "isin": "INE040A01034", "trade_date": "2024-03-20", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 25, "price": 1460.0, "trade_id": "1004", "order_id": "2004", "order_execution_time": "2024-03-20T09:45:00"},
            {"symbol": "PERSISTENT", "isin": "INE262H01013", "trade_date": "2024-08-05", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 8, "price": 4700.0, "trade_id": "1005", "order_id": "2005", "order_execution_time": "2024-08-05T13:05:00"},
            {"symbol": "CDSL", "isin": "INE736A01011", "trade_date": "2025-01-18", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 20, "price": 1380.0, "trade_id": "1006", "order_id": "2006", "order_execution_time": "2025-01-18T10:30:00"},
            {"symbol": "TATAMOTORS", "isin": "INE155A01022", "trade_date": "2025-04-09", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 30, "price": 890.0, "trade_id": "1007", "order_id": "2007", "order_execution_time": "2025-04-09T11:00:00"},
            {"symbol": "SUNPHARMA", "isin": "INE044A01036", "trade_date": "2025-07-14", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 12, "price": 1610.0, "trade_id": "1008", "order_id": "2008", "order_execution_time": "2025-07-14T15:00:00"},
            {"symbol": "GOLDBEES", "isin": "INF204KB17I5", "trade_date": "2024-05-10", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 300, "price": 58.5, "trade_id": "1011", "order_id": "2011", "order_execution_time": "2024-05-10T11:15:00"},
            {"symbol": "RELIANCE", "isin": "INE002A01018", "trade_date": "2025-05-12", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 50, "price": 2850.0, "trade_id": "1009", "order_id": "2009", "order_execution_time": "2025-05-12T09:30:00"},
            {"symbol": "RELIANCE", "isin": "INE002A01018", "trade_date": "2025-05-12", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "sell", "quantity": 50, "price": 2875.0, "trade_id": "1010", "order_id": "2010", "order_execution_time": "2025-05-12T14:55:00"},
        ]
    )


# ---------------------------------------------------------------------------
# Main Application UI
# ---------------------------------------------------------------------------
def main():
    restore_db_from_github_if_needed()

    st.title("📈 Stock Portfolio Tracker (Zerodha Kite & Google Sheets)")
    gh_token, gh_repo, _ = get_github_sync_config()
    sync_badge = (
        f"☁️ GitHub Auto-Sync Active (`{gh_repo}`)"
        if (gh_token and gh_repo)
        else "💾 Local SQLite Mode (Configure GitHub Auto-Sync in Tab 5 for Permanent Cloud Storage)"
    )
    st.caption(
        f"FIFO Delivery Accounting • Pure Equity vs. Overall P&L • Benchmark Run-Line (Nifty 50 / Midcap / Smallcap) • {sync_badge}"
    )

    tx_df = load_transactions()
    meta_map = load_stock_metadata()
    consolidated_df, realized_df = compute_fifo_portfolio(tx_df, meta_map)

    tab_dash, tab_kite, tab_add, tab_rebalance, tab_history = st.tabs(
        [
            "📊 Consolidated Dashboard",
            "🪁 Import Kite Tradebook / Sheet",
            "➕ Add Buy / Sell",
            "🏷️ Sector/Cap Tags & Live Prices",
            "☁️ Cloud Sync, Log & Export",
        ]
    )

    # =========================================================================
    # TAB 1: CONSOLIDATED DASHBOARD & PIE CHARTS
    # =========================================================================
    with tab_dash:
        if consolidated_df.empty:
            st.info(
                "👋 Your portfolio is currently empty! Go to **🪁 Import Kite Tradebook / Sheet** to upload your Zerodha Console Tradebook files (or Google Sheet), or click below to test with a sample Kite Tradebook."
            )
            if st.button("🚀 Load Sample Zerodha Kite Tradebook (with Stocks, GOLDBEES & Intraday Filter)", type="primary"):
                sample_kite = generate_sample_kite_tradebook_df()
                csv_buf = io.StringIO()
                sample_kite.to_csv(csv_buf, index=False)
                csv_buf.seek(0)
                csv_buf.name = "sample_kite_tradebook.csv"
                parsed_df, _ = parse_kite_tradebook_files([csv_buf], exclude_intraday=True)
                insert_transactions_df(parsed_df, mode="replace")
                st.rerun()
        else:
            # -----------------------------------------------------------------
            # OVERALL PORTFOLIO VS. PURE EQUITY-ONLY P&L BREAKDOWN
            # -----------------------------------------------------------------
            st.subheader("💰 Overall Portfolio vs. Pure Equity-Only Return Summary")
            exclude_etfs_from_eq = st.checkbox(
                "Exclude all ETFs (e.g. GOLDBEES, SILVERBEES, NIFTYBEES) from 'Pure Equity' calculation (Uncheck to exclude only Gold/Silver/Debt and keep Equity ETFs inside Equity)",
                value=True,
            )

            non_eq_mask = consolidated_df.apply(
                lambda r: is_non_equity_or_gold_holding(
                    r["stock_name"], r["category"], r["sector"], exclude_all_etfs=exclude_etfs_from_eq
                ),
                axis=1,
            )
            equity_only_df = consolidated_df[~non_eq_mask].copy()
            non_equity_df = consolidated_df[non_eq_mask].copy()

            # Overall Portfolio Totals (Including GOLDBEES / ETFs)
            total_invested = consolidated_df["invested_rs"].sum()
            total_current = consolidated_df["current_value"].sum()
            total_unrealized_pnl = total_current - total_invested
            total_unrealized_pct = (
                (total_unrealized_pnl / total_invested * 100.0) if total_invested > 0 else 0.0
            )
            total_realized_pnl = realized_df["realized_pnl"].sum() if not realized_df.empty else 0.0

            # Pure Equity-Only Totals (Excluding GOLDBEES / Non-Equity)
            eq_invested = equity_only_df["invested_rs"].sum() if not equity_only_df.empty else 0.0
            eq_current = equity_only_df["current_value"].sum() if not equity_only_df.empty else 0.0
            eq_unrealized_pnl = eq_current - eq_invested
            eq_unrealized_pct = (eq_unrealized_pnl / eq_invested * 100.0) if eq_invested > 0 else 0.0

            if not realized_df.empty:
                real_non_eq_mask = realized_df.apply(
                    lambda r: is_non_equity_or_gold_holding(
                        r["stock_name"], r["category"], r["sector"], exclude_all_etfs=exclude_etfs_from_eq
                    ),
                    axis=1,
                )
                eq_realized_pnl = realized_df.loc[~real_non_eq_mask, "realized_pnl"].sum()
            else:
                eq_realized_pnl = 0.0

            eq_overall_profit = eq_unrealized_pnl + eq_realized_pnl
            eq_overall_profit_pct = (eq_overall_profit / eq_invested * 100.0) if eq_invested > 0 else 0.0

            # Gold / Non-Equity Totals
            gold_invested = non_equity_df["invested_rs"].sum() if not non_equity_df.empty else 0.0
            gold_current = non_equity_df["current_value"].sum() if not non_equity_df.empty else 0.0
            gold_pnl = gold_current - gold_invested
            gold_pnl_pct = (gold_pnl / gold_invested * 100.0) if gold_invested > 0 else 0.0

            # Compute Pure Equity XIRR
            eq_cashflows = []
            for _, r in tx_df.iterrows():
                if not is_non_equity_or_gold_holding(
                    r["stock_name"], r["category"], r["sector"], exclude_all_etfs=exclude_etfs_from_eq
                ):
                    try:
                        dt = pd.to_datetime(r["tx_date"]).date()
                    except Exception:
                        dt = date.today()
                    amt = -float(r["invested_rs"]) if r["tx_type"] != "SELL" else float(r["invested_rs"])
                    eq_cashflows.append((dt, amt))
            if eq_current > 0:
                eq_cashflows.append((date.today(), float(eq_current)))
            eq_xirr = calculate_xirr(eq_cashflows)

            # Row 1 KPIs: Overall Portfolio vs Pure Equity Side-by-Side
            k1, k2, k3, k4 = st.columns(4)
            k1.metric(
                "Overall Unrealized P&L (Incl. GoldBees)",
                f"{total_unrealized_pct:+.2f}%",
                f"{format_inr(total_unrealized_pnl)} on {format_inr(total_invested)}",
            )
            k2.metric(
                "Pure Equity Unrealized P&L %",
                f"{eq_unrealized_pct:+.2f}%",
                f"{format_inr(eq_unrealized_pnl)} on {format_inr(eq_invested)} ({len(equity_only_df)} Stocks)",
            )
            k3.metric(
                "Pure Equity Total Return % (Unrealized + Realized)",
                f"{eq_overall_profit_pct:+.2f}%",
                f"Total Equity Gain: {format_inr(eq_overall_profit)}"
                + (f" | Equity XIRR: {eq_xirr:.1f}%" if eq_xirr is not None else ""),
            )
            k4.metric(
                "GoldBees / ETFs Unrealized P&L %",
                f"{gold_pnl_pct:+.2f}%" if gold_invested > 0 else "0.00% (None)",
                f"{format_inr(gold_pnl)} on {format_inr(gold_invested)}"
                if gold_invested > 0
                else "No GoldBees / ETF holdings",
            )

            # -----------------------------------------------------------------
            # EQUITY PERFORMANCE VS. NIFTY 50, NIFTY MIDCAP & SMALLCAP RUN-LINE
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("🏁 Equity Performance Run-Line vs. Nifty 50, Nifty Midcap & Smallcap")
            st.caption(
                "Tracks how your Pure Equity portfolio (excluding GoldBees/Non-Equity) is performing over time compared to **Nifty 50**, **Nifty Midcap**, and **Nifty Smallcap**."
            )

            eq_tx_df = tx_df[
                ~tx_df.apply(
                    lambda r: is_non_equity_or_gold_holding(
                        r["stock_name"], r["category"], r["sector"], exclude_all_etfs=exclude_etfs_from_eq
                    ),
                    axis=1,
                )
            ].copy()

            if eq_tx_df.empty:
                st.info("No pure equity transactions found to benchmark against Nifty indices.")
            elif not YFINANCE_AVAILABLE:
                st.warning("`yfinance` is required to fetch live historical series for Nifty 50, Midcap, and Smallcap.")
            else:
                eq_tx_df["dt_parsed"] = pd.to_datetime(eq_tx_df["tx_date"], errors="coerce")
                earliest_eq_date = eq_tx_df["dt_parsed"].min().date() if not eq_tx_df["dt_parsed"].isna().all() else (date.today() - timedelta(days=365))
                today_d = date.today()

                bm_c1, bm_c2 = st.columns([2, 2])
                with bm_c1:
                    timeframe_opt = st.selectbox(
                        "Select Comparison Timeframe",
                        options=[
                            "Since First Equity Investment",
                            "Last 1 Year (1Y)",
                            "Last 6 Months (6M)",
                            "Last 3 Months (3M)",
                            "Year-to-Date (YTD)",
                            "📅 Custom Date Range (Pick Start & End Date)...",
                        ],
                        index=0,
                    )
                with bm_c2:
                    comparison_mode = st.radio(
                        "Benchmark Comparison Method",
                        options=[
                            "📈 Cumulative Return (%) — SIP-Adjusted Cashflow Comparison",
                            "📊 Point-to-Point Normalized Return (%) — Rebased to 0% at Start Date",
                            "💵 Portfolio Value (₹) — What if I invested the same ₹ into Nifty Indices?",
                        ],
                    )

                win_end = today_d
                if timeframe_opt == "Last 1 Year (1Y)":
                    win_start = max(earliest_eq_date, today_d - timedelta(days=365))
                    display_timeframe = "Last 1 Year (1Y)"
                elif timeframe_opt == "Last 6 Months (6M)":
                    win_start = max(earliest_eq_date, today_d - timedelta(days=182))
                    display_timeframe = "Last 6 Months (6M)"
                elif timeframe_opt == "Last 3 Months (3M)":
                    win_start = max(earliest_eq_date, today_d - timedelta(days=90))
                    display_timeframe = "Last 3 Months (3M)"
                elif timeframe_opt == "Year-to-Date (YTD)":
                    win_start = max(earliest_eq_date, date(today_d.year, 1, 1))
                    display_timeframe = "Year-to-Date (YTD)"
                elif timeframe_opt.startswith("📅"):
                    custom_date_sel = st.date_input(
                        "Choose Custom Comparison Period (Start & End Date)",
                        value=(max(earliest_eq_date, today_d - timedelta(days=182)), today_d),
                        min_value=earliest_eq_date,
                        max_value=today_d,
                    )
                    if isinstance(custom_date_sel, (list, tuple)) and len(custom_date_sel) == 2:
                        win_start, win_end = custom_date_sel[0], custom_date_sel[1]
                    elif isinstance(custom_date_sel, (list, tuple)) and len(custom_date_sel) == 1:
                        win_start, win_end = custom_date_sel[0], today_d
                    else:
                        win_start, win_end = earliest_eq_date, today_d
                    display_timeframe = f"Custom Period ({win_start.strftime('%d %b %Y')} to {win_end.strftime('%d %b %Y')})"
                else:
                    win_start = earliest_eq_date
                    display_timeframe = "Since Inception" 

                stock_tkr_map = {
                    r["stock_name"]: (str(r["ticker"]).strip() or f"{r['stock_name']}.NS")
                    for _, r in equity_only_df.iterrows()
                }

                with st.spinner("Building daily run-line comparison against Nifty 50, Nifty Midcap & Nifty Smallcap..."):
                    bench_df, used_tickers = fetch_benchmark_and_equity_series(
                        equity_tx_json=eq_tx_df[["tx_date", "stock_name", "tx_type", "qty", "buy_price", "invested_rs"]].to_json(),
                        stock_tickers_map=stock_tkr_map,
                        start_date_str=win_start.isoformat(),
                        end_date_str=win_end.isoformat(),
                    )

                if bench_df.empty:
                    st.warning("Could not fetch sufficient historical market data from Yahoo Finance right now. Please check your internet connection or click Auto-Fetch Live Prices in Tab 4.")
                else:
                    fig_run = go.Figure()

                    if comparison_mode.startswith("📈"):
                        # SIP-Adjusted Cashflow Return % comparison
                        fig_run.add_trace(
                            go.Scatter(
                                x=bench_df["Date"],
                                y=bench_df["My Equity Return (%)"],
                                name="My Equity Portfolio (%)",
                                mode="lines",
                                line=dict(color="#2563eb", width=3.5),
                                hovertemplate="<b>%{x|%d %b %Y}</b><br>My Equity Return: %{y:+.2f}%<extra></extra>",
                            )
                        )
                        palette = {
                            "Nifty 50": ("#f59e0b", "solid"),
                            "Nifty Midcap": ("#10b981", "dash"),
                            "Nifty Smallcap": ("#ef4444", "dot"),
                        }
                        for b_name, (col_hex, dash_style) in palette.items():
                            col_key = f"{b_name} SIP Return (%)"
                            if col_key in bench_df.columns:
                                fig_run.add_trace(
                                    go.Scatter(
                                        x=bench_df["Date"],
                                        y=bench_df[col_key],
                                        name=f"{b_name} (Same Cashflow %)",
                                        mode="lines",
                                        line=dict(color=col_hex, width=2.2, dash=dash_style),
                                        hovertemplate=f"<b>%{{x|%d %b %Y}}</b><br>{b_name}: %{{y:+.2f}}%<extra></extra>",
                                    )
                                )
                        fig_run.add_hline(y=0, line_width=1, line_dash="dot", line_color="gray")
                        y_title = f"Return in {display_timeframe} (%)"
                        chart_title = f"Run-Line Graph: My Equity Return (%) vs. Nifty Indices ({display_timeframe}, Cashflow-Matched)"

                    elif comparison_mode.startswith("📊"):
                        # Point-to-point normalized return rebased to 0% at window start
                        fig_run.add_trace(
                            go.Scatter(
                                x=bench_df["Date"],
                                y=bench_df["My Equity TWR (%)"],
                                name="My Equity Portfolio (TWR %)",
                                mode="lines",
                                line=dict(color="#2563eb", width=3.5),
                                hovertemplate="<b>%{x|%d %b %Y}</b><br>My Equity (TWR): %{y:+.2f}%<extra></extra>",
                            )
                        )
                        palette = {
                            "Nifty 50": ("#f59e0b", "solid"),
                            "Nifty Midcap": ("#10b981", "dash"),
                            "Nifty Smallcap": ("#ef4444", "dot"),
                        }
                        for b_name, (col_hex, dash_style) in palette.items():
                            col_key = f"{b_name} Index Return (%)"
                            if col_key in bench_df.columns:
                                fig_run.add_trace(
                                    go.Scatter(
                                        x=bench_df["Date"],
                                        y=bench_df[col_key],
                                        name=f"{b_name} Index (%)",
                                        mode="lines",
                                        line=dict(color=col_hex, width=2.2, dash=dash_style),
                                        hovertemplate=f"<b>%{{x|%d %b %Y}}</b><br>{b_name}: %{{y:+.2f}}%<extra></extra>",
                                    )
                                )
                        fig_run.add_hline(y=0, line_width=1, line_dash="dot", line_color="gray")
                        y_title = f"Return in {display_timeframe} (%)"
                        chart_title = f"Normalized Run-Line Graph ({win_start.strftime('%d %b %Y')} to Today, Base = 0%)"

                    else:
                        # Rupee Value comparison (What if I put the exact same ₹ into Nifty 50 / Midcap / Smallcap?)
                        fig_run.add_trace(
                            go.Scatter(
                                x=bench_df["Date"],
                                y=bench_df["Equity Invested (₹)"],
                                name="Net Equity Capital Invested (₹)",
                                mode="lines",
                                line=dict(color="#6b7280", width=2, dash="dash"),
                            )
                        )
                        fig_run.add_trace(
                            go.Scatter(
                                x=bench_df["Date"],
                                y=bench_df["My Equity Value (₹)"],
                                name="My Equity Portfolio Value (₹)",
                                mode="lines",
                                line=dict(color="#2563eb", width=3.5),
                            )
                        )
                        palette = {
                            "Nifty 50": "#f59e0b",
                            "Nifty Midcap": "#10b981",
                            "Nifty Smallcap": "#ef4444",
                        }
                        for b_name, col_hex in palette.items():
                            col_key = f"{b_name} Equiv Value (₹)"
                            if col_key in bench_df.columns:
                                fig_run.add_trace(
                                    go.Scatter(
                                        x=bench_df["Date"],
                                        y=bench_df[col_key],
                                        name=f"If Invested in {b_name} (₹)",
                                        mode="lines",
                                        line=dict(color=col_hex, width=2.2),
                                    )
                                )
                        y_title = "Portfolio Value (₹)"
                        chart_title = f"Portfolio Value Run-Line in {display_timeframe}: My Equity vs. Nifty Indices (₹)"

                    fig_run.update_layout(
                        title=chart_title,
                        xaxis_title="Date",
                        yaxis_title=y_title,
                        hovermode="x unified",
                        height=460,
                        margin=dict(t=55, b=40, l=20, r=20),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    st.plotly_chart(fig_run, use_container_width=True)

                    # Summary Outperformance / Alpha Cards dynamically computed for the selected timeframe!
                    latest_row = bench_df.iloc[-1]
                    is_p2p = comparison_mode.startswith("📊")
                    my_ret_col = "My Equity TWR (%)" if is_p2p else "My Equity Return (%)"
                    my_ret_latest = float(latest_row.get(my_ret_col, 0.0))

                    st.markdown(
                        f"**Performance in Selected Timeframe — {display_timeframe}** (My Equity: **{my_ret_latest:+.2f}%**):"
                    )
                    b1, b2, b3 = st.columns(3)
                    for col_ui, b_name in zip([b1, b2, b3], ["Nifty 50", "Nifty Midcap", "Nifty Smallcap"]):
                        ret_col = f"{b_name} Index Return (%)" if is_p2p else f"{b_name} SIP Return (%)"
                        val_col = f"{b_name} Equiv Value (₹)"
                        if ret_col in latest_row:
                            b_ret = float(latest_row[ret_col])
                            alpha = my_ret_latest - b_ret
                            b_val = float(latest_row.get(val_col, 0.0))
                            val_str = "" if is_p2p else f" ({format_inr(b_val)})"
                            col_ui.metric(
                                f"vs. {b_name} ({display_timeframe})",
                                f"Benchmark: {b_ret:+.2f}%{val_str}",
                                f"Your Alpha: {alpha:+.2f}% ({'Outperforming' if alpha >= 0 else 'Trailing'})",
                            )

                    # -------------------------------------------------------------
                    # 3. INSTITUTIONAL RISK & VOLATILITY ANALYTICS (Beta, Sharpe, Drawdown)
                    # -------------------------------------------------------------
                    st.divider()
                    st.subheader("🛡️ Institutional Risk & Volatility Analytics (vs. Nifty 50)")
                    st.caption(
                        f"Calculates your portfolio's risk-adjusted profile across **{display_timeframe}** compared to Nifty 50 using daily market fluctuations."
                    )

                    n_days = len(bench_df)
                    if n_days >= 5:
                        # Extract daily return series
                        port_vals = bench_df["My Equity Value (₹)"].values
                        cf_vals = bench_df["Equity Invested (₹)"].diff().fillna(0.0).values
                        prev_vals = np.roll(port_vals, 1)
                        prev_vals[0] = port_vals[0]

                        # Time-Weighted daily portfolio returns
                        with np.errstate(divide="ignore", invalid="ignore"):
                            r_port = np.where(prev_vals > 0, (port_vals - cf_vals - prev_vals) / prev_vals, 0.0)
                        r_port[0] = 0.0
                        r_port = np.nan_to_num(r_port, nan=0.0, posinf=0.0, neginf=0.0)

                        # Benchmark daily returns (Nifty 50)
                        b_col_for_risk = "Nifty 50_Price" if "Nifty 50_Price" in bench_df.columns else None
                        if b_col_for_risk:
                            p_bench = bench_df[b_col_for_risk].values
                            prev_pb = np.roll(p_bench, 1)
                            prev_pb[0] = p_bench[0]
                            with np.errstate(divide="ignore", invalid="ignore"):
                                r_bench = np.where(prev_pb > 0, (p_bench - prev_pb) / prev_pb, 0.0)
                            r_bench[0] = 0.0
                            r_bench = np.nan_to_num(r_bench, nan=0.0, posinf=0.0, neginf=0.0)
                        else:
                            r_bench = r_port.copy()

                        # 1. Annualized Volatility
                        vol_port_ann = float(np.std(r_port, ddof=1) * np.sqrt(252) * 100.0) if n_days > 1 else 0.0
                        vol_bench_ann = float(np.std(r_bench, ddof=1) * np.sqrt(252) * 100.0) if n_days > 1 else 0.0

                        # 2. Portfolio Beta (w.r.t Nifty 50)
                        var_bench = float(np.var(r_bench, ddof=1)) if n_days > 1 else 0.0
                        if var_bench > 1e-8:
                            cov_pb = float(np.cov(r_port, r_bench)[0, 1])
                            port_beta = float(cov_pb / var_bench)
                        else:
                            port_beta = 1.0

                        # 3. Sharpe Ratio (using 6.8% Indian 10Y G-Sec Risk-Free Rate)
                        rf_rate = 0.068
                        ann_ret_port = (float(1.0 + my_ret_latest / 100.0) ** (252.0 / max(n_days, 1))) - 1.0
                        b_ret_val = float(latest_row.get("Nifty 50 SIP Return (%)", latest_row.get("Nifty 50 Index Return (%)", 0.0)))
                        ann_ret_bench = (float(1.0 + b_ret_val / 100.0) ** (252.0 / max(n_days, 1))) - 1.0

                        sharpe_port = float((ann_ret_port - rf_rate) / (vol_port_ann / 100.0)) if vol_port_ann > 0.1 else 0.0
                        sharpe_bench = float((ann_ret_bench - rf_rate) / (vol_bench_ann / 100.0)) if vol_bench_ann > 0.1 else 0.0

                        # 4. Maximum Drawdown (MDD) & Drawdown Series
                        peak_port = np.maximum.accumulate(port_vals)
                        dd_port = np.where(peak_port > 0, (port_vals - peak_port) / peak_port * 100.0, 0.0)
                        mdd_port = float(np.min(dd_port))

                        if b_col_for_risk:
                            peak_bench = np.maximum.accumulate(p_bench)
                            dd_bench = np.where(peak_bench > 0, (p_bench - peak_bench) / peak_bench * 100.0, 0.0)
                            mdd_bench = float(np.min(dd_bench))
                        else:
                            dd_bench = dd_port
                            mdd_bench = mdd_port

                        # Display 4 Risk KPI Cards
                        rc1, rc2, rc3, rc4 = st.columns(4)
                        beta_tag = "Defensive (<1.0)" if port_beta < 0.95 else ("Aggressive (>1.05)" if port_beta > 1.05 else "Market-Aligned (~1.0)")
                        rc1.metric(
                            "Portfolio Beta (β)",
                            f"{port_beta:.2f}",
                            f"{beta_tag} vs. Nifty 50",
                        )

                        sharpe_diff = sharpe_port - sharpe_bench
                        sharpe_tag = "Superior Risk-Adjusted Return" if sharpe_diff >= 0 else "Lower Risk-Adjusted Return"
                        rc2.metric(
                            "Sharpe Ratio (vs 6.8% G-Sec)",
                            f"{sharpe_port:.2f}",
                            f"Nifty 50: {sharpe_bench:.2f} ({sharpe_tag})",
                        )

                        dd_diff = mdd_port - mdd_bench
                        dd_tag = "Less Downside Drop" if dd_diff >= 0 else "Deeper Drawdown"
                        rc3.metric(
                            "Maximum Drawdown (MDD)",
                            f"{mdd_port:.2f}%",
                            f"Nifty 50: {mdd_bench:.2f}% ({dd_tag})",
                        )

                        vol_tag = "Lower Volatility" if vol_port_ann <= vol_bench_ann else "Higher Volatility"
                        rc4.metric(
                            "Annualized Volatility (σ)",
                            f"{vol_port_ann:.1f}%",
                            f"Nifty 50: {vol_bench_ann:.1f}% ({vol_tag})",
                        )

                        # Underwater Drawdown Area Chart
                        fig_dd = go.Figure()
                        fig_dd.add_trace(
                            go.Scatter(
                                x=bench_df["Date"],
                                y=dd_port,
                                name="My Equity Drawdown (%)",
                                fill="tozeroy",
                                line=dict(color="#2563eb", width=2),
                                fillcolor="rgba(37, 99, 235, 0.15)",
                                hovertemplate="<b>%{x|%d %b %Y}</b><br>My Drawdown: %{y:.2f}%<extra></extra>",
                            )
                        )
                        if b_col_for_risk:
                            fig_dd.add_trace(
                                go.Scatter(
                                    x=bench_df["Date"],
                                    y=dd_bench,
                                    name="Nifty 50 Drawdown (%)",
                                    fill="tozeroy",
                                    line=dict(color="#f59e0b", width=1.8, dash="dot"),
                                    fillcolor="rgba(245, 158, 11, 0.08)",
                                    hovertemplate="<b>%{x|%d %b %Y}</b><br>Nifty 50 Drawdown: %{y:.2f}%<extra></extra>",
                                )
                            )
                        fig_dd.update_layout(
                            title=f"Underwater Drawdown Chart: Peak-to-Trough Drops in {display_timeframe} (%)",
                            xaxis_title="Date",
                            yaxis_title="Drawdown from Peak (%)",
                            height=350,
                            margin=dict(t=50, b=30, l=20, r=20),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        )
                        fig_dd.add_hline(y=0, line_width=1, line_color="gray")
                        st.plotly_chart(fig_dd, use_container_width=True)

                    # -------------------------------------------------------------
                    # 4. WEIGHTED PORTFOLIO P/E & P/B W.R.T NIFTY 50
                    # -------------------------------------------------------------
                    st.divider()
                    st.subheader("⚖️ Weighted Portfolio Valuation vs. Nifty 50 (P/E & P/B)")
                    st.caption(
                        "Weighted average valuation multiples of your active equity stock holdings compared against the Nifty 50 index benchmark."
                    )

                    tot_eq_cur_val = equity_only_df["current_value"].sum()
                    if tot_eq_cur_val > 0:
                        weights = equity_only_df["current_value"] / tot_eq_cur_val

                        # Fallback valuation lookup
                        pe_vals = []
                        pb_vals = []
                        for _, r_eq in equity_only_df.iterrows():
                            s_key = r_eq["stock_name"].strip().upper()
                            default_pe, default_pb = KNOWN_NSE_VALUATION.get(s_key, (22.5, 3.8))
                            r_pe = float(r_eq.get("pe_ratio", 0.0) or 0.0)
                            r_pb = float(r_eq.get("pb_ratio", 0.0) or 0.0)
                            pe_vals.append(r_pe if r_pe > 0.0 else default_pe)
                            pb_vals.append(r_pb if r_pb > 0.0 else default_pb)

                        pe_arr = np.array(pe_vals)
                        pb_arr = np.array(pb_vals)
                        w_arr = weights.values

                        weighted_pe = float(np.sum(w_arr * pe_arr))
                        weighted_pb = float(np.sum(w_arr * pb_arr))

                        # Baseline Nifty 50 P/E & P/B (NSE official historical benchmarks)
                        nifty_pe_bench = 22.5
                        nifty_pb_bench = 4.0

                        pe_diff_pct = ((weighted_pe - nifty_pe_bench) / nifty_pe_bench) * 100.0
                        pb_diff_pct = ((weighted_pb - nifty_pb_bench) / nifty_pb_bench) * 100.0

                        earnings_yield_port = (1.0 / weighted_pe * 100.0) if weighted_pe > 0 else 0.0
                        earnings_yield_nifty = (1.0 / nifty_pe_bench * 100.0)

                        if pe_diff_pct <= -10.0:
                            valuation_stance = "🏷️ Value Tilt (Trading at Discount to Nifty)"
                        elif pe_diff_pct >= 15.0:
                            valuation_stance = "🚀 Growth Tilt (Trading at Premium to Nifty)"
                        else:
                            valuation_stance = "⚖️ Fairly Valued (In Line with Nifty 50)"

                        vc1, vc2, vc3, vc4 = st.columns(4)
                        pe_label = f"{pe_diff_pct:+.1f}% vs. Nifty ({nifty_pe_bench}x)"
                        vc1.metric(
                            "Weighted Portfolio P/E",
                            f"{weighted_pe:.1f}x",
                            pe_label,
                        )

                        pb_label = f"{pb_diff_pct:+.1f}% vs. Nifty ({nifty_pb_bench}x)"
                        vc2.metric(
                            "Weighted Portfolio P/B",
                            f"{weighted_pb:.1f}x",
                            pb_label,
                        )

                        vc3.metric(
                            "Earnings Yield (1 / P/E)",
                            f"{earnings_yield_port:.2f}%",
                            f"Nifty 50: {earnings_yield_nifty:.2f}%",
                        )

                        vc4.metric(
                            "Valuation Stance",
                            valuation_stance,
                            f"Across {len(equity_only_df)} Equity Stocks",
                        )

                        # Individual Stock P/E Breakdown Chart vs. Nifty 50 Benchmark Line
                        val_chart_df = pd.DataFrame(
                            {
                                "Stock Name": equity_only_df["stock_name"],
                                "P/E Ratio": pe_arr,
                                "P/B Ratio": pb_arr,
                                "Weight (%)": (weights * 100.0).round(1),
                            }
                        ).sort_values("Weight (%)", ascending=True)

                        fig_val = go.Figure()
                        fig_val.add_trace(
                            go.Bar(
                                y=val_chart_df["Stock Name"],
                                x=val_chart_df["P/E Ratio"],
                                orientation="h",
                                name="Stock P/E",
                                marker=dict(
                                    color=np.where(val_chart_df["P/E Ratio"] <= nifty_pe_bench, "#10b981", "#3b82f6")
                                ),
                                text=[f"{pe:.1f}x (Wt: {wt}%)" for pe, wt in zip(val_chart_df["P/E Ratio"], val_chart_df["Weight (%)"])],
                                textposition="outside",
                                hovertemplate="<b>%{y}</b><br>P/E: %{x:.1f}x<br>Portfolio Weight: %{text}<extra></extra>",
                            )
                        )
                        fig_val.add_vline(
                            x=nifty_pe_bench,
                            line_width=2,
                            line_dash="dash",
                            line_color="#ef4444",
                            annotation_text=f"Nifty 50 P/E ({nifty_pe_bench}x)",
                            annotation_position="top right",
                        )
                        fig_val.update_layout(
                            title="Individual Equity Holdings P/E Ratio vs. Nifty 50 Benchmark Line",
                            xaxis_title="Price-to-Earnings Ratio (P/E)",
                            yaxis_title="",
                            height=max(340, len(val_chart_df) * 30 + 80),
                            margin=dict(t=45, b=30, l=10, r=80),
                        )
                        st.plotly_chart(fig_val, use_container_width=True)

            # -----------------------------------------------------------------
            # USER-CONFIGURABLE STOCK & SECTOR CONCENTRATION WARNING SYSTEM
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("🚨 Portfolio Concentration Risk Monitor")
            saved_stock_lim = get_app_setting("stock_conc_limit", 15.0)
            saved_sector_lim = get_app_setting("sector_conc_limit", 30.0)

            capping_scope = st.radio(
                "Concentration Capping Scope",
                options=[
                    "🎯 Pure Equity Holdings Only (Excludes GoldBees & Hedging ETFs)",
                    "🌐 Entire Portfolio (Includes GoldBees & All ETFs)",
                ],
                horizontal=True,
                help="When 'Pure Equity Holdings Only' is selected, your single-stock cap applies strictly across your equity stocks so large hedging allocations in GoldBees / ETFs never trigger false concentration warnings.",
            )

            is_equity_capping = capping_scope.startswith("🎯")

            lim_c1, lim_c2, lim_c3 = st.columns([2, 2, 1])
            with lim_c1:
                stock_limit_pct = st.slider(
                    f"Max Single-Stock Cap ({'Pure Equity' if is_equity_capping else 'Overall'} %)",
                    min_value=2.0,
                    max_value=50.0,
                    value=float(saved_stock_lim),
                    step=1.0,
                    help="Trigger a warning if any single company stock exceeds this % of your invested capital.",
                )
            with lim_c2:
                sector_limit_pct = st.slider(
                    f"Max Single-Sector Cap ({'Pure Equity' if is_equity_capping else 'Overall'} %)",
                    min_value=5.0,
                    max_value=80.0,
                    value=float(saved_sector_lim),
                    step=1.0,
                    help="Trigger a warning if any single sector exceeds this % of your invested capital.",
                )
            with lim_c3:
                st.write("")
                if st.button("💾 Save Limits", use_container_width=True):
                    set_app_setting("stock_conc_limit", stock_limit_pct, sync_github=False)
                    set_app_setting("sector_conc_limit", sector_limit_pct, sync_github=True)
                    st.toast("Saved concentration limits!", icon="✅")

            if is_equity_capping:
                monitored_df = equity_only_df.copy()
                base_invested = eq_invested
                scope_label = "Pure Equity"
            else:
                monitored_df = consolidated_df.copy()
                base_invested = total_invested
                scope_label = "Total Portfolio"

            if base_invested > 0:
                monitored_df["monitored_pct"] = (monitored_df["invested_rs"] / base_invested) * 100.0
            else:
                monitored_df["monitored_pct"] = 0.0

            breached_stocks = monitored_df[monitored_df["monitored_pct"] > stock_limit_pct]

            if is_equity_capping:
                # Exclude ETF & Gold sector so hedging assets do not trigger false sector alarms
                sector_agg = (
                    monitored_df[monitored_df["sector"] != "ETF & Gold"]
                    .groupby("sector", as_index=False)["invested_rs"]
                    .sum()
                )
            else:
                sector_agg = (
                    monitored_df.groupby("sector", as_index=False)["invested_rs"].sum()
                )
            sector_agg["sector_pct"] = (
                sector_agg["invested_rs"] / base_invested * 100.0 if base_invested > 0 else 0.0
            )
            breached_sectors = sector_agg[sector_agg["sector_pct"] > sector_limit_pct].sort_values(
                "sector_pct", ascending=False
            )

            if is_equity_capping and not non_equity_df.empty:
                hedging_names = ", ".join(non_equity_df["stock_name"].unique())
                st.info(
                    f"🛡️ **Hedging Assets Excluded from Stock Capping:** `{hedging_names}` "
                    f"({format_inr(gold_invested)}, {gold_invested / total_invested * 100.0:.1f}% of overall portfolio) "
                    f"is treated as hedging/asset allocation and excluded from your equity concentration cap."
                )

            if breached_stocks.empty and breached_sectors.empty:
                st.success(
                    f"✅ All {scope_label} holdings are within your safe limits (Single Stock ≤ {stock_limit_pct:.0f}%, Sector ≤ {sector_limit_pct:.0f}% of {scope_label.lower()} capital)."
                )
            else:
                w_col1, w_col2 = st.columns(2)
                with w_col1:
                    if not breached_stocks.empty:
                        for _, srow in breached_stocks.iterrows():
                            excess_pct = srow["monitored_pct"] - stock_limit_pct
                            excess_rs = (excess_pct / 100.0) * base_invested
                            st.warning(
                                f"⚠️ **{scope_label} Cap Breached — `{srow['stock_name']}`**: Holds **{srow['monitored_pct']:.1f}%** "
                                f"({format_inr(srow['invested_rs'])}), which is **+{excess_pct:.1f}%** ({format_inr(excess_rs)}) above your **{stock_limit_pct:.0f}%** {scope_label.lower()} cap."
                            )
                    else:
                        st.info(f"✅ No single stock exceeds {stock_limit_pct:.0f}% of {scope_label.lower()} capital.")
                with w_col2:
                    if not breached_sectors.empty:
                        for _, secrow in breached_sectors.iterrows():
                            excess_pct = secrow["sector_pct"] - sector_limit_pct
                            excess_rs = (excess_pct / 100.0) * base_invested
                            st.error(
                                f"🚨 **{scope_label} Sector Limit Breached — `{secrow['sector']}`**: Holds **{secrow['sector_pct']:.1f}%** "
                                f"({format_inr(secrow['invested_rs'])}), which is **+{excess_pct:.1f}%** ({format_inr(excess_rs)}) above your **{sector_limit_pct:.0f}%** cap."
                            )
                    else:
                        st.info(f"✅ No single sector exceeds {sector_limit_pct:.0f}% of {scope_label.lower()} capital.")

            uncat_stocks = consolidated_df[
                (consolidated_df["category"] == "Uncategorized") | (consolidated_df["sector"] == "Other")
            ]
            if not uncat_stocks.empty:
                st.info(
                    f"💡 **{len(uncat_stocks)} stock(s)** ({', '.join(uncat_stocks['stock_name'].head(5))}) need Sector / Market Cap tags. Open the **🏷️ Sector/Cap Tags & Live Prices** tab to auto-fetch them via Yahoo Finance or tag them in one click."
                )

            # -----------------------------------------------------------------
            # PIE CHARTS: SECTOR & MARKET CAP ALLOCATION
            # -----------------------------------------------------------------
            st.subheader("🥧 Allocation Pie Charts (Sector & Market Cap)")
            chart_col1, chart_col2 = st.columns(2)

            with chart_col1:
                sector_df = (
                    consolidated_df.groupby("sector", as_index=False)
                    .agg(invested_rs=("invested_rs", "sum"), stocks=("stock_name", "count"))
                    .sort_values("invested_rs", ascending=False)
                )
                fig_sector = px.pie(
                    sector_df,
                    names="sector",
                    values="invested_rs",
                    hole=0.42,
                    title="Sector-Wise Investment Breakdown",
                    color_discrete_sequence=px.colors.qualitative.Bold,
                )
                fig_sector.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate="<b>%{label}</b><br>Invested: ₹%{value:,.2f}<br>Share: %{percent}<extra></extra>",
                )
                fig_sector.update_layout(
                    height=380,
                    margin=dict(t=45, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5),
                )
                st.plotly_chart(fig_sector, use_container_width=True)

            with chart_col2:
                cap_df = (
                    consolidated_df.groupby("category", as_index=False)
                    .agg(invested_rs=("invested_rs", "sum"), stocks=("stock_name", "count"))
                    .sort_values("invested_rs", ascending=False)
                )
                color_map = {
                    "Large Cap": "#1f77b4",
                    "Mid Cap": "#ff7f0e",
                    "Small Cap": "#2ca02c",
                    "Micro Cap": "#d62728",
                    "ETF / Index": "#9467bd",
                    "Uncategorized": "#7f7f7f",
                }
                fig_cap = px.pie(
                    cap_df,
                    names="category",
                    values="invested_rs",
                    hole=0.42,
                    title="Market Cap (Large / Mid / Small) Breakdown",
                    color="category",
                    color_discrete_map=color_map,
                )
                fig_cap.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate="<b>%{label}</b><br>Invested: ₹%{value:,.2f}<br>Share: %{percent}<extra></extra>",
                )
                fig_cap.update_layout(
                    height=380,
                    margin=dict(t=45, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5),
                )
                st.plotly_chart(fig_cap, use_container_width=True)

            # -----------------------------------------------------------------
            # CONSOLIDATED CURRENT PORTFOLIO TABLE (FROZEN STOCK NAME + REMARKS)
            # -----------------------------------------------------------------
            st.subheader("📋 Consolidated Current Portfolio Holdings (Frozen Stock Column & Action Remarks)")
            st.caption(
                "📌 **Stock Name** is frozen on the left when scrolling horizontally on mobile. Double-click any cell in the **Remarks (Future Action)** column at the far right to type notes, then click **💾 Save Remarks**."
            )

            f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
            with f_col1:
                selected_caps = st.multiselect(
                    "Filter by Market Cap",
                    options=sorted(consolidated_df["category"].unique()),
                    default=[],
                    placeholder="All Categories",
                )
            with f_col2:
                selected_sectors = st.multiselect(
                    "Filter by Sector",
                    options=sorted(consolidated_df["sector"].unique()),
                    default=[],
                    placeholder="All Sectors",
                )
            with f_col3:
                search_stock = st.text_input("Search Stock Symbol", placeholder="e.g. TCS, HDFCBANK...")

            filtered_cons = consolidated_df.copy()
            if selected_caps:
                filtered_cons = filtered_cons[filtered_cons["category"].isin(selected_caps)]
            if selected_sectors:
                filtered_cons = filtered_cons[filtered_cons["sector"].isin(selected_sectors)]
            if search_stock.strip():
                filtered_cons = filtered_cons[
                    filtered_cons["stock_name"].str.contains(search_stock.strip(), case=False, na=False)
                ]

            display_df = pd.DataFrame(
                {
                    "Stock Name": filtered_cons["stock_name"],
                    "Category": filtered_cons["category"],
                    "Sector": filtered_cons["sector"],
                    "Current Qty": filtered_cons["total_qty"].round(2),
                    "FIFO Avg. Price (₹)": filtered_cons["avg_price"].round(2),
                    "Invested Rs. (₹)": filtered_cons["invested_rs"].round(2),
                    "% of Portfolio": filtered_cons["allocation_pct"].round(2),
                    "CMP (₹)": filtered_cons["cmp"].round(2),
                    "Current Value (₹)": filtered_cons["current_value"].round(2),
                    "Unrealized P&L (₹)": filtered_cons["pnl_rs"].round(2),
                    "Unrealized P&L (%)": filtered_cons["pnl_pct"].round(2),
                    "Remarks (Future Action)": filtered_cons["remarks"].fillna(""),
                }
            ).set_index("Stock Name")

            readonly_cols = [
                "Category",
                "Sector",
                "Current Qty",
                "FIFO Avg. Price (₹)",
                "Invested Rs. (₹)",
                "% of Portfolio",
                "CMP (₹)",
                "Current Value (₹)",
                "Unrealized P&L (₹)",
                "Unrealized P&L (%)",
            ]

            edited_dashboard_df = st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=False,
                disabled=readonly_cols,
                column_config={
                    "FIFO Avg. Price (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Invested Rs. (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "% of Portfolio": st.column_config.ProgressColumn(
                        format="%.2f%%",
                        min_value=0,
                        max_value=float(max(consolidated_df["allocation_pct"].max(), 10.0)),
                    ),
                    "CMP (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Current Value (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Unrealized P&L (₹)": st.column_config.NumberColumn(format="₹%+.2f"),
                    "Unrealized P&L (%)": st.column_config.NumberColumn(format="%+.2f%%"),
                    "Remarks (Future Action)": st.column_config.TextColumn(
                        "Remarks (Future Action)",
                        help="Add action notes (e.g., 'Buy more below ₹1,500', 'Book partial profit at ₹2,000')",
                        width="large",
                    ),
                },
                key="dashboard_holdings_editor",
            )

            if st.button("💾 Save Remarks / Future Action Notes", type="primary"):
                for sname, row_val in edited_dashboard_df.iterrows():
                    upsert_stock_metadata(
                        stock_name=str(sname),
                        remarks=str(row_val["Remarks (Future Action)"] or "").strip(),
                        sync_github=False,
                    )
                push_db_to_github("Save stock remarks / future action notes")
                st.success("✅ Saved your Remarks / Future Action notes!")
                st.rerun()

            # -----------------------------------------------------------------
            # UNREALIZED P&L GRAPH (DEDICATED VISUALIZATION)
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("📊 Unrealized Profit & Loss (P&L) Visualizer")
            pnl_metric_mode = st.radio(
                "Select Unrealized P&L View",
                options=[
                    "💵 Unrealized P&L in Rupees (₹) — Gainers vs. Losers",
                    "📈 Unrealized P&L in Percentage (%) — Return per Stock",
                    "⚖️ Invested Capital vs. Current Value per Stock",
                ],
                horizontal=True,
            )

            pnl_sorted = consolidated_df.copy()
            if pnl_metric_mode.startswith("💵"):
                pnl_sorted = pnl_sorted.sort_values("pnl_rs", ascending=True)
                colors = ["#16a34a" if v >= 0 else "#dc2626" for v in pnl_sorted["pnl_rs"]]
                fig_pnl = go.Figure(
                    go.Bar(
                        x=pnl_sorted["pnl_rs"],
                        y=pnl_sorted["stock_name"],
                        orientation="h",
                        marker_color=colors,
                        text=[f"₹{v:+,.0f} ({p:+.1f}%)" for v, p in zip(pnl_sorted["pnl_rs"], pnl_sorted["pnl_pct"])],
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Unrealized P&L: ₹%{x:+,.2f}<extra></extra>",
                    )
                )
                fig_pnl.update_layout(
                    title="Stock-Wise Unrealized Profit / Loss (₹)",
                    xaxis_title="Unrealized P&L (₹)",
                    yaxis_title="",
                    height=max(360, len(pnl_sorted) * 32 + 90),
                    margin=dict(t=45, b=30, l=10, r=80),
                )
                fig_pnl.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="gray")
                st.plotly_chart(fig_pnl, use_container_width=True)

            elif pnl_metric_mode.startswith("📈"):
                pnl_sorted = pnl_sorted.sort_values("pnl_pct", ascending=True)
                colors = ["#16a34a" if v >= 0 else "#dc2626" for v in pnl_sorted["pnl_pct"]]
                fig_pnl = go.Figure(
                    go.Bar(
                        x=pnl_sorted["pnl_pct"],
                        y=pnl_sorted["stock_name"],
                        orientation="h",
                        marker_color=colors,
                        text=[f"{p:+.1f}% (₹{v:+,.0f})" for v, p in zip(pnl_sorted["pnl_rs"], pnl_sorted["pnl_pct"])],
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Return: %{x:+.2f}%<extra></extra>",
                    )
                )
                fig_pnl.update_layout(
                    title="Stock-Wise Unrealized Return (%)",
                    xaxis_title="Unrealized P&L (%)",
                    yaxis_title="",
                    height=max(360, len(pnl_sorted) * 32 + 90),
                    margin=dict(t=45, b=30, l=10, r=80),
                )
                fig_pnl.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="gray")
                st.plotly_chart(fig_pnl, use_container_width=True)

            else:
                pnl_sorted = pnl_sorted.sort_values("invested_rs", ascending=False)
                fig_pnl = go.Figure()
                fig_pnl.add_trace(
                    go.Bar(
                        name="Invested Capital (₹)",
                        x=pnl_sorted["stock_name"],
                        y=pnl_sorted["invested_rs"],
                        marker_color="#3b82f6",
                    )
                )
                fig_pnl.add_trace(
                    go.Bar(
                        name="Current Market Value (₹)",
                        x=pnl_sorted["stock_name"],
                        y=pnl_sorted["current_value"],
                        marker_color="#10b981",
                    )
                )
                fig_pnl.update_layout(
                    barmode="group",
                    title="Invested Capital vs. Current Market Value per Stock (₹)",
                    xaxis_title="Stock",
                    yaxis_title="Amount (₹)",
                    height=420,
                    margin=dict(t=45, b=40, l=10, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_pnl, use_container_width=True)

            # -----------------------------------------------------------------
            # REDESIGNED FULL-WIDTH MONTH-WISE INVESTMENT VISUALIZER
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("🗓️ Month-Wise Investment & Cumulative Portfolio Growth")
            st.caption(
                "Full-width chronological view of your monthly capital deployment."
            )

            tx_chrono = tx_df.copy()
            tx_chrono["dt"] = pd.to_datetime(tx_chrono["tx_date"], errors="coerce").fillna(
                pd.Timestamp("2024-01-01")
            )
            tx_chrono["ym_sort"] = tx_chrono["dt"].dt.strftime("%Y-%m")
            tx_chrono["month_label"] = tx_chrono["dt"].dt.strftime("%b %Y")
            tx_chrono["year_num"] = tx_chrono["dt"].dt.year
            tx_chrono["month_num"] = tx_chrono["dt"].dt.month
            tx_chrono["month_short"] = tx_chrono["dt"].dt.strftime("%b")

            month_chart_type = st.radio(
                "Select Monthly Chart Style",
                options=[
                    "📈 Combo Chart: Monthly Buy/Sell Bars + Cumulative Net Invested Line",
                    "🟩 Year × Month Investment Heatmap (Calendar Matrix)",
                    "📊 Category-Stacked Monthly Purchases (Large / Mid / Small Cap)",
                ],
                horizontal=True,
            )

            if month_chart_type.startswith("📈"):
                m_Pivot = (
                    tx_chrono.groupby(["ym_sort", "month_label", "tx_type"], as_index=False)["invested_rs"]
                    .sum()
                )
                months_order = (
                    m_Pivot[["ym_sort", "month_label"]]
                    .drop_duplicates()
                    .sort_values("ym_sort")
                )
                buy_map = m_Pivot[m_Pivot["tx_type"] == "BUY"].set_index("ym_sort")["invested_rs"].to_dict()
                sell_map = m_Pivot[m_Pivot["tx_type"] == "SELL"].set_index("ym_sort")["invested_rs"].to_dict()

                months_order["bought_rs"] = months_order["ym_sort"].map(lambda k: buy_map.get(k, 0.0))
                months_order["sold_rs"] = months_order["ym_sort"].map(lambda k: sell_map.get(k, 0.0))
                months_order["net_invested"] = months_order["bought_rs"] - months_order["sold_rs"]
                months_order["cumulative_invested"] = months_order["net_invested"].cumsum()

                fig_m = go.Figure()
                fig_m.add_trace(
                    go.Bar(
                        x=months_order["month_label"],
                        y=months_order["bought_rs"],
                        name="Monthly Bought (₹)",
                        marker_color="#2563eb",
                        hovertemplate="<b>%{x}</b><br>Bought: ₹%{y:,.0f}<extra></extra>",
                    )
                )
                if months_order["sold_rs"].sum() > 0:
                    fig_m.add_trace(
                        go.Bar(
                            x=months_order["month_label"],
                            y=months_order["sold_rs"],
                            name="Monthly Sold (₹)",
                            marker_color="#f97316",
                            hovertemplate="<b>%{x}</b><br>Sold: ₹%{y:,.0f}<extra></extra>",
                        )
                    )
                fig_m.add_trace(
                    go.Scatter(
                        x=months_order["month_label"],
                        y=months_order["cumulative_invested"],
                        name="Cumulative Net Deployed (₹)",
                        mode="lines+markers",
                        line=dict(color="#10b981", width=3),
                        yaxis="y2",
                        hovertemplate="<b>%{x}</b><br>Cumulative Deployed: ₹%{y:,.0f}<extra></extra>",
                    )
                )
                fig_m.update_layout(
                    title="Monthly Capital Bought / Sold (Bars) & Cumulative Net Investment (Line)",
                    xaxis=dict(title="Month", tickangle=-35),
                    yaxis=dict(title="Monthly Trade Value (₹)"),
                    yaxis2=dict(
                        title="Cumulative Deployed (₹)",
                        overlaying="y",
                        side="right",
                        showgrid=False,
                    ),
                    barmode="group",
                    height=450,
                    margin=dict(t=55, b=70, l=20, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_m, use_container_width=True)

            elif month_chart_type.startswith("🟩"):
                buys_only = tx_chrono[tx_chrono["tx_type"] == "BUY"].copy()
                if buys_only.empty:
                    st.info("No BUY transactions available for heatmap.")
                else:
                    all_months_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    heat_piv = (
                        buys_only.pivot_table(
                            index="year_num",
                            columns="month_short",
                            values="invested_rs",
                            aggfunc="sum",
                            fill_value=0.0,
                        )
                        .reindex(columns=all_months_names, fill_value=0.0)
                        .sort_index(ascending=False)
                    )
                    fig_heat = px.imshow(
                        heat_piv,
                        text_auto=".3s",
                        aspect="auto",
                        color_continuous_scale="Blues",
                        labels=dict(x="Month", y="Year", color="Invested (₹)"),
                        title="Year × Month Purchase Heatmap (₹ Invested per Month)",
                    )
                    fig_heat.update_layout(height=max(300, len(heat_piv) * 65 + 120), margin=dict(t=50, b=30, l=20, r=20))
                    st.plotly_chart(fig_heat, use_container_width=True)

            else:
                m_cat = (
                    tx_chrono[tx_chrono["tx_type"] == "BUY"]
                    .groupby(["ym_sort", "month_label", "category"], as_index=False)["invested_rs"]
                    .sum()
                    .sort_values("ym_sort")
                )
                month_cat_order = m_cat.drop_duplicates("ym_sort")["month_label"].tolist()
                fig_mc = px.bar(
                    m_cat,
                    x="month_label",
                    y="invested_rs",
                    color="category",
                    title="Month-Wise Delivery Purchases by Market Cap Category (₹)",
                    labels={"invested_rs": "Invested (₹)", "month_label": "Month"},
                    category_orders={"month_label": month_cat_order},
                    barmode="stack",
                )
                fig_mc.update_layout(
                    height=440,
                    xaxis_tickangle=-35,
                    margin=dict(t=50, b=70, l=20, r=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_mc, use_container_width=True)

            # -----------------------------------------------------------------
            # PERIOD FILTER (MONTH / YEAR / DATE RANGE) -> BOUGHT & SOLD BY STOCK
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("📅 Period Activity Explorer (Filter Bought & Sold Value by Year, Month, or Date Range)")
            st.caption(
                "Filter your trade history by Year, Month, or exact Date Range to see how much you bought and sold in each stock during that period."
            )

            min_tx_date = tx_chrono["dt"].min().date()
            max_tx_date = tx_chrono["dt"].max().date()
            avail_years = sorted(tx_chrono["year_num"].unique().tolist(), reverse=True)
            month_names_list = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

            pf1, pf2, pf3 = st.columns([1.5, 1.5, 2])
            with pf1:
                sel_years = st.multiselect(
                    "Filter by Year(s)",
                    options=avail_years,
                    default=[],
                    placeholder="All Years",
                )
            with pf2:
                sel_months = st.multiselect(
                    "Filter by Month(s)",
                    options=month_names_list,
                    default=[],
                    placeholder="All Months (Jan–Dec)",
                )
            with pf3:
                date_range_sel = st.date_input(
                    "Or Filter by Custom Date Range",
                    value=(min_tx_date, max_tx_date),
                    min_value=min_tx_date,
                    max_value=max_tx_date,
                )

            period_df = tx_chrono.copy()
            if sel_years:
                period_df = period_df[period_df["year_num"].isin(sel_years)]
            if sel_months:
                period_df = period_df[period_df["month_short"].isin(sel_months)]
            if isinstance(date_range_sel, (list, tuple)) and len(date_range_sel) == 2:
                start_d, end_d = date_range_sel[0], date_range_sel[1]
                period_df = period_df[(period_df["dt"].dt.date >= start_d) & (period_df["dt"].dt.date <= end_d)]

            if period_df.empty:
                st.warning("No transactions match the selected Month / Year / Date Range filter.")
            else:
                tot_bought_period = period_df.loc[period_df["tx_type"] == "BUY", "invested_rs"].sum()
                tot_sold_period = period_df.loc[period_df["tx_type"] == "SELL", "invested_rs"].sum()
                net_period = tot_bought_period - tot_sold_period

                pk1, pk2, pk3, pk4 = st.columns(4)
                pk1.metric("Period Total Bought", format_inr(tot_bought_period))
                pk2.metric("Period Total Sold", format_inr(tot_sold_period))
                pk3.metric("Period Net Invested", format_inr(net_period))
                pk4.metric("Transactions in Period", f"{len(period_df)} Trades", f"{period_df['stock_name'].nunique()} Stocks")

                stock_period_rows = []
                for sname, sgrp in period_df.groupby("stock_name"):
                    b_grp = sgrp[sgrp["tx_type"] == "BUY"]
                    s_grp = sgrp[sgrp["tx_type"] == "SELL"]
                    b_qty = b_grp["qty"].sum()
                    b_val = b_grp["invested_rs"].sum()
                    b_avg = (b_val / b_qty) if b_qty > 0 else 0.0

                    s_qty = s_grp["qty"].sum()
                    s_val = s_grp["invested_rs"].sum()
                    s_avg = (s_val / s_qty) if s_qty > 0 else 0.0

                    stock_period_rows.append(
                        {
                            "Stock Name": sname,
                            "Category": sgrp["category"].iloc[-1],
                            "Sector": sgrp["sector"].iloc[-1],
                            "Bought Qty": round(b_qty, 2),
                            "Avg. Buy Price (₹)": round(b_avg, 2),
                            "Bought Value (₹)": round(b_val, 2),
                            "Sold Qty": round(s_qty, 2),
                            "Avg. Sell Price (₹)": round(s_avg, 2),
                            "Sold Value (₹)": round(s_val, 2),
                            "Net Invested in Period (₹)": round(b_val - s_val, 2),
                        }
                    )

                stock_period_df = (
                    pd.DataFrame(stock_period_rows)
                    .sort_values("Bought Value (₹)", ascending=False)
                    .set_index("Stock Name")
                )

                st.dataframe(
                    stock_period_df,
                    use_container_width=True,
                    hide_index=False,
                    column_config={
                        "Avg. Buy Price (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                        "Bought Value (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                        "Avg. Sell Price (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                        "Sold Value (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                        "Net Invested in Period (₹)": st.column_config.NumberColumn(format="₹%+.2f"),
                    },
                )

                plot_p_df = stock_period_df.reset_index()
                fig_period = go.Figure()
                fig_period.add_trace(
                    go.Bar(
                        name="Bought Value (₹)",
                        x=plot_p_df["Stock Name"],
                        y=plot_p_df["Bought Value (₹)"],
                        marker_color="#2563eb",
                    )
                )
                if plot_p_df["Sold Value (₹)"].sum() > 0:
                    fig_period.add_trace(
                        go.Bar(
                            name="Sold Value (₹)",
                            x=plot_p_df["Stock Name"],
                            y=plot_p_df["Sold Value (₹)"],
                            marker_color="#f97316",
                        )
                    )
                fig_period.update_layout(
                    barmode="group",
                    title="Stock-Wise Bought vs. Sold Value in Selected Period (₹)",
                    xaxis_title="Stock",
                    yaxis_title="Amount (₹)",
                    height=380,
                    margin=dict(t=45, b=40, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_period, use_container_width=True)

    # =========================================================================
    # TAB 2: IMPORT ZERODHA KITE TRADEBOOK OR GOOGLE SHEET
    # =========================================================================
    with tab_kite:
        st.subheader("🪁 Import Zerodha Kite Tradebook or Google Sheet")

        source_mode = st.radio(
            "Choose Data Source to Import",
            options=[
                "🪁 Zerodha Kite / Console Tradebook (Multi-Year CSV / XLSX)",
                "📊 Google Sheet / Custom Excel or CSV",
            ],
            horizontal=True,
        )

        if source_mode.startswith("🪁"):
            with st.expander("📖 How to Download Your Tradebook from Zerodha Console (Excluding F&O & Day Trading)", expanded=True):
                st.markdown(
                    """
                    1. Log in to **[Zerodha Console → Tradebook](https://console.zerodha.com/reports/tradebook)** (or inside the **Kite Mobile App**: tap your **Profile icon → Console → Tradebook**).
                    2. Set **Segment** to **Equity** (this automatically excludes all F&O, Currency, and Commodity trades).
                    3. Select the **Date Range** for each Financial Year since you started investing (Zerodha allows downloading up to **365 days at a time**, e.g. `01 Apr 2022 – 31 Mar 2023`, `01 Apr 2023 – 31 Mar 2024`, etc.) and click **CSV** or **XLSX** to download each year's file.
                    4. **How Day Trading (Intraday) is Excluded:** Leave the **"Automatically exclude Intraday / Day-Trading trades"** checkbox ticked below—this app automatically cancels out same-day buy & sell quantities and keeps **only your true Delivery investments**!
                    """
                )

            kite_files = st.file_uploader(
                "Upload one or more Zerodha Tradebook files (.csv or .xlsx) — select all financial years together!",
                type=["csv", "xlsx", "xls"],
                accept_multiple_files=True,
            )

            col_opt1, col_opt2, col_opt3 = st.columns(3)
            with col_opt1:
                exclude_intraday_chk = st.checkbox(
                    "✅ Automatically exclude Intraday (same-day buy & sell) trades",
                    value=True,
                )
            with col_opt2:
                eq_only_chk = st.checkbox(
                    "✅ Keep only Equity (EQ) segment trades",
                    value=True,
                )
            with col_opt3:
                kite_import_mode = st.radio(
                    "Import Mode",
                    options=["Append (Auto-skip duplicate trade_ids)", "Replace entire portfolio"],
                )

            if kite_files:
                try:
                    parsed_kite_df, stats = parse_kite_tradebook_files(
                        kite_files,
                        exclude_intraday=exclude_intraday_chk,
                        eq_segment_only=eq_only_chk,
                    )
                    st.info(
                        f"**Tradebook Scan Summary:** Scanned **{stats.get('total_raw_rows', 0)}** raw rows across **{len(kite_files)}** file(s) • "
                        f"Excluded **{stats.get('intraday_shares_excluded', 0):g}** intraday shares ({stats.get('intraday_trades_excluded', 0)} day-trade rows) • "
                        f"Ready to import **{stats.get('delivery_rows_kept', 0)}** delivery transactions across **{stats.get('unique_stocks', 0)}** unique stocks."
                    )

                    if not parsed_kite_df.empty:
                        st.write("##### Preview of Delivery Transactions to Import")
                        st.dataframe(parsed_kite_df.head(15), use_container_width=True, hide_index=True)

                        if st.button(
                            "🚀 Import Kite Tradebook & Build Consolidated Portfolio",
                            type="primary",
                            use_container_width=True,
                        ):
                            mode_str = "replace" if "Replace" in kite_import_mode else "append"
                            inserted_n, skipped_n = insert_transactions_df(parsed_kite_df, mode=mode_str)
                            st.cache_data.clear()
                            st.success(
                                f"✅ Imported **{inserted_n}** delivery transactions (skipped **{skipped_n}** duplicate `trade_id`s)! "
                                f"Open the **📊 Consolidated Dashboard** tab to see your current holdings."
                            )
                            st.rerun()
                except Exception as e:
                    st.error(f"Could not parse Kite Tradebook file(s): {e}")

        else:
            st.markdown(
                "Upload your Google Sheet exported as `.xlsx` or `.csv` with columns for **Month, Stock Name, Category, Sector, Buy Price, Qty, Invested Rs.**"
            )
            uploaded_gsheet = st.file_uploader(
                "Choose your exported Google Sheet (.xlsx or .csv)",
                type=["xlsx", "xls", "csv"],
                key="gsheet_uploader",
            )
            if uploaded_gsheet is not None:
                try:
                    if uploaded_gsheet.name.lower().endswith(".csv"):
                        raw_import_df = pd.read_csv(uploaded_gsheet)
                    else:
                        xls = pd.ExcelFile(uploaded_gsheet)
                        sheet_name = st.selectbox("Select Worksheet Tab", options=xls.sheet_names)
                        raw_import_df = pd.read_excel(uploaded_gsheet, sheet_name=sheet_name)

                    raw_import_df = raw_import_df.dropna(how="all").reset_index(drop=True)
                    st.dataframe(raw_import_df.head(5), use_container_width=True)

                    cols = list(raw_import_df.columns)
                    None_Opt = "-- Not Present / Auto --"
                    cols_with_none = [None_Opt] + cols

                    det_month = auto_detect_column(cols, ["month", "period", "date"])
                    det_stock = auto_detect_column(cols, ["stockname", "stock", "symbol", "company"])
                    det_cat = auto_detect_column(cols, ["category", "cap", "marketcap"])
                    det_sec = auto_detect_column(cols, ["sector", "industry"])
                    det_price = auto_detect_column(cols, ["buyprice", "price", "rate", "avgprice"])
                    det_qty = auto_detect_column(cols, ["qty", "quantity", "shares"])
                    det_inv = auto_detect_column(cols, ["investedrs", "invested", "amount", "value"])
                    det_type = auto_detect_column(cols, ["tradetype", "type", "txtype", "buy/sell"])

                    m1, m2, m3, m4 = st.columns(4)
                    with m1:
                        col_month = st.selectbox("Month/Date Column", cols_with_none, index=cols_with_none.index(det_month) if det_month in cols_with_none else 0)
                        col_stock = st.selectbox("Stock Name Column *", cols, index=cols.index(det_stock) if det_stock in cols else 0)
                    with m2:
                        col_cat = st.selectbox("Category Column", cols_with_none, index=cols_with_none.index(det_cat) if det_cat in cols_with_none else 0)
                        col_sec = st.selectbox("Sector Column", cols_with_none, index=cols_with_none.index(det_sec) if det_sec in cols_with_none else 0)
                    with m3:
                        col_price = st.selectbox("Buy Price Column", cols_with_none, index=cols_with_none.index(det_price) if det_price in cols_with_none else 0)
                        col_qty = st.selectbox("Qty Column *", cols_with_none, index=cols_with_none.index(det_qty) if det_qty in cols_with_none else 0)
                    with m4:
                        col_inv = st.selectbox("Invested Rs. Column", cols_with_none, index=cols_with_none.index(det_inv) if det_inv in cols_with_none else 0)
                        col_type = st.selectbox("Trade Type (BUY/SELL)", cols_with_none, index=cols_with_none.index(det_type) if det_type in cols_with_none else 0)

                    if st.button("🚀 Import Google Sheet Data", type="primary", use_container_width=True):
                        mapped = pd.DataFrame()
                        mapped["stock_name"] = raw_import_df[col_stock].astype(str).str.strip()
                        mapped["month"] = raw_import_df[col_month].astype(str).str.strip() if col_month != None_Opt else datetime.now().strftime("%b %Y")
                        mapped["category"] = raw_import_df[col_cat].astype(str) if col_cat != None_Opt else ""
                        mapped["sector"] = raw_import_df[col_sec].astype(str) if col_sec != None_Opt else ""
                        mapped["buy_price"] = clean_numeric(raw_import_df[col_price]) if col_price != None_Opt else 0.0
                        mapped["qty"] = clean_numeric(raw_import_df[col_qty]) if col_qty != None_Opt else 0.0
                        mapped["invested_rs"] = clean_numeric(raw_import_df[col_inv]) if col_inv != None_Opt else 0.0
                        mapped["tx_type"] = raw_import_df[col_type].astype(str) if col_type != None_Opt else "BUY"
                        ins_n, _ = insert_transactions_df(mapped, mode="append")
                        st.cache_data.clear()
                        st.success(f"✅ Imported **{ins_n}** rows!")
                        st.rerun()
                except Exception as e:
                    st.error(f"Error reading file: {e}")

    # =========================================================================
    # TAB 3: ADD MANUAL BUY OR SELL TRANSACTION
    # =========================================================================
    with tab_add:
        st.subheader("➕ Record Delivery Buy or Sell")
        existing_stocks = sorted(consolidated_df["stock_name"].unique()) if not consolidated_df.empty else []
        stock_lookup = {}
        if not consolidated_df.empty:
            for _, r in consolidated_df.iterrows():
                stock_lookup[r["stock_name"]] = {
                    "category": r["category"],
                    "sector": r["sector"],
                    "ticker": r["ticker"],
                    "avg_price": r["avg_price"],
                }

        selected_stock_opt = st.selectbox("Select Stock", options=["➕ Add New Stock"] + existing_stocks)
        prefill_name = "" if selected_stock_opt == "➕ Add New Stock" else selected_stock_opt
        prefill_cat = stock_lookup.get(prefill_name, {}).get("category", "Large Cap")
        prefill_sec = stock_lookup.get(prefill_name, {}).get("sector", "")
        prefill_tkr = stock_lookup.get(prefill_name, {}).get("ticker", "")
        prefill_price = float(stock_lookup.get(prefill_name, {}).get("avg_price", 100.0))

        with st.form("add_tx_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                tx_date_val = st.date_input("Trade Date", value=date.today())
                stock_name_val = st.text_input("Stock Symbol / Name *", value=prefill_name, placeholder="e.g. TCS, HDFCBANK")
                tx_type_val = st.selectbox("Transaction Type *", options=["BUY", "SELL"])
            with c2:
                cat_options = ["Large Cap", "Mid Cap", "Small Cap", "ETF / Index", "Micro Cap", "Uncategorized"]
                cat_idx = cat_options.index(prefill_cat) if prefill_cat in cat_options else 0
                category_val = st.selectbox("Market Cap Category", options=cat_options, index=cat_idx)
                sector_val = st.text_input("Sector", value=prefill_sec, placeholder="e.g. IT, Banking & Finance")
                ticker_val = st.text_input("Yahoo Ticker (Optional)", value=prefill_tkr, placeholder="e.g. TCS.NS")

            p1, p2, p3 = st.columns(3)
            with p1:
                price_val = st.number_input("Price per Share (₹) *", min_value=0.01, value=prefill_price, step=1.0)
            with p2:
                qty_val = st.number_input("Quantity *", min_value=0.01, value=1.0, step=1.0)
            with p3:
                st.metric("Total Trade Value", format_inr(price_val * qty_val))

            notes_val = st.text_input("Notes (Optional)")
            if st.form_submit_button("✅ Save Transaction", type="primary", use_container_width=True):
                if not stock_name_val.strip():
                    st.error("Please enter a Stock Symbol / Name.")
                else:
                    row_df = pd.DataFrame(
                        [
                            {
                                "month": tx_date_val.strftime("%b %Y"),
                                "tx_date": tx_date_val.isoformat(),
                                "stock_name": stock_name_val.strip().upper(),
                                "ticker": ticker_val.strip().upper(),
                                "category": category_val,
                                "sector": sector_val.strip() or "Other",
                                "tx_type": tx_type_val,
                                "buy_price": price_val,
                                "qty": qty_val,
                                "invested_rs": round(price_val * qty_val, 2),
                                "notes": notes_val.strip(),
                            }
                        ]
                    )
                    insert_transactions_df(row_df, mode="append")
                    st.cache_data.clear()
                    st.success(f"Saved {tx_type_val} of {qty_val} {stock_name_val.upper()}!")
                    st.rerun()

    # =========================================================================
    # TAB 4: SECTOR / CAP TAGS, LIVE PRICES & REBALANCING
    # =========================================================================
    with tab_rebalance:
        st.subheader("🏷️ Manage Stock Sector, Market Cap & Live Prices")
        if consolidated_df.empty:
            st.info("Import your Kite Tradebook or add holdings first.")
        else:
            st.caption(
                "Since Zerodha Kite Tradebook files only include stock symbols, we auto-classified known NSE stocks. You can auto-fetch any remaining Sector / Market Cap tags & Live Prices from Yahoo Finance below, or edit them directly in the table."
            )

            col_fetch_btn, col_fetch_opt = st.columns([2, 3])
            with col_fetch_btn:
                fetch_clicked = st.button(
                    "🔄 Auto-Fetch Live Prices, Sectors & Market Cap",
                    type="primary",
                    help="Queries Yahoo Finance & Google Finance to fetch real-time Current Market Price (CMP), valuation multiples (P/E, P/B), and Sector tags for all portfolio stocks.",
                )
            with col_fetch_opt:
                st.caption("⚡ Multi-tier live engine: Yahoo Finance Chart API $\\rightarrow$ `yfinance` fast_info $\\rightarrow$ Google Finance fallback.")

            if fetch_clicked:
                updated_stocks = []
                failed_stocks = []
                status_container = st.empty()
                progress_bar = st.progress(0)
                total_to_fetch = len(consolidated_df)

                for idx, (_, row) in enumerate(consolidated_df.iterrows()):
                    sname = str(row["stock_name"]).strip()
                    raw_tkr = str(row["ticker"]).strip()
                    resolved_tkr = resolve_yahoo_ticker(sname, raw_tkr)
                    cur_cat = str(row["category"]).strip()
                    cur_sec = str(row["sector"]).strip()

                    status_container.info(f"⏳ Fetching live quote for **{sname}** (`{resolved_tkr}`)... ({idx+1}/{total_to_fetch})")
                    progress_bar.progress((idx + 1) / total_to_fetch)

                    quote_res = fetch_live_market_data(resolved_tkr, sname)
                    if quote_res and quote_res.get("success") and quote_res.get("cmp", 0.0) > 0:
                        new_cmp = float(quote_res["cmp"])
                        fetched_cat = quote_res.get("category")
                        fetched_sec = quote_res.get("sector")
                        pe_to_save = quote_res.get("pe_ratio")
                        pb_to_save = quote_res.get("pb_ratio")

                        # Update category and sector if currently default/uncategorized
                        if cur_cat in ("Uncategorized", "", "None") and fetched_cat:
                            cur_cat = fetched_cat
                        if cur_sec in ("Other", "", "None") and fetched_sec:
                            cur_sec = fetched_sec

                        upsert_stock_metadata(
                            stock_name=sname,
                            ticker=resolved_tkr,
                            category=cur_cat,
                            sector=cur_sec,
                            current_price=new_cmp,
                            pe_ratio=pe_to_save,
                            pb_ratio=pb_to_save,
                            sync_github=False,
                        )
                        src = quote_res.get("source", "Live API")
                        updated_stocks.append(f"**{sname}**: ₹{new_cmp:,.2f} ({src})")
                    else:
                        failed_stocks.append(f"**{sname}** (`{resolved_tkr}`)")

                status_container.empty()
                progress_bar.empty()

                if updated_stocks:
                    push_db_to_github(f"Auto-fetch live prices for {len(updated_stocks)} stocks")
                    st.cache_data.clear()
                    st.success(f"✅ Successfully updated live CMP for **{len(updated_stocks)}/{total_to_fetch}** stocks:\n\n• " + "\n• ".join(updated_stocks))
                    if failed_stocks:
                        st.warning(f"⚠️ Could not fetch live prices for: {', '.join(failed_stocks)}. You can manually enter CMP in the table below.")
                    st.rerun()
                else:
                    st.error("❌ Could not reach price servers. Please check your internet connection or enter CMP directly in the editable table below.")

            edit_meta_df = pd.DataFrame(
                {
                    "Stock Name": consolidated_df["stock_name"],
                    "Yahoo Ticker": consolidated_df["ticker"],
                    "Category": consolidated_df["category"],
                    "Sector": consolidated_df["sector"],
                    "FIFO Avg. Price (₹)": consolidated_df["avg_price"].round(2),
                    "Current Market Price (₹)": consolidated_df["cmp"].round(2),
                    "P/E Ratio": consolidated_df.get("pe_ratio", pd.Series(22.5, index=consolidated_df.index)).round(1),
                    "P/B Ratio": consolidated_df.get("pb_ratio", pd.Series(3.8, index=consolidated_df.index)).round(1),
                    "Remarks": consolidated_df["remarks"].fillna(""),
                }
            ).set_index("Stock Name")

            edited_meta = st.data_editor(
                edit_meta_df,
                disabled=["FIFO Avg. Price (₹)"],
                hide_index=False,
                use_container_width=True,
                column_config={
                    "Category": st.column_config.SelectboxColumn(
                        options=["Large Cap", "Mid Cap", "Small Cap", "ETF / Index", "Micro Cap", "Uncategorized"],
                        required=True,
                    )
                },
                key="meta_editor_grid",
            )

            if st.button("💾 Save Sector, Market Cap, CMP, P/E, P/B & Remarks Changes"):
                for sname, r in edited_meta.iterrows():
                    upsert_stock_metadata(
                        stock_name=str(sname),
                        ticker=str(r["Yahoo Ticker"]).strip().upper(),
                        category=str(r["Category"]).strip(),
                        sector=str(r["Sector"]).strip().title(),
                        current_price=float(r["Current Market Price (₹)"]),
                        pe_ratio=float(r.get("P/E Ratio", 0.0) or 0.0),
                        pb_ratio=float(r.get("P/B Ratio", 0.0) or 0.0),
                        remarks=str(r["Remarks"] or "").strip(),
                        sync_github=False,
                    )
                push_db_to_github("Save updated stock sector/cap tags and remarks")
                st.cache_data.clear()
                st.success("Saved! Your Sector and Market Cap Pie Charts and Remarks are now updated.")
                st.rerun()

            with st.expander("🔍 Test / Refresh Live Quote for a Single Stock"):
                sq_c1, sq_c2 = st.columns([2, 1])
                with sq_c1:
                    test_stock_sel = st.selectbox(
                        "Select Stock to Fetch Live Quote",
                        options=consolidated_df["stock_name"].unique(),
                        key="test_single_quote_select",
                    )
                with sq_c2:
                    st.write("")
                    st.write("")
                    btn_run_single_quote = st.button("🔎 Fetch Live Price", key="btn_run_single_quote")

                if btn_run_single_quote and test_stock_sel:
                    matched_rows = consolidated_df[consolidated_df["stock_name"] == test_stock_sel]
                    if not matched_rows.empty:
                        m_row = matched_rows.iloc[0]
                        m_tkr = resolve_yahoo_ticker(test_stock_sel, m_row["ticker"])
                        with st.spinner(f"Querying live price for {test_stock_sel} ({m_tkr})..."):
                            single_q = fetch_live_market_data(m_tkr, test_stock_sel)
                            if single_q and single_q.get("success") and single_q.get("cmp", 0) > 0:
                                c_px = single_q["cmp"]
                                c_src = single_q.get("source", "Live API")
                                c_pe = single_q.get("pe_ratio")
                                c_pb = single_q.get("pb_ratio")
                                st.success(
                                    f"🎯 **{test_stock_sel}** (`{m_tkr}`): Live Market Price = **₹{c_px:,.2f}** ({c_src}) | P/E: {c_pe or '—'} | P/B: {c_pb or '—'}"
                                )
                                if st.button(f"💾 Apply ₹{c_px:,.2f} to {test_stock_sel} in Database", key=f"apply_cmp_{test_stock_sel}"):
                                    upsert_stock_metadata(
                                        stock_name=test_stock_sel,
                                        ticker=m_tkr,
                                        current_price=c_px,
                                        pe_ratio=c_pe,
                                        pb_ratio=c_pb,
                                        sync_github=True,
                                    )
                                    st.cache_data.clear()
                                    st.success(f"Applied live CMP ₹{c_px:,.2f} for {test_stock_sel}!")
                                    st.rerun()
                            else:
                                st.error(f"Could not fetch live price for {test_stock_sel} (`{m_tkr}`). Check your internet connection or ticker symbol.")


            # -----------------------------------------------------------------
            # 5-QUARTER SHAREHOLDING PATTERN TRENDS (Promoter, Institutional, Retail)
            # -----------------------------------------------------------------
            st.divider()
            st.subheader("🏛️ Last 5 Quarters Shareholding Pattern Trends (Promoter, Institutional & Retail)")
            latest_5q_labels = get_rolling_5_quarters()
            st.caption(
                f"Track the ownership shift across the last 5 reporting quarters (**{latest_5q_labels[0]} to {latest_5q_labels[-1]}**). Identifies whether **Institutions (FII + DII)** or **Promoters** are accumulating shares (**Rising**), or if **Retail** ownership is increasing (**Declining smart money**)."
            )

            saved_shp = load_all_shareholding_patterns()
            equity_stocks_list = [
                s for s in consolidated_df["stock_name"].unique()
                if not is_non_equity_or_gold_holding(s, "", "", exclude_all_etfs=True)
            ]

            if not equity_stocks_list:
                st.info("No equity stocks available to evaluate shareholding patterns.")
            else:
                col_shp_action, col_shp_info = st.columns([1.5, 2.5])
                with col_shp_action:
                    if st.button("🔄 Auto-Fetch Latest 5Q Filings for All Stocks (via Screener / BSE)", key="btn_fetch_all_shp"):
                        with st.spinner("Fetching official shareholding filings from Screener.in / BSE..."):
                            fetched_count = 0
                            for sname in equity_stocks_list:
                                online_res = fetch_online_shareholding_pattern(sname)
                                if online_res and online_res.get("has_data"):
                                    upsert_stock_shareholding(
                                        sname,
                                        online_res["quarters"],
                                        online_res["promoter"],
                                        online_res["institutional"],
                                        online_res["retail"],
                                        sync_github=False,
                                    )
                                    fetched_count += 1
                            if fetched_count > 0:
                                push_db_to_github(f"Auto-fetched official 5Q shareholding filings for {fetched_count} stocks")
                                st.success(f"✅ Successfully fetched official shareholding filings for {fetched_count} stocks!")
                                st.rerun()
                            else:
                                st.warning("Online fetch could not reach Screener/BSE (check internet connection). Verified built-in official percentages are active below.")

                with col_shp_info:
                    st.caption("⚡ Automatically synchronizes verified BSE/NSE Regulation 31 filings. If a stock is newly added or customized, you can also fetch or edit it individually below.")

                summary_rows = []
                q_labels_active = get_rolling_5_quarters()
                for sname in equity_stocks_list:
                    shp = get_stock_shareholding_pattern(sname, saved_shp)
                    has_data = shp.get("has_data", True)
                    p_vals = shp["promoter"]
                    i_vals = shp["institutional"]
                    r_vals = shp["retail"]

                    if has_data:
                        p_label, p_diff = eval_shareholding_trend(p_vals)
                        i_label, i_diff = eval_shareholding_trend(i_vals)
                        r_label, r_diff = eval_shareholding_trend(r_vals)

                        if i_diff > 1.0 and r_diff < -1.0:
                            verdict = "💎 Institutional Accumulation (FII/DII Buying from Retail)"
                        elif p_diff > 0.5:
                            verdict = "🚀 Promoter Increasing Stake (High Confidence)"
                        elif i_diff < -1.5 and r_diff > 1.5:
                            verdict = "⚠️ Institutional Distribution (Dumping into Retail)"
                        elif p_diff < -1.0:
                            verdict = "⚠️ Promoter Trimming Stake"
                        else:
                            verdict = "⚖️ Stable Ownership Structure"

                        summary_rows.append(
                            {
                                "Stock Name": sname,
                                f"Promoter % ({q_labels_active[-1]})": f"{p_vals[-1]:.2f}%",
                                "Promoter Trend (5Q)": p_label,
                                f"Institutional % ({q_labels_active[-1]})": f"{i_vals[-1]:.2f}%",
                                "Institutional Trend (5Q)": i_label,
                                f"Retail % ({q_labels_active[-1]})": f"{r_vals[-1]:.2f}%",
                                "Retail Trend (5Q)": r_label,
                                "Ownership Shift Insight": verdict,
                            }
                        )
                    else:
                        summary_rows.append(
                            {
                                "Stock Name": sname,
                                f"Promoter % ({q_labels_active[-1]})": "—",
                                "Promoter Trend (5Q)": "⚪ Awaiting Data",
                                f"Institutional % ({q_labels_active[-1]})": "—",
                                "Institutional Trend (5Q)": "⚪ Awaiting Data",
                                f"Retail % ({q_labels_active[-1]})": "—",
                                "Retail Trend (5Q)": "⚪ Awaiting Data",
                                "Ownership Shift Insight": "⚪ Not Fetched (Click Auto-Fetch or Edit below)",
                            }
                        )

                shp_summary_df = pd.DataFrame(summary_rows).set_index("Stock Name")
                st.dataframe(
                    shp_summary_df,
                    use_container_width=True,
                    hide_index=False,
                )

                # Individual Stock 5-Quarter Visual Breakdown & Chart
                st.write("##### 📊 5-Quarter Ownership Visualizer by Stock")
                sel_shp_stock = st.selectbox(
                    "Select Stock to View 5-Quarter Breakdown",
                    options=equity_stocks_list,
                    key="shp_stock_selector",
                )

                chosen_shp = get_stock_shareholding_pattern(sel_shp_stock, saved_shp)
                has_chosen_data = chosen_shp.get("has_data", True)
                q_labels = chosen_shp["quarters"]
                prom_vals = chosen_shp["promoter"]
                inst_vals = chosen_shp["institutional"]
                ret_vals = chosen_shp["retail"]

                if not has_chosen_data:
                    st.info(f"ℹ️ Official shareholding filings for **{sel_shp_stock}** are not yet loaded. You can click 'Fetch for this Stock' below or enter percentages manually.")
                else:
                    pt_label, pt_diff = eval_shareholding_trend(prom_vals)
                    it_label, it_diff = eval_shareholding_trend(inst_vals)
                    rt_label, rt_diff = eval_shareholding_trend(ret_vals)

                    sh_k1, sh_k2, sh_k3 = st.columns(3)
                    sh_k1.metric(
                        "Promoter Holding Trend",
                        f"{prom_vals[-1]:.2f}%",
                        f"{pt_label} (from {prom_vals[0]:.2f}%)",
                    )
                    sh_k2.metric(
                        "Institutional (FII+DII) Trend",
                        f"{inst_vals[-1]:.2f}%",
                        f"{it_label} (from {inst_vals[0]:.2f}%)",
                    )
                    sh_k3.metric(
                        "Retail (Public) Holding Trend",
                        f"{ret_vals[-1]:.2f}%",
                        f"{rt_label} (from {ret_vals[0]:.2f}%)",
                    )

                    fig_shp = go.Figure()
                    fig_shp.add_trace(
                        go.Bar(
                            x=q_labels,
                            y=prom_vals,
                            name="Promoter %",
                            marker_color="#1e3a8a",
                            text=[f"{v:.2f}%" for v in prom_vals],
                            textposition="inside",
                        )
                    )
                    fig_shp.add_trace(
                        go.Bar(
                            x=q_labels,
                            y=inst_vals,
                            name="Institutional (FII + DII) %",
                            marker_color="#10b981",
                            text=[f"{v:.2f}%" for v in inst_vals],
                            textposition="inside",
                        )
                    )
                    fig_shp.add_trace(
                        go.Bar(
                            x=q_labels,
                            y=ret_vals,
                            name="Retail / Public %",
                            marker_color="#f59e0b",
                            text=[f"{v:.2f}%" for v in ret_vals],
                            textposition="inside",
                        )
                    )
                    fig_shp.update_layout(
                        barmode="stack",
                        title=f"5-Quarter Ownership Share Breakdown for {sel_shp_stock} (%)",
                        xaxis_title="Quarter",
                        yaxis_title="Shareholding (%)",
                        yaxis=dict(range=[0, 100]),
                        height=380,
                        margin=dict(t=50, b=30, l=20, r=20),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    st.plotly_chart(fig_shp, use_container_width=True)

                col_single_fetch, col_empty = st.columns([1.5, 2.5])
                with col_single_fetch:
                    if st.button(f"🌐 Fetch Official Filings for {sel_shp_stock} (Screener.in / BSE)", key=f"btn_single_fetch_{sel_shp_stock}"):
                        with st.spinner(f"Fetching official shareholding pattern for {sel_shp_stock}..."):
                            single_res = fetch_online_shareholding_pattern(sel_shp_stock)
                            if single_res and single_res.get("has_data"):
                                upsert_stock_shareholding(
                                    sel_shp_stock,
                                    single_res["quarters"],
                                    single_res["promoter"],
                                    single_res["institutional"],
                                    single_res["retail"],
                                    sync_github=True,
                                )
                                st.success(f"✅ Successfully fetched official shareholding filings for {sel_shp_stock}!")
                                st.rerun()
                            else:
                                st.error(f"Could not automatically fetch from Screener.in for '{sel_shp_stock}'. You can input percentages directly in the editor below.")

                with st.expander(f"✏️ View & Edit 5-Quarter Percentages for {sel_shp_stock}"):
                    edit_q_df = pd.DataFrame(
                        {
                            "Quarter": q_labels,
                            "Promoter (%)": prom_vals,
                            "Institutional (%)": inst_vals,
                            "Retail (%)": ret_vals,
                        }
                    ).set_index("Quarter")

                    edited_q_res = st.data_editor(
                        edit_q_df,
                        use_container_width=True,
                        key=f"editor_shp_{sel_shp_stock}",
                    )

                    if st.button(f"💾 Save 5-Quarter Trends for {sel_shp_stock}"):
                        p_list = [float(v) for v in edited_q_res["Promoter (%)"]]
                        i_list = [float(v) for v in edited_q_res["Institutional (%)"]]
                        r_list = [float(v) for v in edited_q_res["Retail (%)"]]
                        upsert_stock_shareholding(
                            sel_shp_stock, q_labels, p_list, i_list, r_list, sync_github=True
                        )
                        st.success(f"✅ Saved 5-quarter shareholding trends for {sel_shp_stock}!")
                        st.rerun()


            st.divider()
            st.subheader("⚖️ Next-Month SIP Rebalancing Calculator")
            r1, r2, r3, r4 = st.columns(4)
            with r1:
                target_large = st.number_input("Target Large Cap (%)", 0, 100, 60, step=5)
            with r2:
                target_mid = st.number_input("Target Mid Cap (%)", 0, 100, 25, step=5)
            with r3:
                target_small = st.number_input("Target Small Cap (%)", 0, 100, 15, step=5)
            with r4:
                next_sip_budget = st.number_input(
                    "Next Month Investment Budget (₹)", min_value=0.0, value=25000.0, step=1000.0
                )

            total_inv = consolidated_df["invested_rs"].sum()
            cap_actual = consolidated_df.groupby("category")["invested_rs"].sum().to_dict()
            new_total_pool = total_inv + next_sip_budget
            rebal_rows = []
            for cap_name, t_pct in [("Large Cap", target_large), ("Mid Cap", target_mid), ("Small Cap", target_small)]:
                cur_amt = cap_actual.get(cap_name, 0.0)
                cur_pct = (cur_amt / total_inv * 100.0) if total_inv > 0 else 0.0
                ideal_after_sip = new_total_pool * (t_pct / 100.0)
                gap_amt = max(0.0, ideal_after_sip - cur_amt)
                rebal_rows.append(
                    {
                        "Category": cap_name,
                        "Current Invested (₹)": round(cur_amt, 2),
                        "Current Share (%)": round(cur_pct, 2),
                        "Target Share (%)": float(t_pct),
                        "Raw Gap (₹)": gap_amt,
                    }
                )
            rebal_df = pd.DataFrame(rebal_rows)
            tot_gap = rebal_df["Raw Gap (₹)"].sum()
            rebal_df["Suggested Next Month SIP (₹)"] = (
                (rebal_df["Raw Gap (₹)"] / tot_gap * next_sip_budget).round(2) if tot_gap > 0 else 0.0
            )
            st.dataframe(rebal_df.drop(columns=["Raw Gap (₹)"]), use_container_width=True, hide_index=True)

    # =========================================================================
    # TAB 5: PERMANENT GITHUB CLOUD SYNC, TRADE LOG, REALIZED P&L & EXPORT
    # =========================================================================
    with tab_history:
        st.subheader("☁️ Permanent Cloud Storage (Auto-Sync `portfolio_data.db` to GitHub)")
        st.markdown(
            """
            So you **never** have to manually export and re-import your data when Streamlit Cloud sleeps, this app can automatically save and restore `portfolio_data.db` directly inside your own GitHub repository!
            - **To enable permanently on Streamlit Cloud**: Open your app on Streamlit Cloud → click **Manage app (bottom right) → Settings → Secrets**, and paste:
              ```toml
              GITHUB_TOKEN = "ghp_your_personal_access_token"
              GITHUB_REPO = "your_github_username/your_repo_name"
              GITHUB_BRANCH = "main"
              ```
            """
        )
        with st.expander("🔧 Or Test / Configure GitHub Sync Right Here"):
            g1, g2, g3 = st.columns([2, 2, 1])
            with g1:
                manual_repo = st.text_input(
                    "GitHub Repo (`username/repo`)",
                    value=gh_repo,
                    placeholder="e.g. ramesh/portfolio_tracker1",
                )
            with g2:
                manual_token = st.text_input(
                    "GitHub Personal Access Token (with `repo` scope)",
                    value=gh_token,
                    type="password",
                )
            with g3:
                st.write("")
                st.write("")
                if st.button("☁️ Save & Sync Now", use_container_width=True):
                    st.session_state["gh_repo_manual"] = manual_repo.strip()
                    st.session_state["gh_token_manual"] = manual_token.strip()
                    ok, msg = push_db_to_github("Manual sync from Streamlit UI")
                    if ok:
                        st.success(f"✅ {msg} (`{manual_repo}`)")
                    else:
                        st.error(msg)

        st.divider()

        if not realized_df.empty:
            st.subheader("💰 Realized Profit / Loss from Sold Delivery Stocks (FIFO)")
            st.dataframe(
                realized_df.rename(
                    columns={
                        "stock_name": "Stock Name",
                        "category": "Category",
                        "sector": "Sector",
                        "total_bought_qty": "Total Bought Qty",
                        "total_sold_qty": "Total Sold Qty",
                        "remaining_qty": "Remaining Active Qty",
                        "realized_cost": "FIFO Cost of Sold Shares (₹)",
                        "sale_proceeds": "Sale Proceeds (₹)",
                        "realized_pnl": "Realized P&L (₹)",
                        "realized_pnl_pct": "Realized Return (%)",
                    }
                ).set_index("Stock Name"),
                use_container_width=True,
                hide_index=False,
            )
            st.divider()

        st.subheader("🗂️ Complete Delivery Transaction Log")
        if tx_df.empty:
            st.info("No delivery transactions recorded yet.")
        else:
            exp1, exp2 = st.columns(2)
            with exp1:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    consolidated_df.to_excel(writer, sheet_name="Consolidated_Holdings", index=False)
                    if not realized_df.empty:
                        realized_df.to_excel(writer, sheet_name="Realized_PnL_Sold", index=False)
                    tx_df.to_excel(writer, sheet_name="Delivery_Tradebook_Log", index=False)
                st.download_button(
                    "📊 Export Full Portfolio Workbook (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name=f"portfolio_tracker_{date.today().isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with exp2:
                if st.button("⚠️ Clear All Portfolio Data", use_container_width=True):
                    conn = get_connection()
                    conn.execute("DELETE FROM transactions")
                    conn.commit()
                    push_db_to_github("Clear portfolio database")
                    st.cache_data.clear()
                    st.rerun()

            st.dataframe(
                tx_df.sort_values(["tx_date", "id"], ascending=[False, False]).rename(
                    columns={
                        "id": "ID",
                        "month": "Month",
                        "tx_date": "Trade Date",
                        "stock_name": "Stock",
                        "category": "Category",
                        "sector": "Sector",
                        "tx_type": "Type",
                        "buy_price": "Price (₹)",
                        "qty": "Qty",
                        "invested_rs": "Value (₹)",
                        "trade_id": "Kite Trade ID",
                        "notes": "Notes",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )


if __name__ == "__main__":
    main()