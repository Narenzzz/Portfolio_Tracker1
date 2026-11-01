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
        "SELECT stock_name, ticker, category, sector, current_price, COALESCE(remarks, '') FROM stock_metadata"
    )
    return {
        row[0]: {
            "ticker": row[1],
            "category": row[2],
            "sector": row[3],
            "current_price": float(row[4] or 0.0),
            "remarks": str(row[5] or ""),
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
    sync_github: bool = True,
) -> None:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT ticker, category, sector, current_price, COALESCE(remarks, '') FROM stock_metadata WHERE stock_name = ?",
        (stock_name,),
    )
    existing = cur.fetchone()
    if existing:
        tkr = ticker if ticker is not None else existing[0]
        cat = category if category is not None else existing[1]
        sec = sector if sector is not None else existing[2]
        px = float(current_price) if current_price is not None else float(existing[3] or 0.0)
        rem = remarks if remarks is not None else str(existing[4] or "")
    else:
        default_cat, default_sec = lookup_symbol_metadata(stock_name)
        tkr = ticker if ticker is not None else f"{stock_name.upper()}.NS"
        cat = category if category is not None else default_cat
   