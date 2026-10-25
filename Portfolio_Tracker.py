"""
Mobile-Friendly Stock Portfolio Tracker (Streamlit App)
=======================================================
Supports:
1. Direct Zerodha Kite / Console Tradebook Import (Multi-year CSV/XLSX upload):
   - Automatically filters out Intraday / Day-Trading trades (same-day Buy & Sell netting)
   - Deduplicates overlapping tradebooks using Zerodha `trade_id`
   - Uses FIFO (First-In, First-Out) delivery accounting (matching Zerodha Console) to handle
     historical BUY and SELL delivery trades, computing:
       * Current Active Holdings (Total Remaining Qty, FIFO Avg. Buy Price, Current Invested Rs., % Share)
       * Realized P&L from past sold delivery stocks
   - Auto-maps NSE/BSE symbols to Sector & Market Cap Category (Large / Mid / Small Cap) using a
     built-in Indian stock directory + optional Yahoo Finance (`yfinance`) auto-enrichment.
2. Google Sheet Import (.xlsx, .csv, or public URL) & Manual Monthly Entries.
3. Permanent Cloud Persistence (No Reset on Streamlit Sleep!):
   - Automatically backs up and restores `portfolio_data.db` directly to/from your GitHub repository
     whenever you add or import trades (configured via Streamlit Secrets: GITHUB_TOKEN & GITHUB_REPO).

How to Run:
    pip install streamlit pandas plotly openpyxl yfinance numpy
    streamlit run portfolio_tracker.py
"""

import base64
import io
import json
import re
import sqlite3
import urllib.error
import urllib.request
from collections import deque
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# Optional live price & metadata fetching via yfinance
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
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

    # If local DB already has rows, no need to overwrite
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
    whenever transactions or metadata change.
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

    # Check if file already exists in GitHub to obtain its current `sha`
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
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.commit()
    return conn


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
    cur.execute("SELECT stock_name, ticker, category, sector, current_price FROM stock_metadata")
    return {
        row[0]: {
            "ticker": row[1],
            "category": row[2],
            "sector": row[3],
            "current_price": float(row[4] or 0.0),
        }
        for row in cur.fetchall()
    }


def upsert_stock_metadata(
    stock_name: str,
    ticker: str | None = None,
    category: str | None = None,
    sector: str | None = None,
    current_price: float | None = None,
    sync_github: bool = True,
) -> None:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT ticker, category, sector, current_price FROM stock_metadata WHERE stock_name = ?",
        (stock_name,),
    )
    existing = cur.fetchone()
    if existing:
        tkr = ticker if ticker is not None else existing[0]
        cat = category if category is not None else existing[1]
        sec = sector if sector is not None else existing[2]
        px = float(current_price) if current_price is not None else float(existing[3] or 0.0)
    else:
        default_cat, default_sec = lookup_symbol_metadata(stock_name)
        tkr = ticker if ticker is not None else f"{stock_name.upper()}.NS"
        cat = category if category is not None else default_cat
        sec = sector if sector is not None else default_sec
        px = float(current_price) if current_price is not None else 0.0

    conn.execute(
        """
        INSERT INTO stock_metadata (stock_name, ticker, category, sector, current_price, updated_at)
        VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(stock_name) DO UPDATE SET
            ticker = excluded.ticker,
            category = excluded.category,
            sector = excluded.sector,
            current_price = excluded.current_price,
            updated_at = CURRENT_TIMESTAMP
        """,
        (stock_name, tkr, cat, sec, px),
    )
    if category is not None and sector is not None:
        conn.execute(
            "UPDATE transactions SET category = ?, sector = ?, ticker = ? WHERE UPPER(stock_name) = UPPER(?)",
            (cat, sec, tkr, stock_name),
        )
    conn.commit()
    if sync_github:
        push_db_to_github("Update stock metadata in portfolio_data.db")


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
# Zerodha Kite Tradebook Parser & Intraday Filter
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
            df_raw = pd.read_excel(f)
            if "symbol" not in [str(c).lower().strip() for c in df_raw.columns] and "tradingsymbol" not in [
                str(c).lower().strip() for c in df_raw.columns
            ]:
                header_idx = None
                for idx, r in df_raw.head(25).iterrows():
                    row_vals = [str(v).lower().strip() for v in r.values]
                    if ("symbol" in row_vals or "tradingsymbol" in row_vals) and "trade_date" in row_vals:
                        header_idx = idx
                        break
                if header_idx is not None:
                    df_part = df_raw.iloc[header_idx + 1 :].copy()
                    df_part.columns = [str(v).strip() for v in df_raw.iloc[header_idx].values]
                else:
                    df_part = df_raw
            else:
                df_part = df_raw
        raw_frames.append(df_part)

    if not raw_frames:
        return pd.DataFrame(), {}

    combined = pd.concat(raw_frames, ignore_index=True)
    combined.columns = [str(c).strip().lower() for c in combined.columns]

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
        tkr = meta.get("ticker") or grp["ticker"].iloc[-1] or f"{stock_key}.NS"
        cached_cmp = float(meta.get("current_price") or 0.0)

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
            {"symbol": "RELIANCE", "isin": "INE002A01018", "trade_date": "2025-05-12", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "buy", "quantity": 50, "price": 2850.0, "trade_id": "1009", "order_id": "2009", "order_execution_time": "2025-05-12T09:30:00"},
            {"symbol": "RELIANCE", "isin": "INE002A01018", "trade_date": "2025-05-12", "exchange": "NSE", "segment": "EQ", "series": "EQ", "trade_type": "sell", "quantity": 50, "price": 2875.0, "trade_id": "1010", "order_id": "2010", "order_execution_time": "2025-05-12T14:55:00"},
        ]
    )


# ---------------------------------------------------------------------------
# Main Application UI
# ---------------------------------------------------------------------------
def main():
    # Automatically restore portfolio_data.db from GitHub if container just woke up
    restore_db_from_github_if_needed()

    st.title("📈 Stock Portfolio Tracker (Zerodha Kite & Google Sheets)")
    gh_token, gh_repo, _ = get_github_sync_config()
    sync_badge = (
        f"☁️ GitHub Auto-Sync Active (`{gh_repo}`)"
        if (gh_token and gh_repo)
        else "💾 Local SQLite Mode (Configure GitHub Auto-Sync in Tab 5 for Permanent Cloud Storage)"
    )
    st.caption(
        f"FIFO Delivery Accounting • Automatic Intraday Filter • {sync_badge}"
    )

    tx_df = load_transactions()
    meta_map = load_stock_metadata()
    consolidated_df, realized_df = compute_fifo_portfolio(tx_df, meta_map)

    # Top Navigation Tabs
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
            if st.button("🚀 Load Sample Zerodha Kite Tradebook (with Buys, Sells & Intraday Filter)", type="primary"):
                sample_kite = generate_sample_kite_tradebook_df()
                csv_buf = io.StringIO()
                sample_kite.to_csv(csv_buf, index=False)
                csv_buf.seek(0)
                csv_buf.name = "sample_kite_tradebook.csv"
                parsed_df, _ = parse_kite_tradebook_files([csv_buf], exclude_intraday=True)
                insert_transactions_df(parsed_df, mode="replace")
                st.rerun()
        else:
            total_invested = consolidated_df["invested_rs"].sum()
            total_current = consolidated_df["current_value"].sum()
            total_unrealized_pnl = total_current - total_invested
            total_unrealized_pct = (
                (total_unrealized_pnl / total_invested * 100.0) if total_invested > 0 else 0.0
            )
            total_realized_pnl = realized_df["realized_pnl"].sum() if not realized_df.empty else 0.0
            unique_stocks = len(consolidated_df)
            unique_sectors = consolidated_df["sector"].nunique()

            cashflows = []
            for _, r in tx_df.iterrows():
                try:
                    dt = pd.to_datetime(r["tx_date"]).date()
                except Exception:
                    dt = date.today()
                amt = -float(r["invested_rs"]) if r["tx_type"] != "SELL" else float(r["invested_rs"])
                cashflows.append((dt, amt))
            cashflows.append((date.today(), float(total_current)))
            xirr_val = calculate_xirr(cashflows)

            k1, k2, k3, k4 = st.columns(4)
            k1.metric(
                "Current Active Invested",
                format_inr(total_invested),
                f"{unique_stocks} Active Stocks ({unique_sectors} Sectors)",
            )
            k2.metric(
                "Current Portfolio Value",
                format_inr(total_current),
                f"{total_unrealized_pnl:+,.2f} ({total_unrealized_pct:+.2f}%)",
            )
            k3.metric(
                "Realized P&L (Sold Shares)",
                format_inr(total_realized_pnl),
                f"XIRR: {xirr_val:.1f}% p.a." if xirr_val is not None else "FIFO Matched",
            )
            top_stock = consolidated_df.iloc[0]
            k4.metric(
                "Largest Holding",
                f"{top_stock['stock_name']}",
                f"{top_stock['allocation_pct']:.1f}% of portfolio",
            )

            uncat_stocks = consolidated_df[
                (consolidated_df["category"] == "Uncategorized") | (consolidated_df["sector"] == "Other")
            ]
            if not uncat_stocks.empty:
                st.info(
                    f"💡 **{len(uncat_stocks)} stock(s)** ({', '.join(uncat_stocks['stock_name'].head(5))}) need Sector / Market Cap tags. Open the **🏷️ Sector/Cap Tags & Live Prices** tab to auto-fetch them via Yahoo Finance or tag them in one click."
                )

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
                    margin=dict(t=45, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
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
                    margin=dict(t=45, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                )
                st.plotly_chart(fig_cap, use_container_width=True)

            st.subheader("📋 Consolidated Current Portfolio Holdings (FIFO Adjusted)")
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
                    "FIFO Avg. Buy Price (₹)": filtered_cons["avg_price"].round(2),
                    "Invested Rs. (₹)": filtered_cons["invested_rs"].round(2),
                    "% of Total Investment": filtered_cons["allocation_pct"].round(2),
                    "CMP (₹)": filtered_cons["cmp"].round(2),
                    "Current Value (₹)": filtered_cons["current_value"].round(2),
                    "Unrealized P&L (%)": filtered_cons["pnl_pct"].round(2),
                }
            )

            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "FIFO Avg. Buy Price (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Invested Rs. (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "% of Total Investment": st.column_config.ProgressColumn(
                        format="%.2f%%",
                        min_value=0,
                        max_value=float(max(consolidated_df["allocation_pct"].max(), 10.0)),
                    ),
                    "CMP (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Current Value (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Unrealized P&L (%)": st.column_config.NumberColumn(format="%+.2f%%"),
                },
            )

            st.subheader("📈 Stock Allocation Weights & Monthly Net Deployment")
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                fig_bar = px.bar(
                    consolidated_df.sort_values("allocation_pct", ascending=True),
                    x="allocation_pct",
                    y="stock_name",
                    color="category",
                    orientation="h",
                    title="Current Stock Weight (% of Active Portfolio)",
                    labels={"allocation_pct": "Portfolio Share (%)", "stock_name": "Stock"},
                    text=consolidated_df.sort_values("allocation_pct", ascending=True)[
                        "allocation_pct"
                    ].apply(lambda x: f"{x:.1f}%"),
                )
                fig_bar.update_layout(margin=dict(t=45, b=20, l=10, r=20), yaxis_title="")
                st.plotly_chart(fig_bar, use_container_width=True)

            with b_col2:
                monthly_buys = (
                    tx_df[tx_df["tx_type"] == "BUY"]
                    .groupby(["month", "category"], as_index=False)["invested_rs"]
                    .sum()
                )
                fig_month = px.bar(
                    monthly_buys,
                    x="month",
                    y="invested_rs",
                    color="category",
                    title="Month-Wise Delivery Purchases (₹)",
                    labels={"invested_rs": "Bought (₹)", "month": "Month"},
                    barmode="stack",
                )
                fig_month.update_layout(margin=dict(t=45, b=20, l=10, r=10))
                st.plotly_chart(fig_month, use_container_width=True)

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

            if YFINANCE_AVAILABLE and st.button(
                "🔄 Auto-Fetch Live Prices, Sectors & Market Cap from Yahoo Finance",
                type="primary",
            ):
                updated_count = 0
                with st.spinner("Querying Yahoo Finance for NSE/BSE stocks..."):
                    for _, row in consolidated_df.iterrows():
                        sname = row["stock_name"]
                        tkr = str(row["ticker"]).strip() or f"{sname}.NS"
                        cur_cat = row["category"]
                        cur_sec = row["sector"]
                        try:
                            yt = yf.Ticker(tkr)
                            hist = yt.history(period="5d")
                            latest_px = float(hist["Close"].iloc[-1]) if not hist.empty else float(row["cmp"])
                            if cur_cat == "Uncategorized" or cur_sec == "Other":
                                info = yt.info or {}
                                yf_sec = info.get("sector") or info.get("industry")
                                if yf_sec and cur_sec == "Other":
                                    cur_sec = str(yf_sec).strip()
                                mcap = info.get("marketCap")
                                if mcap and cur_cat == "Uncategorized":
                                    if mcap >= 200_000_000_000:
                                        cur_cat = "Large Cap"
                                    elif mcap >= 50_000_000_000:
                                        cur_cat = "Mid Cap"
                                    else:
                                        cur_cat = "Small Cap"
                            upsert_stock_metadata(sname, tkr, cur_cat, cur_sec, latest_px, sync_github=False)
                            updated_count += 1
                        except Exception:
                            pass
                push_db_to_github("Auto-fetch Yahoo Finance live prices & sectors")
                st.success(f"Updated live prices & metadata for {updated_count} stocks!")
                st.rerun()

            edit_meta_df = pd.DataFrame(
                {
                    "Stock Name": consolidated_df["stock_name"],
                    "Yahoo Ticker": consolidated_df["ticker"],
                    "Category": consolidated_df["category"],
                    "Sector": consolidated_df["sector"],
                    "FIFO Avg. Price (₹)": consolidated_df["avg_price"].round(2),
                    "Current Market Price (₹)": consolidated_df["cmp"].round(2),
                }
            )

            edited_meta = st.data_editor(
                edit_meta_df,
                disabled=["Stock Name", "FIFO Avg. Price (₹)"],
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Category": st.column_config.SelectboxColumn(
                        options=["Large Cap", "Mid Cap", "Small Cap", "ETF / Index", "Micro Cap", "Uncategorized"],
                        required=True,
                    )
                },
                key="meta_editor_grid",
            )

            if st.button("💾 Save Sector, Market Cap & CMP Changes"):
                for _, r in edited_meta.iterrows():
                    upsert_stock_metadata(
                        stock_name=r["Stock Name"],
                        ticker=str(r["Yahoo Ticker"]).strip().upper(),
                        category=str(r["Category"]).strip(),
                        sector=str(r["Sector"]).strip().title(),
                        current_price=float(r["Current Market Price (₹)"]),
                        sync_github=False,
                    )
                push_db_to_github("Save updated stock sector/cap tags")
                st.success("Saved! Your Sector and Market Cap Pie Charts are now updated.")
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
                    placeholder="e.g. ramesh/my-stock-tracker",
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
                ),
                use_container_width=True,
                hide_index=True,
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