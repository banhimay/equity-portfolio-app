import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, date, timedelta, timezone
import os
import html as html_lib
import re

def detect_mobile():
    """Phone detection from the browser's User-Agent. Add ?view=mobile or ?view=desktop
    to the web address to force a layout."""
    try:
        forced = st.query_params.get("view")
        if forced == "mobile":
            return True
        if forced == "desktop":
            return False
        ua = st.context.headers.get("User-Agent", "") or ""
        return bool(re.search(r"Mobile|Android|iPhone|iPod", ua, re.I))
    except Exception:
        return False

IS_MOBILE = detect_mobile()

st.set_page_config(page_title="Equity Portfolio Dashboard", layout="wide",
                   initial_sidebar_state="collapsed" if IS_MOBILE else "expanded")

EXCEL_FILE = "Equity_Portfolio_Dashboard.xlsx"
SL_COLS = ['TYPE', 'Stock', 'Buy_Date', 'Buy_Price', 'SL_Price']
DATABASE_URL = os.environ.get("DATABASE_URL", "")   # set on Render -> trailing stop-losses live in Postgres
PASSWORD = os.environ.get("PORTFOLIO_PASSWORD", "admin")  # override via env var in production

# --- AUTHENTICATION ---
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        st.title("🔒 Portfolio Access")
        pwd = st.text_input("Enter Password", type="password")
        if st.button("Login"):
            if pwd == PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        return False
    return True

if not check_password():
    st.stop()

st.markdown(
    "<style>"
    ".kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin:4px 0 8px 0;}"
    ".kpi-grid.sm{grid-template-columns:repeat(auto-fit,minmax(105px,1fr));}"
    ".kpi-grid.narrow{grid-template-columns:repeat(auto-fit,minmax(125px,175px));justify-content:center;}"
    ".kpi-grid.narrow .kpi{text-align:center;}"
    "[class*='st-key-hp_']{display:none !important;}"
    "div[data-testid='stElementContainer']:has(iframe[height='0']){display:none;}"
    "a.stk-link{color:#2f80ed !important;text-decoration:underline;cursor:pointer;font-weight:600;}"
    "@media(max-width:640px){.kpi-grid.narrow{grid-template-columns:repeat(3,1fr);gap:5px;}.kpi-grid.narrow .kpi{padding:6px 4px;}.kpi-grid.narrow .kpi-l{font-size:10px;}.kpi-grid.narrow .kpi-v{font-size:.9rem;}}"
    ".kpi{border:1px solid rgba(128,128,128,.30);border-radius:10px;padding:8px 10px;}"
    ".kpi-l{font-size:12px;opacity:.7;}"
    ".kpi-v{font-size:1.2rem;font-weight:700;line-height:1.3;}"
    ".kpi-s{font-size:12px;font-weight:600;}"
    ".mc{border:1px solid rgba(128,128,128,.35);border-radius:10px;padding:8px 10px;margin:6px 0;}"
    ".mc.tot{border-width:2px;}"
    ".mc-h{display:flex;justify-content:space-between;align-items:baseline;font-weight:700;font-size:15px;}"
    ".mc-h small{font-weight:600;opacity:.75;font-size:11px;margin-left:6px;}"
    ".mc-g{display:grid;grid-template-columns:1fr 1fr;gap:2px 12px;font-size:12.5px;margin-top:4px;}"
    ".mc-g span{opacity:.65;}"
    ".mc-g b{font-weight:600;}"
    ".cmp-wrap{-webkit-overflow-scrolling:touch;}"
    "@media (max-width:640px){"
    ".block-container,[data-testid='stMainBlockContainer']{padding:4.5rem 0.6rem 3rem 0.6rem !important;}"
    "h2{font-size:1.35rem !important;}h3{font-size:1.2rem !important;}h4{font-size:1.05rem !important;}"
    "[data-testid='stVerticalBlock']{gap:0.5rem !important;}"
    "}"
    "</style>",
    unsafe_allow_html=True
)

# --- DATA LOADING AND SAVING ---
def _load_data_xlsx():
    if os.path.exists(EXCEL_FILE):
        tb = pd.read_excel(EXCEL_FILE, sheet_name="Trade_Book")
        tb = tb.loc[:, ~tb.columns.str.contains('^Unnamed')].dropna(subset=['Stock'])

        try:
            co = pd.read_excel(EXCEL_FILE, sheet_name="CMP_Override")
            co = co.dropna(subset=['Stock'])
        except Exception:
            co = pd.DataFrame(columns=['Stock', 'Manual CMP', 'Note'])
    else:
        tb = pd.DataFrame(columns=['TYPE', 'Stock', 'Date of Trxn', 'Price', 'Unit', 'Amount'])
        co = pd.DataFrame(columns=['Stock', 'Manual CMP', 'Note'])
    return tb, co

def _db_write_trades(cur, tb, co):
    """Replace-all write of the Trade Book and CMP overrides (single user, inside one transaction)."""
    cur.execute("DELETE FROM trades")
    rows = []
    for _, r in tb.iterrows():
        d = pd.to_datetime(r['Date of Trxn']).strftime('%Y-%m-%d')
        rows.append((str(r['TYPE']).strip().upper(), str(r['Stock']).strip(), d, float(r['Price']), float(r['Unit'])))
    if rows:
        cur.executemany("INSERT INTO trades (type, stock, trade_date, price, unit) VALUES (%s,%s,%s,%s,%s)", rows)
    cur.execute("DELETE FROM cmp_override")
    crow = []
    for _, r in co.iterrows():
        if pd.isna(r.get('Stock')) or pd.isna(r.get('Manual CMP')):
            continue
        note = r.get('Note')
        crow.append((str(r['Stock']).strip(), float(r['Manual CMP']), None if pd.isna(note) else str(note)))
    if crow:
        cur.executemany("INSERT INTO cmp_override (stock, manual_cmp, note) VALUES (%s,%s,%s)", crow)

@st.cache_data(ttl=600)
def load_data():
    if DATABASE_URL:
        try:
            db_init()
            with db_connect() as conn, conn.cursor() as cur:
                cur.execute("SELECT type, stock, trade_date, price::float, unit::float FROM trades ORDER BY id")
                tb = pd.DataFrame(cur.fetchall(), columns=['TYPE', 'Stock', 'Date of Trxn', 'Price', 'Unit'])
                cur.execute("SELECT stock, manual_cmp::float, note FROM cmp_override ORDER BY id")
                co = pd.DataFrame(cur.fetchall(), columns=['Stock', 'Manual CMP', 'Note'])
        except Exception as exc:
            st.error(f"Could not read the trade database: {exc}")
            st.stop()
        tb['Date of Trxn'] = pd.to_datetime(tb['Date of Trxn'])
        if len(tb) and (tb['Unit'] % 1 == 0).all():
            tb['Unit'] = tb['Unit'].astype('int64')
        tb['Amount'] = tb['Price'] * tb['Unit']
        return tb, co
    return _load_data_xlsx()

# ---- trailing stop-loss ("Shaktiman", tranche 4+): Postgres when DATABASE_URL is set, else the xlsx ----
def db_connect():
    import psycopg2
    return psycopg2.connect(DATABASE_URL, connect_timeout=10)

@st.cache_resource
def db_init():
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("""CREATE TABLE IF NOT EXISTS manual_sl (
            type text NOT NULL, stock text NOT NULL, buy_date text NOT NULL, buy_price numeric(14,2) NOT NULL,
            sl_price numeric(14,2) NOT NULL, updated_at timestamptz NOT NULL DEFAULT now(),
            PRIMARY KEY (type, stock, buy_date, buy_price))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS manual_sl_history (
            id serial PRIMARY KEY, type text NOT NULL, stock text NOT NULL, buy_date text NOT NULL,
            buy_price numeric(14,2) NOT NULL, sl_price numeric(14,2) NOT NULL,
            changed_at timestamptz NOT NULL DEFAULT now())""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trades (
            id serial PRIMARY KEY, type text NOT NULL, stock text NOT NULL, trade_date text NOT NULL,
            price numeric(14,4) NOT NULL, unit numeric(14,4) NOT NULL)""")
        cur.execute("""CREATE TABLE IF NOT EXISTS cmp_override (
            id serial PRIMARY KEY, stock text NOT NULL, manual_cmp numeric(14,4) NOT NULL, note text)""")
        cur.execute("CREATE TABLE IF NOT EXISTS app_meta (key text PRIMARY KEY, value text)")
        cur.execute("""CREATE TABLE IF NOT EXISTS weekly_unrealized (
            friday date NOT NULL, type text NOT NULL, stock text NOT NULL, units numeric(14,4) NOT NULL,
            cost numeric(14,2) NOT NULL, close_price numeric(14,4) NOT NULL, unrealized numeric(14,2) NOT NULL,
            saved_at timestamptz NOT NULL DEFAULT now(), PRIMARY KEY (friday, type, stock))""")
        cur.execute("SELECT 1 FROM app_meta WHERE key='trades_seeded'")
        if cur.fetchone() is None:   # first run: copy the Trade Book from the repo's xlsx once
            tb, co = _load_data_xlsx()
            _db_write_trades(cur, tb, co)
            cur.execute("INSERT INTO app_meta (key, value) VALUES ('trades_seeded', now()::text)")
        cur.execute("SELECT 1 FROM app_meta WHERE key='kabra_sale_added'")
        if cur.fetchone() is None:   # one-time: KABRAEXTRU sale of 23-Oct-2026 that was missing from the app
            cur.execute("SELECT 1 FROM trades WHERE stock='KABRAEXTRU' AND trade_date='2026-10-23' AND unit<0")
            if cur.fetchone() is None:
                cur.execute("INSERT INTO trades (type, stock, trade_date, price, unit) VALUES ('SATELLITE','KABRAEXTRU','2026-10-23',678.10,-37)")
            cur.execute("INSERT INTO app_meta (key, value) VALUES ('kabra_sale_added', now()::text)")
        cur.execute("""CREATE TABLE IF NOT EXISTS cash_entries (
            id serial PRIMARY KEY, account text NOT NULL, kind text NOT NULL, entry_date date, label text,
            amount numeric(14,2) NOT NULL)""")
        cur.execute("""CREATE TABLE IF NOT EXISTS account_cash (
            account text PRIMARY KEY, cash_balance numeric(14,2) NOT NULL, updated_at timestamptz NOT NULL DEFAULT now())""")
        cur.execute("SELECT 1 FROM app_meta WHERE key='cashflow_seeded'")
        if cur.fetchone() is None:   # first run: load the two accounts' history supplied by the user, once
            for acct, kind, d, amt in CASH_SEED_DATED:
                cur.execute("INSERT INTO cash_entries (account, kind, entry_date, amount) VALUES (%s,%s,%s,%s)",
                            (acct, kind, d, amt))
            for acct, label, amt in CASH_SEED_CHARGES:
                cur.execute("INSERT INTO cash_entries (account, kind, label, amount) VALUES (%s,'CHARGE',%s,%s)",
                            (acct, label, amt))
            for acct, amt in CASH_SEED_BALANCE.items():
                cur.execute("INSERT INTO account_cash (account, cash_balance) VALUES (%s,%s)", (acct, amt))
            cur.execute("INSERT INTO app_meta (key, value) VALUES ('cashflow_seeded', now()::text)")
    return True

ACCOUNTS = ["YWA415", "JAE735"]
# one-time starting data (from the user's sheet): (account, kind, date, amount)
_YWA_DEP = [("2025-06-24", 70000), ("2025-07-28", 10000), ("2025-09-03", 20000), ("2025-09-22", 10000),
            ("2025-10-06", 10000), ("2025-10-21", 20000), ("2025-11-03", 5000), ("2025-12-15", 10000),
            ("2025-12-26", 15000), ("2026-01-20", 50000), ("2026-04-20", 45000), ("2026-05-08", 40000),
            ("2026-05-22", 30000), ("2026-06-05", 20000), ("2026-06-29", 20000), ("2026-08-21", 50000)]
_YWA_BACK = [("2025-10-03", 85), ("2026-01-02", 4371), ("2026-04-18", 12627)]
_JAE_DEP = [("2025-09-26", 1000), ("2025-10-10", 25000), ("2025-10-21", 35000), ("2025-11-21", 50000),
            ("2025-11-28", 2000), ("2025-12-18", 7000), ("2026-02-09", 30000), ("2026-04-17", 70000),
            ("2026-05-15", 30000), ("2026-08-17", 50000)]
_JAE_BACK = [("2025-10-03", 911), ("2026-01-02", 6979)]
CASH_SEED_DATED = ([("YWA415", "DEPOSIT", d, v) for d, v in _YWA_DEP] + [("YWA415", "TRANSFER_BACK", d, v) for d, v in _YWA_BACK] +
                   [("JAE735", "DEPOSIT", d, v) for d, v in _JAE_DEP] + [("JAE735", "TRANSFER_BACK", d, v) for d, v in _JAE_BACK])
CASH_SEED_CHARGES = [("YWA415", "2025-26", 1490), ("YWA415", "2026-27", 1112),
                     ("JAE735", "2025-26", 719), ("JAE735", "2026-27", 943)]
CASH_SEED_BALANCE = {"YWA415": 18651, "JAE735": 10684}

@st.cache_data(ttl=600)
def load_cash():
    """(entries, balances) for the Cash Flow page. Needs the database."""
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("SELECT id, account, kind, entry_date, label, amount::float FROM cash_entries ORDER BY entry_date NULLS LAST, id")
        ent = pd.DataFrame(cur.fetchall(), columns=['id', 'account', 'kind', 'date', 'label', 'amount'])
        cur.execute("SELECT account, cash_balance::float FROM account_cash")
        bal = dict(cur.fetchall())
    ent['date'] = pd.to_datetime(ent['date'])
    return ent, bal

def cash_add_entry(account, kind, d, amount):
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("INSERT INTO cash_entries (account, kind, entry_date, amount) VALUES (%s,%s,%s,%s)", (account, kind, d, amount))
    load_cash.clear()

def cash_delete_entry(entry_id):
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM cash_entries WHERE id=%s", (int(entry_id),))
    load_cash.clear()

def cash_set_charge(account, label, amount, kind='CHARGE'):
    """One figure per account, charge head and financial year: update it if it exists, else add it.
    kind 'CHARGE' = Statutory & Broker Charges, 'DP_CHARGE' = DP Charges."""
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("UPDATE cash_entries SET amount=%s WHERE account=%s AND kind=%s AND label=%s", (amount, account, kind, label))
        if cur.rowcount == 0:
            cur.execute("INSERT INTO cash_entries (account, kind, label, amount) VALUES (%s,%s,%s,%s)", (account, kind, label, amount))
    load_cash.clear()

def cash_get_override():
    """(exposed, equity_value) typed in by the user, or None to use the app's own holdings."""
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("SELECT key, value FROM app_meta WHERE key IN ('cf_exposed', 'cf_equity')")
        d = dict(cur.fetchall())
    return (float(d['cf_exposed']), float(d['cf_equity'])) if 'cf_exposed' in d and 'cf_equity' in d else None

def cash_set_override(exposed, equity):
    """Pass None to go back to the app's holdings."""
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM app_meta WHERE key IN ('cf_exposed', 'cf_equity')")
        if exposed is not None:
            cur.execute("INSERT INTO app_meta (key, value) VALUES ('cf_exposed', %s), ('cf_equity', %s)", (str(exposed), str(equity)))

def cash_set_balance(account, amount):
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("""INSERT INTO account_cash (account, cash_balance) VALUES (%s,%s)
                       ON CONFLICT (account) DO UPDATE SET cash_balance=EXCLUDED.cash_balance, updated_at=now()""", (account, amount))
    load_cash.clear()

@st.cache_data(ttl=600)
def load_manual_sl():
    if DATABASE_URL:
        try:
            db_init()
            with db_connect() as conn, conn.cursor() as cur:
                cur.execute("SELECT type, stock, buy_date, buy_price::float, sl_price::float FROM manual_sl")
                return pd.DataFrame(cur.fetchall(), columns=SL_COLS)
        except Exception as exc:
            st.error(f"Could not read the stop-loss database: {exc}")
            return pd.DataFrame(columns=SL_COLS)
    try:
        sl = pd.read_excel(EXCEL_FILE, sheet_name="Manual_SL")
        return sl.dropna(subset=['Stock'])
    except Exception:
        return pd.DataFrame(columns=SL_COLS)

def save_manual_sl(sel_type, sel_stock, entries):
    """entries = {(buy_date, buy_price): sl_price}; a price of 0 clears that tranche's stop-loss.
    Postgres: upsert + append a history row whenever the value actually changes."""
    if DATABASE_URL:
        db_init()
        with db_connect() as conn, conn.cursor() as cur:
            for (d, p), v in entries.items():
                cur.execute("SELECT sl_price::float FROM manual_sl WHERE type=%s AND stock=%s AND buy_date=%s AND buy_price=%s",
                            (sel_type, sel_stock, d, p))
                row = cur.fetchone()
                old = row[0] if row else 0.0
                if abs(old - v) < 0.005:
                    continue
                if v > 0:
                    cur.execute("""INSERT INTO manual_sl (type, stock, buy_date, buy_price, sl_price) VALUES (%s,%s,%s,%s,%s)
                                   ON CONFLICT (type, stock, buy_date, buy_price)
                                   DO UPDATE SET sl_price = EXCLUDED.sl_price, updated_at = now()""",
                                (sel_type, sel_stock, d, p, v))
                else:
                    cur.execute("DELETE FROM manual_sl WHERE type=%s AND stock=%s AND buy_date=%s AND buy_price=%s",
                                (sel_type, sel_stock, d, p))
                cur.execute("INSERT INTO manual_sl_history (type, stock, buy_date, buy_price, sl_price) VALUES (%s,%s,%s,%s,%s)",
                            (sel_type, sel_stock, d, p, v))
        load_manual_sl.clear()
        return
    # fallback: keep it in the xlsx (resets whenever the app is redeployed)
    cur_sl = load_manual_sl()
    if cur_sl.empty:
        keep = cur_sl
    else:
        hit = cur_sl.apply(lambda x: x['TYPE'] == sel_type and x['Stock'] == sel_stock and
                           (str(x['Buy_Date'])[:10], round(float(x['Buy_Price']), 2)) in entries, axis=1)
        keep = cur_sl[~hit]
    add = pd.DataFrame([{'TYPE': sel_type, 'Stock': sel_stock, 'Buy_Date': d, 'Buy_Price': p, 'SL_Price': v}
                        for (d, p), v in entries.items() if v > 0], columns=SL_COLS)
    save_data(trades_df, overrides_df, pd.concat([keep, add], ignore_index=True))

@st.cache_data(ttl=600)
def load_sl_history(sel_type, sel_stock):
    if not DATABASE_URL:
        return pd.DataFrame()
    try:
        db_init()
        with db_connect() as conn, conn.cursor() as cur:
            cur.execute("""SELECT changed_at, buy_date, buy_price::float, sl_price::float FROM manual_sl_history
                           WHERE type=%s AND stock=%s ORDER BY changed_at DESC, id DESC LIMIT 100""", (sel_type, sel_stock))
            return pd.DataFrame(cur.fetchall(), columns=['changed_at', 'buy_date', 'buy_price', 'sl_price'])
    except Exception:
        return pd.DataFrame()

def save_data(tb_df, co_df, sl_df=None):
    if DATABASE_URL:   # trades + CMP overrides live in Postgres (survive redeploys)
        db_init()
        with db_connect() as conn, conn.cursor() as cur:
            _db_write_trades(cur, tb_df, co_df)
        st.cache_data.clear()
        return
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        tb_df.to_excel(writer, sheet_name="Trade_Book", index=False)
        co_df.to_excel(writer, sheet_name="CMP_Override", index=False)
        (sl_df if sl_df is not None else load_manual_sl()).to_excel(writer, sheet_name="Manual_SL", index=False)
    st.cache_data.clear()

def rekey_manual_sl(old, new):
    """A buy trade was edited: carry its manual stop-loss to the new (type, stock, date, price) key.
    old/new = (type, stock, 'YYYY-MM-DD', price)."""
    if old == new:
        return
    if DATABASE_URL:
        db_init()
        with db_connect() as conn, conn.cursor() as cur:
            cur.execute("""UPDATE manual_sl SET type=%s, stock=%s, buy_date=%s, buy_price=%s, updated_at=now()
                           WHERE type=%s AND stock=%s AND buy_date=%s AND buy_price=%s""", (*new, *old))
        load_manual_sl.clear()
        return
    cur_sl = load_manual_sl()
    if cur_sl.empty:
        return
    hit = cur_sl.apply(lambda x: (x['TYPE'], x['Stock'], str(x['Buy_Date'])[:10], round(float(x['Buy_Price']), 2)) == old, axis=1)
    if hit.any():
        cur_sl = cur_sl.copy()
        cur_sl.loc[hit, ['TYPE', 'Stock', 'Buy_Date', 'Buy_Price']] = [new[0], new[1], new[2], new[3]]
        save_data(trades_df, overrides_df, cur_sl)

trades_df, overrides_df = load_data()
manual_sl_df = load_manual_sl()

# --- LIVE PRICE FETCHING ENGINE ---
def to_yf_symbol(ticker: str) -> str:
    """Normalize a stock name to a yfinance-compatible NSE symbol.
    Strips an 'NSE:' prefix if present (used elsewhere for GOOGLEFINANCE-style
    references) and appends '.NS' if not already present."""
    t = ticker.strip().upper()
    if t.startswith("NSE:"):
        t = t[len("NSE:"):]
    if not t.endswith(".NS"):
        t = f"{t}.NS"
    return t

def _valid(x):
    return x is not None and not pd.isna(x)

@st.cache_data(ttl=300)
def fetch_live_prices(tickers, overrides_df):
    """Returns (price_map, prev_close_map). prev_close is yesterday's closing price
    (same idea as GOOGLEFINANCE 'closeyest'); None when it cannot be fetched."""
    price_map = {}
    prev_map = {}
    manual_map = {}
    if not overrides_df.empty:
        manual_map = dict(zip(overrides_df['Stock'], overrides_df['Manual CMP']))

    for ticker in tickers:
        price, prev = None, None
        try:
            stock_data = yf.Ticker(to_yf_symbol(ticker))
            try:
                fast_info = stock_data.fast_info
                price = fast_info.last_price
                prev = getattr(fast_info, 'previous_close', None)
            except Exception:
                pass

            if not _valid(price) or not _valid(prev):
                try:
                    closes = stock_data.history(period="5d")['Close'].dropna()
                    if not closes.empty:
                        if not _valid(price):
                            price = closes.iloc[-1]
                        if not _valid(prev) and len(closes) >= 2:
                            prev = closes.iloc[-2]
                except Exception:
                    pass
        except Exception:
            pass

        if ticker in manual_map and pd.notna(manual_map[ticker]):
            price = manual_map[ticker]

        price_map[ticker] = float(price) if _valid(price) else 0.0
        prev_map[ticker] = float(prev) if _valid(prev) else None

    return price_map, prev_map

# --- FIFO CALCULATOR ENGINE ---
def compute_portfolio(df, overrides_df, net_deployed=0.0):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df['Date of Trxn'] = pd.to_datetime(df['Date of Trxn'])
    df = df.sort_values(by='Date of Trxn', kind='stable').copy()

    unique_tickers = df['Stock'].unique().tolist()
    live_prices, prev_closes = fetch_live_prices(unique_tickers, overrides_df)

    inventory = {}
    realized_pl = {}

    for _, row in df.iterrows():
        stk = str(row['Stock']).strip().upper()
        ptype = str(row['TYPE']).strip().upper()
        key = (ptype, stk)
        qty = float(row['Unit'])
        price = float(row['Price'])

        if key not in realized_pl:
            realized_pl[key] = 0.0
        if key not in inventory:
            inventory[key] = []

        if qty > 0:
            inventory[key].append({'qty': qty, 'price': price})
        elif qty < 0:
            sell_qty = abs(qty)
            while sell_qty > 0 and inventory[key]:
                oldest = inventory[key][0]
                if oldest['qty'] <= sell_qty:
                    realized_pl[key] += oldest['qty'] * (price - oldest['price'])
                    sell_qty -= oldest['qty']
                    inventory[key].pop(0)
                else:
                    realized_pl[key] += sell_qty * (price - oldest['price'])
                    oldest['qty'] -= sell_qty
                    sell_qty = 0

    rows = []
    all_keys = set(inventory.keys()).union(set(realized_pl.keys()))

    for key in all_keys:
        ptype, stk = key
        lots = inventory.get(key, [])
        units = sum(l['qty'] for l in lots)
        total_cost = sum(l['qty'] * l['price'] for l in lots)
        avg_cost = (total_cost / units) if units > 0 else 0.0
        cmp = live_prices.get(stk, 0.0)
        prev_close = prev_closes.get(stk)
        has_day = units > 0 and prev_close is not None and cmp > 0
        yest_close = prev_close if units > 0 and prev_close is not None else float('nan')
        day_change = (cmp - prev_close) * units if has_day else float('nan')
        total_val = units * cmp
        unrealized_pl = total_val - total_cost if units > 0 else 0.0
        r_pl = realized_pl.get(key, 0.0)
        net_gain = unrealized_pl + r_pl

        rows.append({
            'Type': ptype,
            'Stock': stk,
            'Units': int(units),
            'Avg_Cost': avg_cost,
            'Total_Cost': total_cost,
            'CMP': cmp,
            'Yest_Close': yest_close,
            'Day_Change': day_change,
            'Total_Value': total_val,
            'Unrealized_PL': unrealized_pl,
            'Realized_PL': r_pl,
            'Total_Net_Gain': net_gain
        })

    portfolio = pd.DataFrame(rows)
    if not portfolio.empty:
        # Allocation_% = what is tied up in this stock (Total_Cost) as a share of the money put in (Net Deployed)
        if net_deployed > 0:
            portfolio['Allocation_%'] = portfolio['Total_Cost'] / net_deployed
        else:   # no cash-flow data available: fall back to share of current portfolio value
            total_portfolio_val = portfolio['Total_Value'].sum()
            portfolio['Allocation_%'] = portfolio['Total_Value'] / total_portfolio_val if total_portfolio_val > 0 else 0.0

    return portfolio, df

def get_net_deployed():
    """Deposits - Transferred Back, both accounts (Cash Flow tab). 0 if the database is not available."""
    if not DATABASE_URL:
        return 0.0
    try:
        ent, _ = load_cash()
        return float(ent.loc[ent['kind'] == 'DEPOSIT', 'amount'].sum() - ent.loc[ent['kind'] == 'TRANSFER_BACK', 'amount'].sum())
    except Exception:
        return 0.0

portfolio_df, processed_trades = compute_portfolio(trades_df, overrides_df, get_net_deployed())

# --- INDIVIDUAL STOCK TRANCHE + STOP-LOSS ENGINE ---
def get_remaining_lots(df, ptype, stock):
    """Re-runs FIFO for one (Type, Stock) and returns only the lots still held,
    in chronological (buy) order, each as {date, price, qty}."""
    sub = df[(df['TYPE'].str.strip().str.upper() == ptype) &
             (df['Stock'].str.strip().str.upper() == stock)].copy()
    sub['Date of Trxn'] = pd.to_datetime(sub['Date of Trxn'])
    sub = sub.sort_values(by='Date of Trxn')

    lots = []  # {date, price, remaining}
    for _, row in sub.iterrows():
        qty = float(row['Unit'])
        price = float(row['Price'])
        if qty >= 0:
            lots.append({'date': row['Date of Trxn'], 'price': price, 'remaining': qty})
        else:
            to_sell = abs(qty)
            for lot in lots:
                if to_sell <= 0:
                    break
                if lot['remaining'] <= 0:
                    continue
                consume = min(lot['remaining'], to_sell)
                lot['remaining'] -= consume
                to_sell -= consume

    return [l for l in lots if l['remaining'] > 0]

def build_tranche_table(lots, cmp_price, manual_sl=None):
    """Builds the tranche-by-tranche breakdown with weighted avg, growth %,
    P/L, and the stop-loss rule:
      Tranche 1        -> 10% below its own buy price
      Tranche 2        -> price that keeps TOTAL risk in rupees the same as Tranche 1's risk
      Tranche 3        -> breakeven (stop loss = cumulative weighted avg cost)
      Tranche 4 onward -> "Shaktiman": stop-loss price is entered manually
                          (manual_sl = {(buy_date_str, buy_price_rounded): price})
    """
    rows = []
    cum_qty = 0.0
    cum_investment = 0.0
    prev_buy_price = None
    prev_risk_amount = None

    for i, lot in enumerate(lots):
        tranche_no = i + 1
        qty = lot['remaining']
        price = lot['price']
        investment = qty * price

        cum_qty += qty
        cum_investment += investment
        weighted_avg = cum_investment / cum_qty if cum_qty > 0 else 0.0

        is_manual = False
        date_str = lot['date'].strftime('%Y-%m-%d')
        if tranche_no == 1:
            sl_rule = "10% below buy"
            sl_price = price * 0.90
        elif tranche_no == 2:
            sl_rule = "Residual risk"
            sl_price = weighted_avg - (prev_risk_amount / cum_qty)
        elif tranche_no == 3:
            sl_rule = "Breakeven"
            sl_price = weighted_avg
        else:
            sl_rule = "Shaktiman"
            is_manual = True
            sl_price = (manual_sl or {}).get((date_str, round(price, 2)), float('nan'))

        risk_amount = (weighted_avg - sl_price) * cum_qty

        growth_from_prev = ((price - prev_buy_price) / prev_buy_price) if prev_buy_price else None
        growth_from_cmp = ((cmp_price - price) / price) if price else None
        tranche_pl = qty * (cmp_price - price)

        rows.append({
            'Tranche': tranche_no,
            'Buy Date': date_str,
            'Buy Price': price,
            'Qty': qty,
            'Investment': investment,
            'Cum. Qty': cum_qty,
            'Weighted Avg Cost': weighted_avg,
            'Growth % vs Prev Tranche': growth_from_prev,
            'Growth % vs CMP': growth_from_cmp,
            'Stop-Loss Rule': sl_rule,
            'Stop-Loss Price': sl_price,
            'Risk Amount (₹)': risk_amount,
            'Tranche P/L (₹)': tranche_pl,
            'Manual SL': is_manual,
        })

        prev_buy_price = price
        prev_risk_amount = risk_amount if pd.notna(risk_amount) else prev_risk_amount

    return pd.DataFrame(rows)

def build_fifo_calc_table(df, ptype, stock):
    """FIFO_Calc-style listing for one (Type, Stock): every transaction, newest first.
    Cum_Buy runs from the top row down; Cum_Sold is total ever sold; Depleted is how much
    of a buy lot has been consumed by FIFO-matched sales."""
    sub = df[(df['TYPE'].str.strip().str.upper() == ptype) &
             (df['Stock'].str.strip().str.upper() == stock)].copy()
    sub['Date of Trxn'] = pd.to_datetime(sub['Date of Trxn'])
    sub = sub.sort_values(by='Date of Trxn', kind='stable').reset_index(drop=True)

    depleted = [0.0] * len(sub)
    lots = []  # {i, remaining}
    for i, row in sub.iterrows():
        qty = float(row['Unit'])
        if qty >= 0:
            lots.append({'i': i, 'remaining': qty})
        else:
            to_sell = abs(qty)
            for lot in lots:
                if to_sell <= 0:
                    break
                if lot['remaining'] <= 0:
                    continue
                consume = min(lot['remaining'], to_sell)
                lot['remaining'] -= consume
                depleted[lot['i']] += consume
                to_sell -= consume

    total_sold = abs(sub.loc[sub['Unit'].astype(float) < 0, 'Unit'].astype(float).sum())
    rows = []
    for i, row in sub.iterrows():
        qty = float(row['Unit'])
        price = float(row['Price'])
        dep = 0.0 if qty < 0 else depleted[i]
        rem = 0.0 if qty < 0 else qty - dep
        rows.append({
            'Type': ptype, 'Stock': stock,
            'Buy_Date': row['Date of Trxn'],
            'Buy_Price': price, 'Buy_Qty': qty,
            'Cum_Sold': total_sold, 'Depleted': dep,
            'Rem_Qty': rem, 'Rem_Value': rem * price,
        })
    out = pd.DataFrame(rows).iloc[::-1].reset_index(drop=True)  # newest first
    out.insert(5, 'Cum_Buy', out['Buy_Qty'].cumsum())
    return out

def fmt_num(v, decimals=2):
    s = f"{v:,.{decimals}f}"
    return s.rstrip('0').rstrip('.') if decimals > 0 and '.' in s else s

def render_compact_table(headers, rows, left_cols=(), auto_width=False, total_last=False, font_px=12, fit=False, shrink=False):
    """Small-font HTML table that fits in one view. rows = list of lists of
    (text, css_color_or_None). total_last=True styles the last row as a total row.
    Works in light and dark themes."""
    css = ("<style>"
           ".cmp-wrap{overflow-x:auto;}"
           "table.cmp{border-collapse:collapse;line-height:1.25;}"
           "table.cmp th,table.cmp td{padding:3px 7px;white-space:nowrap;text-align:right;"
           "border-bottom:1px solid rgba(128,128,128,0.30);}"
           "table.cmp th{background:rgba(128,128,128,0.18);font-weight:700;text-align:center;}"
           "table.cmp td.l{text-align:left;}"
           "table.cmp tr.tot td{font-weight:800;border-top:2px solid rgba(128,128,128,0.65);"
           "background:rgba(128,128,128,0.10);}"
           "</style>")
    width = "width:auto;" if (auto_width or shrink) else "width:100%;"
    if fit:   # squeeze into the page width: tighter cells, headers may wrap (breaking after "_")
        css = css.replace("</style>",
            "table.cmp.fit th,table.cmp.fit td{padding:3px 4px;}"
            "table.cmp.fit th{white-space:normal;line-height:1.15;}"
            "</style>")
        if shrink:   # columns only as wide as their content, table does not stretch to the page
            css = css.replace("</style>", "table.cmp.fit th,table.cmp.fit td{padding:2px 3px;}</style>")
    head = "".join(f"<th>{html_lib.escape(h).replace('_', '_<wbr>') if fit else html_lib.escape(h)}</th>" for h in headers)
    body = ""
    for ri, r in enumerate(rows):
        tds = ""
        for ci, cell in enumerate(r):
            text, color, *rest = cell
            bg = rest[0] if rest else None
            cls = " class='l'" if ci in left_cols else ""
            if bg:
                style = f" style='background:{bg}; color:{color or '#000'}; font-weight:700;'"
            else:
                style = f" style='color:{color}; font-weight:600;'" if color else ""
            if isinstance(text, tuple) and text and text[0] == "LINK":   # ("LINK", pill_key, stock)
                inner = (f"<a href='#' class='stk-link' data-pill='{html_lib.escape(str(text[1]), quote=True)}' "
                         f"data-stock='{html_lib.escape(str(text[2]), quote=True)}'>{html_lib.escape(str(text[2]))}</a>")
            else:
                inner = html_lib.escape(str(text))
            tds += f"<td{cls}{style}>{inner}</td>"
        tr_cls = " class='tot'" if (total_last and ri == len(rows) - 1) else ""
        body += f"<tr{tr_cls}>{tds}</tr>"
    st.markdown(
        f"{css}<div class='cmp-wrap'><table class='cmp{" fit" if fit else ""}' style='{width}font-size:{font_px}px;'>"
        f"<thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>",
        unsafe_allow_html=True
    )

def fmt_date(d):
    return f"{d.day}-{d.strftime('%b-%Y')}"

def build_closed_lots_table(df, ptype, stock):
    """Every sale matched FIFO against the buy lots it consumed, one row per matched piece.
    'Cycle' counts rounds of buying: it goes up when the stock is bought again after the
    position had been fully squared off."""
    sub = df[(df['TYPE'].str.strip().str.upper() == ptype) &
             (df['Stock'].str.strip().str.upper() == stock)].copy()
    sub['Date of Trxn'] = pd.to_datetime(sub['Date of Trxn'])
    sub = sub.sort_values(by='Date of Trxn', kind='stable')

    lots = []   # {date, price, remaining}
    rows = []
    cycle = 1
    net = 0.0
    flat_after_sell = False
    for _, row in sub.iterrows():
        qty = float(row['Unit'])
        price = float(row['Price'])
        if qty >= 0:
            if flat_after_sell:
                cycle += 1
                flat_after_sell = False
            lots.append({'date': row['Date of Trxn'], 'price': price, 'remaining': qty})
            net += qty
        else:
            to_sell = abs(qty)
            for lot in lots:
                if to_sell <= 0:
                    break
                if lot['remaining'] <= 0:
                    continue
                consume = min(lot['remaining'], to_sell)
                invested = consume * lot['price']
                proceeds = consume * price
                rows.append({
                    'Cycle': cycle,
                    'Buy Date': lot['date'], 'Buy Price': lot['price'],
                    'Sell Date': row['Date of Trxn'], 'Sell Price': price,
                    'Qty': consume, 'Invested': invested, 'Proceeds': proceeds,
                    'Realized P/L': proceeds - invested,
                    'Return %': (proceeds - invested) / invested if invested else None,
                    'Days Held': (row['Date of Trxn'] - lot['date']).days,
                })
                lot['remaining'] -= consume
                to_sell -= consume
            net += qty
            if net <= 1e-9:
                flat_after_sell = True
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(['Sell Date', 'Buy Date'], ascending=False, kind='stable').reset_index(drop=True)
    return out

UP, DOWN = "#1B8A3A", "#C62828"

def money(v, dec=2):
    if v is None or pd.isna(v):
        return "-"
    return f"{'-' if v < 0 else ''}₹{abs(v):,.{dec}f}"

def pl_color(v):
    if v is None or pd.isna(v) or v == 0:
        return None
    return UP if v > 0 else DOWN

def esc(x):
    return html_lib.escape(str(x))

def kpi_grid(items, small=False, narrow=False):
    """items: (label, value, sub_text_or_None, sub_color_or_None). Wraps to 2 per row on phones."""
    cells = ""
    for label, value, sub, sub_color, *rest in items:
        value_color = rest[0] if rest else None
        bg = rest[1] if len(rest) > 1 else None
        kpi_style = f" style='background:{bg};border-color:{bg.replace('.20', '.75')};'" if bg else ""
        sub_html = f"<div class='kpi-s' style='color:{sub_color or 'inherit'};'>{esc(sub)}</div>" if sub else ""
        cells += (f"<div class='kpi'{kpi_style}><div class='kpi-l'>{esc(label)}</div>"
                  f"<div class='kpi-v' style='color:{value_color or 'inherit'};'>{esc(value)}</div>{sub_html}</div>")
    st.markdown(f"<div class='kpi-grid{' sm' if small else ''}{' narrow' if narrow else ''}'>{cells}</div>", unsafe_allow_html=True)

def card_html(title=None, right=None, right_color=None, items=(), badge=None, total=False, bg=None):
    """One mobile card: optional header (title left, headline value right) + a 2-column grid of
    label/value pairs. items: (label, value, color_or_None)."""
    head = ""
    if title is not None:
        badge_html = f"<small>{esc(badge)}</small>" if badge else ""
        head = (f"<div class='mc-h'><div>{esc(title)}{badge_html}</div>"
                f"<div style='color:{right_color or 'inherit'};'>{esc(right if right is not None else '')}</div></div>")
    grid = ""
    for label, value, color in items:
        st_attr = f" style='color:{color};'" if color else ""
        grid += f"<div><span>{esc(label)}</span> <b{st_attr}>{esc(value)}</b></div>"
    grid_html = f"<div class='mc-g'>{grid}</div>" if grid else ""
    bg_attr = f" style='background:{bg}; color:#000;'" if bg else ""
    return f"<div class='mc{' tot' if total else ''}'{bg_attr}>{head}{grid_html}</div>"

def render_cards(cards):
    st.markdown("".join(cards), unsafe_allow_html=True)

# --- NAVIGATION ---
# ---- weekly Unrealized_PL snapshots: one per Friday close, held stocks only (Postgres) ----
WEEKLY_AFTER_CLOSE = (15, 45)   # IST: a Friday counts as "closed" after 3:45 pm (NSE closes 3:30 pm)

def ist_now():
    return datetime.now(timezone(timedelta(hours=5, minutes=30)))

def completed_fridays(n):
    """The last n Fridays whose market close is over, oldest first."""
    now = ist_now()
    d = now.date() - timedelta(days=(now.date().weekday() - 4) % 7)
    if d == now.date() and (now.hour, now.minute) < WEEKLY_AFTER_CLOSE:
        d -= timedelta(days=7)
    return [d - timedelta(days=7 * i) for i in range(n)][::-1]

def weekly_hist_closes(stocks, start, end):
    """{stock: Series of daily closes indexed by date} between start and end (inclusive)."""
    out = {}
    for stk in stocks:
        try:
            h = yf.Ticker(to_yf_symbol(stk)).history(start=start.isoformat(), end=(end + timedelta(days=1)).isoformat())['Close'].dropna()
            idx = pd.to_datetime(h.index)
            if getattr(idx, 'tz', None) is not None:
                idx = idx.tz_localize(None)
            h.index = idx.normalize()
            out[stk] = h
        except Exception:
            out[stk] = pd.Series(dtype=float)
    return out

WEEKLY_KEEP = 12   # Fridays kept: the latest one + the 11 before it

def held_now():
    """{(TYPE, STOCK)} of the positions open today."""
    if portfolio_df.empty:
        return set()
    h = portfolio_df[portfolio_df['Units'] > 0]
    return {(str(r['Type']).upper(), str(r['Stock']).strip().upper()) for _, r in h.iterrows()}

def weekly_compute(fridays, trades, overrides, held=None):
    """Unrealized_PL on each Friday's close for the stocks held that day (FIFO as of that date).
    Returns (rows, missing): a Friday with any unpriced holding goes in `missing` and is not returned,
    so a partial week is never saved."""
    rows, missing = [], {}
    if not fridays or trades.empty:
        return rows, missing
    t = trades.copy()
    t['Date of Trxn'] = pd.to_datetime(t['Date of Trxn'])
    stocks = sorted(t['Stock'].astype(str).str.strip().unique())
    hist = weekly_hist_closes(stocks, min(fridays) - timedelta(days=10), max(fridays))
    manual = {}
    if overrides is not None and not overrides.empty:
        manual = dict(zip(overrides['Stock'], overrides['Manual CMP']))
    for f in fridays:
        sub = t[t['Date of Trxn'] <= pd.Timestamp(f)]
        week, bad = [], []
        for (ptype, stk), _g in sub.groupby([sub['TYPE'].str.strip().str.upper(), sub['Stock'].astype(str).str.strip()]):
            if held is not None and (ptype, stk.upper()) not in held:
                continue                      # already sold: not tracked here
            lots = get_remaining_lots(sub, ptype, stk.upper())
            if not lots:
                continue
            units = sum(l['remaining'] for l in lots)
            cost = sum(l['remaining'] * l['price'] for l in lots)
            h = hist.get(stk, pd.Series(dtype=float))
            h = h[h.index <= pd.Timestamp(f)]
            if len(h) and (pd.Timestamp(f) - h.index[-1]).days <= 6:
                close = float(h.iloc[-1])        # the Friday close (or the last trading day before it)
            elif stk in manual and pd.notna(manual[stk]):
                close = float(manual[stk])       # price fetch failed: use your CMP override
            else:
                bad.append(stk)
                continue
            week.append((f, ptype, stk, float(units), float(cost), close, float(units * close - cost)))
        if bad:
            missing[f] = bad
        else:
            rows += week
    return rows, missing

def weekly_load():
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("SELECT friday, type, stock, units::float, cost::float, close_price::float, unrealized::float "
                    "FROM weekly_unrealized ORDER BY friday, type, stock")
        return pd.DataFrame(cur.fetchall(), columns=['friday', 'type', 'stock', 'units', 'cost', 'close', 'unrealized'])

def weekly_save(rows, replace=False):
    if not rows:
        return
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        if replace:
            for f in sorted({r[0] for r in rows}):
                cur.execute("DELETE FROM weekly_unrealized WHERE friday=%s", (f,))
        cur.executemany("INSERT INTO weekly_unrealized (friday, type, stock, units, cost, close_price, unrealized) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s) ON CONFLICT (friday, type, stock) DO NOTHING", rows)

def weekly_prune(window_start, held):
    """Keep only the last WEEKLY_KEEP Fridays and only the stocks held today."""
    db_init()
    with db_connect() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM weekly_unrealized WHERE friday < %s", (window_start,))
        if held:
            cur.execute("DELETE FROM weekly_unrealized WHERE (type, upper(stock)) NOT IN %s", (tuple(sorted(held)),))
        else:
            cur.execute("DELETE FROM weekly_unrealized")

def weekly_catch_up():
    """Save any Friday close of the last WEEKLY_KEEP weeks that is not stored yet (held stocks only), then drop older weeks and sold stocks."""
    held = held_now()
    window = completed_fridays(WEEKLY_KEEP)
    saved = weekly_load()
    have = {pd.Timestamp(x).date() for x in saved['friday']} if len(saved) else set()
    todo = [f for f in window if f not in have]
    missing = {}
    if todo:
        rows, missing = weekly_compute(todo, processed_trades, overrides_df, held)
        weekly_save(rows)
    weekly_prune(window[0], held)
    return missing

# ---- Core & Satellite Audit (same columns as the Core Audit / Satellite Audit tabs of the sheet) ----
def months_between(start, today):
    """Whole months from start to today (Excel DATEDIF(...,"M"))."""
    m = (today.year - start.year) * 12 + (today.month - start.month)
    return m - (1 if today.day < start.day else 0)

def audit_rows(ptype):
    """One row per open position of this type: start date = first buy date of the lots still held (FIFO)."""
    out = []
    today = ist_now().date()
    if portfolio_df.empty:
        return out
    sub = portfolio_df[(portfolio_df['Type'] == ptype) & (portfolio_df['Units'] > 0)].sort_values('Stock')
    for _, r in sub.iterrows():
        lots = get_remaining_lots(processed_trades, ptype, str(r['Stock']).strip().upper())
        start = min(l['date'] for l in lots).date() if lots else None
        inv, cur = float(r['Total_Cost']), float(r['Total_Value'])
        out.append({'Stock': r['Stock'], 'Start': start, 'Qty': int(r['Units']), 'Avg': float(r['Avg_Cost']),
                    'CMP': float(r['CMP']), 'Inv': inv, 'Cur': cur,
                    'HP': months_between(start, today) if start else None,
                    'ROI': (cur / inv - 1) if inv else 0.0})
    return out

def audit_cash_total():
    if not DATABASE_URL:
        return 0.0
    try:
        _, bal = load_cash()
        return float(sum(bal.get(a, 0.0) for a in ACCOUNTS))
    except Exception:
        return 0.0

def render_audit(title, ptype, other_type):
    own, oth = audit_rows(ptype), audit_rows(other_type)
    cash = audit_cash_total()
    inv_all = sum(r['Inv'] for r in own) + sum(r['Inv'] for r in oth) + cash
    cur_all = sum(r['Cur'] for r in own) + sum(r['Cur'] for r in oth) + cash
    pct = lambda a, b: "-" if not b else f"{a / b:.1%}"
    st.markdown(f"#### {title} ({len(own)})")
    if not own:
        st.caption("No open positions.")
        return
    if IS_MOBILE:
        cards = []
        for i, r in enumerate(own, start=1):
            cards.append(card_html(
                title=f"{i}. {r['Stock']}", right=f"{r['ROI']:+.1%}", right_color=pl_color(r['ROI']),
                items=[("Start", fmt_date(pd.Timestamp(r['Start'])) if r['Start'] else "-", None), ("HP (month)", "-" if r['HP'] is None else str(r['HP']), None),
                       ("Qty", f"{r['Qty']:,}", None), ("Avg buy", money(r['Avg']), None), ("CMP", money(r['CMP']), None),
                       ("Invested", money(r['Inv']), None), ("Current", money(r['Cur']), None),
                       ("IC", pct(r['Inv'], inv_all), None), ("CC", pct(r['Cur'], cur_all), None)]))
        render_cards(cards)
    else:
        rows = []
        for i, r in enumerate(own, start=1):
            rows.append([(i, None), (r['Stock'], None), (fmt_date(pd.Timestamp(r['Start'])) if r['Start'] else "-", None),
                         (f"{r['Qty']:,}", None), (money(r['Avg']), None), (money(r['CMP']), None),
                         (money(r['Inv']), None), (money(r['Cur']), None),
                         ("-" if r['HP'] is None else r['HP'], None),
                         (pct(r['Inv'], inv_all), None), (pct(r['Cur'], cur_all), None),
                         (f"{r['ROI']:.1%}", pl_color(r['ROI']))])
        render_compact_table(
            ['Sr. No.', 'STOCKS', 'Start Date', 'QTY', 'Avg Buy Price', 'CMP', 'Invested Value', 'Current Value',
             'HP (month)', 'IC (%)', 'CC (%)', 'ROI (%)'],
            rows, left_cols=(1,), font_px=12, fit=True, shrink=True)

def render_audit_split():
    """One Portfolio split table for the whole page: Core, Satellite, Cash, Total."""
    core, sat = audit_rows("CORE"), audit_rows("SATELLITE")
    cash = audit_cash_total()
    inv_c, cur_c = sum(r['Inv'] for r in core), sum(r['Cur'] for r in core)
    inv_s, cur_s = sum(r['Inv'] for r in sat), sum(r['Cur'] for r in sat)
    inv_all, cur_all = inv_c + inv_s + cash, cur_c + cur_s + cash
    pct = lambda a, b: "-" if not b else f"{a / b:.1%}"
    roi = lambda a, b: ("-", None) if not a else (f"{(b / a - 1):.1%}", pl_color(b / a - 1))
    summ = []
    for label, a, b in (("Core", inv_c, cur_c), ("Satellite", inv_s, cur_s), ("Cash", cash, cash), ("Total", inv_all, cur_all)):
        summ.append([(label, None), (money(a), None), (money(b), None), (pct(a, inv_all), None), (pct(b, cur_all), None), roi(a, b)])
    st.markdown("#### Portfolio split")
    if IS_MOBILE:
        render_cards([card_html(title=r[0][0], right=r[5][0], right_color=r[5][1],
                                items=[("Invested", r[1][0], None), ("Current", r[2][0], None),
                                       ("IC", r[3][0], None), ("CC", r[4][0], None)],
                                total=(r[0][0] == "Total")) for r in summ])
    else:
        render_compact_table(['Portfolio split', 'Invested Value', 'Current Value', 'IC (%)', 'CC (%)', 'ROI (%)'],
                             summ, left_cols=(0,), total_last=True, font_px=12, fit=True, shrink=True)

PAGES = ["Dashboard Overview", "Core Holdings", "Satellite Holdings",
         "Profit/Loss", "Weekly Unrealized P/L", "Core & Satellite Audit", "Individual Stock", "Trailing Stop Loss", "Trade Book & Add Trade", "Cash Flow", "CMP Overrides"]

def do_logout():
    st.session_state["authenticated"] = False

def do_refresh():
    st.cache_data.clear()

NAV_LABELS = {
    "Dashboard Overview": "🏠 Overview", "Core Holdings": "🎯 Core", "Satellite Holdings": "🚀 Satellite",
    "Profit/Loss": "📉 Profit/Loss", "Weekly Unrealized P/L": "📅 Weekly", "Core & Satellite Audit": "🧾 Audit", "Individual Stock": "🔍 Stock", "Trailing Stop Loss": "🛡️ Trailing SL", "Trade Book & Add Trade": "📒 Trades", "Cash Flow": "💰 Cash Flow",
    "CMP Overrides": "⚙️ CMP",
}

def _nav_changed():
    """Phone menu: tapping the highlighted button again would clear the choice; keep the current page."""
    v = st.session_state.get("nav_page")
    if v is None:
        st.session_state["nav_page"] = st.session_state.get("_last_page", PAGES[0])
    else:
        st.session_state["_last_page"] = v

# a page name saved by an older version of the menu (e.g. "Trade Book" / "Add Trade") would break the menu
for _k in ("nav_page", "_last_page"):
    if _k in st.session_state and st.session_state[_k] not in PAGES:
        st.session_state[_k] = PAGES[0]

if IS_MOBILE:
    st.session_state.setdefault("nav_page", PAGES[0])
    st.session_state.setdefault("_last_page", st.session_state["nav_page"])
    # every menu link is visible at once (wraps onto a few rows) — one tap to switch page
    page = st.pills("Menu", PAGES, selection_mode="single", key="nav_page",
                    format_func=lambda p: NAV_LABELS[p], label_visibility="collapsed",
                    on_change=_nav_changed)
    if page is None:
        page = st.session_state.get("_last_page", PAGES[0])
    with st.expander("🔄 Refresh prices / Logout"):
        st.button("🔄 Refresh Prices", on_click=do_refresh, use_container_width=True)
        st.button("Logout", on_click=do_logout, use_container_width=True)
else:
    st.sidebar.title("📌 Menu")
    page = st.sidebar.radio("Go to", PAGES, key="nav_page")

    if st.sidebar.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

    if st.sidebar.button("🔄 Refresh Prices"):
        st.cache_data.clear()
        st.rerun()

# --- KPI METRICS CARDS ---
if not portfolio_df.empty:
    curr_holdings = portfolio_df[portfolio_df['Units'] > 0]
    total_val = curr_holdings['Total_Value'].sum()
    total_cost = curr_holdings['Total_Cost'].sum()
    unrealized_pl = curr_holdings['Unrealized_PL'].sum()
    realized_pl = portfolio_df['Realized_PL'].sum()

    unreal_pct = (unrealized_pl / total_cost * 100) if total_cost > 0 else 0
    kpi_grid([
        ("Current Value", money(total_val), None, None),
        ("Total Invested", money(total_cost), None, None),
        ("Unrealized P&L", money(unrealized_pl), f"{unreal_pct:+.2f}%", pl_color(unrealized_pl)),
        ("Realized P&L", money(realized_pl), None, None),
    ])

st.markdown("---")

# --- PAGES ---
FORMAT_DICT = {
    'Avg_Cost': '₹{:.2f}', 'Total_Cost': '₹{:.2f}', 'CMP': '₹{:.2f}',
    'Total_Value': '₹{:.2f}', 'Unrealized_PL': '₹{:.2f}',
    'Realized_PL': '₹{:.2f}', 'Total_Net_Gain': '₹{:.2f}', 'Allocation_%': '{:.1%}',
    'Yest_Close': '₹{:.2f}', 'Day_Change': '₹{:.2f}'
}

def render_totals_bar(df, color):
    """Bottom-of-page totals as four small tiles."""
    unreal, real, net = df['Unrealized_PL'].sum(), df['Realized_PL'].sum(), df['Total_Net_Gain'].sum()
    kpi_grid([
        ("Total Cost", money(df['Total_Cost'].sum()), None, None),
        ("Unrealized P&L", money(unreal), None, None, pl_color(unreal)),
        ("Realized P&L", money(real), None, None, pl_color(real)),
        ("Total Net Gain", money(net), None, None, pl_color(net)),
    ], small=True)

def go_to_stock(label):
    """Button callback: jump to the Individual Stock page with this stock pre-selected."""
    st.session_state["nav_page"] = "Individual Stock"
    st.session_state["_last_page"] = "Individual Stock"
    st.session_state["ind_stock_choice"] = label

def stock_label(ptype, stock, is_open=True):
    """Label used by the Individual Stock selector."""
    return f"{ptype} — {stock}" + ("" if is_open else "  (closed)")

def _link_listener():
    """One tiny script (per page load) that turns clicks on a table's stock link into a click on the matching hidden chip."""
    import streamlit.components.v1 as components
    components.html("""<script>
const P = window.parent, d = P.document;
if (P.__stkH) d.removeEventListener('click', P.__stkH, true);
P.__stkH = function (e) {
  const a = e.target.closest && e.target.closest('a.stk-link');
  if (!a) return;
  e.preventDefault();
  const box = d.querySelector('.st-key-hp_' + a.dataset.pill);
  if (!box) return;
  const b = Array.from(box.querySelectorAll('button')).find(x => x.innerText.trim() === a.dataset.stock);
  if (b) b.click();
};
d.addEventListener('click', P.__stkH, true);
</script>""", height=0)

def stock_link_pills(ptype, stocks, key, closed=()):
    """Desktop: a row of clickable stock chips; picking one opens its Individual Stock page."""
    if not stocks:
        return
    closed = set(closed)

    def _open():
        picked = st.session_state.get(key)
        st.session_state[key] = None          # reset so the same chip can be used again later
        if picked:
            go_to_stock(stock_label(ptype, picked, picked not in closed))

    # the chips stay in the page (hidden); the stock names inside the tables click them
    with st.container(key=f"hp_{key}"):
        st.pills("🔍 Open a stock's analysis", list(stocks), selection_mode="single", key=key, on_change=_open)
    _link_listener()

def render_holdings_grid(df, ptype):
    """Core / Satellite holdings. Phone: a tappable stock link + small card per holding.
    Computer: the same compact bordered table as the Dashboard Overview, with a row of
    stock buttons above it that open the Individual Stock analysis."""
    if IS_MOBILE:
        for i, (_, r) in enumerate(df.iterrows(), start=1):
            st.button(f"{i}.  {r['Stock']}", key=f"hold_{ptype}_{i}_{r['Stock']}", type="tertiary",
                      on_click=go_to_stock, args=(f"{ptype} — {r['Stock']}",))
            render_cards([card_html(items=[
                ("CMP", money(r['CMP']), None), ("Units", f"{int(r['Units']):,}", None),
                ("Avg", money(r['Avg_Cost']), None), ("Invested", money(r['Total_Cost']), None),
                ("Value", money(r['Total_Value']), None), ("Unrealized", money(r['Unrealized_PL']), pl_color(r['Unrealized_PL'])),
                ("Realized", money(r['Realized_PL']), None), ("Net gain", money(r['Total_Net_Gain']), None),
                ("Alloc", f"{r['Allocation_%']:.1%}", None),
            ])])
        render_cards([card_html(title="Total", right=money(df['Total_Value'].sum()), total=True,
                                items=[("Total value", money(df['Total_Value'].sum()), None),
                                       ("Allocation", f"{df['Allocation_%'].sum():.1%}", None)])])
        return

    stock_link_pills(ptype, df['Stock'].tolist(), f"open_stock_{ptype}")
    render_overview_open("", df, heading=False, sum_value_alloc=True, fit=True, ptype=ptype)

def day_change_total(df):
    s = df['Day_Change']
    return s.sum() if s.notna().any() else float('nan')

def tranche_count(ptype, stock, cmp_):
    lots = get_remaining_lots(processed_trades, ptype, stock)
    if not lots:
        return None
    return int(len(build_tranche_table(lots, float(cmp_))))

def render_overview_open(title, df, heading=True, sum_value_alloc=False, link_type=None, fit=False, ptype=None):
    pill_key = f"ov_open_{link_type}" if link_type else (f"open_stock_{ptype}" if ptype else None)
    if heading:
        st.markdown(f"#### {title} ({len(df)})")
    if df.empty:
        st.caption("No open positions.")
        return
    rows = []
    cards = []
    if link_type and not IS_MOBILE:
        stock_link_pills(link_type, df['Stock'].tolist(), f"ov_open_{link_type}")
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        # CMP is coloured by today's move: above yesterday's close = green, below = red
        cmp_color = None
        if pd.notna(r['Yest_Close']) and r['CMP'] > 0 and r['CMP'] != r['Yest_Close']:
            cmp_color = UP if r['CMP'] > r['Yest_Close'] else DOWN
        cards.append(card_html(
            title=f"{i}. {r['Stock']}", right=money(r['CMP']), right_color=cmp_color,
            items=[("Units", f"{int(r['Units']):,}", None), ("Avg", money(r['Avg_Cost']), None),
                   ("Yest", money(r['Yest_Close']), None), ("1 Day", money(r['Day_Change']), pl_color(r['Day_Change'])),
                   ("Invested", money(r['Total_Cost']), None), ("Value", money(r['Total_Value']), None),
                   ("Unrealized", money(r['Unrealized_PL']), None), ("Realized", money(r['Realized_PL']), None),
                   ("Net gain", money(r['Total_Net_Gain']), None), ("Alloc", f"{r['Allocation_%']:.1%}", None)]))
        rows.append([
            (i, None), (("LINK", pill_key, r['Stock']) if pill_key else r['Stock'], None), (f"{int(r['Units']):,}", None),
            (money(r['Avg_Cost']), None), (money(r['Total_Cost']), None),
            (money(r['CMP']), cmp_color), (money(r['Yest_Close']), None),
            (money(r['Day_Change']), pl_color(r['Day_Change'])),
            (money(r['Total_Value']), None), (money(r['Unrealized_PL']), None),
            (money(r['Realized_PL']), None), (money(r['Total_Net_Gain']), None),
            (f"{r['Allocation_%']:.1%}", None),
            ("—" if (tc := tranche_count(ptype or link_type or r.get('Type'), r['Stock'], r['CMP'])) is None else tc, None),
        ])
    tot = day_change_total(df)
    val_sum, alloc_sum, cost_sum = df['Total_Value'].sum(), df['Allocation_%'].sum(), df['Total_Cost'].sum()
    unreal_sum = df['Unrealized_PL'].sum()
    if IS_MOBILE:
        extra = [("Total invested", money(cost_sum), None), ("Total value", money(val_sum), None),
                 ("Unrealized", money(unreal_sum), pl_color(unreal_sum)), ("Allocation", f"{alloc_sum:.1%}", None)]
        total_card = card_html(title="Total 1_Day_Change", right=money(tot), right_color=pl_color(tot),
                               items=extra, total=True)
        if link_type:
            for i, ((_, r), card) in enumerate(zip(df.iterrows(), cards), start=1):
                st.button(f"{i}.  {r['Stock']}", key=f"ovb_{link_type}_{i}_{r['Stock']}", type="tertiary",
                          on_click=go_to_stock, args=(stock_label(link_type, r['Stock']),))
                render_cards([card])
            render_cards([total_card])
        else:
            render_cards(cards + [total_card])
        return
    tail = [(money(val_sum), None), (money(unreal_sum), pl_color(unreal_sum)), ("", None), ("", None),
            (f"{alloc_sum:.1%}", None), ("", None)]
    rows.append([("", None), ("Total", None), ("", None), ("", None), (money(cost_sum), None), ("", None), ("", None),
                 (money(tot), pl_color(tot))] + tail)
    render_compact_table(
        ['#', 'Stock', 'Units', 'Avg_Cost', 'Total_Cost', 'CMP', 'Yest_Close', '1_Day_Change',
         'Total_Value', 'Unrealized_PL', 'Realized_PL', 'Total_Net_Gain', 'Allocation_%', 'Tranche'],
        rows, left_cols=(1,), total_last=True, font_px=12 if (link_type or fit) else 13, fit=bool(link_type) or fit,
        shrink=bool(link_type)
    )

def render_day_change_summary(core_df, sat_df):
    core_t, sat_t = day_change_total(core_df), day_change_total(sat_df)
    both = [x for x in (core_t, sat_t) if pd.notna(x)]
    grand = sum(both) if both else float('nan')
    kpi_grid([
        ("Core 1_Day_Change", money(core_t), None, None, pl_color(core_t)),
        ("Satellite 1_Day_Change", money(sat_t), None, None, pl_color(sat_t)),
        ("Total 1_Day_Change", money(grand), None, None, pl_color(grand)),
    ], small=True)

def render_overview_booked(title, df, trades):
    """P/L booked = every stock with at least one sale: fully squared off and partially sold."""
    ptype_b = str(df['Type'].iloc[0]) if len(df) else ""
    rows_data = []
    for _, r in df.iterrows():
        lots = build_closed_lots_table(trades, r['Type'], r['Stock'])
        if lots.empty:
            continue
        cost, proceeds, pl = lots['Invested'].sum(), lots['Proceeds'].sum(), lots['Realized P/L'].sum()
        rows_data.append({
            'Stock': r['Stock'],
            'Status': 'Squared off' if r['Units'] == 0 else 'Partial',
            'Units_Left': int(r['Units']), 'Qty_Sold': lots['Qty'].sum(),
            'Cost': cost, 'Proceeds': proceeds, 'PL': pl,
            'Ret': (pl / cost) if cost else float('nan'),
            'Last_Sale': lots['Sell Date'].max(),
        })
    if not rows_data:
        st.markdown(f"#### {title} (0)")
        st.caption("No sales booked yet.")
        return
    rows_data.sort(key=lambda d: d['Last_Sale'], reverse=True)
    if IS_MOBILE:
        t_cost = sum(d['Cost'] for d in rows_data)
        t_proc = sum(d['Proceeds'] for d in rows_data)
        t_pl = sum(d['PL'] for d in rows_data)
        cards = []
        for i, d in enumerate(rows_data, start=1):
            c = pl_color(d['PL'])
            cards.append(card_html(
                title=f"{i}. {d['Stock']}", badge=d['Status'], right=money(d['PL']), right_color=c,
                items=[("Qty sold", f"{d['Qty_Sold']:,.0f}", None), ("Units left", f"{d['Units_Left']:,}", None),
                       ("Cost", money(d['Cost']), None), ("Proceeds", money(d['Proceeds']), None),
                       ("Return", "-" if pd.isna(d['Ret']) else f"{d['Ret']:.1%}", c),
                       ("Last sale", fmt_date(d['Last_Sale']), None)]))
        cards.append(card_html(
            title="Total", right=money(t_pl), right_color=pl_color(t_pl), total=True,
            items=[("Cost", money(t_cost), None), ("Proceeds", money(t_proc), None),
                   ("Return", f"{t_pl / t_cost:.1%}" if t_cost else "-", pl_color(t_pl))]))
        with st.expander(f"{title} ({len(rows_data)})  ·  {money(t_pl)}", expanded=False):
            for i, d in enumerate(rows_data, start=1):
                st.button(f"{i}.  {d['Stock']}", key=f"bkb_{ptype_b}_{i}_{d['Stock']}", type="tertiary",
                          on_click=go_to_stock, args=(stock_label(ptype_b, d['Stock'], d['Units_Left'] > 0),))
                render_cards([cards[i - 1]])
            render_cards([cards[-1]])
        return
    st.markdown(f"#### {title} ({len(rows_data)})")
    stock_link_pills(ptype_b, [d['Stock'] for d in rows_data], f"bk_open_{ptype_b}",
                     closed=[d['Stock'] for d in rows_data if d['Units_Left'] == 0])
    rows = []
    for i, d in enumerate(rows_data, start=1):
        c = pl_color(d['PL'])
        rows.append([
            (i, None), (("LINK", f"bk_open_{ptype_b}", d['Stock']), None), (d['Status'], None), (f"{d['Units_Left']:,}", None),
            (f"{d['Qty_Sold']:,.0f}", None), (money(d['Cost']), None), (money(d['Proceeds']), None),
            (money(d['PL']), c), ("-" if pd.isna(d['Ret']) else f"{d['Ret']:.1%}", c),
            (fmt_date(d['Last_Sale']), None),
        ])
    t_cost = sum(d['Cost'] for d in rows_data)
    t_proc = sum(d['Proceeds'] for d in rows_data)
    t_pl = sum(d['PL'] for d in rows_data)
    rows.append([("", None), ("Total", None), ("", None), ("", None),
                 (f"{sum(d['Qty_Sold'] for d in rows_data):,.0f}", None),
                 (money(t_cost), None), (money(t_proc), None), (money(t_pl), pl_color(t_pl)),
                 (f"{t_pl / t_cost:.1%}" if t_cost else "-", pl_color(t_pl)), ("", None)])
    render_compact_table(
        ['#', 'Stock', 'Status', 'Units_Left', 'Qty_Sold', 'Cost_of_Sold', 'Sale_Proceeds',
         'Realized_PL', 'Return_%', 'Last_Sale'],
        rows, left_cols=(1, 2), total_last=True, font_px=12, fit=True, shrink=True
    )

def render_pl_booked(title, df, is_loss):
    """Squared-off stocks only. Avg_Cost = average buy price of the units sold (FIFO).
    The Realized_PL total at the bottom is highlighted (red for losses, green for profits)."""
    st.markdown(f"#### {title} ({len(df)})")
    if df.empty:
        st.caption("None yet.")
        return
    total = df['Realized_PL'].sum()
    tint, ink = ("#FDE2E2", "#B71C1C") if is_loss else ("#DFF5E1", "#1B5E20")
    if IS_MOBILE:
        cards = []
        for i, (_, r) in enumerate(df.iterrows(), start=1):
            c = pl_color(r['Realized_PL'])
            cards.append(card_html(
                title=f"{i}. {r['Stock']}", badge=r['Type'], right=money(r['Realized_PL']), right_color=c,
                items=[("Avg cost", money(r['Avg_Cost']), None), ("CMP", money(r['CMP']), None)]))
        cards.append(card_html(title="Total Realized_PL", right=money(total), right_color=ink,
                               total=True, bg=tint))
        render_cards(cards)
        return
    rows = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        rows.append([
            (i, None), (r['Type'], None), (r['Stock'], None),
            (money(r['Avg_Cost']), None), (money(r['Total_Cost']), None), (money(r['CMP']), None),
            (money(r['Total_Value']), None), (money(r['Unrealized_PL']), None),
            (money(r['Realized_PL']), pl_color(r['Realized_PL'])), (money(r['Total_Net_Gain']), None),
        ])
    rows.append([("", None), ("", None), ("Total", None), ("", None), ("", None), ("", None),
                 ("", None), ("", None), (money(total), ink, tint), ("", None)])
    render_compact_table(
        ['#', 'Type', 'Stock', 'Avg_Cost', 'Total_Cost', 'CMP', 'Total_Value',
         'Unrealized_PL', 'Realized_PL', 'Total_Net_Gain'],
        rows, left_cols=(1, 2), total_last=True, font_px=13
    )

def render_pl_partial(title, df, trades, is_loss=False):
    """Stocks that are still held but have already sold some units (partially realized), at a profit or a loss.
    Figures cover only the units sold (FIFO); Avg_Cost = average buy price of those units."""
    rows_data = []
    for _, r in df.iterrows():
        lots = build_closed_lots_table(trades, r['Type'], r['Stock'])
        if lots.empty:
            continue
        qty, cost, proceeds, pl = lots['Qty'].sum(), lots['Invested'].sum(), lots['Proceeds'].sum(), lots['Realized P/L'].sum()
        rows_data.append({'Type': r['Type'], 'Stock': r['Stock'], 'Units_Left': int(r['Units']), 'Qty': qty,
                          'Avg': (cost / qty) if qty else 0.0, 'Cost': cost, 'Proceeds': proceeds, 'PL': pl,
                          'Ret': (pl / cost) if cost else float('nan')})
    rows_data.sort(key=lambda d: d['PL'], reverse=not is_loss)
    st.markdown(f"#### {title} ({len(rows_data)})")
    if not rows_data:
        st.caption("None yet.")
        return
    t_qty = sum(d['Qty'] for d in rows_data)
    t_cost = sum(d['Cost'] for d in rows_data)
    t_proc = sum(d['Proceeds'] for d in rows_data)
    t_pl = sum(d['PL'] for d in rows_data)
    tint, ink = ("#FDE2E2", "#B71C1C") if is_loss else ("#DFF5E1", "#1B5E20")
    kind = "loss" if is_loss else "profit"
    if IS_MOBILE:
        for i, d in enumerate(rows_data, start=1):
            st.button(f"{i}.  {d['Stock']}", key=f"ppb_{kind}_{d['Type']}_{i}_{d['Stock']}", type="tertiary",
                      on_click=go_to_stock, args=(stock_label(d['Type'], d['Stock']),))
            render_cards([card_html(badge=d['Type'], items=[
                ("Realized", money(d['PL']), pl_color(d['PL'])), ("Return", "-" if pd.isna(d['Ret']) else f"{d['Ret']:.1%}", None),
                ("Units left", f"{d['Units_Left']:,}", None), ("Qty sold", f"{d['Qty']:,.0f}", None),
                ("Avg cost", money(d['Avg']), None), ("Proceeds", money(d['Proceeds']), None)])])
        render_cards([card_html(title="Total Realized_PL", right=money(t_pl), right_color=ink, total=True, bg=tint,
                                items=[("Qty sold", f"{t_qty:,.0f}", None), ("Cost of sold", money(t_cost), None),
                                       ("Proceeds", money(t_proc), None)])])
        return
    # one chip row per type would mix labels, so link by the type each stock belongs to
    for ptype in sorted({d['Type'] for d in rows_data}):
        stock_link_pills(ptype, [d['Stock'] for d in rows_data if d['Type'] == ptype], f"pp_open_{kind}_{ptype}")
    rows = []
    for i, d in enumerate(rows_data, start=1):
        rows.append([(i, None), (d['Type'], None), (("LINK", f"pp_open_{kind}_{d['Type']}", d['Stock']), None), (f"{d['Units_Left']:,}", None),
                     (f"{d['Qty']:,.0f}", None), (money(d['Avg']), None), (money(d['Cost']), None),
                     (money(d['Proceeds']), None), (money(d['PL']), pl_color(d['PL'])),
                     ("-" if pd.isna(d['Ret']) else f"{d['Ret']:.1%}", pl_color(d['PL']))])
    rows.append([("", None), ("", None), ("Total", None), ("", None), (f"{t_qty:,.0f}", None), ("", None),
                 (money(t_cost), None), (money(t_proc), None), (money(t_pl), ink, tint),
                 (f"{t_pl / t_cost:.1%}" if t_cost else "-", ink)])
    render_compact_table(
        ['#', 'Type', 'Stock', 'Units_Left', 'Qty_Sold', 'Avg_Cost', 'Cost_of_Sold', 'Sale_Proceeds',
         'Realized_PL', 'Return_%'],
        rows, left_cols=(1, 2), total_last=True, font_px=12, fit=True, shrink=True
    )

def trailing_sl_table(ptype):
    """One row per open position of this type: last entered ('trailing') stop-loss vs CMP."""
    out = []
    sub = portfolio_df[(portfolio_df['Type'] == ptype) & (portfolio_df['Units'] > 0)].sort_values('Stock')
    for _, r in sub.iterrows():
        lots = get_remaining_lots(processed_trades, ptype, r['Stock'])
        if not lots:
            continue
        manual = {}
        if not manual_sl_df.empty:
            ms = manual_sl_df[(manual_sl_df['TYPE'] == ptype) & (manual_sl_df['Stock'] == r['Stock'])]
            for _, mr in ms.iterrows():
                manual[(str(mr['Buy_Date'])[:10], round(float(mr['Buy_Price']), 2))] = float(mr['SL_Price'])
        tdf = build_tranche_table(lots, float(r['CMP']), manual_sl=manual)
        entered = tdf['Stop-Loss Price'].dropna()
        tsl = float(entered.iloc[-1]) if len(entered) else float('nan')
        cmp_ = float(r['CMP'])
        out.append({'Stock': r['Stock'], 'WAC': float(tdf['Weighted Avg Cost'].iloc[-1]), 'Tranche': int(len(tdf)), 'Trailing_SL': tsl, 'CMP': cmp_,
                    'Diff': (cmp_ - tsl) if pd.notna(tsl) else float('nan'),
                    'Diff_pct': ((cmp_ - tsl) / tsl) if pd.notna(tsl) and tsl else float('nan')})
    return out

def render_trailing_sl(title, ptype):
    data = sorted(trailing_sl_table(ptype), key=lambda d: (-d['Tranche'], d['Stock']))   # most tranches first
    st.markdown(f"#### {title} ({len(data)})")
    if not data:
        st.caption("No open positions.")
        return
    if IS_MOBILE:
        for i, d in enumerate(data, start=1):
            c = None if pd.isna(d['Diff']) else (UP if d['Diff'] >= 0 else DOWN)
            st.button(f"{i}.  {d['Stock']}", key=f"tslb_{ptype}_{i}_{d['Stock']}", type="tertiary",
                      on_click=go_to_stock, args=(stock_label(ptype, d['Stock']),))
            render_cards([card_html(items=[
                ("Wtd Avg Cost", money(d['WAC'], 0), None), ("Tranche", str(d['Tranche']), None),
                ("CMP", money(d['CMP'], 0), None),
                ("Trailing SL", "—" if pd.isna(d['Trailing_SL']) else money(d['Trailing_SL'], 0), None),
                ("Difference from CMP", "—" if pd.isna(d['Diff']) else f"{money(d['Diff'], 0)} ({d['Diff_pct']:+.0%})", c)])])
        return
    stock_link_pills(ptype, [d['Stock'] for d in data], f"tsl_open_{ptype}")
    rows = []
    for i, d in enumerate(data, start=1):
        c = None if pd.isna(d['Diff']) else (UP if d['Diff'] >= 0 else DOWN)
        rows.append([
            (i, None), (("LINK", f"tsl_open_{ptype}", d['Stock']), None), (money(d['WAC'], 0), None), (d['Tranche'], None),
            ("—" if pd.isna(d['Trailing_SL']) else money(d['Trailing_SL'], 0), None),
            (money(d['CMP'], 0), None),
            ("—" if pd.isna(d['Diff']) else f"{money(d['Diff'], 0)} ({d['Diff_pct']:+.0%})", c),
        ])
    render_compact_table(['SL', 'STOCK', 'Weighted Avg Cost', 'Tranche', 'Trailing SL', 'CMP', 'Difference from CMP'], rows,
                         left_cols=(1,), auto_width=True, font_px=13)

# ---- weekly Unrealized_PL snapshots: catch up once per session (Friday closes are fetched from price history) ----
_weekly_missing = {}
if DATABASE_URL and not portfolio_df.empty:
    _wk_flag = f"wk_checked_{completed_fridays(1)[0]}"
    if not st.session_state.get(_wk_flag):
        try:
            _weekly_missing = weekly_catch_up()
            st.session_state[_wk_flag] = True
        except Exception as _exc:
            _weekly_missing = {"error": [str(_exc)]}

# ---- one menu entry for both: choose Trade Book or Add Trade inside it ----
if page == "Trade Book & Add Trade":
    _sec = st.segmented_control("Section", ["📒 Trade Book", "➕ Add Trade"], default="📒 Trade Book",
                                key="tb_section", label_visibility="collapsed")
    page = "Add Trade" if _sec == "➕ Add Trade" else "Trade Book"

if page == "Dashboard Overview":
    st.subheader("📊 Dashboard Overview")
    if portfolio_df.empty:
        st.info("No trades yet. Add one from the 'Trade Book & Add Trade' page.")
    else:
        open_df = portfolio_df[portfolio_df['Units'] > 0]
        core_open = open_df[open_df['Type'] == 'CORE']
        sat_open = open_df[open_df['Type'] == 'SATELLITE']
        render_overview_open("🎯 Core Holdings", core_open, link_type="CORE")
        render_overview_open("🚀 Satellite Holdings", sat_open, link_type="SATELLITE")
        render_day_change_summary(core_open, sat_open)
        render_overview_booked("✅ Core P/L Booked", portfolio_df[portfolio_df['Type'] == 'CORE'], processed_trades)
        render_overview_booked("✅ Satellite P/L Booked", portfolio_df[portfolio_df['Type'] == 'SATELLITE'], processed_trades)

elif page == "Core Holdings":
    st.subheader("🎯 Core Holdings")
    core = portfolio_df[(portfolio_df['Type'] == 'CORE') & (portfolio_df['Units'] > 0)].drop(columns=['Type']).reset_index(drop=True)
    core.index = core.index + 1
    render_holdings_grid(core, 'CORE')
    render_totals_bar(core, '#1F77B4')

elif page == "Satellite Holdings":
    st.subheader("🚀 Satellite Holdings")
    sat = portfolio_df[(portfolio_df['Type'] == 'SATELLITE') & (portfolio_df['Units'] > 0)].drop(columns=['Type']).reset_index(drop=True)
    sat.index = sat.index + 1
    render_holdings_grid(sat, 'SATELLITE')
    render_totals_bar(sat, '#D62728')

elif page == "Profit/Loss":
    st.subheader("📉 Profit / Loss Booked")
    squared = portfolio_df[portfolio_df['Units'] == 0].copy()
    avg_costs = []
    for _, r in squared.iterrows():
        lots_done = build_closed_lots_table(processed_trades, r['Type'], r['Stock'])
        qty = lots_done['Qty'].sum() if not lots_done.empty else 0
        avg_costs.append(lots_done['Invested'].sum() / qty if qty else 0.0)
    squared['Avg_Cost'] = avg_costs
    render_pl_booked("🔻 Booked Losses & Closed Positions",
                     squared[squared['Realized_PL'] < 0].sort_values('Realized_PL'), is_loss=True)
    render_pl_booked("🟢 Booked Profits & Closed Positions",
                     squared[squared['Realized_PL'] > 0].sort_values('Realized_PL', ascending=False), is_loss=False)
    held_partial = portfolio_df[(portfolio_df['Units'] > 0) & (portfolio_df['Realized_PL'] > 0)]
    render_pl_partial("🟡 Partially Realized Profits (still holding)", held_partial, processed_trades)
    held_partial_loss = portfolio_df[(portfolio_df['Units'] > 0) & (portfolio_df['Realized_PL'] < 0)]
    render_pl_partial("🟠 Partially Realized Losses (still holding)", held_partial_loss, processed_trades, is_loss=True)

elif page == "Individual Stock":
    st.subheader("🔍 Individual Stock Analysis")
    GREEN, RED = "#1B8A3A", "#C62828"

    if portfolio_df.empty:
        st.info("No positions to analyze.")
    else:
        all_pos = portfolio_df.copy()
        all_pos['Open'] = all_pos['Units'] > 0
        all_pos = all_pos.sort_values(['Open', 'Stock'], ascending=[False, True])
        options = {}
        for _, r in all_pos.iterrows():
            label = f"{r['Type']} — {r['Stock']}" + ("" if r['Open'] else "  (closed)")
            options[label] = (r['Type'], r['Stock'])
        if st.session_state.get("ind_stock_choice") not in options:
            st.session_state.pop("ind_stock_choice", None)
        choice = st.selectbox("Select a stock (open positions first, then closed)", list(options.keys()),
                              key="ind_stock_choice")
        sel_type, sel_stock = options[choice]

        fifo_df = build_fifo_calc_table(processed_trades, sel_type, sel_stock)
        fifo_rows = []
        for _, r in fifo_df.iterrows():
            fifo_rows.append([
                (r['Type'], None), (r['Stock'], None), (fmt_date(r['Buy_Date']), None),
                (fmt_num(r['Buy_Price']), None), (fmt_num(r['Buy_Qty'], 0), None),
                (fmt_num(r['Cum_Buy'], 0), None), (fmt_num(r['Cum_Sold'], 0), None),
                (fmt_num(r['Depleted'], 0), None), (fmt_num(r['Rem_Qty'], 0), None),
                (fmt_num(r['Rem_Value']), None),
            ])
        render_compact_table(
            ['Type', 'Stock', 'Buy_Date', 'Buy_Price', 'Buy_Qty', 'Cum_Buy',
             'Cum_Sold', 'Depleted', 'Rem_Qty', 'Rem_Value'],
            fifo_rows, left_cols=(0, 1), auto_width=True
        )

        if IS_MOBILE:
            st.caption("↔ Swipe a table sideways to see all its columns.")

        lots = get_remaining_lots(processed_trades, sel_type, sel_stock)
        pos_row = all_pos[(all_pos['Type'] == sel_type) & (all_pos['Stock'] == sel_stock)].iloc[0]
        cmp_price = float(pos_row['CMP'])

        if not lots:
            st.info("No open position — this stock is fully squared off. Its realized results are below.")
        else:
            stock_manual = {}
            if not manual_sl_df.empty:
                msub = manual_sl_df[(manual_sl_df['TYPE'] == sel_type) & (manual_sl_df['Stock'] == sel_stock)]
                for _, mr in msub.iterrows():
                    stock_manual[(str(mr['Buy_Date'])[:10], round(float(mr['Buy_Price']), 2))] = float(mr['SL_Price'])
            tranche_df = build_tranche_table(lots, cmp_price, manual_sl=stock_manual)
            # Trailing SL = the last stop-loss that was entered: the latest manual (Shaktiman) value if any,
            # otherwise the formula SL of the last tranche among 1-3
            _sl_entered = tranche_df['Stop-Loss Price'].dropna()
            trailing_sl = float(_sl_entered.iloc[-1]) if len(_sl_entered) else float('nan')

            total_qty = tranche_df['Cum. Qty'].iloc[-1]
            total_investment = tranche_df['Investment'].sum()
            weighted_avg = tranche_df['Weighted Avg Cost'].iloc[-1]
            current_sl_price = tranche_df['Stop-Loss Price'].iloc[-1]
            current_risk = tranche_df['Risk Amount (₹)'].iloc[-1]
            sl_pending = pd.isna(current_sl_price)
            current_val = total_qty * cmp_price
            current_pl = current_val - total_investment
            current_pl_pct = (current_pl / total_investment) if total_investment > 0 else 0

            if sl_pending:
                sl_value, sl_sub, sl_sub_color = "—", "Enter the Shaktiman SL below", "#B26A00"
            elif current_risk > 0:
                sl_value, sl_sub, sl_sub_color = money(current_sl_price), f"Risk ₹{current_risk:,.2f}", DOWN
            elif current_risk < 0:
                sl_value, sl_sub, sl_sub_color = money(current_sl_price), f"Locks in ₹{abs(current_risk):,.2f} profit", UP
            else:
                sl_value, sl_sub, sl_sub_color = money(current_sl_price), "Risk-free (breakeven)", UP
            kpi_grid([
                ("CMP", money(cmp_price), None, None),
                ("Weighted Avg Cost", money(weighted_avg), None, None),
                ("Unrealized P&L", money(current_pl), f"{current_pl_pct*100:+.2f}%", pl_color(current_pl)),
            ])
            if not sl_pending and current_sl_price > 0:
                sl_gap_pct = (cmp_price - current_sl_price) / current_sl_price * 100
                sl_pct_txt = f"CMP {abs(sl_gap_pct):.1f}% {'above' if sl_gap_pct >= 0 else 'below'} SL"
                sl_pct_col = UP if sl_gap_pct >= 0 else DOWN
                sl_tile = ("Current Stop-Loss", sl_value, sl_pct_txt, sl_pct_col, sl_pct_col)
            else:
                sl_tile = ("Current Stop-Loss", sl_value, sl_sub, sl_sub_color, sl_sub_color)
            kpi_grid([
                ("Total Invested Amount", money(total_investment), None, None),
                ("Current Value", money(current_val), None, None),
                sl_tile,
            ], small=True)

            st.markdown("#### Tranche Breakdown & Stop-Loss (open position)")
            MANUAL_BG = "#FFE08A"   # yellow = enter this cell by hand
            t_rows = []
            for _, r in tranche_df.iterrows():
                gp, gc = r['Growth % vs Prev Tranche'], r['Growth % vs CMP']
                pl = r['Tranche P/L (₹)']
                sl_p, risk = r['Stop-Loss Price'], r['Risk Amount (₹)']
                trail_pl = r['Qty'] * (trailing_sl - r['Buy Price']) if pd.notna(trailing_sl) else float('nan')
                if r['Manual SL']:
                    sl_cell = (("enter ↓" if pd.isna(sl_p) else fmt_num(sl_p)), None, MANUAL_BG)
                else:
                    sl_cell = (fmt_num(sl_p), None)
                t_rows.append([
                    (int(r['Tranche']), None), (r['Buy Date'], None),
                    (fmt_num(r['Buy Price']), None), (fmt_num(r['Qty'], 0), None),
                    (fmt_num(r['Investment'], 0), None), (fmt_num(r['Cum. Qty'], 0), None),
                    (fmt_num(r['Weighted Avg Cost']), None),
                    ("—" if pd.isna(gp) else f"{gp:.1%}", None),
                    ("—" if pd.isna(gc) else f"{gc:.1%}", GREEN if gc >= 0 else RED),
                    (r['Stop-Loss Rule'], None),
                    sl_cell,
                    ("—" if pd.isna(risk) else fmt_num(risk, 0), None),
                    (fmt_num(pl, 0), GREEN if pl >= 0 else RED),
                    ("—" if pd.isna(trailing_sl) else fmt_num(trailing_sl), None),
                    ("—" if pd.isna(trailing_sl) else fmt_num(trail_pl, 0),
                     None if pd.isna(trailing_sl) else (GREEN if trail_pl >= 0 else RED)),
                ])
            render_compact_table(
                ['#', 'Date', 'Buy ₹', 'Qty', 'Inv ₹', 'Cum Qty', 'Avg ₹',
                 'Δ Prev', 'Δ CMP', 'SL Rule', 'SL ₹', 'Risk ₹', 'P/L ₹', 'Trailing SL ₹', 'P/L @ Trail SL ₹'],
                t_rows, left_cols=(1, 9), font_px=11
            )
            if not pd.isna(trailing_sl):
                trail_total = float((tranche_df['Qty'] * (trailing_sl - tranche_df['Buy Price'])).sum())
                st.caption(f"Trailing SL ₹{trailing_sl:,.2f} = the last stop-loss entered. P/L @ Trail SL = "
                           f"Qty × (Trailing SL − Buy ₹): what each tranche makes or loses if the stop is hit. "
                           f"Total if hit: {'-' if trail_total < 0 else ''}₹{abs(trail_total):,.0f}.")
            st.caption("SL rule: Tranche 1 = 10% below its buy price · Tranche 2 = same ₹ risk as "
                       "Tranche 1 · Tranche 3 = breakeven (avg cost of all tranches) · "
                       "Tranche 4+ = Shaktiman — stop-loss entered by hand (yellow cell).")

            manual_rows = tranche_df[tranche_df['Manual SL']]
            if not manual_rows.empty:
                st.markdown("##### ✍️ Shaktiman stop-loss — enter by hand")
                with st.form(f"sl_form_{sel_type}_{sel_stock}"):
                    entries = {}
                    for _, mr in manual_rows.iterrows():
                        existing = mr['Stop-Loss Price']
                        entries[(mr['Buy Date'], round(float(mr['Buy Price']), 2))] = st.number_input(
                            f"Tranche {int(mr['Tranche'])} — bought {mr['Buy Date']} @ ₹{mr['Buy Price']:,.2f}",
                            min_value=0.0, value=0.0 if pd.isna(existing) else float(existing),
                            step=0.05, format="%.2f",
                            key=f"sl_{sel_type}_{sel_stock}_{mr['Buy Date']}_{mr['Buy Price']}")
                    if st.form_submit_button("💾 Save stop-loss"):
                        try:
                            save_manual_sl(sel_type, sel_stock, entries)
                            load_sl_history.clear()
                        except Exception as exc:
                            st.error(f"Could not save the stop-loss: {exc}")
                        else:
                            st.rerun()
                if DATABASE_URL:
                    st.caption("Saved in the database — your trailing stop-loss values survive every deploy.")
                else:
                    st.caption("⚠️ No database connected: values are kept in the app file and reset on deploy.")
                hist = load_sl_history(sel_type, sel_stock)
                if not hist.empty:
                    with st.expander(f"🕘 Stop-loss history ({len(hist)} changes)"):
                        h_rows = []
                        for _, hr in hist.iterrows():
                            ts = pd.Timestamp(hr['changed_at']).tz_convert('Asia/Kolkata')
                            h_rows.append([
                                (f"{ts.day}-{ts.strftime('%b-%Y %H:%M')}", None),
                                (f"{hr['buy_date']} @ ₹{hr['buy_price']:,.2f}", None),
                                ("cleared" if hr['sl_price'] == 0 else money(hr['sl_price']), None),
                            ])
                        render_compact_table(['When (IST)', 'Tranche (bought)', 'Stop-loss ₹'], h_rows,
                                             left_cols=(0, 1), auto_width=True)

            st.markdown("#### Profit Target Ladder")
            pcts = [10, 15, 20, 25]
            render_compact_table(
                ['Profit target'] + [f"+{p}%" for p in pcts],
                [[("Target price ₹", None)] + [(fmt_num(weighted_avg * (1 + p / 100)), None) for p in pcts]],
                left_cols=(0,), auto_width=True
            )

            st.markdown("#### Risk : Reward")
            reward_15pct = total_investment * 0.15
            if sl_pending:
                st.caption("Enter the Shaktiman stop-loss above to see Risk : Reward.")
            elif current_risk > 0:
                ratio = reward_15pct / current_risk
                st.markdown(
                    f"<p style='font-size:20px; font-weight:800; color:#8B0000;'>"
                    f"1 : {ratio:.1f}  <span style='font-size:14px; font-weight:400; color:#555;'>"
                    f"(risk ₹{current_risk:,.0f} vs. reward at 15% profit target)</span></p>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    "<p style='font-size:20px; font-weight:800; color:#008000;'>"
                    "Position is risk-free — stop-loss is at or above breakeven.</p>",
                    unsafe_allow_html=True
                )

        st.markdown("#### Closed Positions — Realized P/L")
        closed_df = build_closed_lots_table(processed_trades, sel_type, sel_stock)
        if closed_df.empty:
            st.caption("No shares have been sold yet for this stock.")
        else:
            c_rows = []
            for _, r in closed_df.iterrows():
                pl, ret = r['Realized P/L'], r['Return %']
                col = GREEN if pl >= 0 else RED
                c_rows.append([
                    (int(r['Cycle']), None),
                    (fmt_date(r['Buy Date']), None), (fmt_num(r['Buy Price']), None),
                    (fmt_date(r['Sell Date']), None), (fmt_num(r['Sell Price']), None),
                    (fmt_num(r['Qty'], 0), None),
                    (fmt_num(r['Invested'], 0), None), (fmt_num(r['Proceeds'], 0), None),
                    (fmt_num(pl, 0), col),
                    ("—" if pd.isna(ret) else f"{ret:.1%}", col),
                    (int(r['Days Held']), None),
                ])
            render_compact_table(
                ['Cycle', 'Buy Date', 'Buy ₹', 'Sell Date', 'Sell ₹', 'Qty',
                 'Invested ₹', 'Proceeds ₹', 'Realized P/L ₹', 'Return', 'Days'],
                c_rows, left_cols=(1, 3)
            )
            st.caption("Each row is a sale matched FIFO to the buy lot it closed. Cycle 2, 3… means the stock "
                       "was bought again after the earlier position was fully squared off.")
            tot_pl = closed_df['Realized P/L'].sum()
            tot_inv = closed_df['Invested'].sum()
            tot_col = GREEN if tot_pl >= 0 else RED
            st.markdown(
                f"<p style='font-size:20px; font-weight:800; color:{tot_col};'>"
                f"Total Realized P/L: ₹{tot_pl:,.2f} "
                f"<span style='font-size:14px; font-weight:400;'>"
                f"({len(closed_df)} closed lot{'s' if len(closed_df) != 1 else ''}, {(tot_pl / tot_inv if tot_inv else 0):.1%} on ₹{tot_inv:,.0f} invested)"
                f"</span></p>",
                unsafe_allow_html=True
            )

elif page == "Trailing Stop Loss":
    st.subheader("🛡️ Trailing Stop Loss")
    if portfolio_df.empty:
        st.info("No positions yet.")
    else:
        render_trailing_sl("🚀 Satellite Holdings", "SATELLITE")
        render_trailing_sl("🎯 Core Holdings", "CORE")
        st.caption("Trailing SL = the last stop-loss entered for the stock: your latest Shaktiman value "
                   "(tranche 4+), otherwise the rule-based SL of its last tranche. Tranche = number of open "
                   "tranches. Difference from CMP = CMP − Trailing SL (green = CMP above the stop, red = below). Stocks are listed with the most tranches first.")

elif page == "Weekly Unrealized P/L":
    st.subheader("📅 Weekly Unrealized P/L")
    st.caption("Unrealized_PL of the stocks you held at each Friday's market close (FIFO cost, closing price of that day). "
               "Only stocks you hold now are shown; the last 12 Fridays are kept.")
    if not DATABASE_URL:
        st.info("Weekly snapshots are stored in the database, which is not connected here.")
    else:
        if _weekly_missing:
            _bad = "; ".join(f"{k}: {', '.join(v)}" for k, v in _weekly_missing.items())
            st.warning(f"Could not price some stocks, so these weeks were not saved yet (will retry): {_bad}")
        wk = weekly_load()
        _held = held_now()
        wk = wk[[ (str(t).upper(), str(x).strip().upper()) in _held for t, x in zip(wk['type'], wk['stock'])]] if len(wk) else wk
        if wk.empty:
            st.info("No weekly data saved yet.")
        else:
            wk['friday'] = pd.to_datetime(wk['friday'])
            all_f = sorted(wk['friday'].unique(), reverse=True)
            shown = all_f[:WEEKLY_KEEP]
            tot_by_f = wk.groupby('friday')['unrealized'].sum()
            def _lbl(f):
                return pd.Timestamp(f).strftime("%d-%b-%y")
            if IS_MOBILE:
                pick = st.selectbox("Friday", [ _lbl(f) for f in all_f ], index=0)
                f_sel = next(f for f in all_f if _lbl(f) == pick)
                day = wk[wk['friday'] == f_sel]
                cards = []
                for ptype in ("SATELLITE", "CORE"):
                    for _, r in day[day['type'] == ptype].sort_values('stock').iterrows():
                        cards.append(card_html(title=r['stock'], badge=ptype.title(), right=money(r['unrealized']),
                                               right_color=pl_color(r['unrealized']),
                                               items=[("Units", f"{r['units']:,.0f}", None), ("Cost", money(r['cost']), None),
                                                      ("Close", money(r['close']), None)]))
                cards.append(card_html(title="Total", right=money(float(tot_by_f[f_sel])),
                                       right_color=pl_color(float(tot_by_f[f_sel])), total=True, items=[]))
                render_cards(cards)
            else:
                hdr = ['#', 'Stock'] + [_lbl(f) for f in shown]
                for ptype, ttl in (("SATELLITE", "🚀 Satellite Holdings"), ("CORE", "🎯 Core Holdings")):
                    sub = wk[(wk['type'] == ptype) & (wk['friday'].isin(shown))]
                    if sub.empty:
                        continue
                    st.markdown(f"#### {ttl}")
                    piv = sub.pivot_table(index='stock', columns='friday', values='unrealized', aggfunc='sum')
                    order = sorted(piv.index)
                    trs = []
                    for i, stk in enumerate(order, start=1):
                        cells = [(i, None), (stk, None)]
                        for f in shown:
                            v = piv.loc[stk].get(f, float('nan'))
                            cells.append(("–", None) if pd.isna(v) else (money(v, 0), pl_color(v)))
                        trs.append(cells)
                    tots = [("", None), ("Total", None)]
                    for f in shown:
                        tv = float(sub[sub['friday'] == f]['unrealized'].sum()) if (sub['friday'] == f).any() else float('nan')
                        tots.append(("–", None) if pd.isna(tv) else (money(tv, 0), pl_color(tv)))
                    trs.append(tots)
                    render_compact_table(hdr, trs, left_cols=(1,), total_last=True, font_px=12, fit=True, shrink=True)
                st.markdown("#### 📊 All holdings")
                tr1 = [("", None), ("Unrealized P/L", None)]
                tr2 = [("", None), ("Change vs previous Friday", None)]
                for f in shown:
                    tv = float(tot_by_f[f])
                    tr1.append((money(tv, 0), pl_color(tv)))
                    fi = all_f.index(f)
                    if fi + 1 < len(all_f):
                        dv = tv - float(tot_by_f[all_f[fi + 1]])
                        tr2.append((money(dv, 0), pl_color(dv)))
                    else:
                        tr2.append(("–", None))
                render_compact_table(hdr, [tr1, tr2], left_cols=(1,), font_px=12, fit=True, shrink=True)
        with st.expander("⚙️ Manage snapshots"):
            st.caption("Snapshots are taken automatically for each completed Friday. Use these only to repair or extend history.")
            c1, c2 = st.columns(2)
            if c1.button("Recompute latest Friday (replace)"):
                _f = completed_fridays(1)
                _rows, _miss = weekly_compute(_f, processed_trades, overrides_df, held_now())
                weekly_save(_rows, replace=True)
                st.success(f"Saved {len(_rows)} stocks for {_f[0]:%d-%b-%Y}.") if _rows else st.warning(f"Not saved; unpriced: {_miss}")
            if c2.button("Fill any missing Fridays (last 12 weeks)"):
                _miss = weekly_catch_up()
                st.success("Up to date.") if not _miss else st.warning(f"Not saved (unpriced stocks): { {k.isoformat(): v for k, v in _miss.items()} }")

elif page == "Core & Satellite Audit":
    st.subheader("🧾 Core & Satellite Audit")
    if portfolio_df.empty:
        st.info("No positions yet.")
    else:
        render_audit("🎯 Core Audit", "CORE", "SATELLITE")
        render_audit("🚀 Satellite Audit", "SATELLITE", "CORE")
        render_audit_split()
        st.caption("Start Date = first buy date of the lots still held (FIFO). HP = whole months held. "
                   "IC = Invested Value ÷ total invested (Core + Satellite + Cash). CC = Current Value ÷ total current value. "
                   "ROI = Current ÷ Invested − 1. Cash = cash balance from the Cash Flow tab.")

elif page == "Trade Book":
    st.subheader("📒 Trade Book")
    if trades_df.empty:
        st.info("No trades recorded yet. Switch to ➕ Add Trade above to record one.")
    else:
        tb_view = trades_df.copy()
        tb_view['_id'] = tb_view.index
        tb_view['Date of Trxn'] = pd.to_datetime(tb_view['Date of Trxn'])
        tb_view['Amount'] = tb_view['Price'].astype(float) * tb_view['Unit'].astype(float)
        tb_view['Action'] = tb_view['Unit'].apply(lambda u: 'BUY' if float(u) > 0 else 'SELL')

        f1, f2, f3 = st.columns(3)
        type_sel = f1.selectbox("Type", ["All", "CORE", "SATELLITE"])
        stock_sel = f2.selectbox("Stock", ["All"] + sorted(tb_view['Stock'].astype(str).unique().tolist()))
        action_sel = f3.selectbox("Action", ["All", "BUY", "SELL"])

        if type_sel != "All":
            tb_view = tb_view[tb_view['TYPE'] == type_sel]
        if stock_sel != "All":
            tb_view = tb_view[tb_view['Stock'].astype(str) == stock_sel]
        if action_sel != "All":
            tb_view = tb_view[tb_view['Action'] == action_sel]

        tb_view = tb_view.sort_values('Date of Trxn', ascending=False).reset_index(drop=True)
        tb_view.index = tb_view.index + 1
        tb_view = tb_view[['TYPE', 'Stock', 'Date of Trxn', 'Action', 'Price', 'Unit', 'Amount', '_id']]

        # ---- table: tick a row's checkbox, then use ✏️ Edit / 🗑️ Delete just below ----
        if st.session_state.get("tb_msg"):
            st.success(st.session_state.pop("tb_msg"))

        disp = tb_view.drop(columns=['_id'])
        ev = st.dataframe(
            disp.style.format({
                'Date of Trxn': lambda d: d.strftime('%d-%b-%Y'),
                'Price': '₹{:,.2f}', 'Unit': '{:,.0f}', 'Amount': '₹{:,.2f}'
            }),
            use_container_width=True, on_select="rerun", selection_mode="single-row", key=f"tb_table_{st.session_state.get('tb_ver', 0)}"
        )
        sel_rows = ev.selection.rows if ev and ev.selection else []
        sel_id = int(tb_view['_id'].iloc[sel_rows[0]]) if sel_rows and sel_rows[0] < len(tb_view) else None
        if sel_id is None:
            st.caption("👆 Tick the box at the left of a row to edit or delete that trade.")
        else:
            sr = tb_view.iloc[sel_rows[0]]
            st.markdown(f"Selected: **{esc(str(sr['Stock']))}** · {sr['Date of Trxn'].strftime('%d-%b-%Y')} · "
                        f"{sr['Action']} {abs(float(sr['Unit'])):,.0f} @ ₹{float(sr['Price']):,.2f}")
            bb1, bb2, _sp = st.columns([1, 1, 3])
            if bb1.button("✏️ Edit", key="tb_btn_edit", use_container_width=True):
                st.session_state["tb_edit"] = sel_id
                st.session_state["tb_del"] = None
                st.rerun()
            if bb2.button("🗑️ Delete", key="tb_btn_del", use_container_width=True):
                st.session_state["tb_del"] = sel_id
                st.session_state["tb_edit"] = None
                st.rerun()

        edit_id = st.session_state.get("tb_edit")
        del_id = st.session_state.get("tb_del")
        if edit_id is not None and edit_id not in trades_df.index:
            edit_id = st.session_state["tb_edit"] = None
        if del_id is not None and del_id not in trades_df.index:
            del_id = st.session_state["tb_del"] = None

        if edit_id is not None:
            r = trades_df.loc[edit_id]
            try:
                r_date = pd.to_datetime(r['Date of Trxn']).date()
            except Exception:
                r_date = datetime.now().date()
            with st.form(f"tb_edit_form_{edit_id}"):
                st.markdown(f"**✏️ Edit trade — {esc(str(r['Stock']))}**")
                e1, e2 = st.columns(2)
                e_type = e1.selectbox("Type", ["CORE", "SATELLITE"], index=0 if str(r['TYPE']).upper() == "CORE" else 1)
                e_stock = e2.text_input("Stock Symbol", value=str(r['Stock']))
                e_date = e1.date_input("Transaction Date", r_date)
                e_price = e2.number_input("Price per Unit (₹)", min_value=0.01, step=0.01, value=float(r['Price']))
                e_units = e1.number_input("Units (+ Buy, − Sell)", step=1, value=int(float(r['Unit'])))
                b1, b2 = st.columns(2)
                do_save = b1.form_submit_button("💾 Save changes", type="primary")
                do_cancel = b2.form_submit_button("Cancel")
            if do_cancel:
                st.session_state["tb_edit"] = None
                st.rerun()
            if do_save:
                e_stock = e_stock.upper().strip()
                if not e_stock or e_units == 0:
                    st.warning("Stock symbol is required and Units cannot be 0.")
                else:
                    old_key = (str(r['TYPE']).upper(), str(r['Stock']).strip(), pd.to_datetime(r['Date of Trxn']).strftime('%Y-%m-%d'), round(float(r['Price']), 2))
                    new_key = (e_type, e_stock, e_date.strftime('%Y-%m-%d'), round(float(e_price), 2))
                    upd = trades_df.copy()
                    upd.loc[edit_id, ['TYPE', 'Stock', 'Date of Trxn', 'Price', 'Unit', 'Amount']] = \
                        [e_type, e_stock, e_date.strftime('%Y-%m-%d'), e_price, e_units, e_price * e_units]
                    save_data(upd, overrides_df)
                    if float(r['Unit']) > 0 and int(e_units) > 0:
                        try:
                            rekey_manual_sl(old_key, new_key)
                        except Exception as exc:
                            st.session_state["tb_msg"] = f"Trade updated, but its stop-loss could not be moved: {exc}"
                    st.session_state["tb_edit"] = None
                    st.session_state["tb_ver"] = st.session_state.get("tb_ver", 0) + 1
                    st.session_state.setdefault("tb_msg", f"Updated trade for {e_stock}.")
                    st.rerun()

        if del_id is not None:
            r = trades_df.loc[del_id]
            st.warning(
                f"Delete this trade? **{esc(str(r['Stock']))}** · {pd.to_datetime(r['Date of Trxn']).strftime('%d-%b-%Y')} · "
                f"{'BUY' if float(r['Unit']) > 0 else 'SELL'} {abs(float(r['Unit'])):,.0f} @ ₹{float(r['Price']):,.2f}. "
                f"This changes your holdings and P/L.")
            d1, d2 = st.columns(2)
            if d1.button("🗑️ Yes, delete", type="primary", key="tb_del_yes"):
                upd = trades_df.drop(index=del_id).reset_index(drop=True)
                save_data(upd, overrides_df)
                st.session_state["tb_del"] = None
                st.session_state["tb_ver"] = st.session_state.get("tb_ver", 0) + 1
                st.session_state["tb_msg"] = f"Deleted trade for {str(r['Stock'])}."
                st.rerun()
            if d2.button("Cancel", key="tb_del_no"):
                st.session_state["tb_del"] = None
                st.rerun()

        buys = tb_view[tb_view['Action'] == 'BUY']['Amount'].sum()
        sells = -tb_view[tb_view['Action'] == 'SELL']['Amount'].sum()
        st.markdown(
            f"""
            <div style='display:flex; justify-content:space-between; flex-wrap:wrap;
                        padding:14px 18px; margin-top:10px; background-color:#F5F5F5;
                        border-top:3px solid #6A1B9A; border-radius:4px;'>
                <div style='font-size:20px; font-weight:800; color:#6A1B9A;'>Trades: {len(tb_view)}</div>
                <div style='font-size:20px; font-weight:800; color:#6A1B9A;'>Total Bought: ₹{buys:,.2f}</div>
                <div style='font-size:20px; font-weight:800; color:#6A1B9A;'>Total Sold: ₹{sells:,.2f}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.download_button(
            "⬇️ Download as CSV",
            tb_view.drop(columns=['_id']).to_csv(index=False).encode('utf-8'),
            file_name="trade_book.csv",
            mime="text/csv"
        )

elif page == "Add Trade":
    st.subheader("📝 Record New Trade")
    if st.session_state.get("trade_saved_msg"):
        st.success(st.session_state.pop("trade_saved_msg"))
    known_symbols = sorted(trades_df['Stock'].astype(str).str.strip().str.upper().unique().tolist()) \
        if not trades_df.empty else []
    with st.form("new_trade_form"):
        c1, c2 = st.columns(2)
        with c1:
            p_type = st.selectbox("Type", ["CORE", "SATELLITE"])
            stock = st.selectbox(
                "Stock Symbol — type 2–3 letters to search",
                known_symbols, index=None, accept_new_options=True,
                placeholder="e.g. type BHE → BHEL  (new stock? type its full symbol)")
            stock = (stock or "").upper().strip()
            date = st.date_input("Transaction Date", datetime.now())
        with c2:
            price = st.number_input("Price per Unit (₹)", min_value=0.01, step=0.01)
            units = st.number_input("Units (+ for Buy, - for Sell)", step=1)

        submitted = st.form_submit_button("Save Transaction")
        if submitted and not stock:
            st.warning("Please pick or type a stock symbol first.")
        if submitted and stock:
            new_row = pd.DataFrame([{
                'TYPE': p_type,
                'Stock': stock,
                'Date of Trxn': date.strftime('%Y-%m-%d'),
                'Price': price,
                'Unit': units,
                'Amount': price * units
            }])
            updated_trades = pd.concat([trades_df, new_row], ignore_index=True)
            save_data(updated_trades, overrides_df)
            st.session_state["trade_saved_msg"] = (
                f"Added trade for {stock} successfully!" if stock in known_symbols else
                f"Added trade for {stock} — this is a NEW symbol (first trade for it). "
                f"If that was a typo, tell me and I'll fix the Trade Book.")
            st.rerun()

elif page == "Cash Flow":
    st.subheader("💰 Cash Flow")
    if not DATABASE_URL:
        st.info("The Cash Flow page needs the database (DATABASE_URL), which is not connected here.")
        st.stop()
    try:
        ent, bal = load_cash()
    except Exception as exc:
        st.error(f"Could not read the cash-flow data: {exc}")
        st.stop()

    def _sum(df, kind):
        return float(df.loc[df['kind'] == kind, 'amount'].sum()) if not df.empty else 0.0

    held = portfolio_df[portfolio_df['Units'] > 0] if not portfolio_df.empty else portfolio_df
    exposed = float(held['Total_Cost'].sum()) if len(held) else 0.0
    equity_val = float(held['Total_Value'].sum()) if len(held) else 0.0
    ovr = cash_get_override()
    if ovr:
        exposed, equity_val = ovr
    dep, back = _sum(ent, 'DEPOSIT'), _sum(ent, 'TRANSFER_BACK')
    bro_chg, dp_chg = _sum(ent, 'CHARGE'), _sum(ent, 'DP_CHARGE')
    chg = bro_chg + dp_chg
    cash_total = float(sum(bal.get(a, 0.0) for a in ACCOUNTS))
    net_dep = dep - back
    total_value = equity_val + cash_total   # worth now: shares at market value + cash
    gross = equity_val + cash_total - net_dep
    net_profit = gross - chg
    ret = (net_profit / net_dep) if net_dep else float('nan')

    st.markdown("#### Both accounts together")
    kpi_grid([
        ("Net Deployed", money(net_dep), None, None, None, "rgba(47,128,237,.20)"),
        ("Total Amount Deployed", money(dep), None, None),
        ("Total Transferred Back", money(back), None, None),
        ("Total Value (equity + cash)", money(total_value), None, None, None, "rgba(0,166,126,.20)"),
    ])
    kpi_grid([
        ("Exposed (Invested)", money(exposed), None, None),
        ("Cash Balance", money(cash_total), None, None),
        ("Equity Value on date", money(equity_val), None, None),
    ])
    kpi_grid([
        ("Statutory & Broker Charges", money(bro_chg), None, None),
        ("DP Charges", money(dp_chg), None, None),
        ("Charges Total", money(chg), None, None),
    ], narrow=True)
    kpi_grid([
        ("Gross Profit", money(gross), None, None, pl_color(gross)),
        ("Net Profit", money(net_profit), None, None, pl_color(net_profit)),
        ("Return on Net Deployed", "-" if pd.isna(ret) else f"{ret:.2%}", None, None, pl_color(net_profit)),
    ])
    st.caption(("Exposed and Equity Value are the figures you typed in (see ✏️ below)." if ovr else
                "Exposed and Equity Value come from the holdings in this app (both accounts together).") +
               " Gross Profit = Equity Value + Cash Balance − Net Deployed. Net Profit = Gross Profit − Statutory & Broker Charges − DP Charges.")

    for acct in ACCOUNTS:
        a_ent = ent[ent['account'] == acct]
        a_dep, a_back = _sum(a_ent, 'DEPOSIT'), _sum(a_ent, 'TRANSFER_BACK')
        a_bro, a_dp = _sum(a_ent, 'CHARGE'), _sum(a_ent, 'DP_CHARGE')
        st.markdown(f"#### {acct}")
        kpi_grid([
            ("Deposits", money(a_dep), None, None),
            ("Transferred Back", money(a_back), None, None),
            ("Net Deployed", money(a_dep - a_back), None, None),
            ("Statutory & Broker Charges", money(a_bro), None, None),
            ("DP Charges", money(a_dp), None, None),
            ("Cash Balance", money(bal.get(acct, 0.0)), None, None),
        ])

        def _dated_rows(kind):
            sub = a_ent[a_ent['kind'] == kind]
            rows = [[(i, None), (fmt_date(r['date']), None), (money(r['amount']), None)]
                    for i, (_, r) in enumerate(sub.iterrows(), start=1)]
            rows.append([("", None), ("Total", None), (money(float(sub['amount'].sum())), None)])
            return rows

        c1, c2, c3 = st.columns([1, 1, 1.5])
        with c1:
            st.markdown("**Deposits**")
            render_compact_table(['#', 'Date', 'Amount'], _dated_rows('DEPOSIT'), left_cols=(1,),
                                 auto_width=True, total_last=True, font_px=13)
        with c2:
            st.markdown("**Transferred back**")
            render_compact_table(['#', 'Date', 'Amount'], _dated_rows('TRANSFER_BACK'), left_cols=(1,),
                                 auto_width=True, total_last=True, font_px=13)
        with c3:
            st.markdown("**Charges** (by financial year)")
            ch = a_ent[a_ent['kind'].isin(['CHARGE', 'DP_CHARGE'])]
            fys = sorted(ch['label'].dropna().unique().tolist())
            ch_rows = []
            for fy in fys:
                b_amt = float(ch.loc[(ch['label'] == fy) & (ch['kind'] == 'CHARGE'), 'amount'].sum())
                d_amt = float(ch.loc[(ch['label'] == fy) & (ch['kind'] == 'DP_CHARGE'), 'amount'].sum())
                ch_rows.append([(fy, None), (money(b_amt), None), (money(d_amt), None), (money(b_amt + d_amt), None)])
            ch_rows.append([("Total", None), (money(a_bro), None), (money(a_dp), None), (money(a_bro + a_dp), None)])
            render_compact_table(['Fin. year', 'Statutory & Broker', 'DP Charges', 'Total'], ch_rows, left_cols=(0,),
                                 auto_width=True, total_last=True, font_px=13)

    if st.session_state.get("cash_msg"):
        st.success(st.session_state.pop("cash_msg"))

    with st.expander("✏️ Add or update entries"):
        with st.form("cash_flow_form", clear_on_submit=True):
            st.markdown("**Record a deposit or a transfer back**")
            f1, f2 = st.columns(2)
            f_acct = f1.selectbox("Account", ACCOUNTS)
            f_kind = f2.selectbox("Type", ["Deposit", "Transferred back"])
            f_date = f1.date_input("Date", datetime.now())
            f_amt = f2.number_input("Amount (₹)", min_value=0.0, step=100.0)
            if st.form_submit_button("Save entry", type="primary"):
                if f_amt <= 0:
                    st.warning("Enter an amount above 0.")
                else:
                    cash_add_entry(f_acct, 'DEPOSIT' if f_kind == "Deposit" else 'TRANSFER_BACK', f_date, f_amt)
                    st.session_state["cash_msg"] = f"Saved {f_kind.lower()} of {money(f_amt)} for {f_acct}."
                    st.rerun()
        with st.form("cash_charge_form", clear_on_submit=True):
            st.markdown("**Charges for a financial year** (saving the same head and year again replaces its amount)")
            g1, g2, g3, g4 = st.columns(4)
            g_acct = g1.selectbox("Account", ACCOUNTS, key="chg_acct")
            g_head = g2.selectbox("Charge head", ["Statutory & Broker Charges", "DP Charges"])
            g_fy = g3.text_input("Financial year", value="2026-27")
            g_amt = g4.number_input("Charges (₹)", min_value=0.0, step=1.0)
            if st.form_submit_button("Save charges"):
                if not g_fy.strip():
                    st.warning("Enter the financial year, e.g. 2026-27.")
                else:
                    cash_set_charge(g_acct, g_fy.strip(), g_amt, 'DP_CHARGE' if g_head == "DP Charges" else 'CHARGE')
                    st.session_state["cash_msg"] = f"Saved {g_head} of {money(g_amt)} for {g_acct}, {g_fy.strip()}."
                    st.rerun()
        with st.form("cash_holdings_form"):
            st.markdown("**Exposed and Equity Value** — by default taken from this app's holdings. "
                        "If your broker accounts hold more than the Trade Book, type the totals here.")
            k1, k2 = st.columns(2)
            k_exp = k1.number_input("Exposed / Invested (₹, both accounts)", min_value=0.0, step=100.0,
                                    value=float(ovr[0]) if ovr else float(exposed))
            k_eq = k2.number_input("Equity Value on date (₹, both accounts)", min_value=0.0, step=100.0,
                                   value=float(ovr[1]) if ovr else float(equity_val))
            kb1, kb2 = st.columns(2)
            if kb1.form_submit_button("Use these figures"):
                cash_set_override(k_exp, k_eq)
                st.session_state["cash_msg"] = "Now using the figures you typed for Exposed and Equity Value."
                st.rerun()
            if kb2.form_submit_button("Use the app's holdings"):
                cash_set_override(None, None)
                st.session_state["cash_msg"] = "Now using the holdings from this app."
                st.rerun()
        with st.form("cash_balance_form"):
            st.markdown("**Cash balance** (type the figure from your broker)")
            h1, h2 = st.columns(2)
            h_acct = h1.selectbox("Account", ACCOUNTS, key="bal_acct")
            h_amt = h2.number_input("Cash balance (₹)", min_value=0.0, step=1.0)
            if st.form_submit_button("Save cash balance"):
                cash_set_balance(h_acct, h_amt)
                st.session_state["cash_msg"] = f"Cash balance for {h_acct} set to {money(h_amt)}."
                st.rerun()

    with st.expander("🗑️ Remove an entry"):
        if ent.empty:
            st.caption("No entries.")
        else:
            names = {int(r['id']): (f"{r['account']} · {'Deposit' if r['kind'] == 'DEPOSIT' else 'Transferred back' if r['kind'] == 'TRANSFER_BACK' else ('DP Charges ' if r['kind'] == 'DP_CHARGE' else 'Statutory & Broker Charges ') + str(r['label'])}"
                                    f"{'' if pd.isna(r['date']) else ' · ' + fmt_date(r['date'])} · {money(r['amount'])}")
                     for _, r in ent.iterrows()}
            pick = st.selectbox("Entry", list(names.keys()), format_func=lambda i: names[i], index=None,
                                placeholder="Pick the entry to remove")
            sure = st.checkbox("Yes, remove it")
            if st.button("🗑️ Remove entry", disabled=not (pick is not None and sure)):
                cash_delete_entry(pick)
                st.session_state["cash_msg"] = f"Removed: {names[pick]}"
                st.rerun()

elif page == "CMP Overrides":
    st.subheader("⚙️ Price Overrides")
    st.info("Add a row here if the automated price lags or fails for a symbol.")
    edited_override = st.data_editor(overrides_df, num_rows="dynamic", use_container_width=True)
    if st.button("Save Overrides"):
        save_data(trades_df, edited_override)
        st.success("Overrides updated!")
        st.rerun()
