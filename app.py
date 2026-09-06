import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
import os

st.set_page_config(page_title="Equity Portfolio Dashboard", layout="wide", initial_sidebar_state="expanded")

EXCEL_FILE = "Equity_Portfolio_Dashboard.xlsx"
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

# --- DATA LOADING AND SAVING ---
@st.cache_data(ttl=600)
def load_data():
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

def save_data(tb_df, co_df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        tb_df.to_excel(writer, sheet_name="Trade_Book", index=False)
        co_df.to_excel(writer, sheet_name="CMP_Override", index=False)
    st.cache_data.clear()

trades_df, overrides_df = load_data()

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

@st.cache_data(ttl=300)
def fetch_live_prices(tickers, overrides_df):
    price_map = {}
    manual_map = {}
    if not overrides_df.empty:
        manual_map = dict(zip(overrides_df['Stock'], overrides_df['Manual CMP']))

    for ticker in tickers:
        if ticker in manual_map and pd.notna(manual_map[ticker]):
            price_map[ticker] = float(manual_map[ticker])
            continue

        try:
            symbol = to_yf_symbol(ticker)
            stock_data = yf.Ticker(symbol)
            fast_info = stock_data.fast_info
            price = fast_info.last_price

            if price is None or pd.isna(price):
                hist = stock_data.history(period="1d")
                price = hist['Close'].iloc[-1] if not hist.empty else 0.0

            price_map[ticker] = float(price)
        except Exception:
            price_map[ticker] = 0.0

    return price_map

# --- FIFO CALCULATOR ENGINE ---
def compute_portfolio(df, overrides_df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df['Date of Trxn'] = pd.to_datetime(df['Date of Trxn'])
    df = df.sort_values(by='Date of Trxn').copy()

    unique_tickers = df['Stock'].unique().tolist()
    live_prices = fetch_live_prices(unique_tickers, overrides_df)

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
            'Total_Value': total_val,
            'Unrealized_PL': unrealized_pl,
            'Realized_PL': r_pl,
            'Total_Net_Gain': net_gain
        })

    portfolio = pd.DataFrame(rows)
    if not portfolio.empty:
        total_portfolio_val = portfolio['Total_Value'].sum()
        portfolio['Allocation_%'] = portfolio['Total_Value'] / total_portfolio_val if total_portfolio_val > 0 else 0.0

    return portfolio, df

portfolio_df, processed_trades = compute_portfolio(trades_df, overrides_df)

# --- NAVIGATION SIDEBAR ---
st.sidebar.title("📌 Menu")
page = st.sidebar.radio("Go to", ["Dashboard Overview", "Core Holdings", "Satellite Holdings", "Loss Booked", "Add Trade", "CMP Overrides"])

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

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Current Value", f"₹{total_val:,.2f}")
    col2.metric("Total Invested", f"₹{total_cost:,.2f}")
    col3.metric("Unrealized P&L", f"₹{unrealized_pl:,.2f}", delta=f"{(unrealized_pl/total_cost*100) if total_cost>0 else 0:.2f}%")
    col4.metric("Realized P&L", f"₹{realized_pl:,.2f}")

st.markdown("---")

# --- PAGES ---
FORMAT_DICT = {
    'Avg_Cost': '₹{:.2f}', 'Total_Cost': '₹{:.2f}', 'CMP': '₹{:.2f}',
    'Total_Value': '₹{:.2f}', 'Unrealized_PL': '₹{:.2f}',
    'Realized_PL': '₹{:.2f}', 'Total_Net_Gain': '₹{:.2f}', 'Allocation_%': '{:.1%}'
}

if page == "Dashboard Overview":
    st.subheader("📊 Complete Portfolio Overview")
    if not portfolio_df.empty:
        st.dataframe(portfolio_df.style.format(FORMAT_DICT), use_container_width=True)
    else:
        st.info("No trades yet. Add one from the 'Add Trade' page.")

elif page == "Core Holdings":
    st.subheader("🎯 Core Holdings")
    core = portfolio_df[(portfolio_df['Type'] == 'CORE') & (portfolio_df['Units'] > 0)]
    st.dataframe(core.style.format(FORMAT_DICT), use_container_width=True)

elif page == "Satellite Holdings":
    st.subheader("🚀 Satellite Holdings")
    sat = portfolio_df[(portfolio_df['Type'] == 'SATELLITE') & (portfolio_df['Units'] > 0)]
    st.dataframe(sat.style.format(FORMAT_DICT), use_container_width=True)

elif page == "Loss Booked":
    st.subheader("🔻 Booked Losses & Closed Positions")
    closed = portfolio_df[(portfolio_df['Units'] == 0) & (portfolio_df['Realized_PL'] < 0)]
    st.dataframe(closed.style.format(FORMAT_DICT), use_container_width=True)

elif page == "Add Trade":
    st.subheader("📝 Record New Trade")
    with st.form("new_trade_form"):
        c1, c2 = st.columns(2)
        with c1:
            p_type = st.selectbox("Type", ["CORE", "SATELLITE"])
            stock = st.text_input("Stock Symbol (e.g. BHEL, ACUTAAS)").upper().strip()
            date = st.date_input("Transaction Date", datetime.now())
        with c2:
            price = st.number_input("Price per Unit (₹)", min_value=0.01, step=0.01)
            units = st.number_input("Units (+ for Buy, - for Sell)", step=1)

        submitted = st.form_submit_button("Save Transaction")
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
            st.success(f"Added trade for {stock} successfully!")
            st.rerun()

elif page == "CMP Overrides":
    st.subheader("⚙️ Price Overrides")
    st.info("Add a row here if the automated price lags or fails for a symbol.")
    edited_override = st.data_editor(overrides_df, num_rows="dynamic", use_container_width=True)
    if st.button("Save Overrides"):
        save_data(trades_df, edited_override)
        st.success("Overrides updated!")
        st.rerun()
