import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
import os
import html as html_lib

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
def compute_portfolio(df, overrides_df):
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
        total_portfolio_val = portfolio['Total_Value'].sum()
        portfolio['Allocation_%'] = portfolio['Total_Value'] / total_portfolio_val if total_portfolio_val > 0 else 0.0

    return portfolio, df

portfolio_df, processed_trades = compute_portfolio(trades_df, overrides_df)

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

def build_tranche_table(lots, cmp_price):
    """Builds the tranche-by-tranche breakdown with weighted avg, growth %,
    P/L, and the stop-loss rule:
      Tranche 1        -> 10% below its own buy price
      Tranche 2        -> price that keeps TOTAL risk in rupees the same as Tranche 1's risk
      Tranche 3 onward -> breakeven (stop loss = cumulative weighted avg cost)
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

        if tranche_no == 1:
            sl_rule = "10% below buy"
            sl_price = price * 0.90
        elif tranche_no == 2:
            sl_rule = "Residual risk"
            sl_price = weighted_avg - (prev_risk_amount / cum_qty)
        else:
            sl_rule = "Breakeven"
            sl_price = weighted_avg

        risk_amount = (weighted_avg - sl_price) * cum_qty

        growth_from_prev = ((price - prev_buy_price) / prev_buy_price) if prev_buy_price else None
        growth_from_cmp = ((cmp_price - price) / price) if price else None
        tranche_pl = qty * (cmp_price - price)

        rows.append({
            'Tranche': tranche_no,
            'Buy Date': lot['date'].strftime('%Y-%m-%d'),
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
        })

        prev_buy_price = price
        prev_risk_amount = risk_amount

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

def render_compact_table(headers, rows, left_cols=(), auto_width=False, total_last=False, font_px=12):
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
    width = "width:auto;" if auto_width else "width:100%;"
    head = "".join(f"<th>{html_lib.escape(h)}</th>" for h in headers)
    body = ""
    for ri, r in enumerate(rows):
        tds = ""
        for ci, (text, color) in enumerate(r):
            cls = " class='l'" if ci in left_cols else ""
            style = f" style='color:{color}; font-weight:600;'" if color else ""
            tds += f"<td{cls}{style}>{html_lib.escape(str(text))}</td>"
        tr_cls = " class='tot'" if (total_last and ri == len(rows) - 1) else ""
        body += f"<tr{tr_cls}>{tds}</tr>"
    st.markdown(
        f"{css}<div class='cmp-wrap'><table class='cmp' style='{width}font-size:{font_px}px;'>"
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

# --- NAVIGATION SIDEBAR ---
st.sidebar.title("📌 Menu")
page = st.sidebar.radio("Go to", ["Dashboard Overview", "Core Holdings", "Satellite Holdings",
                                    "Profit/Loss", "Individual Stock", "Trade Book", "Add Trade", "CMP Overrides"],
                        key="nav_page")

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
    'Realized_PL': '₹{:.2f}', 'Total_Net_Gain': '₹{:.2f}', 'Allocation_%': '{:.1%}',
    'Yest_Close': '₹{:.2f}', 'Day_Change': '₹{:.2f}'
}

def render_totals_bar(df, color):
    st.markdown(
        f"""
        <div style='display:flex; justify-content:space-between; flex-wrap:wrap;
                    padding:14px 18px; margin-top:10px; background-color:#F5F5F5;
                    border-top:3px solid {color}; border-radius:4px;'>
            <div style='font-size:20px; font-weight:800; color:{color};'>
                Total Cost: ₹{df['Total_Cost'].sum():,.2f}
            </div>
            <div style='font-size:20px; font-weight:800; color:{color};'>
                Unrealized P&L: ₹{df['Unrealized_PL'].sum():,.2f}
            </div>
            <div style='font-size:20px; font-weight:800; color:{color};'>
                Realized P&L: ₹{df['Realized_PL'].sum():,.2f}
            </div>
            <div style='font-size:20px; font-weight:800; color:{color};'>
                Total Net Gain: ₹{df['Total_Net_Gain'].sum():,.2f}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def go_to_stock(label):
    """Button callback: jump to the Individual Stock page with this stock pre-selected."""
    st.session_state["nav_page"] = "Individual Stock"
    st.session_state["ind_stock_choice"] = label

def render_holdings_grid(df, ptype):
    """Holdings table whose Stock names are clickable links to the Individual Stock page."""
    heads = ['#', 'Stock', 'Units', 'Avg Cost', 'Total Cost', 'CMP', 'Total Value',
             'Unrealized P/L', 'Realized P/L', 'Net Gain', 'Alloc %']
    widths = [0.5, 1.8, 0.8, 1.2, 1.3, 1.1, 1.4, 1.4, 1.3, 1.3, 0.9]
    left = {'#', 'Stock'}
    hc = st.columns(widths)
    for c, h in zip(hc, heads):
        align = 'left' if h in left else 'right'
        c.markdown(f"<div style='font-weight:700; font-size:13px; text-align:{align};'>{h}</div>",
                   unsafe_allow_html=True)
    st.markdown("<hr style='margin:2px 0 4px 0;'>", unsafe_allow_html=True)

    def cell(col, text):
        col.markdown(f"<div style='font-size:13px; text-align:right;'>{text}</div>", unsafe_allow_html=True)

    for i, (_, r) in enumerate(df.iterrows(), start=1):
        cols = st.columns(widths, vertical_alignment="center")
        cols[0].markdown(f"<div style='font-size:13px;'>{i}</div>", unsafe_allow_html=True)
        cols[1].button(str(r['Stock']), key=f"hold_{ptype}_{i}_{r['Stock']}", type="tertiary",
                       on_click=go_to_stock, args=(f"{ptype} — {r['Stock']}",))
        cell(cols[2], f"{int(r['Units']):,}")
        cell(cols[3], f"₹{r['Avg_Cost']:,.2f}")
        cell(cols[4], f"₹{r['Total_Cost']:,.2f}")
        cell(cols[5], f"₹{r['CMP']:,.2f}")
        cell(cols[6], f"₹{r['Total_Value']:,.2f}")
        cell(cols[7], f"₹{r['Unrealized_PL']:,.2f}")
        cell(cols[8], f"₹{r['Realized_PL']:,.2f}")
        cell(cols[9], f"₹{r['Total_Net_Gain']:,.2f}")
        cell(cols[10], f"{r['Allocation_%']:.1%}")

UP, DOWN = "#1B8A3A", "#C62828"

def money(v, dec=2):
    if v is None or pd.isna(v):
        return "-"
    return f"{'-' if v < 0 else ''}₹{abs(v):,.{dec}f}"

def pl_color(v):
    if v is None or pd.isna(v) or v == 0:
        return None
    return UP if v > 0 else DOWN

def day_change_total(df):
    s = df['Day_Change']
    return s.sum() if s.notna().any() else float('nan')

def render_overview_open(title, df):
    st.markdown(f"#### {title} ({len(df)})")
    if df.empty:
        st.caption("No open positions.")
        return
    rows = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        # CMP is coloured by today's move: above yesterday's close = green, below = red
        cmp_color = None
        if pd.notna(r['Yest_Close']) and r['CMP'] > 0 and r['CMP'] != r['Yest_Close']:
            cmp_color = UP if r['CMP'] > r['Yest_Close'] else DOWN
        rows.append([
            (i, None), (r['Stock'], None), (f"{int(r['Units']):,}", None),
            (money(r['Avg_Cost']), None), (money(r['Total_Cost']), None),
            (money(r['CMP']), cmp_color), (money(r['Yest_Close']), None),
            (money(r['Day_Change']), pl_color(r['Day_Change'])),
            (money(r['Total_Value']), None), (money(r['Unrealized_PL']), None),
            (money(r['Realized_PL']), None), (money(r['Total_Net_Gain']), None),
            (f"{r['Allocation_%']:.1%}", None),
        ])
    tot = day_change_total(df)
    rows.append([("", None), ("Total", None)] + [("", None)] * 5 + [(money(tot), pl_color(tot))] + [("", None)] * 5)
    render_compact_table(
        ['#', 'Stock', 'Units', 'Avg_Cost', 'Total_Cost', 'CMP', 'Yest_Close', '1_Day_Change',
         'Total_Value', 'Unrealized_PL', 'Realized_PL', 'Total_Net_Gain', 'Allocation_%'],
        rows, left_cols=(1,), total_last=True, font_px=13
    )

def render_day_change_summary(core_df, sat_df):
    core_t, sat_t = day_change_total(core_df), day_change_total(sat_df)
    both = [x for x in (core_t, sat_t) if pd.notna(x)]
    grand = sum(both) if both else float('nan')
    render_compact_table(
        ['Core 1_Day_Change', 'Satellite 1_Day_Change', 'Total 1_Day_Change'],
        [[(money(core_t), pl_color(core_t)), (money(sat_t), pl_color(sat_t)), (money(grand), pl_color(grand))]],
        auto_width=True, font_px=14
    )

def render_overview_booked(title, df, trades):
    """P/L booked = every stock with at least one sale: fully squared off and partially sold."""
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
    st.markdown(f"#### {title} ({len(rows_data)})")
    if not rows_data:
        st.caption("No sales booked yet.")
        return
    rows_data.sort(key=lambda d: d['Last_Sale'], reverse=True)
    rows = []
    for i, d in enumerate(rows_data, start=1):
        c = pl_color(d['PL'])
        rows.append([
            (i, None), (d['Stock'], None), (d['Status'], None), (f"{d['Units_Left']:,}", None),
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
        rows, left_cols=(1, 2), total_last=True, font_px=13
    )

if page == "Dashboard Overview":
    st.subheader("📊 Dashboard Overview")
    if portfolio_df.empty:
        st.info("No trades yet. Add one from the 'Add Trade' page.")
    else:
        open_df = portfolio_df[portfolio_df['Units'] > 0]
        core_open = open_df[open_df['Type'] == 'CORE']
        sat_open = open_df[open_df['Type'] == 'SATELLITE']
        render_overview_open("🎯 Core Holdings", core_open)
        render_overview_open("🚀 Satellite Holdings", sat_open)
        render_day_change_summary(core_open, sat_open)
        render_overview_booked("✅ Core P/L Booked", portfolio_df[portfolio_df['Type'] == 'CORE'], processed_trades)
        render_overview_booked("✅ Satellite P/L Booked", portfolio_df[portfolio_df['Type'] == 'SATELLITE'], processed_trades)

elif page == "Core Holdings":
    st.subheader("🎯 Core Holdings")
    core = portfolio_df[(portfolio_df['Type'] == 'CORE') & (portfolio_df['Units'] > 0)].drop(columns=['Type']).reset_index(drop=True)
    core.index = core.index + 1
    st.caption("Click a stock name to open its Individual Stock analysis.")
    render_holdings_grid(core, 'CORE')
    render_totals_bar(core, '#1F77B4')

elif page == "Satellite Holdings":
    st.subheader("🚀 Satellite Holdings")
    sat = portfolio_df[(portfolio_df['Type'] == 'SATELLITE') & (portfolio_df['Units'] > 0)].drop(columns=['Type']).reset_index(drop=True)
    sat.index = sat.index + 1
    st.caption("Click a stock name to open its Individual Stock analysis.")
    render_holdings_grid(sat, 'SATELLITE')
    render_totals_bar(sat, '#D62728')

elif page == "Profit/Loss":
    st.subheader("🔻 Booked Losses & Closed Positions")
    closed = portfolio_df[(portfolio_df['Units'] == 0) & (portfolio_df['Realized_PL'] < 0)]
    st.dataframe(closed.drop(columns=['Yest_Close', 'Day_Change']).style.format(FORMAT_DICT),
                 use_container_width=True)

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

        lots = get_remaining_lots(processed_trades, sel_type, sel_stock)
        pos_row = all_pos[(all_pos['Type'] == sel_type) & (all_pos['Stock'] == sel_stock)].iloc[0]
        cmp_price = float(pos_row['CMP'])

        if not lots:
            st.info("No open position — this stock is fully squared off. Its realized results are below.")
        else:
            tranche_df = build_tranche_table(lots, cmp_price)

            total_qty = tranche_df['Cum. Qty'].iloc[-1]
            total_investment = tranche_df['Investment'].sum()
            weighted_avg = tranche_df['Weighted Avg Cost'].iloc[-1]
            current_sl_price = tranche_df['Stop-Loss Price'].iloc[-1]
            current_risk = tranche_df['Risk Amount (₹)'].iloc[-1]
            current_val = total_qty * cmp_price
            current_pl = current_val - total_investment
            current_pl_pct = (current_pl / total_investment) if total_investment > 0 else 0

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("CMP", f"₹{cmp_price:,.2f}")
            c2.metric("Weighted Avg Cost", f"₹{weighted_avg:,.2f}")
            c3.metric("Unrealized P&L", f"₹{current_pl:,.2f}", delta=f"{current_pl_pct*100:.2f}%")
            c4.metric("Current Stop-Loss", f"₹{current_sl_price:,.2f}",
                      delta=f"Risk ₹{current_risk:,.2f}" if current_risk > 0 else "Risk-free (breakeven)")

            st.markdown("#### Tranche Breakdown & Stop-Loss (open position)")
            t_rows = []
            for _, r in tranche_df.iterrows():
                gp, gc = r['Growth % vs Prev Tranche'], r['Growth % vs CMP']
                pl = r['Tranche P/L (₹)']
                t_rows.append([
                    (int(r['Tranche']), None), (r['Buy Date'], None),
                    (fmt_num(r['Buy Price']), None), (fmt_num(r['Qty'], 0), None),
                    (fmt_num(r['Investment'], 0), None), (fmt_num(r['Cum. Qty'], 0), None),
                    (fmt_num(r['Weighted Avg Cost']), None),
                    ("—" if pd.isna(gp) else f"{gp:.1%}", None),
                    ("—" if pd.isna(gc) else f"{gc:.1%}", GREEN if gc >= 0 else RED),
                    (r['Stop-Loss Rule'], None),
                    (fmt_num(r['Stop-Loss Price']), None),
                    (fmt_num(r['Risk Amount (₹)'], 0), None),
                    (fmt_num(pl, 0), GREEN if pl >= 0 else RED),
                ])
            render_compact_table(
                ['#', 'Date', 'Buy ₹', 'Qty', 'Invested ₹', 'Cum Qty', 'Avg Cost ₹',
                 'Δ vs Prev', 'Δ vs CMP', 'SL Rule', 'SL ₹', 'Risk ₹', 'P/L ₹'],
                t_rows, left_cols=(1, 9)
            )
            st.caption("SL rule: Tranche 1 = 10% below its buy price · Tranche 2 = same ₹ risk as "
                       "Tranche 1 · Tranche 3+ = breakeven (avg cost of all tranches).")

            st.markdown("#### Profit Target Ladder")
            pcts = [10, 15, 20, 25]
            render_compact_table(
                ['Profit target'] + [f"+{p}%" for p in pcts],
                [[("Target price ₹", None)] + [(fmt_num(weighted_avg * (1 + p / 100)), None) for p in pcts]],
                left_cols=(0,), auto_width=True
            )

            st.markdown("#### Risk : Reward")
            reward_15pct = total_investment * 0.15
            if current_risk > 0:
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

elif page == "Trade Book":
    st.subheader("📒 Trade Book")
    if trades_df.empty:
        st.info("No trades recorded yet. Add one from the 'Add Trade' page.")
    else:
        tb_view = trades_df.copy()
        tb_view['Date of Trxn'] = pd.to_datetime(tb_view['Date of Trxn'])
        tb_view['Amount'] = tb_view['Price'].astype(float) * tb_view['Unit'].astype(float)
        tb_view['Action'] = tb_view['Unit'].apply(lambda u: 'BUY' if float(u) > 0 else 'SELL')

        f1, f2, f3 = st.columns(3)
        type_sel = f1.selectbox("Type", ["All", "CORE", "SATELLITE"])
        stock_sel = f2.selectbox("Stock", ["All"] + sorted(tb_view['Stock'].astype(str).unique().tolist()))
        action_sel = f3.selectbox("Action", ["All", "BUY", "SELL"])

        if stock_sel != "All":
            pos = portfolio_df[portfolio_df['Stock'] == str(stock_sel).strip().upper()]
            if type_sel != "All":
                pos = pos[pos['Type'] == type_sel]
            held = pos[pos['Units'] > 0]
            s_val = held['Total_Value'].sum()
            s_cost = held['Total_Cost'].sum()
            s_unreal = held['Unrealized_PL'].sum()
            s_real = pos['Realized_PL'].sum()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Current Value", f"₹{s_val:,.2f}")
            m2.metric("Total Invested", f"₹{s_cost:,.2f}")
            m3.metric("Unrealized P/L", f"₹{s_unreal:,.2f}",
                      delta=f"{(s_unreal/s_cost*100) if s_cost > 0 else 0:.2f}%")
            m4.metric("Realized P/L", f"₹{s_real:,.2f}")

        if type_sel != "All":
            tb_view = tb_view[tb_view['TYPE'] == type_sel]
        if stock_sel != "All":
            tb_view = tb_view[tb_view['Stock'].astype(str) == stock_sel]
        if action_sel != "All":
            tb_view = tb_view[tb_view['Action'] == action_sel]

        tb_view = tb_view.sort_values('Date of Trxn', ascending=False).reset_index(drop=True)
        tb_view.index = tb_view.index + 1
        tb_view = tb_view[['TYPE', 'Stock', 'Date of Trxn', 'Action', 'Price', 'Unit', 'Amount']]

        st.dataframe(
            tb_view.style.format({
                'Date of Trxn': lambda d: d.strftime('%d-%b-%Y'),
                'Price': '₹{:,.2f}', 'Unit': '{:,.0f}', 'Amount': '₹{:,.2f}'
            }),
            use_container_width=True
        )

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
            tb_view.to_csv(index=False).encode('utf-8'),
            file_name="trade_book.csv",
            mime="text/csv"
        )

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
