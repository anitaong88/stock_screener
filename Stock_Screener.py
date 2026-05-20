import requests
import streamlit as st
import pandas as pd
import time
import urllib.parse
import io
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- FMP API Key (stored securely in Streamlit secrets) ---
FMP_API_KEY = st.secrets["FMP_API_KEY"]
FMP_BASE = "https://financialmodelingprep.com/stable"

st.set_page_config(page_title="Anita's Stock Screener", layout="wide")

st.markdown("### Anita's Stock Screener")
st.caption(
    "v1.9 — May 20, 2026 — Added: FMP API Data Checker shows all raw fields per endpoint for your plan tier | "
    "v1.8 — May 20, 2026 — Fixed: PE/ROE calculated from free endpoints (PE=price/EPS, ROE=netIncome/equity) | "
    "v1.7 — May 20, 2026 — Fixed: switched to free-plan FMP endpoints (key-metrics-ttm, volume not avgVolume) | "
    "v1.6 — May 20, 2026 — Fixed: falsy PE/ROE checks replaced with explicit is None checks | "
    "v1.5 — May 20, 2026 — Initial release: per-stock FMP API, rate limiting, free plan compatible"
)
st.write("Filter stocks based on your investment criteria")

# --- Hardcoded Ticker Lists ---
TICKERS = {
    "S&P 500": [
        "AAPL","MSFT","NVDA","AMZN","META","GOOGL","GOOG","BRK-B","LLY","AVGO",
        "JPM","TSLA","UNH","V","XOM","MA","JNJ","PG","COST","HD","MRK","ABBV",
        "CVX","BAC","CRM","NFLX","AMD","PEP","KO","TMO","ACN","MCD","CSCO","WMT",
        "ABT","DHR","ADBE","TXN","LIN","PM","NEE","ORCL","RTX","QCOM","MS","HON",
        "UPS","AMGN","IBM","CAT","GS","BA","SBUX","GE","INTU","SPGI","BLK","AXP",
        "ELV","MDT","DE","GILD","ADI","VRTX","ISRG","MMC","SYK","REGN","ZTS","PLD",
        "CI","CB","ADP","SCHW","C","MO","SO","DUK","TGT","CME","EOG","SLB","BDX",
        "ITW","BSX","NOC","WM","AON","PNC","USB","MMM","FDX","EMR","APD","CL","NSC"
    ],
    "Nasdaq 100": [
        "AAPL","MSFT","NVDA","AMZN","META","GOOGL","GOOG","TSLA","AVGO","ASML",
        "COST","NFLX","AMD","QCOM","ADBE","INTU","TXN","AMGN","PEP","CSCO",
        "ISRG","AMAT","HON","VRTX","REGN","MU","LRCX","ADI","KLAC","PANW",
        "MELI","SNPS","CDNS","CRWD","CEG","FTNT","ORLY","MNST","CTAS","PAYX",
        "MRVL","KDP","AEP","DXCM","ODFL","ROST","FAST","CTSH","TEAM","IDXX",
        "BIIB","ILMN","NXPI","WDAY","PCAR","VRSK","DDOG","ZS","ANSS","ALGN",
        "WBD","FANG","GEHC","GFS","ON","EXC","XEL","SIRI","DLTR","ENPH","LCID"
    ],
    "Dow Jones (DJIA)": [
        "AAPL","AMGN","AXP","BA","CAT","CRM","CSCO","CVX","DIS","DOW",
        "GS","HD","HON","IBM","JNJ","JPM","KO","MCD","MMM","MRK",
        "MSFT","NKE","PG","TRV","UNH","V","VZ","WBA","WMT","RTX"
    ],
    "Russell 2000 (Sample)": [
        "SAIA","WING","BOOT","FORM","MGNI","PRGS","ITRI","ENVA","CSWI","BCPC",
        "MGEE","LANC","NWBI","SFBS","TOWN","WAFD","CVBF","NBTB","FULT","WSFS",
        "CAKE","TXRH","BJRI","RRGB","DENN","CHUY","FRSH","HURN","AMSF","PRIM",
        "ARCB","SAIA","HTLD","ODFL","MRTN","MATX","HUBG","JBHT","WERN","KNX",
        "AGIO","ACAD","FOLD","HRMY","PTGX","RCUS","SERA","XNCR","IGMS","KROS",
        "ADUS","AMED","FWRG","LHCG","MGLN","PDCO","PRSC","QTWO","RCKY","RELY"
    ]
}

# --- Sidebar Filters ---
st.sidebar.header("Filter Criteria")

market = st.sidebar.selectbox("Select Market Index", list(TICKERS.keys()))

pe_max = st.sidebar.number_input("Max PE Ratio", min_value=0.0, value=25.0)
roe_min = st.sidebar.number_input("Min ROE (%)", min_value=0.0, value=10.0)
volume_min = st.sidebar.number_input("Min Daily Volume", min_value=0, value=1000000)
industry = st.sidebar.text_input("Industry (e.g. Technology)", value="")

all_symbols = TICKERS[market]
max_stocks = st.sidebar.slider(
    "Max stocks to screen (speed vs coverage)",
    min_value=10, max_value=len(all_symbols), value=min(50, len(all_symbols)), step=10
)

# --- Helper: Build HTML table for download ---
def build_html(df):
    rows = ""
    for _, row in df.iterrows():
        cells = ""
        for col, v in zip(df.columns, row):
            if col == "Yahoo Finance":
                cells += f'<td><a href="{v}" target="_blank">{v}</a></td>'
            else:
                cells += f"<td>{v}</td>"
        rows += f"<tr>{cells}</tr>\n"
    headers = "".join(f"<th>{c}</th>" for c in df.columns)
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Stock Screener Results</title>
  <style>
    body {{ font-family: Arial, sans-serif; padding: 20px; }}
    h2 {{ color: #333; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th {{ background-color: #4CAF50; color: white; padding: 8px 12px; text-align: left; }}
    td {{ border: 1px solid #ddd; padding: 8px 12px; }}
    tr:nth-child(even) {{ background-color: #f9f9f9; }}
    tr:hover {{ background-color: #f1f1f1; }}
  </style>
</head>
<body>
  <h2>Anita's Stock Screener Results</h2>
  <table>
    <thead><tr>{headers}</tr></thead>
    <tbody>{rows}</tbody>
  </table>
</body>
</html>"""

# --- Helper: Build Google Sheets import URL ---
def build_gsheets_url(csv_data):
    encoded = urllib.parse.quote(csv_data)
    return "https://sheets.new"

# --- Helper: Build Word (.docx) document ---
def build_docx(df):
    doc = Document()

    # Title
    title = doc.add_heading("Anita's Stock Screener Results", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    doc.add_paragraph("")  # spacer

    # Table: header + data rows
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"

    # Header row — green background
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr_cells[i].text = col
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tc = hdr_cells[i]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "4CAF50")
        tcPr.append(shd)

    # Data rows
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)
            row_cells[i].paragraphs[0].runs[0].font.size = Pt(9)

    # Set column widths
    for col in table.columns:
        for cell in col.cells:
            cell.width = Inches(1.2)

    # Save to bytes buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# --- Helper: Fetch stock data per symbol (free plan compatible) ---
# PE = price / EPS (both free). ROE = net income / shareholders equity (both free via income/balance sheet).
def get_stock_data(symbol):
    try:
        # --- Quote: price, volume ---
        quote_url = f"{FMP_BASE}/quote?symbol={symbol}&apikey={FMP_API_KEY}"
        quote_r = requests.get(quote_url, timeout=10).json()

        # --- Profile: companyName, industry ---
        profile_url = f"{FMP_BASE}/profile?symbol={symbol}&apikey={FMP_API_KEY}"
        profile_r = requests.get(profile_url, timeout=10).json()

        if not quote_r or not isinstance(quote_r, list):
            return None, f"No quote data: {str(quote_r)[:200]}"
        if not profile_r or not isinstance(profile_r, list):
            return None, f"No profile data: {str(profile_r)[:200]}"

        quote   = quote_r[0]
        profile = profile_r[0]

        price = quote.get("price")
        vol   = quote.get("volume") or quote.get("avgVolume")

        # --- PE: calculate from price / EPS (free tier) ---
        pe = quote.get("pe")  # try quote first
        if pe is None:
            eps = quote.get("eps")
            if eps and eps != 0 and price:
                pe = round(price / eps, 2)

        # --- ROE: calculate from income statement + balance sheet (free tier) ---
        roe = None
        try:
            inc_url = f"https://financialmodelingprep.com/api/v3/income-statement/{symbol}?limit=1&apikey={FMP_API_KEY}"
            inc_r = requests.get(inc_url, timeout=10).json()
            bal_url = f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{symbol}?limit=1&apikey={FMP_API_KEY}"
            bal_r = requests.get(bal_url, timeout=10).json()
            if inc_r and isinstance(inc_r, list) and bal_r and isinstance(bal_r, list):
                net_income = inc_r[0].get("netIncome")
                equity = bal_r[0].get("totalStockholdersEquity")
                if net_income and equity and equity != 0:
                    roe = round((net_income / equity) * 100, 2)
        except Exception:
            pass

        return {
            "pe":    pe,
            "roe":   roe,
            "vol":   vol,
            "ind":   profile.get("industry", ""),
            "name":  profile.get("companyName", symbol),
            "price": price,
            "_quote_keys": list(quote.keys()),
            "_inc_sample": inc_r[0] if 'inc_r' in dir() and inc_r and isinstance(inc_r, list) else {},
        }, None
    except Exception as e:
        return None, str(e)

# --- API Data Checker ---
st.markdown("---")
st.markdown("#### 🔬 FMP API Data Checker")
st.caption("Pick a ticker and see exactly what your FMP plan returns — no filters applied.")

col_a, col_b = st.columns([1, 3])
with col_a:
    check_symbol = st.text_input("Ticker", value="AAPL").upper().strip()
with col_b:
    st.write("")  # spacer

if st.button("🔍 Check API Data"):
    endpoints = {
        "Quote (/stable/quote)": f"{FMP_BASE}/quote?symbol={check_symbol}&apikey={FMP_API_KEY}",
        "Profile (/stable/profile)": f"{FMP_BASE}/profile?symbol={check_symbol}&apikey={FMP_API_KEY}",
        "Income Statement (v3)": f"https://financialmodelingprep.com/api/v3/income-statement/{check_symbol}?limit=1&apikey={FMP_API_KEY}",
        "Balance Sheet (v3)": f"https://financialmodelingprep.com/api/v3/balance-sheet-statement/{check_symbol}?limit=1&apikey={FMP_API_KEY}",
        "Key Metrics TTM (v3)": f"https://financialmodelingprep.com/api/v3/key-metrics-ttm/{check_symbol}?apikey={FMP_API_KEY}",
        "Ratios TTM (v3)": f"https://financialmodelingprep.com/api/v3/ratios-ttm/{check_symbol}?apikey={FMP_API_KEY}",
    }

    for label, url in endpoints.items():
        with st.expander(f"📡 {label}", expanded=True):
            try:
                r = requests.get(url, timeout=10).json()
                if isinstance(r, list) and len(r) > 0:
                    row = r[0]
                    # Show as a clean table: field | value, highlight non-None values
                    rows = [(k, v) for k, v in row.items()]
                    df_raw = pd.DataFrame(rows, columns=["Field", "Value"])
                    df_raw["Available"] = df_raw["Value"].apply(lambda v: "✅" if v is not None and v != "" and v != 0 else "❌")
                    st.dataframe(df_raw, hide_index=True, use_container_width=True)
                elif isinstance(r, dict) and r.get("Error Message"):
                    st.error(f"API error: {r['Error Message']}")
                elif isinstance(r, dict):
                    st.json(r)
                else:
                    st.warning(f"Empty or unexpected response: {str(r)[:300]}")
            except Exception as e:
                st.error(f"Request failed: {e}")
        time.sleep(0.3)

st.markdown("---")
st.markdown("#### 🔍 Screen Stocks")
st.caption("Once you've confirmed which fields are available above, use the screener below.")

# --- Main Screen Button ---
if st.button("🔍 Screen Stocks"):
    symbols = all_symbols[:max_stocks]
    st.info(f"Screening {len(symbols)} stocks from {market}... (5 API calls per stock: quote, profile, income, balance sheet, rate limited)")

    results = []
    skipped = []
    progress = st.progress(0)
    status = st.empty()

    for i, symbol in enumerate(symbols):
        status.text(f"Checking {symbol} ({i+1}/{len(symbols)})...")

        data, error = get_stock_data(symbol)

        if not data:
            if i == 0:
                st.error(f"Debug — First stock ({symbol}) error: {error}")
            skipped.append(symbol)
            progress.progress((i + 1) / len(symbols))
            time.sleep(0.5)
            continue

        try:
            pe    = data["pe"]
            roe   = data["roe"]
            vol   = data["vol"]
            ind   = data["ind"]
            name  = data["name"]
            price = data["price"]

            if pe is None and roe is None:
                skipped.append(symbol)
                progress.progress((i + 1) / len(symbols))
                time.sleep(0.3)
                continue

            pe_ok  = pe  is not None and pe  <= pe_max
            roe_ok = roe is not None and roe >= roe_min
            vol_ok = vol is not None and vol >= volume_min
            ind_ok = industry.lower() in ind.lower() if industry else True

            if pe_ok and roe_ok and vol_ok and ind_ok:
                results.append({
                    "Symbol":     symbol,
                    "Company":    name,
                    "Price":      f"${price:.2f}" if price else "N/A",
                    "PE Ratio":   f"{pe:.2f}"     if pe    else "N/A",
                    "ROE (%)":    f"{roe:.2f}"    if roe   else "N/A",
                    "Avg Volume": f"{vol:,}"       if vol   else "N/A",
                    "Industry":   ind,
                })

        except Exception:
            skipped.append(symbol)

        progress.progress((i + 1) / len(symbols))
        time.sleep(0.5)

    status.empty()
    progress.empty()

    if results:
        st.success(f"✅ Found {len(results)} matching stocks from {market}!")
        df = pd.DataFrame(results)

        df["Yahoo Finance"] = df["Symbol"].apply(
            lambda s: f"https://finance.yahoo.com/quote/{s}"
        )

        st.dataframe(
            df,
            column_config={
                "Yahoo Finance": st.column_config.LinkColumn(
                    "🔗 More Info",
                    help="Click to open company details on Yahoo Finance",
                    display_text="View",
                )
            },
            hide_index=True,
        )

        st.markdown("### 📥 Download Results")
        st.caption("Choose the format that works best for you:")

        col1, col2, col3, col4 = st.columns(4)

        csv = df.to_csv(index=False)
        with col1:
            st.download_button(
                label="📄 CSV",
                data=csv,
                file_name="screener_results.csv",
                mime="text/csv",
                help="Opens in Excel, Google Sheets, Apple Numbers, or any text editor"
            )
            st.caption("Works everywhere — Excel, Google Sheets, Numbers, Notepad")

        html = build_html(df)
        with col2:
            st.download_button(
                label="🌐 HTML Table",
                data=html,
                file_name="screener_results.html",
                mime="text/html",
                help="Opens as a formatted table in any web browser — links are clickable!"
            )
            st.caption("Clickable links — opens in any web browser")

        docx_bytes = build_docx(df)
        with col3:
            st.download_button(
                label="📝 Word Doc",
                data=docx_bytes,
                file_name="screener_results.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help="Opens in Microsoft Word, Google Docs, or LibreOffice"
            )
            st.caption("Opens in Word, Google Docs, or LibreOffice")

        with col4:
            st.markdown(
                """<a href="https://sheets.new" target="_blank">
                <button style="
                    background-color: #0F9D58;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 14px;
                    width: 100%;
                ">📊 Google Sheets</button></a>""",
                unsafe_allow_html=True
            )
            st.caption("Opens a new Google Sheet — then File → Import the CSV above")

    else:
        st.warning("No stocks matched your criteria. Try relaxing your filters!")

    if skipped:
        st.info(f"ℹ️ {len(skipped)} stocks skipped (no data available): {', '.join(skipped)}")
