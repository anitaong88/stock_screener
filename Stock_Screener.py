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
FMP_V3 = "https://financialmodelingprep.com/api/v3"

st.set_page_config(page_title="Anita's Stock Screener", layout="wide")

st.markdown("### Anita's Stock Screener")
st.caption(
    "v2.0 — Jun 2026 — Fixed: switched all endpoints to FMP API v3 (stable endpoints deprecated) | "
    "Added: clearer skip reasons, combined API calls to stay within free plan limits | "
    "v1.9 — May 2026 — Added FMP API Data Checker"
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
        "ARCB","HTLD","MRTN","MATX","HUBG","JBHT","WERN","KNX",
        "AGIO","ACAD","FOLD","HRMY","PTGX","RCUS","SERA","XNCR","IGMS","KROS",
        "ADUS","AMED","FWRG","LHCG","MGLN","PDCO","PRSC","QTWO","RCKY","RELY"
    ]
}

# --- Sidebar Filters ---
st.sidebar.header("Filter Criteria")

market = st.sidebar.selectbox("Select Market Index", list(TICKERS.keys()))

pe_max     = st.sidebar.number_input("Max PE Ratio",      min_value=0.0, value=25.0)
roe_min    = st.sidebar.number_input("Min ROE (%)",       min_value=0.0, value=10.0)
volume_min = st.sidebar.number_input("Min Daily Volume",  min_value=0,   value=1000000)
industry   = st.sidebar.text_input("Industry (e.g. Technology)", value="")

all_symbols = TICKERS[market]
max_stocks  = st.sidebar.slider(
    "Max stocks to screen (speed vs coverage)",
    min_value=10, max_value=len(all_symbols), value=min(30, len(all_symbols)), step=10
)

st.sidebar.markdown("---")
st.sidebar.caption(
    "⚠️ **Free plan tip:** FMP free plan allows ~250 API calls/day. "
    "Screening uses 3 calls per stock. Keeping 'Max stocks' at 30 uses 90 calls — well within the limit."
)

# --- Helper: Build HTML table for download ---
def build_html(df):
    rows = ""
    for _, row in df.iterrows():
        cells = ""
        for col, v in zip(df.columns, row):
            if col == "Yahoo Finance":
                cells += f'<td><a href="{v}" target="_blank">View</a></td>'
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

# --- Helper: Build Word (.docx) document ---
def build_docx(df):
    doc = Document()

    title = doc.add_heading("Anita's Stock Screener Results", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    doc.add_paragraph("")

    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"

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

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)
            row_cells[i].paragraphs[0].runs[0].font.size = Pt(9)

    for col in table.columns:
        for cell in col.cells:
            cell.width = Inches(1.2)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Core data fetch — v2.0
#
# Uses 3 API calls per stock (down from 5):
#   1. /quote/{symbol}          → price, volume, PE, EPS
#   2. /profile/{symbol}        → company name, industry
#   3. /ratios-ttm/{symbol}     → ROE (returnOnEquityTTM) — single call,
#                                  replaces the old income-statement +
#                                  balance-sheet pair
#
# All three use the stable /api/v3 base that FMP has maintained since 2017.
# The /stable/* base was introduced later and has been subject to breaking
# changes; /api/v3 is the safest choice for free-plan users.
# ---------------------------------------------------------------------------
def get_stock_data(symbol):
    try:
        # 1. Quote
        quote_url = f"{FMP_V3}/quote/{symbol}?apikey={FMP_API_KEY}"
        quote_r   = requests.get(quote_url, timeout=10).json()

        if not quote_r or not isinstance(quote_r, list):
            return None, f"No quote data returned (got: {str(quote_r)[:120]})"

        quote = quote_r[0]

        # Check for API-level error embedded in the response
        if "Error Message" in quote:
            return None, f"API error: {quote['Error Message']}"

        price = quote.get("price")
        vol   = quote.get("volume") or quote.get("avgVolume")
        pe    = quote.get("pe")

        # Fallback: calculate PE from price / EPS if not in quote
        if pe is None:
            eps = quote.get("eps")
            if eps and eps != 0 and price:
                pe = round(price / eps, 2)

        # 2. Profile (company name + industry)
        profile_url = f"{FMP_V3}/profile/{symbol}?apikey={FMP_API_KEY}"
        profile_r   = requests.get(profile_url, timeout=10).json()
        time.sleep(0.3)  # gentle rate limiting between calls

        name = symbol
        ind  = ""
        if profile_r and isinstance(profile_r, list):
            name = profile_r[0].get("companyName", symbol)
            ind  = profile_r[0].get("industry", "")

        # 3. Ratios TTM — single call replaces income + balance sheet pair
        roe = None
        ratios_url = f"{FMP_V3}/ratios-ttm/{symbol}?apikey={FMP_API_KEY}"
        ratios_r   = requests.get(ratios_url, timeout=10).json()
        time.sleep(0.3)

        if ratios_r and isinstance(ratios_r, list) and ratios_r[0]:
            roe_raw = ratios_r[0].get("returnOnEquityTTM")
            if roe_raw is not None:
                roe = round(float(roe_raw) * 100, 2)   # FMP returns as decimal e.g. 0.35 → 35%

        return {
            "pe":    pe,
            "roe":   roe,
            "vol":   vol,
            "ind":   ind,
            "name":  name,
            "price": price,
        }, None

    except requests.exceptions.Timeout:
        return None, "Request timed out — FMP may be slow, try again"
    except requests.exceptions.ConnectionError:
        return None, "Connection error — check your internet connection"
    except Exception as e:
        return None, f"Unexpected error: {str(e)}"


# ---------------------------------------------------------------------------
# FMP API Data Checker (diagnostic tool — unchanged in purpose, updated URLs)
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("#### 🔬 FMP API Data Checker")
st.caption("Pick a ticker and see exactly what your FMP plan returns — no filters applied.")

col_a, col_b = st.columns([1, 3])
with col_a:
    check_symbol = st.text_input("Ticker", value="AAPL").upper().strip()

if st.button("🔍 Check API Data"):
    endpoints = {
        "Quote (/api/v3/quote)":               f"{FMP_V3}/quote/{check_symbol}?apikey={FMP_API_KEY}",
        "Profile (/api/v3/profile)":           f"{FMP_V3}/profile/{check_symbol}?apikey={FMP_API_KEY}",
        "Ratios TTM (/api/v3/ratios-ttm)":     f"{FMP_V3}/ratios-ttm/{check_symbol}?apikey={FMP_API_KEY}",
        "Key Metrics TTM (/api/v3/key-metrics-ttm)": f"{FMP_V3}/key-metrics-ttm/{check_symbol}?apikey={FMP_API_KEY}",
        "Income Statement (/api/v3)":           f"{FMP_V3}/income-statement/{check_symbol}?limit=1&apikey={FMP_API_KEY}",
        "Balance Sheet (/api/v3)":              f"{FMP_V3}/balance-sheet-statement/{check_symbol}?limit=1&apikey={FMP_API_KEY}",
    }

    for label, url in endpoints.items():
        with st.expander(f"📡 {label}", expanded=True):
            try:
                r = requests.get(url, timeout=10).json()
                if isinstance(r, list) and len(r) > 0:
                    row = r[0]
                    df_raw = pd.DataFrame(
                        [(k, v, "✅" if v is not None and v != "" and v != 0 else "❌")
                         for k, v in row.items()],
                        columns=["Field", "Value", "Available"]
                    )
                    st.dataframe(df_raw, hide_index=True, use_container_width=True)
                elif isinstance(r, dict) and r.get("Error Message"):
                    st.error(f"API error: {r['Error Message']}")
                elif isinstance(r, list) and len(r) == 0:
                    st.warning("Empty response — this endpoint may require a paid plan or the symbol is invalid.")
                else:
                    st.json(r)
            except Exception as e:
                st.error(f"Request failed: {e}")
        time.sleep(0.3)

# ---------------------------------------------------------------------------
# Main Screener
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("#### 🔍 Screen Stocks")
st.caption("Stocks are screened using FMP's /api/v3 endpoints (quote, profile, ratios-ttm).")

if st.button("🚀 Screen Stocks"):
    symbols = all_symbols[:max_stocks]
    api_calls = len(symbols) * 3
    st.info(
        f"Screening {len(symbols)} stocks from {market}… "
        f"(3 API calls per stock = ~{api_calls} calls total out of your 250/day free limit)"
    )

    results  = []
    skipped  = []
    skip_log = []   # detailed skip reasons shown at the end
    progress = st.progress(0)
    status   = st.empty()

    for i, symbol in enumerate(symbols):
        status.text(f"Checking {symbol} ({i+1}/{len(symbols)})…")

        data, error = get_stock_data(symbol)

        if not data:
            skipped.append(symbol)
            skip_log.append(f"**{symbol}** — {error}")
            progress.progress((i + 1) / len(symbols))
            time.sleep(0.3)
            continue

        pe    = data["pe"]
        roe   = data["roe"]
        vol   = data["vol"]
        ind   = data["ind"]
        name  = data["name"]
        price = data["price"]

        # Skip if we have absolutely no financial data to filter on
        if pe is None and roe is None:
            skipped.append(symbol)
            skip_log.append(f"**{symbol}** — No PE or ROE data available (may need paid plan)")
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
        else:
            # Log why this stock was filtered out (helpful for debugging)
            reasons = []
            if not pe_ok:
                reasons.append(f"PE={pe} (max {pe_max})" if pe is not None else "PE=N/A")
            if not roe_ok:
                reasons.append(f"ROE={roe}% (min {roe_min}%)" if roe is not None else "ROE=N/A")
            if not vol_ok:
                reasons.append(f"Vol={vol:,} (min {volume_min:,})" if vol is not None else "Vol=N/A")
            if not ind_ok:
                reasons.append(f"Industry='{ind}' (filter: '{industry}')")
            skip_log.append(f"**{symbol}** — filtered out: {', '.join(reasons)}")

        progress.progress((i + 1) / len(symbols))
        time.sleep(0.4)   # ~2.5 stocks/sec; well within FMP rate limits

    status.empty()
    progress.empty()

    # --- Results ---
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

    # --- Skip / filter summary ---
    if skip_log:
        with st.expander(f"ℹ️ Details — {len(skipped)} skipped + filtered stocks", expanded=False):
            st.caption("Skipped = no data returned. Filtered out = data returned but didn't meet your criteria.")
            for line in skip_log:
                st.markdown(line)
