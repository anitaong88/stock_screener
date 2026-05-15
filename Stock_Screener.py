import yfinance as yf
import streamlit as st
import pandas as pd

st.title("Anita's Stock Screener")
st.write("Filter stocks based on your investment criteria")

# --- Market Index Ticker Lists from Wikipedia ---
@st.cache_data(ttl=86400)  # Cache for 24 hours
def get_tickers(market):
    try:
        if market == "S&P 500":
            url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
            table = pd.read_html(url)[0]
            return table["Symbol"].str.replace(".", "-", regex=False).tolist()

        elif market == "Nasdaq 100":
            url = "https://en.wikipedia.org/wiki/Nasdaq-100"
            tables = pd.read_html(url)
            for t in tables:
                if "Ticker" in t.columns:
                    return t["Ticker"].tolist()

        elif market == "Dow Jones (DJIA)":
            url = "https://en.wikipedia.org/wiki/Dow_Jones_Industrial_Average"
            tables = pd.read_html(url)
            for t in tables:
                if "Symbol" in t.columns:
                    return t["Symbol"].tolist()

        elif market == "Russell 2000":
            # Russell 2000 has 2000 stocks — we use a curated sample
            url = "https://en.wikipedia.org/wiki/Russell_2000_Index"
            st.info("ℹ️ Russell 2000 has 2,000 stocks. Screening may take several minutes.")
            tables = pd.read_html(url)
            for t in tables:
                for col in t.columns:
                    if "ticker" in col.lower() or "symbol" in col.lower():
                        return t[col].tolist()
    except Exception as e:
        st.error(f"Could not load tickers for {market}: {e}")
    return []

# --- Sidebar Filters ---
st.sidebar.header("Filter Criteria")

market = st.sidebar.selectbox(
    "Select Market Index",
    ["S&P 500", "Nasdaq 100", "Dow Jones (DJIA)", "Russell 2000"]
)

pe_max = st.sidebar.number_input("Max PE Ratio", min_value=0.0, value=25.0)
roe_min = st.sidebar.number_input("Min ROE (%)", min_value=0.0, value=10.0)
volume_min = st.sidebar.number_input("Min Daily Volume", min_value=0, value=1000000)
industry = st.sidebar.text_input("Industry (e.g. Technology)", value="")

max_stocks = st.sidebar.slider(
    "Max stocks to screen (speed vs coverage)",
    min_value=10, max_value=500, value=100, step=10
)

# --- Main Screen Button ---
if st.button("🔍 Screen Stocks"):
    symbols = get_tickers(market)

    if not symbols:
        st.error("Could not retrieve stock list. Please try another market.")
    else:
        symbols = symbols[:max_stocks]
        st.info(f"Screening {len(symbols)} stocks from {market}...")

        results = []
        progress = st.progress(0)
        status = st.empty()

        for i, symbol in enumerate(symbols):
            status.text(f"Checking {symbol} ({i+1}/{len(symbols)})...")
            try:
                stock = yf.Ticker(symbol)
                info = stock.info

                pe = info.get("trailingPE", None)
                roe = info.get("returnOnEquity", None)
                vol = info.get("averageVolume", None)
                ind = info.get("industry", "")
                name = info.get("longName", symbol)
                price = info.get("currentPrice", None)

                if roe:
                    roe = roe * 100

                pe_ok = pe and pe <= pe_max
                roe_ok = roe and roe >= roe_min
                vol_ok = vol and vol >= volume_min
                ind_ok = industry.lower() in ind.lower() if industry else True

                if pe_ok and roe_ok and vol_ok and ind_ok:
                    results.append({
                        "Symbol": symbol,
                        "Company": name,
                        "Price": f"${price:.2f}" if price else "N/A",
                        "PE Ratio": f"{pe:.2f}" if pe else "N/A",
                        "ROE (%)": f"{roe:.2f}" if roe else "N/A",
                        "Avg Volume": f"{vol:,}" if vol else "N/A",
                        "Industry": ind
                    })
            except:
                pass

            progress.progress((i + 1) / len(symbols))

        status.empty()
        progress.empty()

        if results:
            st.success(f"✅ Found {len(results)} matching stocks from {market}!")
            df = pd.DataFrame(results)
            st.dataframe(df)
            csv = df.to_csv(index=False)
            st.download_button(
                "📥 Download Results as CSV",
                data=csv,
                file_name="screener_results.csv",
                mime="text/csv"
            )
        else:
            st.warning("No stocks matched your criteria. Try relaxing your filters!")