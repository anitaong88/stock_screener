import yfinance as yf
import streamlit as st
import pandas as pd

st.title("Anita's Stock Screener")
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
        "ACLS","ADUS","AFCG","AGIO","ALGT","ALKS","AMCX","AMED","AMPH","AMSF",
        "ANET","ANF","AORT","AROC","ASTE","ATNI","AVAV","AVNS","AXNX","AZTA",
        "BCPC","BDSI","BFST","BGFV","BHB","BHLB","BKE","BLMN","BMRC","BOOT",
        "BPOP","BRSP","BSVN","BV","BYFC","CAKE","CALM","CARG","CARS","CASA",
        "CASH","CBRL","CCOI","CDRE","CENT","CEVA","CHEF","CHUY","CINF","CIVB",
        "CLAR","CLDT","CLFD","CLOV","CMCO","CNMD","CNOB","CNXN","CODI","COHU",
        "COOP","CORE","COUR","CRVL","CRVO","CSBR","CSGS","CSWI","CTBI","CTRE",
        "CTRL","CUTR","CVBF","CVCO","CVLG","CVLT","CWST","DAKT","DCOM","DFIN"
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

# --- Main Screen Button ---
if st.button("🔍 Screen Stocks"):
    symbols = all_symbols[:max_stocks]
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