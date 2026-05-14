import yfinance as yf
import streamlit as st
import pandas as pd

st.title("Anita's Stock Screener")
st.write("Filter stocks based on your investment criteria")

# User Input Criteria
st.sidebar.header("Filter Criteria")

pe_max = st.sidebar.number_input("Max PE Ratio", min_value=0.0, value=25.0)
roe_min = st.sidebar.number_input("Min ROE (%)", min_value=0.0, value=10.0)
volume_min = st.sidebar.number_input("Min Daily Volume", min_value=0, value=1000000)
industry = st.sidebar.text_input("Industry (e.g. Technology)", value="")

# Stock list to screen
st.subheader("Enter Stock Symbols to Screen")
stocks_input = st.text_input("Enter stock symbols separated by commas (e.g. AAPL, MSFT, TSLA)", value="AAPL, MSFT, TSLA, JPM, JNJ")

if st.button("🔍 Screen Stocks"):
    symbols = [s.strip().upper() for s in stocks_input.split(",")]
    results = []
    
    with st.spinner("Fetching stock data..."):
        for symbol in symbols:
            try:
                stock = yf.Ticker(symbol)
                info = stock.info
                
                pe = info.get("trailingPE", None)
                roe = info.get("returnOnEquity", None)
                vol = info.get("averageVolume", None)
                ind = info.get("industry", "")
                name = info.get("longName", symbol)
                price = info.get("currentPrice", None)

                if roe: roe = roe * 100

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

    if results:
        st.success(f"✅ Found {len(results)} matching stocks!")
        st.dataframe(pd.DataFrame(results))
    else:
        st.warning("No stocks matched your criteria. Try adjusting your filters!")