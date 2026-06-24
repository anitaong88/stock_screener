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
    min_value=10, max_value=len(all_symbols), value=min(30, len(all_symbols)),
