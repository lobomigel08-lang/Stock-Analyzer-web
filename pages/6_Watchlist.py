import streamlit as st
from style import render_page_content, page_footer
import deal_tool

render_page_content("👀", "Watchlist", "")

col1, col2 = st.columns([3, 1])
with col1:
    new_ticker = st.text_input("Add a ticker", placeholder="e.g. ZS").strip().upper()
with col2:
    st.write("")
    st.write("")
    if st.button("Add", disabled=not new_ticker):
        try:
            deal_tool.add_to_watchlist(new_ticker)
            st.success(f"Added {new_ticker}")
            st.rerun()
        except ValueError as e:
            st.error(str(e))

st.divider()

watchlist = deal_tool.load_watchlist()
if not watchlist:
    st.info("Your watchlist is empty. Add a ticker above to get started.")
else:
    if st.button("Refresh Prices"):
        st.rerun()

    for ticker_symbol in watchlist:
        try:
            company = deal_tool.resolve_company(ticker_symbol)
            info = company["info"]
            price_divisor = company["price_divisor"]
            sym = deal_tool.get_currency_symbol(company["quote_currency"])
            price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / price_divisor
            prev_close = (info.get("regularMarketPreviousClose") or price) / price_divisor
            change_pct = ((price - prev_close) / prev_close) if prev_close else 0

            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            c1.write(f"**{ticker_symbol}**")
            c2.write(company["company_name"])
            c3.write(f"{sym}{price:,.2f}  ({change_pct:+.1%})")
            if c4.button("Remove", key=f"remove_{ticker_symbol}"):
                deal_tool.remove_from_watchlist(ticker_symbol)
                st.rerun()
        except ValueError as e:
            c1, c2 = st.columns([2, 6])
            c1.write(f"**{ticker_symbol}**")
            c2.error(str(e))

page_footer()