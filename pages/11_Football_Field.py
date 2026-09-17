import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile

render_page_content("🏈", "Football Field Valuation", "DCF, trading comps, and real analyst price targets — combined into one valuation range summary.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL").strip().upper()
run = st.button("Build Football Field", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Building football field for {ticker}..."):
            field = deal_tool.build_football_field(ticker)
            wb = deal_tool.export_football_field_to_excel(field)

        sym = field["currency_symbol"]
        st.success(f"Football Field — {ticker}")
        st.metric("Current Price", f"{sym}{field['current_price']:,.2f}")

        st.subheader("Valuation Range by Method")
        for m in field["methods"]:
            c1, c2 = st.columns([1, 3])
            c1.write(f"**{m['method']}**")
            c2.write(f"{sym}{m['low']:,.2f} — {sym}{m['high']:,.2f}  (mid {sym}{m['mid']:,.2f})")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{ticker}_FootballField.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                st.download_button("Download Excel Model", f.read(), file_name=f"{ticker}_FootballField.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Couldn't build the football field: {e}")

page_footer()