import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile

render_page_content("📋", "Trading Comparables", "Peer benchmarking with percentile bands, live-formula Excel export.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL").strip().upper()
max_peers = st.slider("Max Peers", 3, 10, 5)
run = st.button("Run Comps", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Building comps for {ticker}..."):
            comps = deal_tool.build_comps_table(ticker, max_peers=max_peers)
            wb = deal_tool.export_comps_to_excel(comps, ticker)

        target = comps["target"]
        st.success(f"{target['company_name']} ({ticker})")

        col1, col2, col3 = st.columns(3)
        col1.metric("EV / Revenue", f"{target['ev_revenue']:.2f}x" if target.get("ev_revenue") else "N/A")
        col2.metric("EV / EBITDA", f"{target['ev_ebitda']:.2f}x" if target.get("ev_ebitda") else "N/A")
        col3.metric("P/E", f"{target['pe_ratio']:.2f}x" if target.get("pe_ratio") else "N/A")

        st.subheader("Peer Set")
        rows = [{"Ticker": p["ticker"], "Company": p["company_name"],
                 "EV/Revenue": f"{p['ev_revenue']:.2f}x" if p.get("ev_revenue") else "N/A",
                 "EV/EBITDA": f"{p['ev_ebitda']:.2f}x" if p.get("ev_ebitda") else "N/A",
                 "P/E": f"{p['pe_ratio']:.2f}x" if p.get("pe_ratio") else "N/A"} for p in comps["peers"]]
        if rows:
            st.dataframe(rows, width="stretch", hide_index=True)
        else:
            st.info("No peer data was available for this company.")

        if comps.get("implied_ev_from_revenue") or comps.get("implied_ev_from_ebitda"):
            st.subheader("Implied Valuation")
            c1, c2 = st.columns(2)
            if comps.get("implied_ev_from_revenue"):
                c1.metric("Implied EV (from Revenue multiple)", f"${comps['implied_ev_from_revenue']:,.0f}")
            if comps.get("implied_ev_from_ebitda"):
                c2.metric("Implied EV (from EBITDA multiple)", f"${comps['implied_ev_from_ebitda']:,.0f}")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{ticker}_Comps.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                st.download_button("Download Excel Model", f.read(), file_name=f"{ticker}_Comps.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Couldn't run comps: {e}")

page_footer()