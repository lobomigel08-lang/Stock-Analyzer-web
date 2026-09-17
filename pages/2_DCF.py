import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import io
import os
import tempfile

render_page_content("📈", "DCF Valuation", "Discounted cash flow model with WACC build-up, scenario toggle, and a live Excel export.")

ticker = st.text_input("Ticker", placeholder="e.g. AAPL, ZS, BHP.AX").strip().upper()
run = st.button("Run DCF", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Running DCF for {ticker}..."):
            result = deal_tool.run_dcf(ticker)
            scenarios = deal_tool.run_scenarios(ticker)
            wb = deal_tool.export_dcf_to_excel(result)
            deal_tool.add_scenarios_tab(wb, scenarios)

        sym = result["currency_symbol"]
        st.success(f"{result['company_name']} ({ticker})")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Current Price", f"{sym}{result['current_price']:,.2f}")
        col2.metric("DCF Implied Price", f"{sym}{result['implied_share_price']:,.2f}")
        col3.metric("Upside / (Downside)", f"{result['upside']:+.1%}")
        col4.metric("WACC", f"{result['assumptions_used']['wacc']:.2%}")

        if result["used_fallback_growth"] or result["used_fallback_margin"]:
            st.warning("Limited historical data — some assumptions used generic fallback values.")
        if result["is_negative_margin"]:
            st.warning("This company currently has a negative FCF margin — treat this DCF as one data point, not a verdict.")

        st.subheader("Bear / Base / Bull")
        scen_cols = st.columns(3)
        for col, name in zip(scen_cols, ["Bear", "Base", "Bull"]):
            with col:
                st.metric(name, f"{sym}{scenarios[name]['implied_share_price']:,.2f}", f"{scenarios[name]['upside']:+.1%}")

        st.subheader("Projected Revenue & Free Cash Flow")
        chart_data = {"Revenue": result["projected_revenue"], "FCF": result["projected_fcf"]}
        st.line_chart(chart_data)

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{ticker}_DCF.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                st.download_button("Download Excel Model", f.read(), file_name=f"{ticker}_DCF.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Couldn't run the DCF: {e}")

page_footer()