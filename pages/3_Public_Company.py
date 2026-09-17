import streamlit as st
import deal_tool
import os
import tempfile

st.set_page_config(page_title="Public Company Valuation | Deal Tool", page_icon="🏢", layout="wide")
st.title("🏢 Public Company Valuation")
st.caption("Trading performance, ownership, multiples, analyst consensus, DCF, and comps — combined into one package.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL, BHP.AX").strip().upper()
run = st.button("Run Analysis", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Assembling valuation package for {ticker}..."):
            result = deal_tool.run_public_company_valuation(ticker)
            wb = deal_tool.export_public_company_valuation_to_excel(result)

        sym = result["currency_symbol"]
        st.success(f"{result['company_name']} ({ticker}) — {result['sector']} / {result['industry']}")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Current Price", f"{sym}{result['current_price']:,.2f}")
        if result["dcf"]:
            col2.metric("DCF Implied Price", f"{sym}{result['dcf']['implied_share_price']:,.2f}", f"{result['dcf']['upside']:+.1%}")
        if result["target_mean"]:
            col3.metric("Analyst Target (Mean)", f"{sym}{result['target_mean']:,.2f}")
        if result["market_cap"]:
            col4.metric("Market Cap", f"{sym}{result['market_cap']:,.0f}")

        tab1, tab2, tab3, tab4 = st.tabs(["Trading & Ownership", "Multiples", "Analyst Consensus", "Valuation Synthesis"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Trading Performance**")
                st.write(f"52-week range: {sym}{result['fifty_two_week_low']:,.2f} – {sym}{result['fifty_two_week_high']:,.2f}")
                if result["returns_1y"] is not None:
                    st.write(f"1-year return: {result['returns_1y']:+.1%}")
                if result["beta"] is not None:
                    st.write(f"Beta: {result['beta']:.2f}")
            with c2:
                st.markdown("**Ownership & Float**")
                if result["insider_pct"] is not None:
                    st.write(f"Insider ownership: {result['insider_pct']:.1%}")
                if result["institution_pct"] is not None:
                    st.write(f"Institutional ownership: {result['institution_pct']:.1%}")
                if result["float_shares"]:
                    st.write(f"Float: {result['float_shares']:,.0f} shares")

        with tab2:
            mult_data = {
                "EV / Revenue": result["ev_revenue"], "EV / EBITDA": result["ev_ebitda"],
                "Trailing P/E": result["trailing_pe"], "Forward P/E": result["forward_pe"],
                "Price / Book": result["price_to_book"],
            }
            for label, val in mult_data.items():
                if val is not None:
                    st.write(f"**{label}:** {val:.2f}x")

        with tab3:
            if result["target_mean"]:
                st.write(f"**Rating:** {(result['recommendation_key'] or 'N/A').replace('_', ' ').title()}")
                st.write(f"**Analysts covering:** {result['num_analysts'] or 'N/A'}")
                st.write(f"**Target range:** {sym}{result['target_low']:,.2f} – {sym}{result['target_high']:,.2f}")
            else:
                st.info("No analyst consensus data was returned for this company.")

        with tab4:
            synth_rows = []
            if result["dcf"]:
                synth_rows.append(("DCF", result["dcf"]["implied_share_price"]))
            if result["comps"] and result["comps"].get("implied_ev_from_revenue") and result["shares_outstanding"]:
                nd = (result["comps"]["target"].get("total_debt", 0) or 0) - (result["comps"]["target"].get("cash", 0) or 0)
                synth_rows.append(("Comps (EV/Revenue)", (result["comps"]["implied_ev_from_revenue"] - nd) / result["shares_outstanding"]))
            if result["target_mean"]:
                synth_rows.append(("Analyst Target", result["target_mean"]))
            if synth_rows:
                st.bar_chart({label: val for label, val in synth_rows})
                avg = sum(v for _, v in synth_rows) / len(synth_rows)
                st.metric("Average of Available Methods", f"{sym}{avg:,.2f}",
                           f"{(avg - result['current_price']) / result['current_price']:+.1%}" if result["current_price"] else None)

        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, f"{ticker}_PublicCompanyValuation.xlsx")
            wb.save(xlsx_path)
            with open(xlsx_path, "rb") as f:
                xlsx_bytes = f.read()

            pptx_path = os.path.join(tmpdir, f"{ticker}_PublicCompanyValuation_PitchBook.pptx")
            with st.spinner("Building pitchbook..."):
                deal_tool.build_pitch_book_public_company(result, pptx_path)
            with open(pptx_path, "rb") as f:
                pptx_bytes = f.read()

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button("Download Excel Model", xlsx_bytes, file_name=f"{ticker}_PublicCompanyValuation.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with dl2:
            st.download_button("Download Pitchbook", pptx_bytes, file_name=f"{ticker}_PitchBook.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run the analysis: {e}")