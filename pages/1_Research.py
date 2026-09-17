import streamlit as st
import deal_tool
import os
import tempfile

st.set_page_config(page_title="Research | Deal Tool", page_icon="🔎", layout="wide")
st.title("🔎 Equity Research")
st.caption("A quick pre-call brief: profile, fundamentals, DCF anchor, and recent headlines.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL, BHP.AX").strip().upper()
run = st.button("Run Research", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Researching {ticker}..."):
            research = deal_tool.run_research(ticker)

        company = research["company"]
        snap = research["snapshot"]
        sym = deal_tool.get_currency_symbol(company["quote_currency"])
        st.success(f"{company['company_name']} ({ticker})  ·  {company['exchange']}  ·  {company['sector']} / {company['industry']}")

        col1, col2, col3 = st.columns(3)
        col1.metric("Current Price", f"{sym}{research['current_price']:,.2f}")
        if snap.get("analyst_target"):
            col2.metric("Analyst Target", f"{sym}{snap['analyst_target']:,.2f}", snap.get("recommendation"))
        if research["dcf"]:
            col3.metric("DCF Implied Price", f"{sym}{research['dcf']['implied_share_price']:,.2f}", f"{research['dcf']['upside']:+.1%}")

        st.subheader("Fundamentals")
        f1, f2, f3, f4 = st.columns(4)
        if snap.get("market_cap"):
            f1.write(f"**Market Cap:** {sym}{snap['market_cap']:,.0f}")
        if snap.get("revenue"):
            f2.write(f"**Revenue:** {sym}{snap['revenue']:,.0f}")
        if snap.get("revenue_growth") is not None:
            f3.write(f"**Revenue Growth:** {snap['revenue_growth']:+.1%}")
        if snap.get("pe_ratio"):
            f4.write(f"**P/E (TTM):** {snap['pe_ratio']:.1f}x")

        if research["dcf"]:
            a = research["dcf"]["assumptions_used"]
            st.caption(f"DCF assumptions: {a['starting_growth_rate']:+.1%} → {a['ending_growth_rate']:+.1%} growth, "
                       f"{a['starting_fcf_margin']:+.1%} → {a['ending_fcf_margin']:+.1%} FCF margin, WACC {a['wacc']:.1%}")

        if research["news"]:
            st.subheader("Recent Headlines")
            for item in research["news"]:
                st.write(f"• {item['title']} ({item['publisher']})")

        with tempfile.TemporaryDirectory() as tmpdir:
            pptx_path = os.path.join(tmpdir, f"{ticker}_PitchBook.pptx")
            with st.spinner("Building pitchbook..."):
                deal_tool.build_pitch_book("research", pptx_path, research=research)
            with open(pptx_path, "rb") as f:
                st.download_button("Download Research Pitchbook", f.read(), file_name=f"{ticker}_PitchBook.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run research: {e}")