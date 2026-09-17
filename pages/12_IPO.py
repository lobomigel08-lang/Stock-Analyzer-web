import streamlit as st
import deal_tool
import os
import tempfile

st.set_page_config(page_title="IPO Valuation | Deal Tool", page_icon="🚀", layout="wide")
st.title("🚀 IPO Valuation")
st.caption("Peer-multiple valuation for a pre-IPO company — no trading history to anchor to, so this leans on comps.")

company_name = st.text_input("Company Name", placeholder="e.g. NewCo")
industry = st.selectbox("Industry", list(deal_tool.INDUSTRY_UNIVERSE.keys()))
c1, c2, c3 = st.columns(3)
revenue = c1.number_input("Revenue", min_value=0.0, step=1000000.0)
ebitda = c2.number_input("EBITDA (optional)", min_value=0.0, step=1000000.0)
shares = c3.number_input("Shares Outstanding Post-IPO (optional)", min_value=0.0, step=1000000.0)
illiquidity_discount = st.slider("Illiquidity Discount", 0.0, 0.40, 0.15, step=0.05)

run = st.button("Run Valuation", type="primary", disabled=not (company_name and revenue > 0))

if run:
    try:
        with st.spinner("Building IPO valuation..."):
            ipo = deal_tool.run_ipo_valuation(company_name, industry, revenue, ebitda=ebitda or None,
                                                shares_outstanding_post_ipo=shares or None,
                                                illiquidity_discount=illiquidity_discount)
            wb = deal_tool.export_ipo_valuation_to_excel(ipo)

        st.success(f"{company_name} — {industry}")

        col1, col2, col3 = st.columns(3)
        if ipo["implied_price_per_share"] is not None:
            col1.metric("Blended Price / Share", f"${ipo['implied_price_per_share']:,.2f}")
            col2.metric(f"After {illiquidity_discount:.0%} Discount", f"${ipo['discounted_price_per_share']:,.2f}")
        col3.metric("Peers Used", str(len(ipo["peers"])))

        st.subheader("Peer Set")
        rows = [{"Ticker": p["ticker"], "Company": p["company_name"],
                 "EV/Revenue": f"{p['ev_revenue']:.2f}x" if p.get("ev_revenue") else "N/A",
                 "EV/EBITDA": f"{p['ev_ebitda']:.2f}x" if p.get("ev_ebitda") else "N/A"} for p in ipo["peers"]]
        if rows:
            st.dataframe(rows, width="stretch", hide_index=True)

        st.subheader("Implied Price by Method")
        method_prices = {}
        for label, key in [("EV/Revenue", "implied_price_from_revenue"), ("EV/EBITDA", "implied_price_from_ebitda"), ("P/E", "implied_price_from_pe")]:
            if ipo.get(key) is not None:
                method_prices[label] = ipo[key]
        if method_prices:
            st.bar_chart(method_prices)

        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, f"{company_name.replace(' ', '_')}_IPO.xlsx")
            wb.save(xlsx_path)
            with open(xlsx_path, "rb") as f:
                xlsx_bytes = f.read()
            pptx_path = os.path.join(tmpdir, f"{company_name.replace(' ', '_')}_IPO_PitchBook.pptx")
            deal_tool.build_pitch_book_ipo(ipo, "$", pptx_path)
            with open(pptx_path, "rb") as f:
                pptx_bytes = f.read()

        dl1, dl2 = st.columns(2)
        dl1.download_button("Download Excel Model", xlsx_bytes, file_name=f"{company_name.replace(' ', '_')}_IPO.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dl2.download_button("Download Pitchbook", pptx_bytes, file_name=f"{company_name.replace(' ', '_')}_IPO_PitchBook.pptx",
                             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run the IPO valuation: {e}")