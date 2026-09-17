import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile

render_page_content("💰", "LBO Model", "Entry at market multiple, real circular debt schedule with a 100% cash sweep, MOIC/IRR.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL").strip().upper()
col_a, col_b, col_c = st.columns(3)
leverage = col_a.number_input("Leverage (x EBITDA)", min_value=1.0, max_value=8.0, value=4.5, step=0.5)
hold_years = col_b.number_input("Hold Period (years)", min_value=1, max_value=10, value=5, step=1)
run = st.button("Run LBO", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Running LBO for {ticker}..."):
            result = deal_tool.run_lbo_model(ticker, leverage_multiple=leverage, hold_years=int(hold_years))
            wb = deal_tool.export_lbo_to_excel(result)
            moic, irr = deal_tool._lbo_quick_returns(result, result["entry_multiple"], result["exit_multiple"])

        sym = result["currency_symbol"]
        entry_ev = result["entry_ebitda"] * result["entry_multiple"]
        st.success(f"{result['company_name']} ({ticker})")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Entry EV", f"{sym}{entry_ev:,.0f}")
        col2.metric("Leverage", f"{result['leverage_multiple']:.1f}x")
        col3.metric("MOIC", f"{moic:.2f}x" if moic else "N/A")
        col4.metric("IRR", f"{irr:.1%}" if irr else "N/A")

        st.subheader("Entry Sources & Uses")
        new_debt = result["entry_ebitda"] * result["leverage_multiple"]
        fees = entry_ev * 0.02
        sponsor_equity = (entry_ev + fees) - new_debt
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**New Debt:** {sym}{new_debt:,.0f}")
            st.write(f"**Sponsor Equity:** {sym}{sponsor_equity:,.0f}")
        with c2:
            st.write(f"**Enterprise Value:** {sym}{entry_ev:,.0f}")
            st.write(f"**Transaction Fees:** {sym}{fees:,.0f}")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{ticker}_LBO.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                xlsx_bytes = f.read()
            ppt_path = os.path.join(tmpdir, f"{ticker}_LBO_PitchBook.pptx")
            deal_tool.build_pitch_book_lbo(result, ppt_path)
            with open(ppt_path, "rb") as f:
                ppt_bytes = f.read()

        dl1, dl2 = st.columns(2)
        dl1.download_button("Download Excel Model", xlsx_bytes, file_name=f"{ticker}_LBO.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dl2.download_button("Download Pitchbook", ppt_bytes, file_name=f"{ticker}_LBO_PitchBook.pptx",
                             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run the LBO: {e}")

page_footer()