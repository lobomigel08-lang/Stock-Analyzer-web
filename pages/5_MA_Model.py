import streamlit as st
import deal_tool
import os
import tempfile

st.set_page_config(page_title="M&A Model | Deal Tool", page_icon="🤝", layout="wide")
st.title("🤝 M&A Model")
st.caption("Accretion/dilution, purchase price allocation, and sources & uses.")

col_a, col_b = st.columns(2)
acquirer = col_a.text_input("Acquirer Ticker", placeholder="e.g. PANW").strip().upper()
target = col_b.text_input("Target Ticker", placeholder="e.g. ZS").strip().upper()
premium = st.slider("Offer Premium", 0.0, 0.60, 0.30, step=0.05)
st.caption(f"Premium: {premium:.0%}")
run = st.button("Run M&A Model", type="primary", disabled=not (acquirer and target))

if run and acquirer and target:
    try:
        with st.spinner(f"Modeling {acquirer} acquiring {target}..."):
            deal = deal_tool.run_ma_model(acquirer, target, offer_premium=premium)
            wb = deal_tool.export_ma_model_to_excel(deal)

        tcs, acs = deal["target_currency_symbol"], deal["acquirer_currency_symbol"]
        st.success(f"{deal['acquirer_name']} acquiring {deal['target_name']}")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Offer Price", f"{tcs}{deal['offer_price_per_share']:,.2f}")
        col2.metric("Total Consideration", f"{tcs}{deal['total_consideration']:,.0f}")
        col3.metric("Accretion / (Dilution)", f"{deal['accretion_dilution_pct']:+.1%}" if deal["accretion_dilution_pct"] is not None else "N/A")
        col4.metric("Pro Forma EPS", f"{acs}{deal['pro_forma_eps']:.2f}" if deal["pro_forma_eps"] else "N/A")

        tab1, tab2, tab3 = st.tabs(["Sources & Uses", "Purchase Price Allocation", "Synergies"])
        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Sources**")
                st.write(f"Cash from Balance Sheet: {acs}{deal['cash_from_balance_sheet']:,.0f}")
                st.write(f"New Debt Raised: {acs}{deal['new_debt_raised']:,.0f}")
                st.write(f"Stock Consideration: {tcs}{deal['stock_consideration']:,.0f}")
            with c2:
                st.markdown("**Uses**")
                st.write(f"Purchase of Target Equity: {tcs}{deal['total_consideration']:,.0f}")

        with tab2:
            st.write(f"**Target Book Equity:** {tcs}{deal['target_book_equity']:,.0f}")
            st.write(f"**Excess Purchase Price:** {tcs}{deal['excess_purchase_price']:,.0f}")
            st.write(f"**Intangible Step-Up:** {tcs}{deal['intangible_step_up']:,.0f}")
            st.write(f"**PP&E Step-Up:** {tcs}{deal['ppe_step_up']:,.0f}")
            st.write(f"**Goodwill (residual):** {tcs}{deal['goodwill']:,.0f}")
            st.write(f"**Deferred Tax Liability:** {tcs}{deal['deferred_tax_liability']:,.0f}")

        with tab3:
            st.write(f"**Pretax Synergies:** {tcs}{deal['pretax_synergies']:,.0f}")
            st.write(f"**After-Tax Synergies:** {tcs}{deal['aftertax_synergies']:,.0f}")
            st.write(f"**After-Tax Incremental D&A/Amortization Impact:** -{tcs}{deal['aftertax_incremental_da_amort']:,.0f}")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{acquirer}_{target}_MA.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                xlsx_bytes = f.read()
            ppt_path = os.path.join(tmpdir, f"{acquirer}_{target}_PitchBook.pptx")
            with st.spinner("Building pitchbook (re-runs the model across multiple premiums for the sensitivity table)..."):
                deal_tool.build_pitch_book_ma(deal, ppt_path)
            with open(ppt_path, "rb") as f:
                ppt_bytes = f.read()

        dl1, dl2 = st.columns(2)
        dl1.download_button("Download Excel Model", xlsx_bytes, file_name=f"{acquirer}_{target}_MA.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dl2.download_button("Download Pitchbook", ppt_bytes, file_name=f"{acquirer}_{target}_PitchBook.pptx",
                             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run the M&A model: {e}")