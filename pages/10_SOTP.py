import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile

render_page_content("🧩", "Sum-of-the-Parts Valuation", "Segment financials and multiples are supplied directly — automated data providers don't reliably expose true segment-level data.")

company_name = st.text_input("Company Name", placeholder="e.g. Conglomerate Inc")

if "sotp_segments" not in st.session_state:
    st.session_state.sotp_segments = []

st.subheader("Segments")
with st.form("add_segment", clear_on_submit=True):
    c1, c2, c3, c4 = st.columns(4)
    seg_name = c1.text_input("Segment Name")
    metric_type = c2.selectbox("Metric", ["revenue", "ebitda"])
    metric_value = c3.number_input("Metric Value", min_value=0.0, step=1000000.0)
    multiple = c4.number_input("Multiple (x)", min_value=0.0, step=0.5)
    if st.form_submit_button("Add Segment"):
        if seg_name and metric_value > 0 and multiple > 0:
            st.session_state.sotp_segments.append({"name": seg_name, "metric_type": metric_type,
                                                     "metric_value": metric_value, "multiple": multiple})
        else:
            st.error("Segment name, metric value, and multiple are all required.")

if st.session_state.sotp_segments:
    st.dataframe([{"Segment": s["name"], "Metric": s["metric_type"], "Value": f"${s['metric_value']:,.0f}",
                   "Multiple": f"{s['multiple']:.1f}x"} for s in st.session_state.sotp_segments],
                 width="stretch", hide_index=True)
    if st.button("Clear Segments"):
        st.session_state.sotp_segments = []
        st.rerun()

st.divider()
c1, c2 = st.columns(2)
net_debt = c1.number_input("Consolidated Net Debt", min_value=0.0, step=1000000.0)
shares_outstanding = c2.number_input("Shares Outstanding (optional)", min_value=0.0, step=1000000.0)

run = st.button("Run Valuation", type="primary", disabled=not (company_name and st.session_state.sotp_segments))

if run:
    try:
        with st.spinner("Building sum-of-the-parts valuation..."):
            result = deal_tool.run_sotp_valuation(company_name, st.session_state.sotp_segments,
                                                    net_debt=net_debt, shares_outstanding=shares_outstanding or None)
            wb = deal_tool.export_sotp_to_excel(result)

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Enterprise Value", f"${result['total_ev']:,.0f}")
        col2.metric("Equity Value", f"${result['equity_value']:,.0f}")
        if result["implied_price_per_share"] is not None:
            col3.metric("Implied Price / Share", f"${result['implied_price_per_share']:,.2f}")

        st.subheader("Value Composition")
        st.bar_chart({s["name"]: s["implied_ev"] for s in result["segments"]})

        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = os.path.join(tmpdir, f"{company_name.replace(' ', '_')}_SOTP.xlsx")
            wb.save(xlsx_path)
            with open(xlsx_path, "rb") as f:
                xlsx_bytes = f.read()
            pptx_path = os.path.join(tmpdir, f"{company_name.replace(' ', '_')}_SOTP_PitchBook.pptx")
            deal_tool.build_pitch_book_sotp(result, pptx_path)
            with open(pptx_path, "rb") as f:
                pptx_bytes = f.read()

        dl1, dl2 = st.columns(2)
        dl1.download_button("Download Excel Model", xlsx_bytes, file_name=f"{company_name.replace(' ', '_')}_SOTP.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dl2.download_button("Download Pitchbook", pptx_bytes, file_name=f"{company_name.replace(' ', '_')}_SOTP_PitchBook.pptx",
                             mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        st.error(f"Couldn't run the valuation: {e}")

page_footer()