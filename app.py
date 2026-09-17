import streamlit as st
import deal_tool

st.set_page_config(page_title="Deal Tool", page_icon="📊", layout="wide")

st.title("📊 Deal Tool")
st.caption("Institutional-grade financial analysis, in your browser.")

st.markdown("""
Use the sidebar to navigate between tools. Every page runs the same underlying
models as the command-line version — DCF, comps, LBO, M&A, and more — and
lets you download the generated Excel workbook or PowerPoint pitchbook
directly from the browser.
""")

col1, col2, col3 = st.columns(3)
with col1:
    st.subheader("Valuation")
    st.markdown("- Research\n- DCF\n- Trading Comps\n- Public Company Valuation\n- Sum-of-the-Parts")
with col2:
    st.subheader("Deal Modeling")
    st.markdown("- M&A Model\n- LBO Model\n- IPO Valuation\n- 3-Statement Model\n- Football Field")
with col3:
    st.subheader("Portfolio")
    st.markdown("- Watchlist\n- Portfolio Tracker\n- CGT Tax Report")

st.divider()
st.caption("Pick a page from the sidebar to get started.")