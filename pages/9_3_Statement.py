import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile

render_page_content("📑", "3-Statement Model", "Fully integrated income statement, balance sheet, and cash flow with a genuine circular debt schedule.")

ticker = st.text_input("Ticker", placeholder="e.g. ZS, AAPL").strip().upper()
years = st.slider("Years to Project", 3, 7, 5)
run = st.button("Run Model", type="primary", disabled=not ticker)

if run and ticker:
    try:
        with st.spinner(f"Building 3-statement model for {ticker}..."):
            result = deal_tool.run_three_statement_model(ticker, years_to_project=years)
            wb = deal_tool.export_three_statement_to_excel(result)

        sym = result["currency_symbol"]
        st.success(f"{result['company_name']} ({ticker})")

        final_year = result["years"][-1]
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(f"Year {years} Revenue", f"{sym}{final_year['revenue']:,.0f}")
        col2.metric(f"Year {years} Net Income", f"{sym}{final_year['net_income']:,.0f}")
        col3.metric(f"Year {years} Ending Debt", f"{sym}{final_year['debt_ending']:,.0f}")
        col4.metric("Balance Check", f"{sym}{final_year['balance_check']:,.4f}",
                    "Balances" if abs(final_year["balance_check"]) < 1 else "Review needed")

        st.subheader("Revenue & Net Income Trajectory")
        st.line_chart({"Revenue": [y["revenue"] for y in result["years"]], "Net Income": [y["net_income"] for y in result["years"]]})

        st.subheader("Debt Schedule")
        st.line_chart({"Ending Debt": [y["debt_ending"] for y in result["years"]]})

        with st.expander("Year-by-Year Detail"):
            rows = [{"Year": y["year"], "Revenue": f"{sym}{y['revenue']:,.0f}", "EBIT": f"{sym}{y['ebit']:,.0f}",
                     "Net Income": f"{sym}{y['net_income']:,.0f}", "Cash": f"{sym}{y['cash']:,.0f}",
                     "Debt": f"{sym}{y['debt_ending']:,.0f}", "Balance Check": f"{y['balance_check']:.4f}"}
                    for y in result["years"]]
            st.dataframe(rows, width="stretch", hide_index=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, f"{ticker}_3Statement.xlsx")
            wb.save(path)
            with open(path, "rb") as f:
                st.download_button("Download Excel Model", f.read(), file_name=f"{ticker}_3Statement.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Couldn't run the 3-statement model: {e}")

page_footer()