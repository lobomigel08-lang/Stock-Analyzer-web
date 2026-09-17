import streamlit as st
from style import render_page_content, page_footer
import deal_tool
import os
import tempfile
from datetime import datetime

render_page_content("💼", "Portfolio Tracker", "Australian CGT — FIFO cost-basis matching, 50% discount for holdings over 12 months. Planning estimate only, not tax advice.")

tab_holdings, tab_transactions, tab_tax = st.tabs(["Holdings", "Log a Transaction", "Tax Report"])

with tab_holdings:
    transactions = deal_tool.load_portfolio()
    holdings = deal_tool.compute_current_holdings(transactions)

    if not holdings:
        st.info("No current holdings — log a buy transaction to get started.")
    else:
        total_cost_base = 0
        total_market_value = 0
        rows = []
        for ticker, position in holdings.items():
            try:
                company = deal_tool.resolve_company(ticker)
                info = company["info"]
                sym = deal_tool.get_currency_symbol(company["quote_currency"])
                current_price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / company["price_divisor"]
                market_value = current_price * position["quantity"]
                cost_base = position["total_cost_base"]
                unrealized = market_value - cost_base
                total_cost_base += cost_base
                total_market_value += market_value
                rows.append({"Ticker": ticker, "Quantity": f"{position['quantity']:,.2f}",
                             "Avg Cost": f"{sym}{position['avg_cost_per_share']:,.2f}",
                             "Current": f"{sym}{current_price:,.2f}", "Market Value": f"{sym}{market_value:,.2f}",
                             "Unrealized": f"{sym}{unrealized:,.2f}", "%": f"{(unrealized / cost_base):+.1%}" if cost_base else "N/A"})
            except ValueError as e:
                rows.append({"Ticker": ticker, "Quantity": "Error", "Avg Cost": str(e), "Current": "", "Market Value": "", "Unrealized": "", "%": ""})

        st.dataframe(rows, width='stretch', hide_index=True)
        total_unrealized = total_market_value - total_cost_base
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Market Value", f"${total_market_value:,.2f}")
        c2.metric("Total Cost Base", f"${total_cost_base:,.2f}")
        c3.metric("Total Unrealized", f"${total_unrealized:,.2f}", f"{(total_unrealized / total_cost_base):+.1%}" if total_cost_base else None)

    st.divider()
    st.subheader("Transaction History")
    if transactions:
        st.dataframe([{"Date": t["date"], "Type": t["type"].upper(), "Ticker": t["ticker"],
                        "Qty": t["quantity"], "Price": deal_tool.format_price(t["price_per_share"]),
                        "Fees": f"${t['fees']:,.2f}"} for t in transactions], width='stretch', hide_index=True)
    else:
        st.caption("No transactions recorded yet.")

with tab_transactions:
    with st.form("log_transaction"):
        c1, c2 = st.columns(2)
        txn_ticker = c1.text_input("Ticker").strip().upper()
        txn_type = c2.selectbox("Type", ["buy", "sell"])
        c3, c4, c5 = st.columns(3)
        quantity = c3.number_input("Quantity", min_value=0.0, step=1.0)
        price = c4.number_input("Price per Share", min_value=0.0, step=0.01)
        txn_date = c5.date_input("Date", value=datetime.now())
        submitted = st.form_submit_button("Log Transaction", type="primary")

        if submitted:
            if not txn_ticker or quantity <= 0 or price <= 0:
                st.error("Ticker, quantity, and price are all required.")
            else:
                try:
                    deal_tool.add_transaction(txn_ticker, txn_type, quantity, price, date=txn_date.strftime("%Y-%m-%d"))
                    st.success(f"Logged: {txn_type.upper()} {quantity} {txn_ticker} @ {deal_tool.format_price(price)}")
                except ValueError as e:
                    st.error(str(e))

with tab_tax:
    try:
        realized = deal_tool.compute_realized_gains(deal_tool.load_portfolio())
    except ValueError as e:
        st.error(f"Couldn't compute realized gains: {e}")
        realized = None

    if realized is None:
        pass
    elif not realized:
        st.info("No sell transactions recorded yet — nothing to report.")
    else:
        years = sorted(set(r["financial_year"] for r in realized))
        selected_year = st.selectbox("Financial Year", years, index=len(years) - 1)
        marginal_rate = deal_tool.get_marginal_tax_rate()
        marginal_rate = st.number_input("Marginal Tax Rate (%, optional)", min_value=0.0, max_value=100.0,
                                          value=float(marginal_rate) if marginal_rate else 0.0, step=0.5)
        marginal_rate = marginal_rate if marginal_rate > 0 else None

        summary = deal_tool.summarize_financial_year(realized, selected_year)
        c1, c2, c3 = st.columns(3)
        c1.metric("Short-Term Gains", f"${summary['short_term_gain_gross']:,.2f}")
        c2.metric("Long-Term Gains", f"${summary['long_term_gain_gross']:,.2f}")
        c3.metric("Net Capital Gain", f"${summary['net_capital_gain']:,.2f}")

        if marginal_rate:
            st.metric(f"Estimated Tax Owed at {marginal_rate:.1f}%", f"${summary['net_capital_gain'] * (marginal_rate / 100):,.2f}")

        st.dataframe([{"Ticker": e["ticker"], "Bought": e["buy_date"], "Sold": e["sell_date"],
                        "Qty": e["quantity"], "Gain (AUD)": f"${e['gain_loss_aud']:,.2f}",
                        "Held (days)": e["held_days"], "50% Discount": "Yes" if e["discount_eligible"] else "No"}
                       for e in summary["events"]], width='stretch', hide_index=True)

        if st.button("Generate Excel Tax Report"):
            with tempfile.TemporaryDirectory() as tmpdir:
                with st.spinner("Building report..."):
                    real_path = deal_tool.export_tax_report_to_excel(selected_year, marginal_tax_rate=marginal_rate)
                with open(real_path, "rb") as f:
                    st.download_button("Download Tax Report", f.read(), file_name=os.path.basename(real_path),
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

page_footer()