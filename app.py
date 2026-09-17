import streamlit as st
from style import inject_custom_css, page_footer, NAVY, ACCENT, LIGHT_BG, BORDER, TEXT_MUTED

st.set_page_config(page_title="Catalyst", page_icon="📊", layout="wide")


def home_page():
    inject_custom_css()

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, {NAVY} 0%, #1B4A72 100%); border-radius: 16px;
                padding: 48px 40px; margin-bottom: 20px; border-left: 5px solid {ACCENT};">
        <div style="font-size: 40px; font-weight: 800; color: white; letter-spacing: -0.5px;">Catalyst</div>
        <div style="font-size: 17px; color: #C9D8E8; margin-top: 8px; max-width: 640px;">
            Institutional-grade financial analysis, in your browser. DCF, comps, LBO, M&amp;A, and more —
            run the model, get the live Excel workbook and pitchbook, all in one place.
        </div>
    </div>
    """, unsafe_allow_html=True)

    stats = [("12", "Models & Tools"), ("Live", "Excel Export"), ("Auto", "Pitchbook Generation"), ("Real", "Market Data")]
    stat_cols = st.columns(4)
    for col, (value, label) in zip(stat_cols, stats):
        with col:
            st.markdown(f"""
            <div style="text-align:center; padding: 14px 0;">
                <div style="font-size:26px; font-weight:800; color:{NAVY};">{value}</div>
                <div style="font-size:12px; color:{TEXT_MUTED}; text-transform:uppercase; letter-spacing:0.04em;">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

    def section_label(text):
        st.markdown(f"<div style='font-size:13px; font-weight:700; color:{TEXT_MUTED}; "
                    f"text-transform:uppercase; letter-spacing:0.06em; margin: 8px 0 10px 0;'>{text}</div>",
                    unsafe_allow_html=True)

    def nav_card(col, page_path, icon, title, description):
        with col:
            with st.container(border=True):
                st.markdown(f"<div style='font-size:26px;'>{icon}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='font-weight:700; font-size:16px; color:{NAVY}; margin-top:4px;'>{title}</div>",
                            unsafe_allow_html=True)
                st.markdown(f"<div style='font-size:13px; color:{TEXT_MUTED}; margin: 4px 0 12px 0; min-height:38px;'>"
                            f"{description}</div>", unsafe_allow_html=True)
                st.page_link(page_path, label="Open →")

    section_label("Valuation")
    c1, c2, c3 = st.columns(3)
    nav_card(c1, "pages/1_Research.py", "🔎", "Research", "A quick pre-call brief with a DCF anchor and recent headlines.")
    nav_card(c2, "pages/2_DCF.py", "📈", "DCF", "WACC build-up, scenario toggle, and a live Excel export.")
    nav_card(c3, "pages/8_Comps.py", "📋", "Trading Comps", "Peer benchmarking with percentile bands.")

    c4, c5, c6 = st.columns(3)
    nav_card(c4, "pages/3_Public_Company.py", "🏢", "Public Company Valuation", "Trading, ownership, multiples, and consensus combined.")
    nav_card(c5, "pages/10_SOTP.py", "🧩", "Sum-of-the-Parts", "Segment-level valuation for multi-business companies.")
    nav_card(c6, "pages/11_Football_Field.py", "🏈", "Football Field", "DCF, comps, and analyst targets in one range summary.")

    section_label("Deal Modeling")
    c7, c8, c9 = st.columns(3)
    nav_card(c7, "pages/5_MA_Model.py", "🤝", "M&A Model", "Accretion/dilution, purchase price allocation, sources & uses.")
    nav_card(c8, "pages/4_LBO.py", "💰", "LBO Model", "Real circular debt schedule with a cash sweep, MOIC/IRR.")
    nav_card(c9, "pages/12_IPO.py", "🚀", "IPO Valuation", "Peer-multiple valuation for a pre-IPO company.")

    c10, c11, _ = st.columns(3)
    nav_card(c10, "pages/9_3_Statement.py", "📑", "3-Statement Model", "Fully integrated statements with a circular debt schedule.")

    section_label("Portfolio")
    c12, c13, _ = st.columns(3)
    nav_card(c12, "pages/6_Watchlist.py", "👀", "Watchlist", "Track tickers and see how they're moving.")
    nav_card(c13, "pages/7_Portfolio.py", "💼", "Portfolio Tracker", "Holdings, transactions, and a CGT tax report.")

    page_footer()


pg = st.navigation({
    "": [st.Page(home_page, title="Home", icon="🏠", default=True)],
    "Valuation": [
        st.Page("pages/1_Research.py", title="Research", icon="🔎"),
        st.Page("pages/2_DCF.py", title="DCF", icon="📈"),
        st.Page("pages/8_Comps.py", title="Trading Comps", icon="📋"),
        st.Page("pages/3_Public_Company.py", title="Public Company Valuation", icon="🏢"),
        st.Page("pages/10_SOTP.py", title="Sum-of-the-Parts", icon="🧩"),
        st.Page("pages/11_Football_Field.py", title="Football Field", icon="🏈"),
    ],
    "Deal Modeling": [
        st.Page("pages/5_MA_Model.py", title="M&A Model", icon="🤝"),
        st.Page("pages/4_LBO.py", title="LBO Model", icon="💰"),
        st.Page("pages/12_IPO.py", title="IPO Valuation", icon="🚀"),
        st.Page("pages/9_3_Statement.py", title="3-Statement Model", icon="📑"),
    ],
    "Portfolio": [
        st.Page("pages/6_Watchlist.py", title="Watchlist", icon="👀"),
        st.Page("pages/7_Portfolio.py", title="Portfolio Tracker", icon="💼"),
    ],
})

st.markdown(f"""
<style>
[data-testid="stSidebarNav"] {{ padding-top: 0.5rem; }}
</style>
""", unsafe_allow_html=True)

st.logo("assets/catalyst_logo.svg", size="large", icon_image="assets/catalyst_icon.svg")

pg.run()