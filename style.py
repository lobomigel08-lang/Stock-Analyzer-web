import streamlit as st

NAVY = "#0F2A4A"
ACCENT = "#E8801C"
LIGHT_BG = "#F5F7FA"
BORDER = "#E3E7ED"
TEXT_MUTED = "#5B6472"


def inject_custom_css():
    st.markdown(f"""
    <style>
    /* Tighten default top padding so the page header sits closer to the top */
    .block-container {{
        padding-top: 2.2rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background-color: {NAVY};
    }}
    section[data-testid="stSidebar"] * {{
        color: #E8EDF4 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a {{
        border-radius: 6px;
        margin: 1px 8px;
    }}
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {{
        background-color: rgba(255,255,255,0.08);
    }}
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {{
        background-color: {ACCENT};
    }}

    /* Metric cards */
    div[data-testid="stMetric"] {{
        background-color: {LIGHT_BG};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 14px 16px;
        transition: box-shadow 0.15s ease;
    }}
    div[data-testid="stMetric"]:hover {{
        box-shadow: 0 2px 10px rgba(15,42,74,0.08);
    }}
    div[data-testid="stMetricLabel"] {{
        color: {TEXT_MUTED};
    }}

    /* Bordered containers (used for nav cards and result panels) */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 12px !important;
        transition: box-shadow 0.15s ease, transform 0.15s ease;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
        box-shadow: 0 6px 18px rgba(15,42,74,0.12);
        transform: translateY(-2px);
    }}

    /* Buttons */
    div.stButton > button, div.stDownloadButton > button {{
        border-radius: 8px;
        font-weight: 600;
        transition: box-shadow 0.15s ease;
    }}
    div.stButton > button[kind="primary"] {{
        background-color: {ACCENT};
        border: none;
    }}
    div.stButton > button[kind="primary"]:hover {{
        background-color: {NAVY};
        box-shadow: 0 4px 12px rgba(232,128,28,0.35);
    }}

    /* Tabs */
    button[data-baseweb="tab"] {{
        font-weight: 600;
    }}

    /* Headings */
    h1, h2, h3 {{
        color: {NAVY};
    }}

    /* Hide the default Streamlit hamburger footer branding for a cleaner look */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    </style>
    """, unsafe_allow_html=True)


def page_footer():
    """A small, consistent professional footer at the bottom of every page."""
    st.markdown(f"""
    <div style="margin-top: 48px; padding-top: 16px; border-top: 1px solid {BORDER};
                font-size: 12px; color: {TEXT_MUTED}; text-align: center;">
        Catalyst — Institutional-grade financial analysis
    </div>
    """, unsafe_allow_html=True)


def page_header(icon, title, subtitle):
    """Consistent branded header used at the top of every page, replacing
    plain st.title/st.caption with a styled block."""
    st.markdown(f"""
    <div style="padding: 4px 0 18px 0; border-bottom: 3px solid {ACCENT}; margin-bottom: 24px;">
        <div style="font-size: 30px; font-weight: 700; color: {NAVY};">{icon}&nbsp;&nbsp;{title}</div>
        <div style="font-size: 15px; color: {TEXT_MUTED}; margin-top: 4px;">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


def render_page_content(icon, header_title=None, header_subtitle=None):
    """
    Called at the top of every individual page (NOT app.py, which owns
    st.set_page_config and st.navigation centrally now). Applies the
    shared theme and optionally renders the branded page header.
    """
    inject_custom_css()
    if header_title:
        page_header(icon, header_title, header_subtitle or "")