import sys
import re
import os
import json
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import yfinance as ticker_module
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference


def get_app_data_dir():
    """
    When running as a packaged standalone app (e.g. built with PyInstaller),
    data needs a fixed, predictable, writable location regardless of how
    the app was launched (double-clicked, run from a random folder, etc.)
    — and critically, each person who runs their own copy of the app
    should get their own separate data, not risk colliding with anyone
    else's. `~/DealTool` serves that purpose.

    When running as a normal script (e.g. via GitHub Actions, which reads
    watchlist.json relative to the checked-out repo each run), this keeps
    using the current working directory exactly as before — changing that
    unconditionally would silently break the existing daily-email
    automation, which depends on the repo-relative path.
    """
    if getattr(sys, "frozen", False):
        data_dir = os.path.expanduser("~/DealTool")
        os.makedirs(data_dir, exist_ok=True)
        return data_dir
    return "."


APP_DATA_DIR = get_app_data_dir()

from copy import copy


# ============================================
# MARKET-AWARE DATA LAYER
# Everything below (research, DCF, comps) is built on top of this.
# Solve ticker resolution, currency, and missing-field problems once,
# here — not separately inside every mode.
# ============================================

CURRENCY_SYMBOLS = {
    "USD": "$", "GBP": "£", "EUR": "€", "JPY": "¥", "AUD": "A$",
    "CAD": "C$", "KRW": "₩", "HKD": "HK$", "CHF": "CHF ", "CNY": "¥", "INR": "₹",
}


def get_ticker(ticker_symbol):
    """Single point of contact with yfinance. If the data source ever changes,
    or retry/caching logic is needed later, this is the only place to edit."""
    return ticker_module.Ticker(ticker_symbol)


def get_currency_symbol(currency_code):
    """Falls back to 'CODE ' (e.g. 'SEK ') for currencies without a common symbol,
    rather than defaulting to $ and silently mislabeling foreign figures."""
    return CURRENCY_SYMBOLS.get(currency_code, f"{currency_code} ")


def format_price(value, currency_symbol="$"):
    """
    Never rounds off precision you actually entered, at any price level —
    shows the exact value, with unnecessary trailing zeros trimmed for
    readability. A round price like $131.00 still displays cleanly; a
    price like $131.567 or $1.195 shows every digit, not just the first 2.
    """
    formatted = f"{value:,.10f}".rstrip("0").rstrip(".")
    if "." not in formatted:
        formatted += ".00"
    else:
        decimals = formatted.split(".")[1]
        if len(decimals) < 2:
            formatted += "0" * (2 - len(decimals))
    return f"{currency_symbol}{formatted}"


def safe_get_row(dataframe, possible_row_names, column_index=0, default=None):
    """
    Financial statement row labels vary by market and filer (e.g. 'Total Debt'
    may not exist for some non-US companies, or may be named differently).
    Try each candidate name in order instead of one .loc[] call that throws
    KeyError the moment a label doesn't match exactly.
    """
    if dataframe is None or dataframe.empty:
        return default
    for name in possible_row_names:
        if name in dataframe.index:
            try:
                value = dataframe.loc[name].iloc[column_index]
                if value is not None:
                    return value
            except (IndexError, KeyError):
                continue
    return default


def get_currency_info(info):
    """
    Quote currency (what the share price is in) and financial currency
    (what the statements are in) can differ. The classic trap: many UK
    stocks quote in GBp (pence) while everything else is in GBP — 100 GBp
    = 1 GBP. Missing this silently makes every downstream number wrong
    by a factor of 100.
    """
    quote_currency = info.get("currency", "USD")
    financial_currency = info.get("financialCurrency", quote_currency)

    price_divisor = 100 if quote_currency == "GBp" else 1
    normalized_quote_currency = "GBP" if quote_currency == "GBp" else quote_currency

    return {
        "quote_currency": normalized_quote_currency,
        "financial_currency": financial_currency,
        "price_divisor": price_divisor,
        "currencies_match": normalized_quote_currency == financial_currency,
    }


def resolve_company(ticker_symbol):
    """
    First call every mode should make. Confirms the ticker actually resolves
    to something tradeable (catches typos and wrong exchange suffixes early,
    e.g. 'SAP' vs the correct 'SAP.DE'), and returns currency/market context
    everything downstream needs.
    """
    ticker = get_ticker(ticker_symbol)
    try:
        info = ticker.info or {}
    except Exception:
        # yfinance's own internals can throw directly (a malformed/blocked
        # response, rate limiting, a network hiccup) rather than just
        # returning empty data — without this, that raw library exception
        # propagates unhandled through every single caller in this tool.
        raise ValueError(
            f"Could not resolve '{ticker_symbol}'. Check the symbol and exchange "
            f"suffix — e.g. Tokyo '7203.T', Frankfurt 'SAP.DE', Sydney 'BHP.AX'."
        )

    has_price = info.get("currentPrice") is not None or info.get("regularMarketPrice") is not None
    if not info or not has_price:
        raise ValueError(
            f"Could not resolve '{ticker_symbol}'. Check the symbol and exchange "
            f"suffix — e.g. Tokyo '7203.T', Frankfurt 'SAP.DE', Sydney 'BHP.AX'."
        )

    currency_info = get_currency_info(info)

    return {
        "ticker": ticker_symbol,
        "ticker_obj": ticker,
        "info": info,
        "company_name": info.get("shortName") or info.get("longName") or ticker_symbol,
        "exchange": info.get("exchange", "Unknown"),
        "country": info.get("country", "Unknown"),
        "industry": info.get("industry", "Unknown"),
        "sector": info.get("sector", "Unknown"),
        **currency_info,
    }


# ============================================
# DCF MODEL
# ============================================

REVENUE_ROW_NAMES = ["Total Revenue", "TotalRevenue", "Revenue"]
DEBT_ROW_NAMES = ["Total Debt", "TotalDebt", "Long Term Debt"]
CASH_ROW_NAMES = ["Cash And Cash Equivalents", "CashAndCashEquivalents", "Cash"]
OCF_ROW_NAMES = ["Operating Cash Flow", "Total Cash From Operating Activities", "OperatingCashFlow"]
CAPEX_ROW_NAMES = ["Capital Expenditure", "CapitalExpenditures", "CapitalExpenditure"]

# Generic, conservative fallbacks used only when a company has no usable
# historical data at all (e.g. a very recent IPO). These are deliberately
# modest — NOT tuned to any one company — so an unknown company doesn't
# quietly get treated as a fast-growing, high-margin SaaS business by default.
FALLBACK_GROWTH_RATE = 0.10
FALLBACK_FCF_MARGIN = 0.10


def estimate_dcf_assumptions(ticker_obj, info):
    """
    Derives starting growth/margin assumptions from the company's own
    financials, instead of hardcoding one company's profile (e.g. a mature,
    high-margin SaaS business) as the default for every ticker. Without this,
    running the DCF on an early-stage or capex-heavy company — where current
    revenue is small and FCF margin may be thin or negative — silently
    produces a tiny, meaningless implied value using assumptions that don't
    describe that company at all.
    """
    financials = ticker_obj.financials
    cashflow = ticker_obj.cashflow

    # --- Historical revenue growth (most recent year vs. prior year) ---
    revenue_series = None
    for name in REVENUE_ROW_NAMES:
        if financials is not None and not financials.empty and name in financials.index:
            revenue_series = financials.loc[name].dropna()
            break

    historical_growth = None
    if revenue_series is not None and len(revenue_series) >= 2:
        latest, prior = revenue_series.iloc[0], revenue_series.iloc[1]
        if prior:
            historical_growth = (latest - prior) / abs(prior)

    # --- FCF margin from actual cash flow statement: (operating cash flow - capex) / revenue ---
    fcf_margin = None
    try:
        ocf_row = capex_row = None
        for name in OCF_ROW_NAMES:
            if cashflow is not None and not cashflow.empty and name in cashflow.index:
                ocf_row = cashflow.loc[name]
                break
        for name in CAPEX_ROW_NAMES:
            if cashflow is not None and not cashflow.empty and name in cashflow.index:
                capex_row = cashflow.loc[name]
                break
        if ocf_row is not None and capex_row is not None and revenue_series is not None and len(revenue_series) >= 1:
            ocf = ocf_row.iloc[0]
            capex = capex_row.iloc[0]
            free_cash_flow = ocf + capex if capex < 0 else ocf - capex  # capex is usually stored negative
            most_recent_revenue = revenue_series.iloc[0]
            if most_recent_revenue:
                fcf_margin = free_cash_flow / most_recent_revenue
    except Exception:
        fcf_margin = None

    # --- Fall back to yfinance summary fields, then to generic conservative defaults ---
    growth = historical_growth if historical_growth is not None else info.get("revenueGrowth")
    used_fallback_growth = growth is None
    if growth is None:
        growth = FALLBACK_GROWTH_RATE
    growth = max(min(growth, 1.50), -0.50)  # clamp: no company should be modeled at, say, 400% or -90% forever

    margin = fcf_margin if fcf_margin is not None else info.get("profitMargins")
    used_fallback_margin = margin is None
    if margin is None:
        margin = FALLBACK_FCF_MARGIN
    margin = max(min(margin, 0.50), -0.75)

    # Fade assumptions toward more moderate long-run levels over the
    # projection window, same idea as the original hardcoded defaults,
    # but anchored to this company's own starting point instead of Zscaler's.
    ending_growth = max(min(growth * 0.5, 0.20), 0.03)
    ending_margin = margin if margin > 0.10 else 0.10  # assume a path toward modest profitability, not permanent negative margins

    is_negative_margin = margin < 0

    return {
        "starting_growth_rate": growth,
        "ending_growth_rate": ending_growth,
        "starting_fcf_margin": margin,
        "ending_fcf_margin": ending_margin,
        "used_fallback_growth": used_fallback_growth,
        "used_fallback_margin": used_fallback_margin,
        "is_negative_margin": is_negative_margin,
    }


def run_dcf(ticker_symbol, starting_growth_rate=None, ending_growth_rate=None,
            starting_fcf_margin=None, ending_fcf_margin=None, wacc=0.09,
            terminal_growth_rate=0.03, years_to_project=5):
    """
    Any of the four growth/margin parameters left as None are derived from
    the company's own financials via estimate_dcf_assumptions(), rather than
    defaulting to one hardcoded profile. Pass explicit values to override
    (e.g. for scenario testing) exactly as before.
    """

    company = resolve_company(ticker_symbol)
    ticker = company["ticker_obj"]
    info = company["info"]
    price_divisor = company["price_divisor"]

    income_statement = ticker.financials
    most_recent_revenue = safe_get_row(income_statement, REVENUE_ROW_NAMES)
    if most_recent_revenue is None:
        raise ValueError(f"No revenue data found for '{ticker_symbol}' — this market/filer may report it under a different label.")

    assumptions = estimate_dcf_assumptions(ticker, info)
    if starting_growth_rate is None:
        starting_growth_rate = assumptions["starting_growth_rate"]
    if ending_growth_rate is None:
        ending_growth_rate = assumptions["ending_growth_rate"]
    if starting_fcf_margin is None:
        starting_fcf_margin = assumptions["starting_fcf_margin"]
    if ending_fcf_margin is None:
        ending_fcf_margin = assumptions["ending_fcf_margin"]

    # Purely additive: a real operating-model breakdown (Revenue -> Gross
    # Profit -> SG&A -> EBITDA -> D&A -> EBIT -> NOPAT), sourced from the
    # same real-financials estimator the 3-statement model uses. This does
    # NOT change projected_fcf/implied_share_price below (every existing
    # caller — scenarios, sensitivity, embedded M&A snapshots, football
    # field — keeps working unchanged); it's used by the Excel export to
    # show the actual operating build instead of a blended margin.
    operating_assumptions = estimate_three_statement_assumptions(ticker, info)

    projected_revenue, projected_fcf, discounted_fcf = [], [], []
    current_revenue = most_recent_revenue

    for year in range(1, years_to_project + 1):
        progress = (year - 1) / (years_to_project - 1)

        growth_this_year = starting_growth_rate + (ending_growth_rate - starting_growth_rate) * progress
        current_revenue = current_revenue * (1 + growth_this_year)
        projected_revenue.append(current_revenue)

        margin_this_year = starting_fcf_margin + (ending_fcf_margin - starting_fcf_margin) * progress
        fcf_this_year = current_revenue * margin_this_year
        projected_fcf.append(fcf_this_year)

        discounted_fcf.append(fcf_this_year / ((1 + wacc) ** year))

    sum_of_discounted_fcf = sum(discounted_fcf)

    final_year_fcf = projected_fcf[-1]
    terminal_value = (final_year_fcf * (1 + terminal_growth_rate)) / (wacc - terminal_growth_rate)
    pv_terminal_value = terminal_value / ((1 + wacc) ** years_to_project)

    enterprise_value = sum_of_discounted_fcf + pv_terminal_value

    balance_sheet = ticker.balance_sheet
    total_debt = safe_get_row(balance_sheet, DEBT_ROW_NAMES, default=0) or 0
    cash = safe_get_row(balance_sheet, CASH_ROW_NAMES, default=0) or 0
    net_debt = total_debt - cash

    shares_outstanding = info.get("sharesOutstanding")
    if not shares_outstanding:
        raise ValueError(f"No shares outstanding figure for '{ticker_symbol}' — cannot derive a per-share value.")

    equity_value = enterprise_value - net_debt
    implied_share_price = equity_value / shares_outstanding

    current_price = (info.get("currentPrice") or info.get("regularMarketPrice")) / price_divisor
    upside = (implied_share_price - current_price) / current_price

    market_cap = info.get("marketCap")
    beta = info.get("beta")
    historical_revenue = income_statement.loc[REVENUE_ROW_NAMES[0]].dropna().tolist()[:3] \
        if REVENUE_ROW_NAMES[0] in income_statement.index else []

    return {
        "ticker": ticker_symbol,
        "company_name": company["company_name"],
        "currency": company["financial_currency"],
        "currency_symbol": get_currency_symbol(company["financial_currency"]),
        "projected_revenue": projected_revenue,
        "projected_fcf": projected_fcf,
        "discounted_fcf": discounted_fcf,
        "enterprise_value": enterprise_value,
        "net_debt": net_debt,
        "equity_value": equity_value,
        "shares_outstanding": shares_outstanding,
        "implied_share_price": implied_share_price,
        "current_price": current_price,
        "upside": upside,
        "market_cap": market_cap,
        "beta": beta,
        "total_debt": total_debt,
        "historical_revenue": historical_revenue,  # most recent first
        "assumptions_used": {
            "starting_growth_rate": starting_growth_rate,
            "ending_growth_rate": ending_growth_rate,
            "starting_fcf_margin": starting_fcf_margin,
            "ending_fcf_margin": ending_fcf_margin,
            "wacc": wacc,
            "terminal_growth_rate": terminal_growth_rate,
        },
        "used_fallback_growth": assumptions["used_fallback_growth"],
        "used_fallback_margin": assumptions["used_fallback_margin"],
        "is_negative_margin": assumptions["is_negative_margin"],
    }


def build_football_field(ticker_symbol):
    """
    Pulls together every valuation method that has REAL computed data —
    DCF (Bear/Base/Bull range), Comps (25th/75th percentile bands for
    EV/Revenue and EV/EBITDA), and analyst price targets (a genuine
    yfinance field, not fabricated) — into a single Low/Mid/High summary,
    exactly like a real banking valuation page. Skips any method whose
    underlying data isn't available, rather than filling in a placeholder.
    """
    methods = []
    current_price, currency_symbol = None, "$"

    try:
        scenarios = run_scenarios(ticker_symbol)
        prices = [scenarios[k]["implied_share_price"] for k in ("Bear", "Base", "Bull")]
        methods.append({"method": "DCF (Bear/Base/Bull)", "low": min(prices), "mid": scenarios["Base"]["implied_share_price"], "high": max(prices)})
        current_price = scenarios["Base"]["current_price"]
        currency_symbol = scenarios["Base"]["currency_symbol"]
    except Exception:
        pass

    try:
        comps = build_comps_table(ticker_symbol)
        target = comps["target"]
        net_debt = comps["net_debt"]
        shares = target.get("shares_outstanding") if target else None
        if shares and net_debt is not None:
            for label, band_key, metric_key in [("Comps (EV/Revenue)", "ev_revenue_band", "revenue"),
                                                   ("Comps (EV/EBITDA)", "ev_ebitda_band", "ebitda")]:
                band = comps[band_key]
                metric = target.get(metric_key)
                if band["p25"] and band["p75"] and metric and metric > 0:
                    low = (band["p25"] * metric - net_debt) / shares
                    high = (band["p75"] * metric - net_debt) / shares
                    mid = (band["median"] * metric - net_debt) / shares
                    methods.append({"method": label, "low": min(low, high), "mid": mid, "high": max(low, high)})
    except Exception:
        pass

    try:
        company = resolve_company(ticker_symbol)
        info = company["info"]
        t_low, t_high = info.get("targetLowPrice"), info.get("targetHighPrice")
        t_mid = info.get("targetMedianPrice") or info.get("targetMeanPrice")
        if t_low and t_high:
            methods.append({"method": "Analyst Price Targets", "low": t_low, "mid": t_mid or (t_low + t_high) / 2, "high": t_high})
        if current_price is None:
            current_price = (info.get("currentPrice") or info.get("regularMarketPrice")) / company["price_divisor"]
            currency_symbol = get_currency_symbol(company["financial_currency"])
    except Exception:
        pass

    return {
        "ticker": ticker_symbol,
        "current_price": current_price,
        "currency_symbol": currency_symbol,
        "methods": methods,
    }


def run_scenarios(ticker_symbol):
    """
    Bear/Base/Bull are now built as adjustments RELATIVE to this company's
    own base-case assumptions (derived from its own financials), rather than
    fixed absolute numbers that only made sense for one specific company.
    A company with a 5% base growth rate gets a Bear/Bull spread around 5%,
    not around 23%.
    """
    base = run_dcf(ticker_symbol)
    base_assumptions = base["assumptions_used"]

    def shifted(growth_delta, margin_delta, wacc_delta):
        return run_dcf(
            ticker_symbol,
            starting_growth_rate=base_assumptions["starting_growth_rate"] + growth_delta,
            ending_growth_rate=max(base_assumptions["ending_growth_rate"] + growth_delta * 0.5, 0.01),
            starting_fcf_margin=base_assumptions["starting_fcf_margin"] + margin_delta,
            ending_fcf_margin=base_assumptions["ending_fcf_margin"] + margin_delta,
            wacc=base_assumptions["wacc"] + wacc_delta,
            terminal_growth_rate=base_assumptions["terminal_growth_rate"],
        )

    return {
        "Bear": shifted(growth_delta=-0.08, margin_delta=-0.05, wacc_delta=+0.02),
        "Base": base,
        "Bull": shifted(growth_delta=+0.08, margin_delta=+0.05, wacc_delta=-0.01),
    }


def build_sensitivity_table(ticker_symbol, row_range, col_range, vary="wacc_vs_tg",
                             starting_growth_rate=0.23, ending_growth_rate=0.12,
                             starting_fcf_margin=0.27, ending_fcf_margin=0.32,
                             wacc=0.09, terminal_growth_rate=0.03):
    table = {}
    for row_val in row_range:
        table[row_val] = {}
        for col_val in col_range:
            if vary == "wacc_vs_tg":
                result = run_dcf(ticker_symbol, starting_growth_rate=starting_growth_rate,
                                  ending_growth_rate=ending_growth_rate,
                                  starting_fcf_margin=starting_fcf_margin,
                                  ending_fcf_margin=ending_fcf_margin,
                                  wacc=row_val, terminal_growth_rate=col_val)
            else:
                result = run_dcf(ticker_symbol, starting_growth_rate=col_val,
                                  ending_growth_rate=ending_growth_rate,
                                  starting_fcf_margin=starting_fcf_margin,
                                  ending_fcf_margin=ending_fcf_margin,
                                  wacc=row_val, terminal_growth_rate=terminal_growth_rate)
            table[row_val][col_val] = result["implied_share_price"]
    return table


# ============================================
# COMPS PLATFORM (shared infra + Excel output)
# ============================================

INDUSTRY_UNIVERSE = {
    # --- Technology ---
    "Software - Infrastructure": ["ZS", "CRWD", "PANW", "FTNT", "OKTA", "NET"],
    "Software - Application": ["CRM", "NOW", "WDAY", "ADBE", "INTU", "TEAM"],
    "Semiconductors": ["NVDA", "AMD", "AVGO", "TXN", "QCOM", "MU"],
    "Semiconductor Equipment & Materials": ["ASML", "AMAT", "LRCX", "KLAC"],
    "Consumer Electronics": ["AAPL", "SONY", "HPQ", "DELL"],
    "Internet Content & Information": ["GOOGL", "META", "PINS", "SNAP"],
    "Internet Retail": ["AMZN", "EBAY", "ETSY", "MELI", "JD"],
    "Information Technology Services": ["ACN", "IBM", "INFY", "EPAM"],
    "Communication Equipment": ["CSCO", "ANET", "JNPR", "NOK", "ERIC"],
    "Electronic Components": ["APH", "TEL", "GLW", "JBL"],
    "Computer Hardware": ["DELL", "HPQ", "SMCI", "NTAP"],

    # --- Communication Services ---
    "Entertainment": ["DIS", "NFLX", "WBD", "PARA"],
    "Telecom Services": ["T", "VZ", "TMUS", "CHTR"],

    # --- Consumer Cyclical ---
    "Auto Manufacturers": ["TM", "F", "GM", "HMC", "STLA"],
    "Specialty Retail": ["HD", "LOW", "TJX", "ROST"],
    "Discount Stores": ["WMT", "TGT", "COST", "DG"],
    "Restaurants": ["MCD", "SBUX", "CMG", "YUM", "DPZ"],
    "Apparel Retail": ["TJX", "ROST", "GPS", "ANF"],
    "Footwear & Accessories": ["NKE", "DECK", "SKX", "CROX"],
    "Travel Services": ["BKNG", "EXPE", "TCOM", "ABNB"],
    "Residential Construction": ["DHI", "LEN", "PHM", "NVR"],

    # --- Consumer Defensive ---
    "Beverages - Non-Alcoholic": ["KO", "PEP", "KDP", "MNST"],
    "Beverages - Alcoholic": ["BUD", "STZ", "DEO", "TAP"],
    "Packaged Foods": ["GIS", "KHC", "MDLZ", "HSY", "K"],
    "Household & Personal Products": ["PG", "CL", "KMB", "CHD"],
    "Grocery Stores": ["KR", "ACI", "SFM"],

    # --- Healthcare ---
    "Drug Manufacturers - General": ["JNJ", "PFE", "MRK", "ABBV", "LLY"],
    "Biotechnology": ["AMGN", "GILD", "VRTX", "REGN", "BIIB"],
    "Medical Devices": ["MDT", "SYK", "BSX", "ZBH", "EW"],
    "Medical Instruments & Supplies": ["ISRG", "IDXX", "RMD"],
    "Health Insurance Plans": ["UNH", "ELV", "CVS", "CI", "HUM"],
    "Diagnostics & Research": ["TMO", "DHR", "A", "IQV"],

    # --- Financial Services ---
    "Banks - Diversified": ["JPM", "BAC", "C", "WFC"],
    "Banks - Regional": ["USB", "PNC", "TFC", "FITB"],
    "Insurance - Diversified": ["BRK-B", "AIG", "TRV"],
    "Insurance - Life": ["MET", "PRU", "AFL"],
    "Asset Management": ["BLK", "BX", "KKR", "APO"],
    "Credit Services": ["V", "MA", "AXP", "PYPL", "COF"],
    "Capital Markets": ["GS", "MS", "SCHW"],

    # --- Industrials ---
    "Aerospace & Defense": ["BA", "LMT", "RTX", "NOC", "GD"],
    "Farm & Heavy Construction Machinery": ["CAT", "DE", "CNH"],
    "Specialty Industrial Machinery": ["HON", "ETN", "EMR", "ITW"],
    "Railroads": ["UNP", "CSX", "NSC"],
    "Airlines": ["DAL", "UAL", "AAL", "LUV"],
    "Trucking": ["ODFL", "JBHT", "CHRW"],
    "Waste Management": ["WM", "RSG", "WCN"],
    "Building Products & Equipment": ["CARR", "JCI", "AOS"],

    # --- Energy ---
    "Oil & Gas Integrated": ["XOM", "CVX", "SHEL", "BP"],
    "Oil & Gas E&P": ["EOG", "COP", "PXD", "DVN"],
    "Oil & Gas Midstream": ["KMI", "WMB", "OKE"],
    "Oil & Gas Equipment & Services": ["SLB", "HAL", "BKR"],

    # --- Basic Materials ---
    "Chemicals": ["LIN", "APD", "SHW", "ECL"],
    "Steel": ["NUE", "STLD", "X"],
    "Gold": ["NEM", "GOLD", "AEM"],
    "Copper": ["FCX", "SCCO"],
    "Specialty Chemicals": ["DD", "ALB", "PPG"],

    # --- Real Estate ---
    "REIT - Residential": ["AVB", "EQR", "ESS", "MAA"],
    "REIT - Retail": ["SPG", "O", "REG"],
    "REIT - Industrial": ["PLD", "PSA"],
    "REIT - Office": ["BXP", "VNO"],

    # --- Utilities ---
    "Utilities - Regulated Electric": ["NEE", "DUK", "SO", "D"],
    "Utilities - Diversified": ["EXC", "AEP", "SRE"],
}

# Broader sector-level fallback, used only when a company's exact industry
# isn't in INDUSTRY_UNIVERSE above. Peers within the same sector are a
# weaker match than same-industry peers (e.g. a bank and an insurer are
# both "Financial Services" but not true comps), so this is a safety net
# for coverage, not a substitute for adding the real industry when possible.
SECTOR_FALLBACK_UNIVERSE = {
    "Technology": ["AAPL", "MSFT", "GOOGL", "NVDA", "CRM"],
    "Communication Services": ["GOOGL", "META", "DIS", "T", "VZ"],
    "Consumer Cyclical": ["AMZN", "HD", "MCD", "NKE", "TM"],
    "Consumer Defensive": ["WMT", "PG", "KO", "COST", "PEP"],
    "Healthcare": ["JNJ", "UNH", "PFE", "ABBV", "MRK"],
    "Financial Services": ["JPM", "BAC", "V", "MA", "BRK-B"],
    "Industrials": ["HON", "UNP", "CAT", "RTX", "BA"],
    "Energy": ["XOM", "CVX", "COP", "SLB"],
    "Basic Materials": ["LIN", "SHW", "NEM", "FCX"],
    "Real Estate": ["PLD", "AMT", "EQIX", "SPG"],
    "Utilities": ["NEE", "DUK", "SO", "D"],
}


def get_peer_tickers_for(industry, sector, exclude_ticker=None, max_peers=5):
    """
    Tries exact industry match first (best comps), then falls back to the
    broader sector bucket if the industry isn't covered, so a company in an
    industry we haven't explicitly listed still gets *some* peers instead
    of none. Returns (peer_list, match_quality) where match_quality is
    "industry" or "sector" so callers can note which kind of match was used.
    """
    if industry in INDUSTRY_UNIVERSE:
        peers = [t for t in INDUSTRY_UNIVERSE[industry] if t != exclude_ticker]
        return peers[:max_peers], "industry"

    if sector in SECTOR_FALLBACK_UNIVERSE:
        peers = [t for t in SECTOR_FALLBACK_UNIVERSE[sector] if t != exclude_ticker]
        return peers[:max_peers], "sector"

    return [], "none"


def find_peers(ticker_symbol, max_peers=5):
    company = resolve_company(ticker_symbol)
    industry = company["industry"]
    sector = company["sector"]

    peers, match_quality = get_peer_tickers_for(industry, sector, exclude_ticker=ticker_symbol, max_peers=max_peers)

    if match_quality == "none":
        print(f"No peer list found for industry '{industry}' or sector '{sector}' (ticker: {ticker_symbol}).")
    elif match_quality == "sector":
        print(f"No exact peer list for industry '{industry}' — using broader '{sector}' sector peers instead "
              f"(weaker comps than same-industry, but better than none).")

    return peers


def get_company_multiples(ticker_symbol):
    """
    Every field here is optional-safe: non-US filers frequently lack one
    or more fields on yfinance. Missing data returns None rather than
    crashing the whole comps run. EBIT is approximated from yfinance's
    operatingMargins (a real, reported figure) since a clean EBIT line
    isn't consistently available across filers — flagged as an estimate
    via the field name.
    """
    try:
        company = resolve_company(ticker_symbol)
    except ValueError:
        return None

    info = company["info"]
    market_cap = info.get("marketCap")
    total_debt = info.get("totalDebt", 0) or 0
    cash = info.get("totalCash", 0) or 0
    enterprise_value = info.get("enterpriseValue")
    if enterprise_value is None and market_cap is not None:
        enterprise_value = market_cap + total_debt - cash

    revenue = info.get("totalRevenue")
    ebitda = info.get("ebitda")
    pe_ratio = info.get("trailingPE")
    revenue_growth = info.get("revenueGrowth")
    ebitda_margin = info.get("ebitdaMargins")
    operating_margin = info.get("operatingMargins")  # used as an EBIT-margin proxy
    fcf = info.get("freeCashflow")
    net_income = info.get("netIncomeToCommon")
    shares_outstanding = info.get("sharesOutstanding")

    ebit_est = operating_margin * revenue if (operating_margin is not None and revenue) else None
    fcf_margin = fcf / revenue if (fcf and revenue) else None
    net_leverage = (total_debt - cash) / ebitda if (ebitda and ebitda > 0) else None
    fcf_yield = fcf / market_cap if (fcf and market_cap) else None
    rule_of_40 = (revenue_growth + fcf_margin) if (revenue_growth is not None and fcf_margin is not None) else None

    ev_revenue = enterprise_value / revenue if (enterprise_value and revenue) else None
    ev_ebitda = enterprise_value / ebitda if (enterprise_value and ebitda and ebitda > 0) else None
    ev_ebit = enterprise_value / ebit_est if (enterprise_value and ebit_est and ebit_est > 0) else None

    return {
        "ticker": ticker_symbol,
        "company_name": company["company_name"],
        "market_cap": market_cap,
        "enterprise_value": enterprise_value,
        "total_debt": total_debt,
        "cash": cash,
        "revenue": revenue,
        "ebitda": ebitda,
        "ebit_est": ebit_est,
        "net_income": net_income,
        "fcf": fcf,
        "shares_outstanding": shares_outstanding,
        "pe_ratio": pe_ratio,
        "ev_revenue": ev_revenue,
        "ev_ebitda": ev_ebitda,
        "ev_ebit": ev_ebit,
        "fcf_yield": fcf_yield,
        "revenue_growth": revenue_growth,
        "ebitda_margin": ebitda_margin,
        "operating_margin": operating_margin,
        "fcf_margin": fcf_margin,
        "net_leverage": net_leverage,
        "rule_of_40": rule_of_40,
    }


def _percentile(sorted_values, pct):
    """Linear-interpolation percentile (Excel's PERCENTILE.INC method), so
    the 25th/median/75th bands match what a banker would get typing
    =PERCENTILE() in Excel over the same peer set."""
    if not sorted_values:
        return None
    if len(sorted_values) == 1:
        return sorted_values[0]
    k = (len(sorted_values) - 1) * pct
    f, c = int(k), min(int(k) + 1, len(sorted_values) - 1)
    if f == c:
        return sorted_values[f]
    return sorted_values[f] + (sorted_values[c] - sorted_values[f]) * (k - f)


def _band(values, min_val=None, max_val=None):
    """Returns {p25, median, mean, p75} for a metric across peers, after
    the same outlier filtering used elsewhere (drops negative/extreme
    multiples that would otherwise distort a small peer set)."""
    clean = sorted(v for v in values if v is not None)
    if min_val is not None:
        clean = [v for v in clean if v > min_val]
    if max_val is not None:
        clean = [v for v in clean if v < max_val]
    if not clean:
        return {"p25": None, "median": None, "mean": None, "p75": None}
    return {
        "p25": _percentile(clean, 0.25),
        "median": _percentile(clean, 0.50),
        "mean": sum(clean) / len(clean),
        "p75": _percentile(clean, 0.75),
    }


def build_comps_table(ticker_symbol, max_peers=5):
    """
    Target + peers, each with a full operating/valuation profile, plus
    25th/median/mean/75th percentile bands per multiple (not just a
    single average) and an implied EV -> equity value -> per-share price
    at each band, for each multiple — matching a real comps book, not
    just a flat average table.
    """
    peers = find_peers(ticker_symbol, max_peers=max_peers)

    target_data = get_company_multiples(ticker_symbol)
    peer_data = [p for p in (get_company_multiples(t) for t in peers) if p is not None]

    ev_revenue_band = _band([p["ev_revenue"] for p in peer_data])
    ev_ebitda_band = _band([p["ev_ebitda"] for p in peer_data], min_val=0, max_val=60)
    ev_ebit_band = _band([p["ev_ebit"] for p in peer_data], min_val=0, max_val=80)
    pe_band = _band([p["pe_ratio"] for p in peer_data], min_val=0, max_val=100)

    # Backward-compatible single-average fields (existing callers still use these)
    avg_ev_revenue = ev_revenue_band["mean"]
    avg_ev_ebitda = ev_ebitda_band["mean"]
    avg_pe = pe_band["mean"]

    implied_ev_from_revenue = (
        avg_ev_revenue * target_data["revenue"]
        if (avg_ev_revenue and target_data and target_data["revenue"]) else None
    )
    target_ebitda = target_data.get("ebitda") if target_data else None
    implied_ev_from_ebitda = (
        avg_ev_ebitda * target_ebitda
        if (avg_ev_ebitda and target_ebitda and target_ebitda > 0) else None
    )

    # --- Implied share price at each percentile band, for each multiple ---
    net_debt = (target_data["total_debt"] - target_data["cash"]) if target_data else None

    def implied_price_band(band, metric_value):
        """EV band * target metric -> EV -> equity value -> nothing (no
        share count on this data path) -- returns implied EQUITY VALUE
        band; callers with a share count convert to per-share themselves."""
        if not metric_value or net_debt is None:
            return {"p25": None, "median": None, "mean": None, "p75": None}
        result = {}
        for key, mult in band.items():
            if mult is None:
                result[key] = None
            else:
                result[key] = mult * metric_value - net_debt
        return result

    implied_equity_from_revenue = implied_price_band(ev_revenue_band, target_data["revenue"] if target_data else None)
    implied_equity_from_ebitda = implied_price_band(ev_ebitda_band, target_ebitda)

    return {
        "target": target_data,
        "peers": peer_data,
        "avg_ev_revenue": avg_ev_revenue,
        "avg_ev_ebitda": avg_ev_ebitda,
        "avg_pe": avg_pe,
        "ev_revenue_band": ev_revenue_band,
        "ev_ebitda_band": ev_ebitda_band,
        "ev_ebit_band": ev_ebit_band,
        "pe_band": pe_band,
        "implied_ev_from_revenue": implied_ev_from_revenue,
        "implied_ev_from_ebitda": implied_ev_from_ebitda,
        "implied_equity_from_revenue": implied_equity_from_revenue,
        "implied_equity_from_ebitda": implied_equity_from_ebitda,
        "net_debt": net_debt,
    }


# ============================================
# 3-STATEMENT MODEL
# Full linked Income Statement / Balance Sheet / Cash Flow projection.
# More granular than the DCF (which only projects revenue -> FCF directly):
# this breaks revenue down into margins and working-capital drivers, so
# you can see exactly which line item assumption moves the outcome.
# Cash is the balancing "plug" — debt and dividend policy are held at
# simple, overridable assumptions rather than solved simultaneously with
# cash, which keeps the model transparent at the cost of some realism
# (no revolver draw-down/paydown logic).
# ============================================

GROSS_PROFIT_ROW_NAMES = ["Gross Profit"]
SGA_ROW_NAMES = ["Selling General And Administration", "SellingGeneralAndAdministration"]
TAX_PROVISION_ROW_NAMES = ["Tax Provision"]
PRETAX_INCOME_ROW_NAMES = ["Pretax Income"]
INTEREST_EXPENSE_ROW_NAMES = ["Interest Expense", "Interest Expense Non Operating"]
AR_ROW_NAMES = ["Accounts Receivable", "Receivables", "Net Receivables"]
INVENTORY_ROW_NAMES = ["Inventory"]
AP_ROW_NAMES = ["Accounts Payable", "Payables And Accrued Expenses", "Payables"]
PPE_ROW_NAMES = ["Net PPE", "Properties Plant And Equipment Net"]
EQUITY_ROW_NAMES = ["Stockholders Equity", "Common Stock Equity", "Total Equity Gross Minority Interest"]
DIVIDENDS_PAID_ROW_NAMES = ["Cash Dividends Paid", "Common Stock Dividend Paid"]


# ============================================
# SEC EDGAR DATA ENRICHMENT LAYER
# A real fallback source for US-listed companies, used when yfinance's
# line-item detail is missing or blended. NOT a replacement for yfinance
# (which stays the backbone for market data, non-US coverage, and
# everything already working) — this only fills specific documented gaps:
# more granular income-statement line items, and segment revenue where a
# company happens to tag it. Every function here fails gracefully (returns
# None) rather than raising, since this is meant to be an optional
# enrichment layer a caller can silently fall back past.
# ============================================
from urllib.request import Request as _SecRequest, urlopen as _sec_urlopen

SEC_EDGAR_USER_AGENT = "MigelDealTool lobomigel08@gmail.com"  # SEC requires a real contact string on every request
_SEC_TICKER_TO_CIK_CACHE = None


def _sec_headers():
    return {"User-Agent": SEC_EDGAR_USER_AGENT, "Accept-Encoding": "gzip, deflate"}


def fetch_sec_cik_for_ticker(ticker_symbol):
    """
    Looks up a company's SEC CIK (Central Index Key) from SEC's free,
    single-file ticker-to-CIK mapping. Cached in-process since this file
    covers every US-listed ticker at once — no need to re-fetch per call.
    Returns None (not an exception) if the ticker isn't found, e.g. for
    non-US-listed companies, which simply aren't in SEC's system.
    """
    global _SEC_TICKER_TO_CIK_CACHE
    if _SEC_TICKER_TO_CIK_CACHE is None:
        try:
            req = _SecRequest("https://www.sec.gov/files/company_tickers.json", headers=_sec_headers())
            with _sec_urlopen(req, timeout=10) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            _SEC_TICKER_TO_CIK_CACHE = {
                entry["ticker"].upper(): str(entry["cik_str"]).zfill(10) for entry in raw.values()
            }
        except Exception:
            _SEC_TICKER_TO_CIK_CACHE = {}
    return _SEC_TICKER_TO_CIK_CACHE.get(ticker_symbol.upper())


def fetch_sec_company_facts(cik):
    """
    Fetches the full XBRL "company facts" for a given 10-digit CIK — every
    line item that company has tagged in its filings, across all periods.
    Returns None on any failure (network, invalid CIK, no data) rather
    than raising, since callers treat this as an optional enrichment.
    """
    if not cik:
        return None
    try:
        url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
        req = _SecRequest(url, headers=_sec_headers())
        with _sec_urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def extract_sec_line_item(facts, concept_names, unit="USD"):
    """
    Pulls the most recent ANNUAL (10-K, form='10-K') value for a line item
    from SEC company-facts data, trying each candidate XBRL concept name
    in order — companies tag the same real-world line item under
    different standard names (e.g. CostOfRevenue vs CostOfGoodsSold).
    Returns None if no candidate concept is found, rather than guessing.
    """
    if not facts or "facts" not in facts or "us-gaap" not in facts["facts"]:
        return None
    gaap = facts["facts"]["us-gaap"]
    for concept in concept_names:
        if concept not in gaap:
            continue
        units = gaap[concept].get("units", {})
        values = units.get(unit, [])
        annual_values = [v for v in values if v.get("form") == "10-K" and v.get("fy") and v.get("val") is not None]
        if not annual_values:
            continue
        annual_values.sort(key=lambda v: (v.get("fy", 0), v.get("end", "")), reverse=True)
        return annual_values[0]["val"]
    return None


# Common XBRL concept name variants for each line item — companies tag
# the same economic concept under different standard names.
SEC_COGS_CONCEPTS = ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold"]
SEC_SGA_CONCEPTS = ["SellingGeneralAndAdministrativeExpense", "GeneralAndAdministrativeExpense"]
SEC_REVENUE_CONCEPTS = ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet"]


def enrich_operating_assumptions_from_sec(ticker_symbol, revenue, existing_gross_margin, existing_sga_pct_revenue):
    """
    The actual fallback logic: only queries SEC EDGAR for whatever's
    genuinely MISSING from yfinance's data (existing_* args are None only
    when yfinance couldn't supply them) — never overrides a real yfinance
    value. Returns (gross_margin, sga_pct_revenue, source_notes) where
    source_notes flags which figures came from SEC vs stayed as generic
    fallbacks, so callers can be transparent about data provenance.
    """
    if existing_gross_margin is not None and existing_sga_pct_revenue is not None:
        return existing_gross_margin, existing_sga_pct_revenue, []  # yfinance already had both; no need to call SEC at all

    cik = fetch_sec_cik_for_ticker(ticker_symbol)
    if not cik:
        return existing_gross_margin, existing_sga_pct_revenue, []  # not a US-listed company in SEC's system

    facts = fetch_sec_company_facts(cik)
    if not facts:
        return existing_gross_margin, existing_sga_pct_revenue, []

    source_notes = []
    gross_margin = existing_gross_margin
    sga_pct_revenue = existing_sga_pct_revenue

    if gross_margin is None:
        cogs = extract_sec_line_item(facts, SEC_COGS_CONCEPTS)
        sec_revenue = extract_sec_line_item(facts, SEC_REVENUE_CONCEPTS) or revenue
        if cogs is not None and sec_revenue:
            gross_margin = (sec_revenue - cogs) / sec_revenue
            source_notes.append("Gross margin sourced from SEC EDGAR (yfinance didn't provide it)")

    if sga_pct_revenue is None:
        sga = extract_sec_line_item(facts, SEC_SGA_CONCEPTS)
        sec_revenue = extract_sec_line_item(facts, SEC_REVENUE_CONCEPTS) or revenue
        if sga is not None and sec_revenue:
            sga_pct_revenue = sga / sec_revenue
            source_notes.append("SG&A % sourced from SEC EDGAR (yfinance didn't provide it)")

    return gross_margin, sga_pct_revenue, source_notes



# ============================================
# WIKIPEDIA / WIKIDATA ENRICHMENT
# Free, no API key, no rate limit that matters for this tool's usage.
# Best for qualitative content (logos, short descriptions) — NOT financial
# data. This directly targets a real, visible gap: most companies fall
# back to the pitchbook monogram because yfinance doesn't reliably expose
# a logo URL. Best-effort only — a title mismatch or missing page just
# means no logo/description, never an error.
# ============================================
from urllib.request import Request as _WikiRequest, urlopen as _wiki_urlopen
from urllib.parse import quote as _wiki_quote


def fetch_wikipedia_summary(company_name):
    """Returns {"logo_url": ..., "description": ...} with either field
    possibly None, or None entirely if the page lookup fails outright."""
    try:
        title = _wiki_quote(company_name.replace(" ", "_"))
        url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{title}"
        req = _WikiRequest(url, headers={"User-Agent": "MigelDealTool/1.0 (lobomigel08@gmail.com)"})
        with _wiki_urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return {
            "logo_url": (data.get("thumbnail") or {}).get("source"),
            "description": data.get("extract"),
        }
    except Exception:
        return None


# ============================================
# FINNHUB ENRICHMENT
# Free tier (60 calls/min — generous compared to Alpha Vantage), but
# requires a free API key: set FINNHUB_API_KEY as an environment variable.
# Skips gracefully (returns None) if the key isn't set, exactly like this
# tool's existing EMAIL_ADDRESS/EMAIL_APP_PASSWORD pattern for optional
# features. Used here for: company logo (a real, correctly-labeled field
# in their profile endpoint) and analyst price targets.
# ============================================
from urllib.request import Request as _FinnhubRequest, urlopen as _finnhub_urlopen


def _finnhub_api_key():
    return os.environ.get("FINNHUB_API_KEY")


def fetch_finnhub_profile(ticker_symbol):
    """Company profile including a real logo URL, market cap, industry.
    Returns None if no API key is set or the request fails."""
    api_key = _finnhub_api_key()
    if not api_key:
        return None
    try:
        url = f"https://finnhub.io/api/v1/stock/profile2?symbol={ticker_symbol}&token={api_key}"
        req = _FinnhubRequest(url)
        with _finnhub_urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data if data else None  # Finnhub returns {} for unknown tickers, not an error
    except Exception:
        return None


def fetch_finnhub_price_target(ticker_symbol):
    """Analyst price target consensus. Returns None if unavailable."""
    api_key = _finnhub_api_key()
    if not api_key:
        return None
    try:
        url = f"https://finnhub.io/api/v1/stock/price-target?symbol={ticker_symbol}&token={api_key}"
        req = _FinnhubRequest(url)
        with _finnhub_urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data if data and data.get("targetMean") else None
    except Exception:
        return None


# ============================================
# SIMFIN ENRICHMENT
# Free tier requires a free API key (SIMFIN_API_KEY env var). Standardized
# line-item naming across companies (a real advantage over raw XBRL tags),
# and — unlike SEC EDGAR — covers companies outside the US, which is the
# specific gap this fills in the fallback waterfall below.
#
# HONEST CAVEAT: SimFin's exact API contract may have evolved since this
# was written and can't be verified against a live call from this
# environment (no general network access here — the same constraint
# noted for every other data source in this build). Every field access
# below is defensive (try/except, graceful None), so if SimFin has changed
# their schema, this fails gracefully rather than crashing — but the
# specific endpoint/field names here should be treated as "best effort
# based on their documented v2/v3 API," not verified against a live
# response. Test with your own key before relying on it.
# ============================================
from urllib.request import Request as _SimfinRequest, urlopen as _simfin_urlopen


def _simfin_api_key():
    return os.environ.get("SIMFIN_API_KEY")


def fetch_simfin_income_statement(ticker_symbol):
    """Most recent annual income statement in SimFin's standardized format.
    Returns None if no API key, the ticker isn't covered, or the request
    fails for any reason."""
    api_key = _simfin_api_key()
    if not api_key:
        return None
    try:
        url = f"https://backend.simfin.com/api/v3/companies/statements/compact?ticker={ticker_symbol}&statements=pl&period=fy"
        req = _SimfinRequest(url, headers={"Authorization": api_key})
        with _simfin_urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def extract_simfin_line_items(statement_data, revenue_fallback=None):
    """
    Defensively parses SimFin's compact statement format into
    (gross_margin, sga_pct_revenue) or (None, None) if the data doesn't
    parse as expected — see the module-level caveat above on schema
    uncertainty. Every step is wrapped so an unexpected shape degrades to
    "no data" rather than raising.
    """
    try:
        if not statement_data or not isinstance(statement_data, list):
            return None, None
        company_data = statement_data[0]
        columns = company_data.get("columns", [])
        rows = company_data.get("data", [])
        if not rows:
            return None, None
        latest_row = rows[0]  # most recent period first, per SimFin's documented ordering
        row_dict = dict(zip(columns, latest_row))

        revenue = row_dict.get("Revenue") or revenue_fallback
        cogs = row_dict.get("Cost of Revenue")
        sga = row_dict.get("Selling, General & Administrative")

        gross_margin = None
        if revenue and cogs is not None:
            gross_margin = (revenue - abs(cogs)) / revenue

        sga_pct_revenue = None
        if revenue and sga is not None:
            sga_pct_revenue = abs(sga) / revenue

        return gross_margin, sga_pct_revenue
    except Exception:
        return None, None


# ============================================
# COMBINED ENRICHMENT WATERFALL
# Priority: yfinance (already fetched, free, broadest coverage) -> SEC
# EDGAR (US-listed, primary-source, free, no key) -> SimFin (any market,
# needs a free key) -> generic default. Each step only runs if the
# previous ones left something genuinely missing.
# ============================================

def enrich_operating_assumptions(ticker_symbol, revenue, existing_gross_margin, existing_sga_pct_revenue):
    """
    The full multi-source fallback: tries SEC EDGAR first (free, no key,
    but US-listed only), then SimFin (needs a key, but covers more
    markets), collecting source_notes from whichever source actually
    supplied something. Never overrides a real yfinance value.
    """
    if existing_gross_margin is not None and existing_sga_pct_revenue is not None:
        return existing_gross_margin, existing_sga_pct_revenue, []

    gross_margin, sga_pct_revenue, source_notes = enrich_operating_assumptions_from_sec(
        ticker_symbol, revenue, existing_gross_margin, existing_sga_pct_revenue)

    if gross_margin is None or sga_pct_revenue is None:
        statement = fetch_simfin_income_statement(ticker_symbol)
        if statement:
            sf_gross_margin, sf_sga_pct = extract_simfin_line_items(statement, revenue_fallback=revenue)
            if gross_margin is None and sf_gross_margin is not None:
                gross_margin = sf_gross_margin
                source_notes.append("Gross margin sourced from SimFin (yfinance and SEC EDGAR didn't provide it)")
            if sga_pct_revenue is None and sf_sga_pct is not None:
                sga_pct_revenue = sf_sga_pct
                source_notes.append("SG&A % sourced from SimFin (yfinance and SEC EDGAR didn't provide it)")

    return gross_margin, sga_pct_revenue, source_notes


def fetch_best_available_logo(info, ticker_symbol, company_name):
    """
    Tries, in order: yfinance's own logo field (if present) -> Wikipedia
    (free, no key) -> Finnhub (needs a key, but often has a cleaner logo).
    Returns a logo URL string, or None if nothing was found anywhere —
    callers should fall back to the monogram in that case.
    """
    logo_url = info.get("logo_url") or info.get("logoUrl") if info else None
    if logo_url:
        return logo_url

    wiki = fetch_wikipedia_summary(company_name)
    if wiki and wiki.get("logo_url"):
        return wiki["logo_url"]

    finnhub_profile = fetch_finnhub_profile(ticker_symbol)
    if finnhub_profile and finnhub_profile.get("logo"):
        return finnhub_profile["logo"]

    return None



def estimate_three_statement_assumptions(ticker_obj, info):
    """
    Same philosophy as estimate_dcf_assumptions(): derive every driver from
    the company's own most recent financials rather than one hardcoded
    profile, so the model describes THIS company. Ratios not found default
    to 0 (e.g. a services company with no real inventory) rather than
    crashing — a missing line item just means that assumption has no effect.
    """
    financials = ticker_obj.financials
    balance_sheet = ticker_obj.balance_sheet
    cashflow = ticker_obj.cashflow

    revenue = safe_get_row(financials, REVENUE_ROW_NAMES, default=0) or 0
    gross_profit = safe_get_row(financials, GROSS_PROFIT_ROW_NAMES, default=None)
    sga_raw = safe_get_row(financials, SGA_ROW_NAMES, default=None)
    pretax_income = safe_get_row(financials, PRETAX_INCOME_ROW_NAMES, default=None)
    tax_provision = safe_get_row(financials, TAX_PROVISION_ROW_NAMES, default=None)

    ocf = safe_get_row(cashflow, OCF_ROW_NAMES, default=0) or 0
    capex = safe_get_row(cashflow, CAPEX_ROW_NAMES, default=0) or 0
    da = safe_get_row(cashflow, ["Depreciation And Amortization", "Depreciation Amortization Depletion"], default=0) or 0
    dividends_paid = abs(safe_get_row(cashflow, DIVIDENDS_PAID_ROW_NAMES, default=0) or 0)
    historical_debt_repayment = abs(safe_get_row(cashflow, ["Repayment Of Debt"], default=0) or 0)

    ar = safe_get_row(balance_sheet, AR_ROW_NAMES, default=0) or 0
    inventory = safe_get_row(balance_sheet, INVENTORY_ROW_NAMES, default=0) or 0
    ap = safe_get_row(balance_sheet, AP_ROW_NAMES, default=0) or 0
    total_debt = safe_get_row(balance_sheet, DEBT_ROW_NAMES, default=0) or 0

    # yfinance first; SEC EDGAR only as a fallback for whatever's genuinely
    # missing (never overrides a real yfinance value); a generic default
    # only if ALL real sources came up empty.
    gross_margin_from_yf = (gross_profit / revenue) if (gross_profit is not None and revenue) else None
    sga_pct_from_yf = (sga_raw / revenue) if (sga_raw is not None and revenue) else None
    ticker_symbol_for_sec = getattr(ticker_obj, "ticker", None)
    if ticker_symbol_for_sec and (gross_margin_from_yf is None or sga_pct_from_yf is None):
        gross_margin, sga_pct_revenue, sec_source_notes = enrich_operating_assumptions(
            ticker_symbol_for_sec, revenue, gross_margin_from_yf, sga_pct_from_yf)
    else:
        gross_margin, sga_pct_revenue, sec_source_notes = gross_margin_from_yf, sga_pct_from_yf, []
    gross_margin = gross_margin if gross_margin is not None else 0.40
    sga_pct_revenue = sga_pct_revenue if sga_pct_revenue is not None else 0.20
    da_pct_revenue = (da / revenue) if revenue else 0.03
    capex_pct_revenue = (abs(capex) / revenue) if revenue else 0.03

    if pretax_income and tax_provision is not None and pretax_income > 0:
        tax_rate = max(min(tax_provision / pretax_income, 0.35), 0.05)
    else:
        tax_rate = 0.21  # generic corporate-rate fallback when pretax income is negative/unusable

    # A flat rate applied to the AVERAGE of beginning/ending debt each year
    # (not just beginning balance) — see run_three_statement_model()'s
    # iterative solve for why average-balance interest needs iteration.
    interest_rate_on_debt = 0.05
    ar_pct_revenue = (ar / revenue) if revenue else 0
    inventory_pct_revenue = (inventory / revenue) if revenue else 0
    ap_pct_revenue = (ap / revenue) if revenue else 0
    net_income_approx = pretax_income - tax_provision if (pretax_income is not None and tax_provision is not None) else None
    dividend_payout_ratio = (dividends_paid / net_income_approx) if (net_income_approx and net_income_approx > 0) else 0

    # Derived from what the company actually repaid last year, as a % of
    # its debt balance — a real, if rough, estimate of scheduled
    # amortization, clamped so a one-off large repayment doesn't imply the
    # company retires all its debt in ~2 years.
    mandatory_amortization_pct = (historical_debt_repayment / total_debt) if total_debt else 0
    mandatory_amortization_pct = max(min(mandatory_amortization_pct, 0.25), 0)

    return {
        "gross_margin": gross_margin,
        "sga_pct_revenue": sga_pct_revenue,
        "da_pct_revenue": da_pct_revenue,
        "capex_pct_revenue": capex_pct_revenue,
        "tax_rate": tax_rate,
        "interest_rate_on_debt": interest_rate_on_debt,
        "ar_pct_revenue": ar_pct_revenue,
        "inventory_pct_revenue": inventory_pct_revenue,
        "ap_pct_revenue": ap_pct_revenue,
        "dividend_payout_ratio": dividend_payout_ratio,
        "mandatory_amortization_pct": mandatory_amortization_pct,
        "sec_source_notes": sec_source_notes,
    }


def run_three_statement_model(ticker_symbol, years_to_project=5, starting_growth_rate=None,
                               ending_growth_rate=None, gross_margin=None, sga_pct_revenue=None,
                               da_pct_revenue=None, capex_pct_revenue=None, tax_rate=None,
                               interest_rate_on_debt=None, ar_pct_revenue=None,
                               inventory_pct_revenue=None, ap_pct_revenue=None,
                               dividend_payout_ratio=None, mandatory_amortization_pct=None,
                               min_cash_balance=0, cash_sweep_pct=0.0):
    """
    Projects a full, BALANCING Income Statement, Balance Sheet, and Cash
    Flow Statement for years_to_project years.

    Two things make this "professional-grade" rather than a rough
    approximation:

    1. Balance sheet plug: everything not explicitly modeled (goodwill,
       deferred taxes, minority interest, other current/long-term assets
       and liabilities) is netted into one "Other Net Assets" figure,
       calculated once from the real starting balance sheet and held flat.
       Because every projected change flows through the income statement
       and cash flow statement consistently (double-entry: every debit has
       a matching credit), Total Assets equals Total Liabilities + Equity
       in EVERY projected year — verified explicitly via a Balance Check
       row in the Excel export, not just assumed.

    2. Debt schedule with circularity solved by iteration: interest expense
       depends on the average of beginning/ending debt, but ending debt
       (via any revolver draw or cash sweep) depends on net income, which
       depends on interest expense — a circular reference. Real models hit
       this by either enabling iterative calculation in Excel or accepting
       a beginning-balance approximation. Here we just iterate the
       calculation a fixed number of times per year, which converges
       almost immediately and gives an exact answer rather than an
       approximation.

    A revolver automatically draws if cash would otherwise go below
    min_cash_balance (default 0, i.e. just prevents negative cash), and
    cash_sweep_pct (0 by default — set higher to model a company
    prioritizing debt paydown) sweeps that fraction of any cash above
    min_cash_balance toward paying down debt early, after mandatory
    amortization.
    """
    company = resolve_company(ticker_symbol)
    ticker = company["ticker_obj"]
    info = company["info"]

    financials = ticker.financials
    balance_sheet = ticker.balance_sheet

    most_recent_revenue = safe_get_row(financials, REVENUE_ROW_NAMES)
    if most_recent_revenue is None:
        raise ValueError(f"No revenue data found for '{ticker_symbol}'.")

    dcf_assumptions = estimate_dcf_assumptions(ticker, info)
    three_stmt_assumptions = estimate_three_statement_assumptions(ticker, info)

    if starting_growth_rate is None:
        starting_growth_rate = dcf_assumptions["starting_growth_rate"]
    if ending_growth_rate is None:
        ending_growth_rate = dcf_assumptions["ending_growth_rate"]
    if gross_margin is None:
        gross_margin = three_stmt_assumptions["gross_margin"]
    if sga_pct_revenue is None:
        sga_pct_revenue = three_stmt_assumptions["sga_pct_revenue"]
    if da_pct_revenue is None:
        da_pct_revenue = three_stmt_assumptions["da_pct_revenue"]
    if capex_pct_revenue is None:
        capex_pct_revenue = three_stmt_assumptions["capex_pct_revenue"]
    if tax_rate is None:
        tax_rate = three_stmt_assumptions["tax_rate"]
    if interest_rate_on_debt is None:
        interest_rate_on_debt = three_stmt_assumptions["interest_rate_on_debt"]
    if ar_pct_revenue is None:
        ar_pct_revenue = three_stmt_assumptions["ar_pct_revenue"]
    if inventory_pct_revenue is None:
        inventory_pct_revenue = three_stmt_assumptions["inventory_pct_revenue"]
    if ap_pct_revenue is None:
        ap_pct_revenue = three_stmt_assumptions["ap_pct_revenue"]
    if dividend_payout_ratio is None:
        dividend_payout_ratio = three_stmt_assumptions["dividend_payout_ratio"]
    if mandatory_amortization_pct is None:
        mandatory_amortization_pct = three_stmt_assumptions["mandatory_amortization_pct"]

    # --- Starting balance sheet position (most recent actuals) ---
    starting_cash = safe_get_row(balance_sheet, CASH_ROW_NAMES, default=0) or 0
    starting_debt = safe_get_row(balance_sheet, DEBT_ROW_NAMES, default=0) or 0
    starting_ppe = safe_get_row(balance_sheet, PPE_ROW_NAMES, default=0) or 0
    starting_equity = safe_get_row(balance_sheet, EQUITY_ROW_NAMES, default=0) or 0
    starting_ar = ar_pct_revenue * most_recent_revenue
    starting_inventory = inventory_pct_revenue * most_recent_revenue
    starting_ap = ap_pct_revenue * most_recent_revenue

    # The balancing plug: whatever isn't Cash/AR/Inventory/PP&E on the
    # asset side or AP/Debt/Equity on the other, netted into one figure and
    # held constant. Solving Assets = Liabilities + Equity for this term
    # using the REAL starting balances is what makes every future year
    # balance automatically — see the docstring above.
    other_net_assets = (starting_ap + starting_debt + starting_equity
                         - starting_cash - starting_ar - starting_inventory - starting_ppe)

    years = []
    current_revenue = most_recent_revenue
    cash, debt, ppe, equity = starting_cash, starting_debt, starting_ppe, starting_equity
    ar_bal, inv_bal, ap_bal = starting_ar, starting_inventory, starting_ap

    for year in range(1, years_to_project + 1):
        progress = (year - 1) / (years_to_project - 1) if years_to_project > 1 else 1
        growth_this_year = starting_growth_rate + (ending_growth_rate - starting_growth_rate) * progress

        # --- Income statement items that don't depend on the debt schedule ---
        revenue = current_revenue * (1 + growth_this_year)
        gross_profit = revenue * gross_margin
        sga = revenue * sga_pct_revenue
        da = revenue * da_pct_revenue
        ebit = gross_profit - sga - da

        new_ar = revenue * ar_pct_revenue
        new_inventory = revenue * inventory_pct_revenue
        new_ap = revenue * ap_pct_revenue
        capex = revenue * capex_pct_revenue
        new_ppe = ppe + capex - da
        change_in_nwc = (new_ar - ar_bal) + (new_inventory - inv_bal) - (new_ap - ap_bal)

        beginning_debt = debt
        mandatory_amortization = min(beginning_debt * mandatory_amortization_pct, beginning_debt)

        # --- Iteratively solve the interest-expense / ending-debt circularity ---
        ending_debt_guess = beginning_debt - mandatory_amortization
        for _ in range(10):
            average_debt = (beginning_debt + ending_debt_guess) / 2
            interest_expense = average_debt * interest_rate_on_debt
            pretax_income = ebit - interest_expense
            tax = max(pretax_income, 0) * tax_rate
            net_income = pretax_income - tax
            dividends_paid = max(net_income, 0) * dividend_payout_ratio

            cfo = net_income + da - change_in_nwc
            cash_before_debt_and_dividends = cash + cfo - capex

            available_after_mandatory = cash_before_debt_and_dividends - mandatory_amortization - dividends_paid

            if available_after_mandatory < min_cash_balance:
                # Not enough cash to stay above the minimum -> revolver draws
                # to cover the shortfall.
                revolver_draw = min_cash_balance - available_after_mandatory
                revolver_repay = 0
            else:
                revolver_draw = 0
                excess_cash = available_after_mandatory - min_cash_balance
                revolver_repay = min(excess_cash * cash_sweep_pct, beginning_debt - mandatory_amortization)

            new_ending_debt_guess = beginning_debt - mandatory_amortization - revolver_repay + revolver_draw
            if abs(new_ending_debt_guess - ending_debt_guess) < 1:  # converged to within $1
                ending_debt_guess = new_ending_debt_guess
                break
            ending_debt_guess = new_ending_debt_guess

        ending_debt = ending_debt_guess
        new_cash = cash_before_debt_and_dividends - mandatory_amortization - dividends_paid - revolver_repay + revolver_draw
        new_equity = equity + net_income - dividends_paid

        total_assets = new_cash + new_ar + new_inventory + new_ppe + other_net_assets
        total_liabilities_and_equity = new_ap + ending_debt + new_equity
        balance_check = total_assets - total_liabilities_and_equity

        years.append({
            "year": year,
            "revenue": revenue, "gross_profit": gross_profit, "sga": sga, "da": da,
            "ebit": ebit, "interest_expense": interest_expense, "pretax_income": pretax_income,
            "tax": tax, "net_income": net_income,
            "cash": new_cash, "accounts_receivable": new_ar, "inventory": new_inventory,
            "ppe": new_ppe, "other_net_assets": other_net_assets,
            "accounts_payable": new_ap, "debt_beginning": beginning_debt,
            "mandatory_amortization": mandatory_amortization,
            "revolver_draw": revolver_draw, "revolver_repayment": revolver_repay,
            "debt_ending": ending_debt, "equity": new_equity,
            "total_assets": total_assets, "total_liabilities_and_equity": total_liabilities_and_equity,
            "balance_check": balance_check,
            "cfo": cfo, "cfi": -capex, "cff": -mandatory_amortization - dividends_paid - revolver_repay + revolver_draw,
            "capex": capex, "dividends_paid": dividends_paid,
        })

        current_revenue, cash, debt, ppe, equity = revenue, new_cash, ending_debt, new_ppe, new_equity
        ar_bal, inv_bal, ap_bal = new_ar, new_inventory, new_ap

    return {
        "ticker": ticker_symbol,
        "company_name": company["company_name"],
        "currency": company["financial_currency"],
        "currency_symbol": get_currency_symbol(company["financial_currency"]),
        "years": years,
        "other_net_assets": other_net_assets,
        "most_recent_revenue": most_recent_revenue,
        "starting_cash": starting_cash, "starting_debt": starting_debt,
        "starting_ppe": starting_ppe, "starting_equity": starting_equity,
        "starting_ar": starting_ar, "starting_inventory": starting_inventory, "starting_ap": starting_ap,
        "assumptions_used": {
            "starting_growth_rate": starting_growth_rate, "ending_growth_rate": ending_growth_rate,
            "gross_margin": gross_margin, "sga_pct_revenue": sga_pct_revenue,
            "da_pct_revenue": da_pct_revenue, "capex_pct_revenue": capex_pct_revenue,
            "tax_rate": tax_rate, "interest_rate_on_debt": interest_rate_on_debt,
            "ar_pct_revenue": ar_pct_revenue, "inventory_pct_revenue": inventory_pct_revenue,
            "ap_pct_revenue": ap_pct_revenue, "dividend_payout_ratio": dividend_payout_ratio,
            "mandatory_amortization_pct": mandatory_amortization_pct,
            "min_cash_balance": min_cash_balance, "cash_sweep_pct": cash_sweep_pct,
        },
    }


def export_three_statement_to_excel(result):
    """
    LIVE model matching real 3-statement conventions (single combined
    sheet, exactly like the AFIN1002-style "Financial statement model +
    DCF" reference template): every projected line is an Excel formula.

    The debt schedule has a genuine circular reference — interest expense
    depends on the average of beginning/ending debt, but ending debt
    (after mandatory amortization, dividends, and any revolver draw/sweep)
    depends on net income, which depends on interest expense. Real models
    resolve this with Excel's iterative calculation setting, which is
    what this workbook enables (wb.calculation.iterate = True) — verified
    separately to actually converge to the correct value, not just avoid
    a formula error.
    """
    wb = Workbook()
    fmt = currency_formats(result["currency_symbol"])
    years = result["years"]
    n = len(years)
    a = result["assumptions_used"]

    wb.calculation.iterate = True
    wb.calculation.iterateCount = 100
    wb.calculation.iterateDelta = 0.001

    # ============================================
    # ASSUMPTIONS TAB
    # ============================================
    assump = wb.create_sheet("Assumptions")
    wb.remove(wb["Sheet"])  # drop openpyxl's default blank sheet
    assump.sheet_view.showGridLines = False
    assump["A1"] = f"{result['company_name']} ({result['ticker']}) — 3-Statement Model"
    assump["A1"].font = TITLE_FONT
    assump.merge_cells("A1:C1")
    assump["A2"] = "Blue cells are hardcoded inputs. The debt schedule on the Model tab is genuinely circular " \
                   "(interest depends on debt, debt depends on net income, net income depends on interest) — " \
                   "resolved via Excel's iterative calculation, enabled on this workbook."
    assump["A2"].font = Font(italic=True, size=9, color="666666")

    input_rows = [
        ("Most Recent Revenue (Year 0)", result["most_recent_revenue"], fmt["currency"]),
        ("Starting Revenue Growth Rate", a["starting_growth_rate"], "0.0%"),
        ("Ending Revenue Growth Rate", a["ending_growth_rate"], "0.0%"),
        ("Gross Margin", a["gross_margin"], "0.0%"),
        ("SG&A % of Revenue", a["sga_pct_revenue"], "0.0%"),
        ("D&A % of Revenue", a["da_pct_revenue"], "0.0%"),
        ("CapEx % of Revenue", a["capex_pct_revenue"], "0.0%"),
        ("Tax Rate", a["tax_rate"], "0.0%"),
        ("Interest Rate on Debt", a["interest_rate_on_debt"], "0.00%"),
        ("Accounts Receivable % of Revenue", a["ar_pct_revenue"], "0.0%"),
        ("Inventory % of Revenue", a["inventory_pct_revenue"], "0.0%"),
        ("Accounts Payable % of Revenue", a["ap_pct_revenue"], "0.0%"),
        ("Dividend Payout Ratio", a["dividend_payout_ratio"], "0.0%"),
        ("Mandatory Debt Amortization %", a["mandatory_amortization_pct"], "0.0%"),
        ("Minimum Cash Balance", a["min_cash_balance"], fmt["currency"]),
        ("Cash Sweep %", a["cash_sweep_pct"], "0.0%"),
        ("Years Projected", n, "0"),
        ("Starting Cash", result["starting_cash"], fmt["currency"]),
        ("Starting Debt", result["starting_debt"], fmt["currency"]),
        ("Starting PP&E", result["starting_ppe"], fmt["currency"]),
        ("Starting Equity", result["starting_equity"], fmt["currency"]),
        ("Starting Accounts Receivable", result["starting_ar"], fmt["currency"]),
        ("Starting Inventory", result["starting_inventory"], fmt["currency"]),
        ("Starting Accounts Payable", result["starting_ap"], fmt["currency"]),
    ]
    for i, (label, value, number_fmt) in enumerate(input_rows, start=4):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_input(assump, f"B{i}", value, number_fmt)
    ROW = {label: i for i, (label, *_) in enumerate(input_rows, start=4)}
    assump.column_dimensions["A"].width = 34
    assump.column_dimensions["B"].width = 20

    def A(label):
        return f"Assumptions!$B${ROW[label]}"

    # ============================================
    # MODEL TAB — everything on one sheet, like the reference template
    # ============================================
    model = wb.create_sheet("Model", 0)
    model.sheet_view.showGridLines = False
    model.freeze_panes = "C4"
    model["A1"] = f"{result['ticker']} — {n}-Year 3-Statement Model"
    model["A1"].font = TITLE_FONT
    model.merge_cells(f"A1:{get_column_letter(2 + n)}1")

    year_cols = [get_column_letter(3 + i) for i in range(n)]

    model["A3"] = "Year Index"
    model["B3"] = 0
    prev_col = "B"
    for col in year_cols:
        _write_formula(model, f"{col}3", f"={prev_col}3+1")
        prev_col = col

    def section(row, label):
        model.cell(row=row, column=1, value=label).font = Font(bold=True, size=12, color="1F4E78")

    def line(row, label, formula_fn, number_fmt=None, bold=False, year0_value=None):
        """formula_fn(col, prev_col) -> formula string for projected years (columns C onward).
        year0_value: literal input for column B (Year 0 actual), or None to leave blank."""
        nfmt = number_fmt if number_fmt is not None else fmt["currency"]
        model.cell(row=row, column=1, value=label).font = Font(bold=True) if bold else LABEL_FONT
        if year0_value is not None:
            _write_input(model, f"B{row}", year0_value, nfmt)
        prev_col = "B"
        for col in year_cols:
            _write_formula(model, f"{col}{row}", formula_fn(col, prev_col), nfmt, bold=bold)
            prev_col = col

    # --- Income Statement (part 1: down to EBIT, no debt dependency yet) ---
    section(5, "INCOME STATEMENT")
    line(6, "Revenue Growth Rate", lambda c, p: (
        f"={A('Starting Revenue Growth Rate')}+({A('Ending Revenue Growth Rate')}-{A('Starting Revenue Growth Rate')})"
        f"*({c}3-1)/({A('Years Projected')}-1)" if n > 1 else f"={A('Starting Revenue Growth Rate')}"
    ), "0.0%")
    line(7, "Revenue", lambda c, p: f"={p}7*(1+{c}6)", bold=True, year0_value=result["most_recent_revenue"])
    line(8, "Gross Profit", lambda c, p: f"={c}7*{A('Gross Margin')}")
    line(9, "SG&A", lambda c, p: f"={c}7*{A('SG&A % of Revenue')}")
    line(10, "D&A", lambda c, p: f"={c}7*{A('D&A % of Revenue')}")
    line(11, "EBIT", lambda c, p: f"={c}8-{c}9-{c}10", bold=True)

    # --- Debt Schedule (the circular part) ---
    section(13, "DEBT SCHEDULE")
    line(14, "Beginning Debt", lambda c, p: f"={p}20", year0_value=result["starting_debt"])
    line(15, "Average Debt (Beg./End.)", lambda c, p: f"=({c}14+{c}20)/2")
    line(16, "Interest Expense", lambda c, p: f"={c}15*{A('Interest Rate on Debt')}")
    line(17, "Mandatory Amortization", lambda c, p: f"=MIN({c}14*{A('Mandatory Debt Amortization %')},{c}14)")
    line(18, "Available After Mandatory Amort. & Dividends", lambda c, p: f"={c}27-{c}17-{c}26")
    line(19, "Revolver Draw / (Repayment)", lambda c, p: (
        f"=IF({c}18<{A('Minimum Cash Balance')},{A('Minimum Cash Balance')}-{c}18,"
        f"-MIN(({c}18-{A('Minimum Cash Balance')})*{A('Cash Sweep %')},{c}14-{c}17))"
    ))
    line(20, "Ending Debt", lambda c, p: f"={c}14-{c}17+{c}19", bold=True, year0_value=result["starting_debt"])

    # --- Income Statement (part 2: depends on interest expense above) ---
    line(23, "Pretax Income", lambda c, p: f"={c}11-{c}16")
    line(24, "Tax", lambda c, p: f"=MAX({c}23,0)*{A('Tax Rate')}")
    line(25, "Net Income", lambda c, p: f"={c}23-{c}24", bold=True)
    line(26, "Dividends Paid", lambda c, p: f"=MAX({c}25,0)*{A('Dividend Payout Ratio')}")

    # Cash before financing activities — feeds the debt schedule's
    # availability check above (row 18), closing the circular loop.
    line(27, "Cash Before Debt & Dividends", lambda c, p: f"={p}38+{c}34-{c}33")

    # --- Cash Flow & Working Capital ---
    section(28, "CASH FLOW & WORKING CAPITAL")
    line(29, "Accounts Receivable", lambda c, p: f"={c}7*{A('Accounts Receivable % of Revenue')}", year0_value=result["starting_ar"])
    line(30, "Inventory", lambda c, p: f"={c}7*{A('Inventory % of Revenue')}", year0_value=result["starting_inventory"])
    line(31, "Accounts Payable", lambda c, p: f"={c}7*{A('Accounts Payable % of Revenue')}", year0_value=result["starting_ap"])
    line(32, "Change in NWC", lambda c, p: f"=({c}29-{p}29)+({c}30-{p}30)-({c}31-{p}31)")
    line(33, "CapEx", lambda c, p: f"={c}7*{A('CapEx % of Revenue')}")
    line(34, "Cash Flow from Operations", lambda c, p: f"={c}25+{c}10-{c}32", bold=True)

    # --- Balance Sheet ---
    section(36, "BALANCE SHEET")
    line(37, "PP&E, net", lambda c, p: f"={p}37+{c}33-{c}10", year0_value=result["starting_ppe"])
    line(38, "Cash", lambda c, p: f"={c}27-{c}17-{c}26+{c}19", bold=True, year0_value=result["starting_cash"])
    line(39, "Other Net Assets (plug)", lambda c, p: (
        f"={A('Starting Accounts Payable')}+{A('Starting Debt')}+{A('Starting Equity')}"
        f"-{A('Starting Cash')}-{A('Starting Accounts Receivable')}-{A('Starting Inventory')}-{A('Starting PP&E')}"
    ))
    line(40, "Total Assets", lambda c, p: f"={c}38+{c}29+{c}30+{c}37+{c}39", bold=True)
    line(41, "Equity", lambda c, p: f"={p}41+{c}25-{c}26", year0_value=result["starting_equity"])
    line(42, "Total Liabilities & Equity", lambda c, p: f"={c}31+{c}20+{c}41", bold=True)
    line(43, "Balance Check", lambda c, p: f"={c}40-{c}42", bold=True)

    for col in year_cols:
        model[f"{col}43"].font = Font(bold=True, color="1B7A1B")

    model.column_dimensions["A"].width = 30
    for col in ["B"] + year_cols:
        model.column_dimensions[col].width = 16

    # --- Chart: Revenue vs Net Income ---
    chart = BarChart()
    chart.title = "Projected Revenue vs Net Income"
    chart.style = 10
    chart.y_axis.number_format = fmt["currency"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16
    last_col_idx = 2 + n
    rev_data = Reference(model, min_col=1, max_col=last_col_idx, min_row=7, max_row=7)
    ni_data = Reference(model, min_col=1, max_col=last_col_idx, min_row=25, max_row=25)
    chart.add_data(rev_data, titles_from_data=True, from_rows=True)
    chart.add_data(ni_data, titles_from_data=True, from_rows=True)
    categories = Reference(model, min_col=3, max_col=last_col_idx, min_row=3, max_row=3)
    chart.set_categories(categories)
    model.add_chart(chart, f"{get_column_letter(last_col_idx + 2)}5")

    return wb


def print_three_statement_summary(result):
    divider = "─" * 60
    cs = result["currency_symbol"]
    print(divider)
    print(f"{result['company_name']} ({result['ticker']}) — 3-Statement Model")
    print(divider)
    for y in result["years"]:
        print(f"Year {y['year']}: Revenue {cs}{y['revenue']:,.0f}  |  "
              f"Net Income {cs}{y['net_income']:,.0f}  |  Cash {cs}{y['cash']:,.0f}  |  "
              f"Debt {cs}{y['debt_ending']:,.0f}  |  Balance Check {cs}{y['balance_check']:,.2f}")
    print(divider)




def get_recent_news(ticker_symbol, max_items=5):
    ticker = get_ticker(ticker_symbol)
    try:
        raw_items = ticker.news or []
    except Exception:
        raw_items = []

    cleaned = []
    for item in raw_items[:max_items]:
        content = item.get("content", item)  # yfinance nests differently across versions
        title = content.get("title") or item.get("title")
        if not title:
            continue
        provider = content.get("provider")
        publisher = provider.get("displayName") if isinstance(provider, dict) else item.get("publisher", "Unknown")
        cleaned.append({"title": title, "publisher": publisher or "Unknown"})

    return cleaned


def run_public_company_valuation(ticker_symbol):
    """
    A comprehensive standalone valuation package for an already-public
    company: trading history and volatility, ownership/float, current
    trading multiples, analyst consensus, and the DCF/comps methodology —
    combined into ONE integrated model rather than three separate files.
    Reuses run_dcf() and build_comps_table() rather than re-deriving the
    same valuation twice; if either fails (e.g. insufficient data), the
    rest of the package still builds — this degrades gracefully, not all-or-nothing.
    """
    company = resolve_company(ticker_symbol)
    info = company["info"]
    price_divisor = company["price_divisor"]

    current_price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / price_divisor
    fifty_two_week_high = (info.get("fiftyTwoWeekHigh") or 0) / price_divisor
    fifty_two_week_low = (info.get("fiftyTwoWeekLow") or 0) / price_divisor

    # --- Historical returns & volatility, computed from real price history ---
    hist = _safe_hist(company["ticker_obj"], "5y")
    returns_1y = returns_3y = volatility_annualized = None
    if hist is not None and not hist.empty:
        closes = hist["Close"].dropna()
        if len(closes) > 252:
            returns_1y = float(closes.iloc[-1] / closes.iloc[-252] - 1)
        if len(closes) > 756:
            returns_3y = float(closes.iloc[-1] / closes.iloc[-756] - 1)
        daily_returns = closes.pct_change().dropna()
        if len(daily_returns) > 30:
            volatility_annualized = float(daily_returns.std() * (252 ** 0.5))

    # --- Reuse existing valuation methods rather than re-deriving them ---
    try:
        dcf = run_dcf(ticker_symbol)
        dcf_error = None
    except Exception as e:
        dcf, dcf_error = None, str(e)
    try:
        comps = build_comps_table(ticker_symbol)
    except Exception:
        comps = None

    return {
        "ticker": ticker_symbol,
        "ticker_obj": company["ticker_obj"],
        "business_description": info.get("longBusinessSummary") or info.get("description"),
        "employees": info.get("fullTimeEmployees"),
        "company_name": company["company_name"],
        "currency_symbol": get_currency_symbol(company["quote_currency"]),
        "exchange": company["exchange"], "sector": company["sector"], "industry": company["industry"],
        "current_price": current_price,
        "fifty_two_week_high": fifty_two_week_high, "fifty_two_week_low": fifty_two_week_low,
        "market_cap": info.get("marketCap"), "enterprise_value": info.get("enterpriseValue"),
        "shares_outstanding": info.get("sharesOutstanding"), "float_shares": info.get("floatShares"),
        "insider_pct": info.get("heldPercentInsiders"), "institution_pct": info.get("heldPercentInstitutions"),
        "avg_volume": info.get("averageVolume"), "beta": info.get("beta"),
        "trailing_pe": info.get("trailingPE"), "forward_pe": info.get("forwardPE"),
        "price_to_book": info.get("priceToBook"), "dividend_yield": info.get("dividendYield"),
        "ev_revenue": (info.get("enterpriseValue") / info.get("totalRevenue")) if (info.get("enterpriseValue") and info.get("totalRevenue")) else None,
        "ev_ebitda": (info.get("enterpriseValue") / info.get("ebitda")) if (info.get("enterpriseValue") and info.get("ebitda") and info.get("ebitda") > 0) else None,
        "recommendation_mean": info.get("recommendationMean"), "recommendation_key": info.get("recommendationKey"),
        "num_analysts": info.get("numberOfAnalystOpinions"),
        "target_high": (info.get("targetHighPrice") or 0) / price_divisor if info.get("targetHighPrice") else None,
        "target_low": (info.get("targetLowPrice") or 0) / price_divisor if info.get("targetLowPrice") else None,
        "target_mean": (info.get("targetMeanPrice") or 0) / price_divisor if info.get("targetMeanPrice") else None,
        "target_median": (info.get("targetMedianPrice") or 0) / price_divisor if info.get("targetMedianPrice") else None,
        "returns_1y": returns_1y, "returns_3y": returns_3y, "volatility_annualized": volatility_annualized,
        "dcf": dcf, "dcf_error": dcf_error, "comps": comps,
    }


def print_public_company_valuation(result):
    sym = result["currency_symbol"]
    divider = "─" * 64
    print(divider)
    print(f"{result['company_name']} ({result['ticker']}) — Public Company Valuation")
    print(divider)
    print(f"Current Price: {sym}{result['current_price']:,.2f}   "
          f"52W Range: {sym}{result['fifty_two_week_low']:,.2f} - {sym}{result['fifty_two_week_high']:,.2f}")
    if result["market_cap"]:
        print(f"Market Cap: {sym}{result['market_cap']:,.0f}")
    if result["ev_ebitda"]:
        print(f"EV/EBITDA: {result['ev_ebitda']:.1f}x   Trailing P/E: {result['trailing_pe']:.1f}x" if result["trailing_pe"] else f"EV/EBITDA: {result['ev_ebitda']:.1f}x")
    if result["target_mean"]:
        print(f"Analyst Target (mean): {sym}{result['target_mean']:,.2f}  ({result['num_analysts'] or '?'} analysts)")
    if result["dcf"]:
        print(f"DCF Implied Price: {sym}{result['dcf']['implied_share_price']:,.2f}  ({result['dcf']['upside']:+.1%})")
    print(divider)


def export_public_company_valuation_to_excel(result):
    """
    LIVE model bringing trading data, ownership/float, current multiples,
    and the DCF/comps methodology into one workbook — real inputs in blue,
    every derived figure (returns, valuation synthesis) a formula.
    """
    wb = Workbook()
    fmt = currency_formats(result["currency_symbol"])
    sym = result["currency_symbol"]

    # ============================================
    # SUMMARY TAB
    # ============================================
    summary = wb.active
    summary.title = "Summary"
    summary.sheet_view.showGridLines = False
    summary["A1"] = f"{result['company_name']} ({result['ticker']}) — Public Company Valuation"
    summary["A1"].font = TITLE_FONT
    summary.merge_cells("A1:B1")
    summary["A2"] = "Blue cells are raw market data (inputs); valuation synthesis rows are live formulas."
    summary["A2"].font = Font(italic=True, size=9, color="666666")

    rows = [
        ("Current Price", result["current_price"], fmt["price"]),
        ("52-Week High", result["fifty_two_week_high"], fmt["price"]),
        ("52-Week Low", result["fifty_two_week_low"], fmt["price"]),
        ("Market Cap", result["market_cap"], fmt["currency"]),
        ("Enterprise Value", result["enterprise_value"], fmt["currency"]),
        ("Beta", result["beta"], "0.00"),
        ("1-Year Return", result["returns_1y"], "0.0%"),
        ("Annualized Volatility", result["volatility_annualized"], "0.0%"),
    ]
    r = 4
    for label, value, number_fmt in rows:
        summary.cell(row=r, column=1, value=label).font = LABEL_FONT
        if value is not None:
            _write_input(summary, f"B{r}", value, number_fmt)
        r += 1

    r += 1
    summary.cell(row=r, column=1, value="Valuation Synthesis").font = Font(bold=True, size=12, color="1F4E78")
    r += 1
    synth_start = r
    if result["dcf"]:
        summary.cell(row=r, column=1, value="DCF Implied Price").font = LABEL_FONT
        _write_input(summary, f"B{r}", result["dcf"]["implied_share_price"], fmt["price"])
        r += 1
    if result["comps"] and result["comps"].get("implied_ev_from_revenue") and result["shares_outstanding"]:
        nd = (result["comps"]["target"].get("total_debt", 0) or 0) - (result["comps"]["target"].get("cash", 0) or 0)
        implied_comps_price = (result["comps"]["implied_ev_from_revenue"] - nd) / result["shares_outstanding"]
        summary.cell(row=r, column=1, value="Comps Implied Price (EV/Revenue)").font = LABEL_FONT
        _write_input(summary, f"B{r}", implied_comps_price, fmt["price"])
        r += 1
    if result["target_mean"]:
        summary.cell(row=r, column=1, value="Analyst Target (Mean)").font = LABEL_FONT
        _write_input(summary, f"B{r}", result["target_mean"], fmt["price"])
        r += 1
    synth_end = r - 1

    if synth_end >= synth_start:
        summary.cell(row=r, column=1, value="Average of Available Methods").font = Font(bold=True, color="1F4E78")
        _write_formula(summary, f"B{r}", f"=AVERAGE(B{synth_start}:B{synth_end})", fmt["price"], bold=True)
        r += 1
        summary.cell(row=r, column=1, value="Current Price (for reference)").font = LABEL_FONT
        _write_formula(summary, f"B{r}", f"=B4", fmt["price"])
        r += 1
        summary.cell(row=r, column=1, value="Implied Upside / (Downside)").font = Font(bold=True, color="1F4E78")
        _write_formula(summary, f"B{r}", f"=(B{r - 2}-B{r - 1})/B{r - 1}", fmt["percent"], bold=True)

    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 20

    # ============================================
    # OWNERSHIP & FLOAT TAB
    # ============================================
    own = wb.create_sheet("Ownership & Float")
    own.sheet_view.showGridLines = False
    own["A1"] = f"{result['company_name']} — Ownership & Float"
    own["A1"].font = TITLE_FONT
    own.merge_cells("A1:B1")

    own_rows = [
        ("Shares Outstanding", result["shares_outstanding"], "#,##0"),
        ("Float (Shares)", result["float_shares"], "#,##0"),
        ("Insider Ownership %", result["insider_pct"], "0.0%"),
        ("Institutional Ownership %", result["institution_pct"], "0.0%"),
        ("Average Daily Volume", result["avg_volume"], "#,##0"),
    ]
    r2 = 4
    for label, value, number_fmt in own_rows:
        own.cell(row=r2, column=1, value=label).font = LABEL_FONT
        if value is not None:
            _write_input(own, f"B{r2}", value, number_fmt)
        r2 += 1
    if result["shares_outstanding"] and result["float_shares"]:
        own.cell(row=r2 + 1, column=1, value="Float as % of Shares Outstanding").font = Font(bold=True, color="1F4E78")
        _write_formula(own, f"B{r2 + 1}", "=B5/B4", "0.0%", bold=True)
    own.column_dimensions["A"].width = 30
    own.column_dimensions["B"].width = 20

    # ============================================
    # TRADING MULTIPLES TAB
    # ============================================
    mult = wb.create_sheet("Trading Multiples")
    mult.sheet_view.showGridLines = False
    mult["A1"] = f"{result['company_name']} — Current Trading Multiples"
    mult["A1"].font = TITLE_FONT
    mult.merge_cells("A1:B1")
    mult_rows = [
        ("EV / Revenue", result["ev_revenue"], '0.00"x"'), ("EV / EBITDA", result["ev_ebitda"], '0.00"x"'),
        ("Trailing P/E", result["trailing_pe"], '0.00"x"'), ("Forward P/E", result["forward_pe"], '0.00"x"'),
        ("Price / Book", result["price_to_book"], '0.00"x"'), ("Dividend Yield", result["dividend_yield"], "0.0%"),
    ]
    r3 = 4
    for label, value, number_fmt in mult_rows:
        mult.cell(row=r3, column=1, value=label).font = LABEL_FONT
        if value is not None:
            _write_input(mult, f"B{r3}", value, number_fmt)
        r3 += 1
    mult.column_dimensions["A"].width = 22
    mult.column_dimensions["B"].width = 16

    # ============================================
    # ANALYST CONSENSUS TAB
    # ============================================
    an = wb.create_sheet("Analyst Consensus")
    an.sheet_view.showGridLines = False
    an["A1"] = f"{result['company_name']} — Analyst Consensus"
    an["A1"].font = TITLE_FONT
    an.merge_cells("A1:B1")
    an_rows = [
        ("Recommendation", result["recommendation_key"] or "N/A", "General"),
        ("Recommendation Score (1=Strong Buy, 5=Strong Sell)", result["recommendation_mean"], "0.00"),
        ("Number of Analysts", result["num_analysts"], "0"),
        ("Target Low", result["target_low"], fmt["price"]), ("Target Mean", result["target_mean"], fmt["price"]),
        ("Target Median", result["target_median"], fmt["price"]), ("Target High", result["target_high"], fmt["price"]),
    ]
    r4 = 4
    for label, value, number_fmt in an_rows:
        an.cell(row=r4, column=1, value=label).font = LABEL_FONT
        if value is not None:
            cell = an.cell(row=r4, column=2, value=value)
            if number_fmt != "General":
                cell.number_format = number_fmt
            cell.font = Font(color="0000FF")
            cell.fill = ASSUMPTION_FILL
        r4 += 1
    an.column_dimensions["A"].width = 42
    an.column_dimensions["B"].width = 18

    # Comps and DCF snapshots, reusing the same safe embedding pattern as the M&A model
    if result["dcf"]:
        _add_dcf_snapshot_tab(wb, result["dcf"], "DCF Snapshot")
    if result["comps"]:
        comps_wb = export_comps_to_excel(result["comps"], result["ticker"])
        _copy_sheet_into(wb, comps_wb["Comps"], "Trading Comps")

    return wb

def run_research(ticker_symbol):
    """
    'Read this before a call' view: profile, key financial snapshot, a quick
    DCF for a fair-value anchor, and recent headlines. If the DCF can't run
    (missing data for this market), research mode still returns everything
    else rather than failing outright.
    """
    company = resolve_company(ticker_symbol)
    info = company["info"]
    price_divisor = company["price_divisor"]

    current_price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / price_divisor

    snapshot = {
        "market_cap": info.get("marketCap"),
        "pe_ratio": info.get("trailingPE"),
        "forward_pe": info.get("forwardPE"),
        "revenue": info.get("totalRevenue"),
        "revenue_growth": info.get("revenueGrowth"),
        "gross_margin": info.get("grossMargins"),
        "profit_margin": info.get("profitMargins"),
        "fifty_two_week_high": (info.get("fiftyTwoWeekHigh") or 0) / price_divisor,
        "fifty_two_week_low": (info.get("fiftyTwoWeekLow") or 0) / price_divisor,
        "analyst_target": (info.get("targetMeanPrice") or 0) / price_divisor,
        "recommendation": info.get("recommendationKey", "n/a"),
        "beta": info.get("beta"),
    }

    try:
        dcf_result = run_dcf(ticker_symbol)
        dcf_error = None
    except Exception as e:
        dcf_result = None
        dcf_error = str(e)

    return {
        "company": company,
        "current_price": current_price,
        "snapshot": snapshot,
        "dcf": dcf_result,
        "dcf_error": dcf_error,
        "news": get_recent_news(ticker_symbol),
    }


def print_research_report(research):
    company = research["company"]
    symbol = get_currency_symbol(company["quote_currency"])
    snap = research["snapshot"]
    divider = "─" * 60

    print(divider)
    print(f"{company['company_name']} ({company['ticker']})  ·  {company['exchange']}  ·  {company['country']}")
    print(f"{company['sector']} — {company['industry']}")
    print(divider)

    print(f"Current Price:      {symbol}{research['current_price']:,.2f}")
    if snap["analyst_target"]:
        print(f"Analyst Target:     {symbol}{snap['analyst_target']:,.2f}   ({snap['recommendation']})")
    if snap["fifty_two_week_low"] and snap["fifty_two_week_high"]:
        print(f"52-Week Range:      {symbol}{snap['fifty_two_week_low']:,.2f} - {symbol}{snap['fifty_two_week_high']:,.2f}")

    print()
    print("Fundamentals")
    if snap["market_cap"]:
        print(f"  Market Cap:       {symbol}{snap['market_cap']:,.0f}")
    if snap["revenue"]:
        print(f"  Revenue:          {symbol}{snap['revenue']:,.0f}")
    if snap["revenue_growth"] is not None:
        print(f"  Revenue Growth:   {snap['revenue_growth']:+.1%}")
    if snap["gross_margin"] is not None:
        print(f"  Gross Margin:     {snap['gross_margin']:.1%}")
    if snap["profit_margin"] is not None:
        print(f"  Profit Margin:    {snap['profit_margin']:.1%}")
    if snap["pe_ratio"]:
        print(f"  P/E (trailing):   {snap['pe_ratio']:.1f}x")
    if snap["forward_pe"]:
        print(f"  P/E (forward):    {snap['forward_pe']:.1f}x")

    print()
    print("DCF Fair Value Anchor")
    if research["dcf"]:
        dcf = research["dcf"]
        a = dcf["assumptions_used"]
        print(f"  Implied Price:    {dcf['currency_symbol']}{dcf['implied_share_price']:,.2f}")
        print(f"  Upside/Downside:  {dcf['upside']:+.1%}")
        print(f"  Assumptions used: {a['starting_growth_rate']:+.1%} → {a['ending_growth_rate']:+.1%} growth, "
              f"{a['starting_fcf_margin']:+.1%} → {a['ending_fcf_margin']:+.1%} FCF margin, WACC {a['wacc']:.1%}")
        if dcf["used_fallback_growth"] or dcf["used_fallback_margin"]:
            print("  Note: some assumptions used generic fallback values — insufficient historical data to derive them directly.")
        if dcf["is_negative_margin"]:
            print("  Caution: this company currently has a negative free cash flow margin (e.g. early-stage, "
                  "capex-heavy, or pre-profitability). A DCF anchored to near-term cash flows may understate "
                  "value for companies being priced on future capacity or a longer-term thesis — treat this "
                  "implied price as one data point, not a verdict.")
    else:
        print(f"  Unavailable — {research['dcf_error']}")

    print()
    print("Recent Headlines")
    if research["news"]:
        for item in research["news"]:
            print(f"  • {item['title']}  ({item['publisher']})")
    else:
        print("  No recent headlines found.")
    print(divider)


# ============================================
# EXCEL EXPORT — shared style constants
# ============================================

HEADER_FONT = Font(bold=True, color="FFFFFF", size=12)
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
LABEL_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=16, color="1F4E78")
THIN_BORDER = Border(bottom=Side(style="thin", color="CCCCCC"))
CORNER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


def currency_formats(currency_symbol):
    """Number formats built from the target's own reporting currency,
    instead of hardcoding '$' and mislabeling every non-US model."""
    return {
        "currency": f'"{currency_symbol}"#,##0',
        "price": f'"{currency_symbol}"#,##0.00',
        "percent": '0.0%',
    }


INPUT_FONT = Font(color="0000FF")  # blue = hardcoded input, per standard financial-model convention
ASSUMPTION_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # light yellow


def _write_input(ws, cell_ref, value, number_format=None, bold=False):
    """A hardcoded input cell: blue font, yellow fill — the standard visual
    marker so anyone opening the model instantly knows which cells are
    safe to type over, vs. which are formulas that would break if edited."""
    cell = ws[cell_ref]
    cell.value = value
    cell.font = Font(color="0000FF", bold=bold)
    cell.fill = ASSUMPTION_FILL
    if number_format:
        cell.number_format = number_format
    return cell


def _write_formula(ws, cell_ref, formula, number_format=None, bold=False):
    """A calculated cell: black font, no fill — standard convention for
    'this is derived, don't type over it directly.'"""
    cell = ws[cell_ref]
    cell.value = formula
    cell.font = Font(color="000000", bold=bold)
    if number_format:
        cell.number_format = number_format
    return cell


def export_dcf_to_excel(result):
    """
    IB-grade live model: WACC is DERIVED via CAPM build-up (not typed in
    directly), a scenario toggle (1/2/3 = Bear/Base/Bull) shifts every
    downstream assumption using the same deltas as run_scenarios(), an
    optional mid-year discounting convention is available, and 2 years of
    historical revenue sit alongside the projection for context. A Checks
    tab centralizes every internal consistency check in one place.
    """
    wb = Workbook()
    fmt = currency_formats(result["currency_symbol"])
    a = result["assumptions_used"]
    years_to_project = len(result["projected_revenue"])
    most_recent_revenue = result["projected_revenue"][0] / (1 + a["starting_growth_rate"])
    hist_rev = result.get("historical_revenue") or []
    beta = result.get("beta") or 1.0
    market_cap = result.get("market_cap") or (result["current_price"] * result["shares_outstanding"])
    total_debt = result.get("total_debt") or max(result["net_debt"], 0)

    # ============================================
    # ASSUMPTIONS TAB
    # ============================================
    assump = wb.active
    assump.title = "Assumptions"
    assump.sheet_view.showGridLines = False
    assump["A1"] = f"{result['company_name']} ({result['ticker']}) — DCF Model"
    assump["A1"].font = TITLE_FONT
    assump.merge_cells("A1:C1")
    assump["A2"] = "Blue cells are hardcoded inputs. WACC is DERIVED (CAPM build-up), not typed directly."
    assump["A2"].font = Font(italic=True, size=9, color="666666")

    input_rows = [
        ("Most Recent Revenue (Year 0)", most_recent_revenue, fmt["currency"]),
        ("Historical Revenue (Year -1)", hist_rev[1] if len(hist_rev) > 1 else None, fmt["currency"]),
        ("Historical Revenue (Year -2)", hist_rev[2] if len(hist_rev) > 2 else None, fmt["currency"]),
        ("Base Starting Revenue Growth Rate", a["starting_growth_rate"], "0.0%"),
        ("Base Ending Revenue Growth Rate", a["ending_growth_rate"], "0.0%"),
        ("Base Starting FCF Margin", a["starting_fcf_margin"], "0.0%"),
        ("Base Ending FCF Margin", a["ending_fcf_margin"], "0.0%"),
        ("Terminal (Long-Run) Growth Rate", a["terminal_growth_rate"], "0.00%"),
        ("Years Projected", years_to_project, "0"),
        ("Net Debt", result["net_debt"], fmt["currency"]),
        ("Shares Outstanding", result["shares_outstanding"], "#,##0"),
        ("Current Share Price", result["current_price"], fmt["price"]),
        # --- WACC build-up (CAPM) ---
        ("Market Value of Equity", market_cap, fmt["currency"]),
        ("Total Debt", total_debt, fmt["currency"]),
        ("Risk-Free Rate", 0.045, "0.00%"),
        ("Equity Risk Premium", 0.05, "0.00%"),
        ("Beta", beta, "0.00"),
        ("Pre-Tax Cost of Debt", a["wacc"] * 0.6, "0.00%"),  # rough starting estimate; edit directly
        ("Tax Rate", 0.21, "0.0%"),
        # --- Toggles ---
        ("Mid-Year Convention (1=On, 0=Off)", 1, "0"),
        ("Scenario (1=Bear, 2=Base, 3=Bull)", 2, "0"),
    ]
    for i, (label, value, number_fmt) in enumerate(input_rows, start=4):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        if value is not None:
            _write_input(assump, f"B{i}", value, number_fmt)
    ROW = {label: i for i, (label, *_) in enumerate(input_rows, start=4)}

    def A(label):
        return f"Assumptions!$B${ROW[label]}"

    # --- Derived WACC build-up rows (formulas, not inputs) ---
    calc_row = 4 + len(input_rows) + 1
    assump.cell(row=calc_row - 1, column=1, value="Derived").font = Font(bold=True, size=11, color="1F4E78")
    calc_rows = [
        ("Cost of Equity (CAPM)", f"={A('Risk-Free Rate')}+{A('Beta')}*{A('Equity Risk Premium')}", "0.00%"),
        ("After-Tax Cost of Debt", f"={A('Pre-Tax Cost of Debt')}*(1-{A('Tax Rate')})", "0.00%"),
        ("Weight of Equity (E/V)", f"={A('Market Value of Equity')}/({A('Market Value of Equity')}+{A('Total Debt')})", "0.0%"),
        ("Weight of Debt (D/V)", f"={A('Total Debt')}/({A('Market Value of Equity')}+{A('Total Debt')})", "0.0%"),
    ]
    for i, (label, formula, number_fmt) in enumerate(calc_rows, start=calc_row):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_formula(assump, f"B{i}", formula, number_fmt)
    CALC_ROW = {label: i for i, (label, *_) in enumerate(calc_rows, start=calc_row)}

    wacc_builtup_row = calc_row + len(calc_rows)
    assump.cell(row=wacc_builtup_row, column=1, value="WACC (Built-Up)").font = Font(bold=True)
    _write_formula(assump, f"B{wacc_builtup_row}",
                    f"=Assumptions!$B${CALC_ROW['Weight of Equity (E/V)']}*Assumptions!$B${CALC_ROW['Cost of Equity (CAPM)']}"
                    f"+Assumptions!$B${CALC_ROW['Weight of Debt (D/V)']}*Assumptions!$B${CALC_ROW['After-Tax Cost of Debt']}",
                    "0.00%", bold=True)

    # --- Scenario-adjusted assumptions (same deltas as run_scenarios: Python) ---
    scen_row = wacc_builtup_row + 2
    assump.cell(row=scen_row - 1, column=1, value="Scenario-Adjusted (feeds the DCF tab)").font = Font(bold=True, size=11, color="1F4E78")
    scen_rows = [
        ("Growth Delta (Start)", f"=CHOOSE({A('Scenario (1=Bear, 2=Base, 3=Bull)')},-0.08,0,0.08)", "0.0%"),
        ("Growth Delta (End)", f"=CHOOSE({A('Scenario (1=Bear, 2=Base, 3=Bull)')},-0.04,0,0.04)", "0.0%"),
        ("Margin Delta", f"=CHOOSE({A('Scenario (1=Bear, 2=Base, 3=Bull)')},-0.05,0,0.05)", "0.0%"),
        ("WACC Delta", f"=CHOOSE({A('Scenario (1=Bear, 2=Base, 3=Bull)')},0.02,0,-0.01)", "0.00%"),
    ]
    for i, (label, formula, number_fmt) in enumerate(scen_rows, start=scen_row):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_formula(assump, f"B{i}", formula, number_fmt)
    SCEN = {label: i for i, (label, *_) in enumerate(scen_rows, start=scen_row)}

    final_row = scen_row + len(scen_rows)
    final_rows = [
        ("Starting Growth Rate (Used)", f"={A('Base Starting Revenue Growth Rate')}+Assumptions!$B${SCEN['Growth Delta (Start)']}", "0.0%"),
        ("Ending Growth Rate (Used)", f"={A('Base Ending Revenue Growth Rate')}+Assumptions!$B${SCEN['Growth Delta (End)']}", "0.0%"),
        ("Starting FCF Margin (Used)", f"={A('Base Starting FCF Margin')}+Assumptions!$B${SCEN['Margin Delta']}", "0.0%"),
        ("Ending FCF Margin (Used)", f"={A('Base Ending FCF Margin')}+Assumptions!$B${SCEN['Margin Delta']}", "0.0%"),
        ("WACC (Used)", f"=B{wacc_builtup_row}+Assumptions!$B${SCEN['WACC Delta']}", "0.00%"),
    ]
    for i, (label, formula, number_fmt) in enumerate(final_rows, start=final_row):
        assump.cell(row=i, column=1, value=label).font = Font(bold=True, color="1F4E78")
        _write_formula(assump, f"B{i}", formula, number_fmt, bold=True)
    FINAL = {label: i for i, (label, *_) in enumerate(final_rows, start=final_row)}

    assump.column_dimensions["A"].width = 38
    assump.column_dimensions["B"].width = 20

    def F(label):
        return f"Assumptions!$B${FINAL[label]}"

    # ============================================
    # DCF TAB
    # ============================================
    dcf_tab = wb.create_sheet("DCF")
    dcf_tab.sheet_view.showGridLines = False
    dcf_tab["A1"] = f"{result['ticker']} — {years_to_project}-Year DCF Projection"
    dcf_tab["A1"].font = TITLE_FONT
    dcf_tab.merge_cells("A1:H1")

    year_cols = [get_column_letter(3 + i) for i in range(years_to_project)]

    # Historical revenue reference, shown for context (not part of the live chain)
    if hist_rev:
        hist_text = ", ".join(f"Yr-{i}: {v:,.0f}" for i, v in enumerate(hist_rev[1:], start=1))
        dcf_tab["A2"] = f"Historical revenue (for context): {hist_text}"
        dcf_tab["A2"].font = Font(italic=True, size=9, color="666666")

    dcf_tab["A3"] = "Year Index"
    dcf_tab["B3"] = 0
    prev_col = "B"
    for col in year_cols:
        _write_formula(dcf_tab, f"{col}3", f"={prev_col}3+1")
        prev_col = col

    dcf_tab["A4"] = "Revenue Growth Rate"
    for col in year_cols:
        progress = f"({col}3-1)/(Assumptions!$B${ROW['Years Projected']}-1)" if years_to_project > 1 else "1"
        _write_formula(dcf_tab, f"{col}4",
                        f"={F('Starting Growth Rate (Used)')}+({F('Ending Growth Rate (Used)')}-{F('Starting Growth Rate (Used)')})*{progress}",
                        "0.0%")

    dcf_tab["A5"] = "Revenue"
    _write_formula(dcf_tab, "B5", f"=Assumptions!$B${ROW['Most Recent Revenue (Year 0)']}", fmt["currency"], bold=True)
    for i, col in enumerate(year_cols):
        prior_col = "B" if i == 0 else year_cols[i - 1]
        _write_formula(dcf_tab, f"{col}5", f"={prior_col}5*(1+{col}4)", fmt["currency"], bold=True)

    dcf_tab["A6"] = "FCF Margin"
    for col in year_cols:
        progress = f"({col}3-1)/(Assumptions!$B${ROW['Years Projected']}-1)" if years_to_project > 1 else "1"
        _write_formula(dcf_tab, f"{col}6",
                        f"={F('Starting FCF Margin (Used)')}+({F('Ending FCF Margin (Used)')}-{F('Starting FCF Margin (Used)')})*{progress}",
                        "0.0%")

    dcf_tab["A7"] = "Free Cash Flow"
    for col in year_cols:
        _write_formula(dcf_tab, f"{col}7", f"={col}5*{col}6", fmt["currency"], bold=True)

    dcf_tab["A8"] = "Discount Factor (mid-year convention optional)"
    for col in year_cols:
        _write_formula(dcf_tab, f"{col}8",
                        f"=1/(1+{F('WACC (Used)')})^({col}3-0.5*{A('Mid-Year Convention (1=On, 0=Off)')})",
                        "0.0000")

    dcf_tab["A9"] = "PV of FCF"
    for col in year_cols:
        _write_formula(dcf_tab, f"{col}9", f"={col}7*{col}8", fmt["currency"])

    for c in ["A"] + year_cols:
        dcf_tab.column_dimensions[c].width = 18 if c != "A" else 34

    last_col = year_cols[-1]
    v_row = 12
    dcf_tab.cell(row=v_row, column=1, value="Sum of PV of FCF (Years 1-N)").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row}", f"=SUM(C9:{last_col}9)", fmt["currency"], bold=True)

    dcf_tab.cell(row=v_row + 1, column=1, value="Terminal Value (undiscounted)").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 1}",
                    f"={last_col}7*(1+Assumptions!$B${ROW['Terminal (Long-Run) Growth Rate']})/"
                    f"({F('WACC (Used)')}-Assumptions!$B${ROW['Terminal (Long-Run) Growth Rate']})",
                    fmt["currency"])

    dcf_tab.cell(row=v_row + 2, column=1, value="PV of Terminal Value").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 2}", f"=B{v_row + 1}*{last_col}8", fmt["currency"], bold=True)

    dcf_tab.cell(row=v_row + 3, column=1, value="Enterprise Value").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 3}", f"=B{v_row}+B{v_row + 2}", fmt["currency"], bold=True)

    dcf_tab.cell(row=v_row + 4, column=1, value="Less: Net Debt").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 4}", f"=-Assumptions!$B${ROW['Net Debt']}", fmt["currency"])

    dcf_tab.cell(row=v_row + 5, column=1, value="Equity Value").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 5}", f"=B{v_row + 3}+B{v_row + 4}", fmt["currency"], bold=True)

    dcf_tab.cell(row=v_row + 6, column=1, value="Shares Outstanding").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 6}", f"=Assumptions!$B${ROW['Shares Outstanding']}", "#,##0")

    dcf_tab.cell(row=v_row + 7, column=1, value="Implied Share Price").font = Font(bold=True, size=12, color="1F4E78")
    _write_formula(dcf_tab, f"B{v_row + 7}", f"=B{v_row + 5}/B{v_row + 6}", fmt["price"], bold=True)

    dcf_tab.cell(row=v_row + 8, column=1, value="Current Share Price").font = LABEL_FONT
    _write_formula(dcf_tab, f"B{v_row + 8}", f"=Assumptions!$B${ROW['Current Share Price']}", fmt["price"])

    dcf_tab.cell(row=v_row + 9, column=1, value="Upside / (Downside)").font = Font(bold=True, color="1F4E78")
    _write_formula(dcf_tab, f"B{v_row + 9}", f"=(B{v_row + 7}-B{v_row + 8})/B{v_row + 8}", fmt["percent"], bold=True)

    implied_price_cell = f"DCF!B{v_row + 7}"
    upside_cell = f"DCF!B{v_row + 9}"
    ev_cell = f"DCF!B{v_row + 3}"

    last_col_idx = 2 + years_to_project
    chart2 = BarChart()
    chart2.title = "Projected Revenue vs Free Cash Flow"
    chart2.style = 10
    chart2.y_axis.number_format = fmt["currency"]
    chart2.y_axis.delete = False
    chart2.x_axis.delete = False
    chart2.height, chart2.width = 8, 16
    data = Reference(dcf_tab, min_col=1, max_col=last_col_idx, min_row=5, max_row=5)
    chart2.add_data(data, titles_from_data=True, from_rows=True)
    data_fcf = Reference(dcf_tab, min_col=1, max_col=last_col_idx, min_row=7, max_row=7)
    chart2.add_data(data_fcf, titles_from_data=True, from_rows=True)
    categories = Reference(dcf_tab, min_col=3, max_col=last_col_idx, min_row=3, max_row=3)
    chart2.set_categories(categories)
    dcf_tab.add_chart(chart2, f"{get_column_letter(last_col_idx + 2)}3")

    # ============================================
    # SENSITIVITY TAB — centered on the DERIVED (Used) WACC now
    # ============================================
    sens_tab = wb.create_sheet("Sensitivity")
    sens_tab.sheet_view.showGridLines = False
    sens_tab["A1"] = f"{result['ticker']} — Implied Share Price Sensitivity"
    sens_tab["A1"].font = TITLE_FONT
    sens_tab.merge_cells("A1:H1")
    sens_tab["A2"] = "Rows = WACC, Columns = Terminal Growth Rate. Centered on the DERIVED WACC. Every cell is a live formula."
    sens_tab["A2"].font = Font(italic=True, size=9, color="666666")

    # Replicates the EXACT Excel WACC build-up formula (same risk-free/ERP/
    # tax defaults used in the Assumptions tab) so the sensitivity grid is
    # centered on the actual derived WACC the model uses — not the stale
    # raw wacc value, which would otherwise make the Checks tab fail.
    _risk_free, _erp, _tax = 0.045, 0.05, 0.21
    _cost_of_equity_est = _risk_free + beta * _erp
    _pretax_cod_est = a["wacc"] * 0.6
    _aftertax_cod_est = _pretax_cod_est * (1 - _tax)
    _total_v = (market_cap + total_debt) or 1
    _weight_e = market_cap / _total_v
    _weight_d = total_debt / _total_v
    derived_wacc_est = _weight_e * _cost_of_equity_est + _weight_d * _aftertax_cod_est

    base_wacc, base_g = derived_wacc_est, a["terminal_growth_rate"]
    wacc_values = [base_wacc - 0.02, base_wacc - 0.01, base_wacc, base_wacc + 0.01, base_wacc + 0.02]
    growth_values = [base_g - 0.01, base_g - 0.005, base_g, base_g + 0.005, base_g + 0.01]

    corner = sens_tab.cell(row=4, column=1, value="WACC \\ Term. Growth")
    corner.font = Font(bold=True, size=9, color="FFFFFF")
    corner.fill = HEADER_FILL
    corner.alignment = Alignment(horizontal="center", wrap_text=True)

    for col_idx, g_val in enumerate(growth_values, start=2):
        cell = sens_tab.cell(row=4, column=col_idx, value=g_val)
        cell.number_format = "0.00%"
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    fcf_range = f"DCF!C7:{last_col}7"

    fcf_year_range = f"DCF!C3:{last_col}3"
    mid_year_ref = f"Assumptions!$B${ROW['Mid-Year Convention (1=On, 0=Off)']}"

    for row_idx, wacc_val in enumerate(wacc_values, start=5):
        wacc_cell = sens_tab.cell(row=row_idx, column=1, value=wacc_val)
        wacc_cell.number_format = "0.00%"
        wacc_cell.font, wacc_cell.fill = HEADER_FONT, HEADER_FILL
        wacc_cell.alignment = Alignment(horizontal="center")

        for col_idx in range(2, 2 + len(growth_values)):
            g_ref = f"{get_column_letter(col_idx)}$4"
            w_ref = f"$A{row_idx}"
            # SUMPRODUCT, not NPV() -- NPV() always assumes end-of-year cash
            # flows, so it can't reflect the mid-year convention toggle the
            # main DCF tab applies. This mirrors that toggle exactly.
            formula = (
                f"=(SUMPRODUCT({fcf_range},1/(1+{w_ref})^({fcf_year_range}-0.5*{mid_year_ref}))"
                f"+(DCF!{last_col}7*(1+{g_ref})/({w_ref}-{g_ref}))"
                f"/(1+{w_ref})^(DCF!{last_col}3-0.5*{mid_year_ref})"
                f"-Assumptions!$B${ROW['Net Debt']})"
                f"/Assumptions!$B${ROW['Shares Outstanding']}"
            )
            _write_formula(sens_tab, f"{get_column_letter(col_idx)}{row_idx}", formula, fmt["price"])

    sens_tab.column_dimensions["A"].width = 16
    for col_idx in range(2, 2 + len(growth_values)):
        sens_tab.column_dimensions[get_column_letter(col_idx)].width = 13

    # ============================================
    # CHECKS TAB — every internal consistency check, in one place
    # ============================================
    checks_tab = wb.create_sheet("Checks")
    checks_tab.sheet_view.showGridLines = False
    checks_tab["A1"] = "Model Checks"
    checks_tab["A1"].font = TITLE_FONT
    checks_tab.merge_cells("A1:C1")
    for col, header in enumerate(["Check", "Result", "Status"], start=1):
        cell = checks_tab.cell(row=3, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL

    mid_row = 5 + (len(wacc_values) // 2)
    mid_col_letter = get_column_letter(2 + (len(growth_values) // 2))
    checks = [
        ("Enterprise Value = Sum PV FCF + PV Terminal Value",
         f"=ROUND(B{v_row + 3}-(B{v_row}+B{v_row + 2}),2)",
         f'=IF(ABS(B4)<0.01,"PASS","FAIL")'),
        ("Sensitivity center cell matches main Implied Price (only valid when Scenario = Base)",
         f"=ROUND(Sensitivity!{mid_col_letter}{mid_row}-{implied_price_cell},2)",
         f'=IF({A("Scenario (1=Bear, 2=Base, 3=Bull)")}<>2,"N/A (not Base)",IF(ABS(B5)<0.01,"PASS","FAIL"))'),
        ("WACC weights sum to 100%",
         f"=ROUND(Assumptions!$B${CALC_ROW['Weight of Equity (E/V)']}+Assumptions!$B${CALC_ROW['Weight of Debt (D/V)']}-1,4)",
         f'=IF(ABS(B6)<0.01,"PASS","FAIL")'),
    ]
    for i, (label, formula, status_formula) in enumerate(checks, start=4):
        checks_tab.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_formula(checks_tab, f"B{i}", formula, "0.0000")
        _write_formula(checks_tab, f"C{i}", status_formula)
        checks_tab.cell(row=i, column=3).font = Font(bold=True, color="1B7A1B")
    checks_tab.column_dimensions["A"].width = 48
    checks_tab.column_dimensions["B"].width = 14
    checks_tab.column_dimensions["C"].width = 10

    # ============================================
    # SUMMARY TAB
    # ============================================
    summary = wb.create_sheet("Summary", 0)
    summary.sheet_view.showGridLines = False
    summary["A1"] = f"{result['company_name']} ({result['ticker']}) — DCF Valuation Summary"
    summary["A1"].font = TITLE_FONT
    summary.merge_cells("A1:B1")

    summary_rows = [
        ("Implied Share Price", f"={implied_price_cell}", fmt["price"], True),
        ("Current Share Price", f"=Assumptions!$B${ROW['Current Share Price']}", fmt["price"], False),
        ("Upside / (Downside)", f"={upside_cell}", fmt["percent"], True),
        ("Enterprise Value", f"={ev_cell}", fmt["currency"], False),
        ("Net Debt", f"=Assumptions!$B${ROW['Net Debt']}", fmt["currency"], False),
        ("Shares Outstanding", f"=Assumptions!$B${ROW['Shares Outstanding']}", "#,##0", False),
        ("WACC (Derived)", f"={F('WACC (Used)')}", "0.00%", False),
        ("Scenario", f"=CHOOSE({A('Scenario (1=Bear, 2=Base, 3=Bull)')},\"Bear\",\"Base\",\"Bull\")", "General", False),
        ("Terminal Growth Rate", f"=Assumptions!$B${ROW['Terminal (Long-Run) Growth Rate']}", "0.00%", False),
    ]
    for i, (label, formula, number_fmt, is_key) in enumerate(summary_rows, start=3):
        summary.cell(row=i, column=1, value=label).font = LABEL_FONT if not is_key else Font(bold=True, size=12, color="1F4E78")
        _write_formula(summary, f"B{i}", formula, number_fmt, bold=is_key)
        summary[f"A{i}"].border = THIN_BORDER
        summary[f"B{i}"].border = THIN_BORDER

    note_row = 3 + len(summary_rows) + 1
    summary.cell(row=note_row, column=1, value="Change the Scenario cell on the Assumptions tab (1/2/3) to flip Bear/Base/Bull.")
    summary.cell(row=note_row, column=1).font = Font(italic=True, size=9, color="666666")
    note_row += 1
    if result["used_fallback_growth"] or result["used_fallback_margin"]:
        summary.cell(row=note_row, column=1,
                     value="Note: limited historical data — some assumptions used generic fallback values.")
        summary.cell(row=note_row, column=1).font = Font(italic=True, size=9, color="806000")
        note_row += 1
    if result["is_negative_margin"]:
        summary.cell(row=note_row, column=1,
                     value="Caution: negative FCF margin — likely early-stage/capex-heavy. Treat as one data point, not a verdict.")
        summary.cell(row=note_row, column=1).font = Font(italic=True, size=9, color="9C0000")
        summary.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=4)

    price_chart = BarChart()
    price_chart.title = "Implied Fair Value vs Current Price"
    price_chart.style = 12
    price_chart.height, price_chart.width = 7, 12
    price_chart.y_axis.number_format = fmt["currency"]
    price_chart.y_axis.delete = False
    price_chart.x_axis.delete = False

    summary["D3"], summary["D4"] = "Implied Price", "Current Price"
    _write_formula(summary, "E3", f"={implied_price_cell}", fmt["price"])
    _write_formula(summary, "E4", f"=Assumptions!$B${ROW['Current Share Price']}", fmt["price"])
    price_chart.add_data(Reference(summary, min_col=5, min_row=3, max_row=4))
    price_chart.set_categories(Reference(summary, min_col=4, min_row=3, max_row=4))
    summary.add_chart(price_chart, "D6")

    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 20

    return wb



def add_scenarios_tab(wb, scenarios):
    fmt = currency_formats(next(iter(scenarios.values()))["currency_symbol"])
    tab = wb.create_sheet("Scenarios")
    tab.sheet_view.showGridLines = False

    tab["A1"] = "Bear / Base / Bull Comparison"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:D1")

    for col, header in enumerate(["Metric", "Bear", "Base", "Bull"], start=1):
        cell = tab.cell(row=3, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    metrics = [
        ("Implied Share Price", "implied_share_price", fmt["price"]),
        ("Current Price", "current_price", fmt["price"]),
        ("Upside / Downside", "upside", fmt["percent"]),
    ]
    for row_offset, (label, key, number_fmt) in enumerate(metrics, start=4):
        tab.cell(row=row_offset, column=1, value=label).font = LABEL_FONT
        for col_offset, scenario_name in enumerate(["Bear", "Base", "Bull"], start=2):
            cell = tab.cell(row=row_offset, column=col_offset, value=scenarios[scenario_name][key])
            cell.number_format = number_fmt
            cell.alignment = Alignment(horizontal="center")

    for col_letter, width in zip("ABCD", [22, 16, 16, 16]):
        tab.column_dimensions[col_letter].width = width

    chart = BarChart()
    chart.title = "Implied Share Price: Bear vs Base vs Bull"
    chart.style = 11
    chart.y_axis.number_format = fmt["price"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16
    chart.add_data(Reference(tab, min_col=1, max_col=4, min_row=4, max_row=4), titles_from_data=True, from_rows=True)
    chart.set_categories(Reference(tab, min_col=2, max_col=4, min_row=3, max_row=3))
    tab.add_chart(chart, "F3")


def add_sensitivity_tab(wb, sheet_name, title, sensitivity_data, row_label, col_label,
                         currency_symbol="$", row_fmt='0.0%', col_fmt='0.0%'):
    fmt = currency_formats(currency_symbol)
    tab = wb.create_sheet(sheet_name)
    tab.sheet_view.showGridLines = False
    tab["A1"] = title
    tab["A1"].font = TITLE_FONT

    row_values = list(sensitivity_data.keys())
    col_values = list(next(iter(sensitivity_data.values())).keys())

    corner = tab.cell(row=3, column=1, value=f"{row_label} \\ {col_label}")
    corner.font = Font(bold=True, size=9)
    corner.fill = CORNER_FILL
    corner.alignment = Alignment(horizontal="center", wrap_text=True)

    for col_index, col_val in enumerate(col_values, start=2):
        cell = tab.cell(row=3, column=col_index, value=col_val)
        cell.number_format, cell.font, cell.fill = col_fmt, HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row_index, row_val in enumerate(row_values, start=4):
        header = tab.cell(row=row_index, column=1, value=row_val)
        header.number_format, header.font, header.fill = row_fmt, HEADER_FONT, HEADER_FILL
        header.alignment = Alignment(horizontal="center")
        for col_index, col_val in enumerate(col_values, start=2):
            cell = tab.cell(row=row_index, column=col_index, value=sensitivity_data[row_val][col_val])
            cell.number_format = fmt["price"]
            cell.alignment = Alignment(horizontal="center")

    tab.column_dimensions["A"].width = 14
    for col_index in range(2, 2 + len(col_values)):
        tab.column_dimensions[tab.cell(row=3, column=col_index).column_letter].width = 12


def export_comps_to_excel(comps, ticker_symbol):
    """
    A proper comps book, not a flat multiples table: full operating and
    valuation profile per company (revenue growth, margins, FCF yield,
    Rule of 40, net leverage), plus a percentile-band valuation summary
    (25th/median/mean/75th, not just a single average) with implied EV,
    equity value, AND per-share price at each band for each multiple —
    everything a live formula, all built from real reported figures.
    """
    wb = Workbook()
    tab = wb.active
    tab.title = "Comps"
    tab.sheet_view.showGridLines = False

    target = comps["target"]
    fmt = currency_formats("$")

    tab["A1"] = f"{target['company_name']} ({ticker_symbol}) — Trading Comps"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:N1")
    tab["A2"] = "Blue cells are raw inputs; every multiple, margin, and the valuation summary below are live formulas."
    tab["A2"].font = Font(italic=True, size=9, color="666666")

    headers = ["Ticker", "Company", "Enterprise Value", "Market Cap", "Revenue", "EBITDA", "EBIT (est.)", "FCF",
               "EV/Rev", "EV/EBITDA", "EV/EBIT", "P/E", "Rev Growth", "EBITDA Margin", "FCF Yield", "Net Leverage", "Rule of 40"]
    for col, header in enumerate(headers, start=1):
        cell = tab.cell(row=4, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    target_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    def write_company_row(row, company, is_target=False):
        tab.cell(row=row, column=1, value=company["ticker"])
        tab.cell(row=row, column=2, value=company["company_name"])
        _write_input(tab, f"C{row}", company["enterprise_value"], fmt["currency"], bold=is_target)
        _write_input(tab, f"D{row}", company["market_cap"], fmt["currency"], bold=is_target)
        _write_input(tab, f"E{row}", company["revenue"], fmt["currency"], bold=is_target)
        _write_input(tab, f"F{row}", company["ebitda"], fmt["currency"], bold=is_target)
        _write_input(tab, f"G{row}", company["ebit_est"], fmt["currency"], bold=is_target)
        _write_input(tab, f"H{row}", company["fcf"], fmt["currency"], bold=is_target)
        _write_formula(tab, f"I{row}", f"=IFERROR(C{row}/E{row},\"N/A\")", '0.00"x"', bold=is_target)
        _write_formula(tab, f"J{row}", f"=IFERROR(C{row}/F{row},\"N/A\")", '0.00"x"', bold=is_target)
        _write_formula(tab, f"K{row}", f"=IFERROR(C{row}/G{row},\"N/A\")", '0.00"x"', bold=is_target)
        _write_input(tab, f"L{row}", company["pe_ratio"], '0.00"x"', bold=is_target)
        _write_input(tab, f"M{row}", company["revenue_growth"], "0.0%", bold=is_target)
        _write_formula(tab, f"N{row}", f"=IFERROR(F{row}/E{row},\"N/A\")", "0.0%", bold=is_target)
        _write_formula(tab, f"O{row}", f"=IFERROR(H{row}/D{row},\"N/A\")", "0.0%", bold=is_target)
        # Net Leverage = (Total Debt - Cash) / EBITDA -- approximated as
        # (EV - Market Cap) / EBITDA, since EV - MktCap = Net Debt algebraically.
        _write_formula(tab, f"P{row}", f"=IFERROR((C{row}-D{row})/F{row},\"N/A\")", '0.00"x"', bold=is_target)
        _write_formula(tab, f"Q{row}", f"=IFERROR(M{row}+O{row},\"N/A\")", "0.0%", bold=is_target)
        if is_target:
            for col in range(1, 18):
                tab.cell(row=row, column=col).fill = target_fill
        for col in (1, 2):
            tab.cell(row=row, column=col).alignment = Alignment(horizontal="left")
        for col in range(3, 18):
            tab.cell(row=row, column=col).alignment = Alignment(horizontal="center")

    write_company_row(5, target, is_target=True)
    for i, peer in enumerate(comps["peers"], start=6):
        write_company_row(i, peer)

    first_peer_row, last_peer_row = 6, 5 + len(comps["peers"])

    # --- Percentile bands (25th/median/mean/75th), not a flat average ---
    band_row = last_peer_row + 2
    tab.cell(row=band_row - 1, column=1, value="Peer Valuation Bands").font = Font(bold=True, size=12, color="1F4E78")
    band_headers = ["Multiple", "25th Pctl.", "Median", "Mean", "75th Pctl."]
    for col, h in enumerate(band_headers, start=1):
        cell = tab.cell(row=band_row, column=col, value=h)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    band_defs = [("EV/Revenue", "I"), ("EV/EBITDA", "J"), ("EV/EBIT", "K"), ("P/E", "L")]
    for i, (label, col_letter) in enumerate(band_defs, start=band_row + 1):
        tab.cell(row=i, column=1, value=label).font = LABEL_FONT
        rng = f"{col_letter}{first_peer_row}:{col_letter}{last_peer_row}"
        _write_formula(tab, f"B{i}", f'=IFERROR(PERCENTILE({rng},0.25),"N/A")', '0.00"x"')
        _write_formula(tab, f"C{i}", f'=IFERROR(MEDIAN({rng}),"N/A")', '0.00"x"')
        _write_formula(tab, f"D{i}", f'=IFERROR(AVERAGE({rng}),"N/A")', '0.00"x"')
        _write_formula(tab, f"E{i}", f'=IFERROR(PERCENTILE({rng},0.75),"N/A")', '0.00"x"')
    ev_rev_band_row, ev_ebitda_band_row = band_row + 1, band_row + 2

    # --- Implied valuation: EV -> Equity Value -> Per-Share, at each band ---
    val_row = band_row + len(band_defs) + 2
    tab.cell(row=val_row - 1, column=1, value="Implied Valuation (from EV/Revenue)").font = Font(bold=True, size=12, color="1F4E78")
    val_headers = ["", "25th Pctl.", "Median", "Mean", "75th Pctl."]
    for col, h in enumerate(val_headers, start=1):
        if h:
            cell = tab.cell(row=val_row, column=col, value=h)
            cell.font, cell.fill = HEADER_FONT, HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

    tab.cell(row=val_row + 1, column=1, value="Implied Enterprise Value").font = LABEL_FONT
    tab.cell(row=val_row + 2, column=1, value="Implied Equity Value").font = LABEL_FONT
    tab.cell(row=val_row + 3, column=1, value="Implied Price / Share").font = Font(bold=True, color="1F4E78")
    tab.cell(row=val_row + 5, column=1, value="Target Shares Outstanding").font = LABEL_FONT
    _write_input(tab, f"B{val_row + 5}", target.get("shares_outstanding"), "#,##0")

    for col_letter in ("B", "C", "D", "E"):
        _write_formula(tab, f"{col_letter}{val_row + 1}", f"=IFERROR({col_letter}{ev_rev_band_row}*E5,\"N/A\")", fmt["currency"], bold=True)
        _write_formula(tab, f"{col_letter}{val_row + 2}", f'=IFERROR({col_letter}{val_row + 1}-(C5-D5),"N/A")', fmt["currency"])
        _write_formula(tab, f"{col_letter}{val_row + 3}", f'=IFERROR({col_letter}{val_row + 2}/$B${val_row + 5},"N/A")', fmt["price"], bold=True)

    for col_letter, width in zip("ABCDEFGHIJKLMNOPQ", [10, 22, 15, 15, 14, 13, 13, 13, 9, 10, 9, 8, 10, 12, 10, 11, 10]):
        tab.column_dimensions[col_letter].width = width

    return wb


def export_football_field_to_excel(field):
    """
    A real banking valuation summary page: a Low/Mid/High table per
    method, plus a horizontal range chart — built with the standard
    Excel technique for this (a stacked bar where the first series, the
    "Low" value, is invisible, and the second series, "High minus Low",
    is the visible colored bar) — with the current share price overlaid
    as a reference line.
    """
    wb = Workbook()
    fmt = currency_formats(field["currency_symbol"])
    tab = wb.active
    tab.title = "Football Field"
    tab.sheet_view.showGridLines = False

    tab["A1"] = f"{field['ticker']} — Valuation Summary"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:E1")

    if not field["methods"]:
        tab["A3"] = "No valuation methods had enough data to compute a range for this company."
        tab["A3"].font = Font(italic=True, color="9C0000")
        return wb

    headers = ["Method", "Low", "Mid", "High", "Low (chart helper)"]
    for col, h in enumerate(headers, start=1):
        cell = tab.cell(row=3, column=col, value=h)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for i, m in enumerate(field["methods"], start=4):
        tab.cell(row=i, column=1, value=m["method"]).font = LABEL_FONT
        for col, key in [(2, "low"), (3, "mid"), (4, "high")]:
            cell = tab.cell(row=i, column=col, value=m[key])
            cell.number_format = fmt["price"]
            cell.alignment = Alignment(horizontal="center")
        # Chart helper column: identical to Low, used as the invisible
        # base of the stacked bar so only the Low-to-High range shows.
        _write_formula(tab, f"E{i}", f"=B{i}", fmt["price"])

    last_row = 3 + len(field["methods"])
    if field["current_price"]:
        tab.cell(row=last_row + 2, column=1, value="Current Share Price").font = Font(bold=True, color="1F4E78")
        cell = tab.cell(row=last_row + 2, column=2, value=field["current_price"])
        cell.number_format = fmt["price"]
        cell.font = Font(bold=True, color="1F4E78")

    for col_letter, width in zip("ABCDE", [26, 14, 14, 14, 16]):
        tab.column_dimensions[col_letter].width = width

    # --- Range chart: stacked bar, "Low" invisible + "High-Low" visible ---
    chart = BarChart()
    chart.type = "bar"  # horizontal bars, the standard football-field orientation
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"{field['ticker']} Valuation Range by Method"
    chart.style = 10
    chart.x_axis.number_format = fmt["price"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16

    # Series 1: invisible "Low" base
    low_data = Reference(tab, min_col=5, min_row=3, max_row=last_row)
    chart.add_data(low_data, titles_from_data=True)
    chart.series[0].graphicalProperties.noFill = True

    # Series 2: visible "High - Low" range -- computed as a formula column
    range_col = 6
    tab.cell(row=3, column=range_col, value="Range (chart helper)")
    for i in range(4, last_row + 1):
        _write_formula(tab, f"F{i}", f"=D{i}-B{i}", fmt["price"])
    range_data = Reference(tab, min_col=range_col, min_row=3, max_row=last_row)
    chart.add_data(range_data, titles_from_data=True)
    chart.series[1].graphicalProperties.solidFill = "1F4E78"

    categories = Reference(tab, min_col=1, min_row=4, max_row=last_row)
    chart.set_categories(categories)
    chart.legend = None
    tab.add_chart(chart, "A" + str(last_row + 5))

    return wb


# Reuses run_dcf() and build_comps_table() on both acquirer and target for
# standalone valuation context, then layers deal mechanics on top: offer
# price, sources & uses, synergies, and pro-forma accretion/dilution.
# ============================================

def run_lbo_model(ticker_symbol, entry_multiple=None, leverage_multiple=4.5, hold_years=5, exit_multiple=None):
    """
    Gathers the real company data an LBO needs — entry EBITDA/revenue,
    current EV (used as a sensible default entry multiple anchor) — then
    hands off to the live-formula Excel export for the actual mechanics
    (sources & uses, debt paydown via FCF sweep, exit, returns). The
    entry multiple defaults to the company's OWN current EV/EBITDA
    (a real, market-observed figure), not an arbitrary assumption.
    """
    company = resolve_company(ticker_symbol)
    ticker = company["ticker_obj"]
    info = company["info"]
    price_divisor = company["price_divisor"]

    income_statement = ticker.financials
    most_recent_revenue = safe_get_row(income_statement, REVENUE_ROW_NAMES)
    if most_recent_revenue is None:
        raise ValueError(f"No revenue data found for '{ticker_symbol}'.")

    ebitda = info.get("ebitda")
    if not ebitda or ebitda <= 0:
        raise ValueError(f"No usable EBITDA figure for '{ticker_symbol}' — can't size an LBO entry multiple without it.")

    enterprise_value = info.get("enterpriseValue")
    market_ev_ebitda = enterprise_value / ebitda if enterprise_value else None
    if entry_multiple is None:
        entry_multiple = market_ev_ebitda or 10.0  # generic fallback if EV unavailable
    if exit_multiple is None:
        exit_multiple = entry_multiple  # standard simplifying assumption: exit at the same multiple you entered at

    assumptions = estimate_dcf_assumptions(ticker, info)  # reuse the same growth/margin estimation logic
    three_stmt_assumptions = estimate_three_statement_assumptions(ticker, info)  # for D&A and CapEx specifically

    balance_sheet = ticker.balance_sheet
    existing_debt = safe_get_row(balance_sheet, DEBT_ROW_NAMES, default=0) or 0

    return {
        "ticker": ticker_symbol,
        "company_name": company["company_name"],
        "currency_symbol": get_currency_symbol(company["financial_currency"]),
        "most_recent_revenue": most_recent_revenue,
        "entry_ebitda": ebitda,
        "market_ev_ebitda": market_ev_ebitda,
        "entry_multiple": entry_multiple,
        "exit_multiple": exit_multiple,
        "leverage_multiple": leverage_multiple,
        "hold_years": hold_years,
        "existing_debt": existing_debt,
        "assumptions_used": {
            "starting_growth_rate": assumptions["starting_growth_rate"],
            "ending_growth_rate": assumptions["ending_growth_rate"],
            "starting_fcf_margin": assumptions["starting_fcf_margin"],  # reused as EBITDA margin proxy
            "ending_fcf_margin": assumptions["ending_fcf_margin"],
            "da_pct_revenue": three_stmt_assumptions["da_pct_revenue"],
            "capex_pct_revenue": three_stmt_assumptions["capex_pct_revenue"],
            "tax_rate": three_stmt_assumptions["tax_rate"],
        },
    }


def export_lbo_to_excel(result):
    """
    A genuine LBO model: entry sourced from the company's OWN market
    EV/EBITDA (not an arbitrary assumption), debt sized off a leverage
    multiple, and a real circular debt schedule — interest depends on
    average debt balance, but ending debt (after mandatory amortization
    and a 100% FCF cash sweep, the standard LBO convention) depends on
    net income, which depends on interest. Resolved via Excel's iterative
    calculation, the same technique verified to converge correctly for
    the 3-statement model's debt schedule.
    """
    wb = Workbook()
    fmt = currency_formats(result["currency_symbol"])
    a = result["assumptions_used"]
    n = result["hold_years"]

    wb.calculation.iterate = True
    wb.calculation.iterateCount = 100
    wb.calculation.iterateDelta = 0.001

    # ============================================
    # ASSUMPTIONS TAB
    # ============================================
    assump = wb.active
    assump.title = "Assumptions"
    assump.sheet_view.showGridLines = False
    assump["A1"] = f"{result['company_name']} ({result['ticker']}) — LBO Model"
    assump["A1"].font = TITLE_FONT
    assump.merge_cells("A1:C1")
    assump["A2"] = "Blue cells are hardcoded inputs. Debt paydown uses a 100% FCF cash sweep after mandatory " \
                   "amortization — the standard LBO convention. Circular (interest <-> debt) — resolved via " \
                   "Excel's iterative calculation, enabled on this workbook."
    assump["A2"].font = Font(italic=True, size=9, color="666666")

    input_rows = [
        ("Most Recent Revenue (Year 0)", result["most_recent_revenue"], fmt["currency"]),
        ("Entry EBITDA", result["entry_ebitda"], fmt["currency"]),
        ("Entry EV / EBITDA Multiple", result["entry_multiple"], '0.00"x"'),
        ("Exit EV / EBITDA Multiple", result["exit_multiple"], '0.00"x"'),
        ("Leverage (Debt / EBITDA)", result["leverage_multiple"], '0.00"x"'),
        ("Transaction Fees (% of EV)", 0.02, "0.0%"),
        ("Hold Period (Years)", n, "0"),
        ("Revenue Growth Rate (Start)", a["starting_growth_rate"], "0.0%"),
        ("Revenue Growth Rate (End)", a["ending_growth_rate"], "0.0%"),
        ("EBITDA Margin (Start)", a["starting_fcf_margin"], "0.0%"),
        ("EBITDA Margin (End)", a["ending_fcf_margin"], "0.0%"),
        ("D&A % of Revenue", a["da_pct_revenue"], "0.0%"),
        ("CapEx % of Revenue", a["capex_pct_revenue"], "0.0%"),
        ("Tax Rate", a["tax_rate"], "0.0%"),
        ("Interest Rate on Debt", 0.08, "0.00%"),  # LBO debt typically prices above the 3-statement default
        ("Mandatory Amortization %", 0.05, "0.0%"),
        ("Cash Sweep % (of FCF after mandatory)", 1.00, "0.0%"),
    ]
    for i, (label, value, number_fmt) in enumerate(input_rows, start=4):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_input(assump, f"B{i}", value, number_fmt)
    ROW = {label: i for i, (label, *_) in enumerate(input_rows, start=4)}

    def A(label):
        return f"Assumptions!$B${ROW[label]}"

    # --- Derived: Entry Sources & Uses ---
    calc_row = 4 + len(input_rows) + 1
    assump.cell(row=calc_row - 1, column=1, value="Entry Sources & Uses").font = Font(bold=True, size=11, color="1F4E78")
    calc_rows = [
        ("Entry Enterprise Value", f"={A('Entry EBITDA')}*{A('Entry EV / EBITDA Multiple')}", fmt["currency"]),
        ("New Debt", f"={A('Entry EBITDA')}*{A('Leverage (Debt / EBITDA)')}", fmt["currency"]),
        ("Transaction Fees", f"=B{calc_row}*{A('Transaction Fees (% of EV)')}", fmt["currency"]),
    ]
    for i, (label, formula, number_fmt) in enumerate(calc_rows, start=calc_row):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_formula(assump, f"B{i}", formula, number_fmt)
    CALC = {label: i for i, (label, *_) in enumerate(calc_rows, start=calc_row)}

    total_uses_row = calc_row + len(calc_rows)
    assump.cell(row=total_uses_row, column=1, value="Total Uses (EV + Fees)").font = LABEL_FONT
    _write_formula(assump, f"B{total_uses_row}", f"=B{CALC['Entry Enterprise Value']}+B{CALC['Transaction Fees']}", fmt["currency"])

    sponsor_equity_row = total_uses_row + 1
    assump.cell(row=sponsor_equity_row, column=1, value="Sponsor Equity (plug)").font = Font(bold=True, color="1F4E78")
    _write_formula(assump, f"B{sponsor_equity_row}", f"=B{total_uses_row}-B{CALC['New Debt']}", fmt["currency"], bold=True)

    assump.column_dimensions["A"].width = 38
    assump.column_dimensions["B"].width = 20

    # ============================================
    # MODEL TAB
    # ============================================
    model = wb.create_sheet("Model", 0)
    model.sheet_view.showGridLines = False
    model.freeze_panes = "C4"
    model["A1"] = f"{result['ticker']} — {n}-Year LBO Model"
    model["A1"].font = TITLE_FONT
    model.merge_cells(f"A1:{get_column_letter(2 + n)}1")

    year_cols = [get_column_letter(3 + i) for i in range(n)]

    model["A3"] = "Year Index"
    model["B3"] = 0
    prev_col = "B"
    for col in year_cols:
        _write_formula(model, f"{col}3", f"={prev_col}3+1")
        prev_col = col

    def section(row, label):
        model.cell(row=row, column=1, value=label).font = Font(bold=True, size=12, color="1F4E78")

    def line(row, label, formula_fn, number_fmt=None, bold=False, year0_value=None):
        nfmt = number_fmt if number_fmt is not None else fmt["currency"]
        model.cell(row=row, column=1, value=label).font = Font(bold=True) if bold else LABEL_FONT
        if year0_value is not None:
            _write_input(model, f"B{row}", year0_value, nfmt)
        prev_col = "B"
        for col in year_cols:
            _write_formula(model, f"{col}{row}", formula_fn(col, prev_col), nfmt, bold=bold)
            prev_col = col

    # --- Operating build ---
    section(5, "OPERATING BUILD")
    line(6, "Revenue Growth Rate", lambda c, p: (
        f"={A('Revenue Growth Rate (Start)')}+({A('Revenue Growth Rate (End)')}-{A('Revenue Growth Rate (Start)')})"
        f"*({c}3-1)/({A('Hold Period (Years)')}-1)" if n > 1 else f"={A('Revenue Growth Rate (Start)')}"
    ), "0.0%")
    line(7, "Revenue", lambda c, p: f"={p}7*(1+{c}6)", bold=True, year0_value=result["most_recent_revenue"])
    line(8, "EBITDA Margin", lambda c, p: (
        f"={A('EBITDA Margin (Start)')}+({A('EBITDA Margin (End)')}-{A('EBITDA Margin (Start)')})"
        f"*({c}3-1)/({A('Hold Period (Years)')}-1)" if n > 1 else f"={A('EBITDA Margin (Start)')}"
    ), "0.0%")
    line(9, "EBITDA", lambda c, p: f"={c}7*{c}8", bold=True)
    line(10, "D&A", lambda c, p: f"={c}7*{A('D&A % of Revenue')}")
    line(11, "EBIT", lambda c, p: f"={c}9-{c}10", bold=True)
    line(12, "CapEx", lambda c, p: f"={c}7*{A('CapEx % of Revenue')}")

    # --- Debt schedule (the circular part) ---
    section(14, "DEBT SCHEDULE")
    line(15, "Beginning Debt", lambda c, p: f"={p}20")
    _write_formula(model, "B15", f"=Assumptions!$B${CALC['New Debt']}", fmt["currency"])
    line(16, "Interest Expense", lambda c, p: f"=(({c}15+{c}20)/2)*{A('Interest Rate on Debt')}")
    line(17, "Mandatory Amortization", lambda c, p: f"=MIN({c}15*{A('Mandatory Amortization %')},{c}15)")
    line(18, "FCF Available for Sweep", lambda c, p: f"={c}23-{c}17")
    line(19, "Cash Sweep", lambda c, p: f"=MAX(MIN({c}18*{A('Cash Sweep % (of FCF after mandatory)')},{c}15-{c}17),0)")
    line(20, "Ending Debt", lambda c, p: f"={c}15-{c}17-{c}19", bold=True)
    _write_formula(model, "B20", f"=Assumptions!$B${CALC['New Debt']}", fmt["currency"], bold=True)

    # --- Income statement (part 2, depends on interest) ---
    line(22, "Pretax Income", lambda c, p: f"={c}11-{c}16")
    _write_formula(model, "B22", "=0", fmt["currency"])  # Year 0 is the entry point; no operations that year
    line(23, "Free Cash Flow (before sweep)", lambda c, p: f"=({c}22-MAX({c}22,0)*{A('Tax Rate')})+{c}10-{c}12")
    # Note: FCF here omits change in working capital for simplicity (a real
    # limitation, not an oversight) — the same simplification level as a
    # typical "quick" LBO, not a full diligence model.

    for col in year_cols:
        model.column_dimensions[col].width = 16
    model.column_dimensions["A"].width = 32
    model.column_dimensions["B"].width = 16

    # ============================================
    # RETURNS
    # ============================================
    last_col = year_cols[-1]
    r_row = 26
    model.cell(row=r_row, column=1, value="EXIT & RETURNS").font = Font(bold=True, size=12, color="1F4E78")
    model.cell(row=r_row + 1, column=1, value="Exit EBITDA (Final Year)").font = LABEL_FONT
    _write_formula(model, f"B{r_row + 1}", f"={last_col}9", fmt["currency"])
    model.cell(row=r_row + 2, column=1, value="Exit Enterprise Value").font = LABEL_FONT
    _write_formula(model, f"B{r_row + 2}", f"=B{r_row + 1}*Assumptions!$B${ROW['Exit EV / EBITDA Multiple']}", fmt["currency"])
    model.cell(row=r_row + 3, column=1, value="Less: Exit Debt").font = LABEL_FONT
    _write_formula(model, f"B{r_row + 3}", f"=-{last_col}20", fmt["currency"])
    model.cell(row=r_row + 4, column=1, value="Exit Equity Value").font = Font(bold=True)
    _write_formula(model, f"B{r_row + 4}", f"=B{r_row + 2}+B{r_row + 3}", fmt["currency"], bold=True)
    model.cell(row=r_row + 5, column=1, value="Sponsor Equity (Entry)").font = LABEL_FONT
    _write_formula(model, f"B{r_row + 5}", f"=Assumptions!$B${sponsor_equity_row}", fmt["currency"])
    model.cell(row=r_row + 6, column=1, value="MOIC").font = Font(bold=True, size=12, color="1F4E78")
    _write_formula(model, f"B{r_row + 6}", f"=B{r_row + 4}/B{r_row + 5}", '0.00"x"', bold=True)
    model.cell(row=r_row + 7, column=1, value="IRR").font = Font(bold=True, size=12, color="1F4E78")
    _write_formula(model, f"B{r_row + 7}", f"=B{r_row + 6}^(1/Assumptions!$B${ROW['Hold Period (Years)']})-1", "0.0%", bold=True)

    # --- Chart: Debt paydown over the hold period ---
    chart = BarChart()
    chart.title = "Debt Paydown Over Hold Period"
    chart.style = 10
    chart.y_axis.number_format = fmt["currency"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16
    last_col_idx = 2 + n
    debt_data = Reference(model, min_col=1, max_col=last_col_idx, min_row=20, max_row=20)
    chart.add_data(debt_data, titles_from_data=True, from_rows=True)
    categories = Reference(model, min_col=3, max_col=last_col_idx, min_row=3, max_row=3)
    chart.set_categories(categories)
    model.add_chart(chart, f"{get_column_letter(last_col_idx + 2)}5")

    # ============================================
    # CHECKS TAB
    # ============================================
    checks_tab = wb.create_sheet("Checks")
    checks_tab.sheet_view.showGridLines = False
    checks_tab["A1"] = "Model Checks"
    checks_tab["A1"].font = TITLE_FONT
    checks_tab.merge_cells("A1:C1")
    for col, header in enumerate(["Check", "Result", "Status"], start=1):
        cell = checks_tab.cell(row=3, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL

    checks = [
        ("Sources = Uses at entry", f"=ROUND((Assumptions!$B${CALC['New Debt']}+Assumptions!$B${sponsor_equity_row})-Assumptions!$B${total_uses_row},2)"),
        ("Ending debt never negative in any year",
         f"=ROUND(MIN(C20:{last_col}20),2)"),
    ]
    for i, (label, formula) in enumerate(checks, start=4):
        checks_tab.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_formula(checks_tab, f"B{i}", formula, "0.0000")
        if i == 4:
            _write_formula(checks_tab, f"C{i}", f'=IF(ABS(B{i})<0.01,"PASS","FAIL")')
        else:
            _write_formula(checks_tab, f"C{i}", f'=IF(B{i}>=-0.01,"PASS","FAIL")')
        checks_tab.cell(row=i, column=3).font = Font(bold=True, color="1B7A1B")
    checks_tab.column_dimensions["A"].width = 40
    checks_tab.column_dimensions["B"].width = 16
    checks_tab.column_dimensions["C"].width = 10

    # ============================================
    # SUMMARY TAB
    # ============================================
    summary = wb.create_sheet("Summary", 0)
    summary.sheet_view.showGridLines = False
    summary["A1"] = f"{result['company_name']} ({result['ticker']}) — LBO Summary"
    summary["A1"].font = TITLE_FONT
    summary.merge_cells("A1:B1")

    summary_rows = [
        ("MOIC", f"Model!B{r_row + 6}", '0.00"x"', True),
        ("IRR", f"Model!B{r_row + 7}", "0.0%", True),
        ("Entry EV", f"Assumptions!B{CALC['Entry Enterprise Value']}", fmt["currency"], False),
        ("Sponsor Equity", f"Assumptions!B{sponsor_equity_row}", fmt["currency"], False),
        ("New Debt", f"Assumptions!B{CALC['New Debt']}", fmt["currency"], False),
        ("Leverage", f"Assumptions!B{ROW['Leverage (Debt / EBITDA)']}", '0.00"x"', False),
        ("Exit Equity Value", f"Model!B{r_row + 4}", fmt["currency"], False),
        ("Hold Period", f"Assumptions!B{ROW['Hold Period (Years)']}", "0", False),
    ]
    for i, (label, ref, number_fmt, is_key) in enumerate(summary_rows, start=3):
        summary.cell(row=i, column=1, value=label).font = LABEL_FONT if not is_key else Font(bold=True, size=12, color="1F4E78")
        _write_formula(summary, f"B{i}", f"={ref}", number_fmt, bold=is_key)
        summary[f"A{i}"].border = THIN_BORDER
        summary[f"B{i}"].border = THIN_BORDER

    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 20

    return wb


def run_ma_model(acquirer_ticker, target_ticker, offer_premium=0.30, pct_cash=0.5,
                  pct_stock=0.5, synergy_pct_of_target_revenue=0.03, cost_of_new_debt=0.06,
                  tax_rate=0.21, foregone_cash_return=0.02):
    """
    pct_cash + pct_stock must sum to 1.0 — the split of consideration between
    cash and newly issued acquirer stock.

    synergy_pct_of_target_revenue: run-rate cost/revenue synergies as a % of
    target revenue — deliberately simple (a single assumption you can shock),
    not a bottom-up synergy build.

    Cash consideration first draws down the acquirer's own balance-sheet
    cash; anything beyond that is modeled as new debt at cost_of_new_debt.
    """
    if abs((pct_cash + pct_stock) - 1.0) > 1e-6:
        raise ValueError("pct_cash + pct_stock must sum to 1.0")

    acquirer = resolve_company(acquirer_ticker)
    target = resolve_company(target_ticker)
    acquirer_info = acquirer["info"]
    target_info = target["info"]

    # --- Standalone valuation context (reused, not reimplemented) ---
    try:
        acquirer_dcf = run_dcf(acquirer_ticker)
    except Exception:
        acquirer_dcf = None
    try:
        target_dcf = run_dcf(target_ticker)
    except Exception:
        target_dcf = None

    target_comps = build_comps_table(target_ticker)

    # --- Deal terms ---
    target_current_price = (target_info.get("currentPrice") or target_info.get("regularMarketPrice") or 0) / target["price_divisor"]
    acquirer_current_price = (acquirer_info.get("currentPrice") or acquirer_info.get("regularMarketPrice") or 0) / acquirer["price_divisor"]

    offer_price_per_share = target_current_price * (1 + offer_premium)
    target_shares = target_info.get("sharesOutstanding")
    if not target_shares:
        raise ValueError(f"No shares outstanding for target '{target_ticker}' — cannot size the deal.")
    acquirer_shares = acquirer_info.get("sharesOutstanding")
    if not acquirer_shares:
        raise ValueError(f"No shares outstanding for acquirer '{acquirer_ticker}' — cannot size stock issuance.")

    total_consideration = offer_price_per_share * target_shares
    cash_consideration = total_consideration * pct_cash
    stock_consideration = total_consideration * pct_stock
    new_shares_issued = stock_consideration / acquirer_current_price if acquirer_current_price else 0

    # --- Sources & uses ---
    acquirer_cash_on_hand = safe_get_row(acquirer["ticker_obj"].balance_sheet, CASH_ROW_NAMES, default=0) or 0
    cash_from_balance_sheet = min(cash_consideration, acquirer_cash_on_hand)
    new_debt_raised = max(cash_consideration - acquirer_cash_on_hand, 0)

    # --- Synergies ---
    target_revenue = safe_get_row(target["ticker_obj"].financials, REVENUE_ROW_NAMES, default=0) or 0
    pretax_synergies = target_revenue * synergy_pct_of_target_revenue
    aftertax_synergies = pretax_synergies * (1 - tax_rate)

    # --- Financing costs ---
    aftertax_interest_on_new_debt = new_debt_raised * cost_of_new_debt * (1 - tax_rate)
    aftertax_foregone_income_on_cash = cash_from_balance_sheet * foregone_cash_return * (1 - tax_rate)

    # --- Purchase Price Allocation ---
    # Standard practice: excess of purchase price over book value of net
    # assets acquired gets allocated first to identifiable intangibles and
    # a step-up in PP&E fair value (both real, if judgment-driven,
    # assumptions), with whatever remains as goodwill (the residual, not
    # a separate estimate). The step-ups create NEW tax-deductible-in-book
    # (but not in tax basis) D&A/amortization, which is why a deferred tax
    # liability arises — and that incremental non-cash expense genuinely
    # reduces pro forma net income, which this now reflects in the
    # accretion/dilution math below, not just displayed separately.
    target_book_equity = safe_get_row(target["ticker_obj"].balance_sheet, EQUITY_ROW_NAMES, default=0) or 0
    excess_purchase_price = max(total_consideration - target_book_equity, 0)
    intangible_pct_of_excess = 0.40  # a common rule-of-thumb split; a real deal would size this from a valuation
    ppe_stepup_pct_of_excess = 0.10
    intangible_step_up = excess_purchase_price * intangible_pct_of_excess
    ppe_step_up = excess_purchase_price * ppe_stepup_pct_of_excess
    goodwill = excess_purchase_price - intangible_step_up - ppe_step_up  # residual, not separately estimated

    intangible_useful_life_years = 10
    ppe_useful_life_years = 15
    incremental_amortization = intangible_step_up / intangible_useful_life_years
    incremental_da_from_stepup = ppe_step_up / ppe_useful_life_years
    total_incremental_da_amort = incremental_amortization + incremental_da_from_stepup
    aftertax_incremental_da_amort = total_incremental_da_amort * (1 - tax_rate)

    deferred_tax_liability = (intangible_step_up + ppe_step_up) * tax_rate

    # --- Net income (yfinance summary field first, EPS * shares fallback) ---
    def net_income_for(info, shares):
        ni = info.get("netIncomeToCommon")
        if ni is None:
            eps = info.get("trailingEps")
            ni = eps * shares if (eps and shares) else None
        return ni

    acquirer_net_income = net_income_for(acquirer_info, acquirer_shares)
    target_net_income = net_income_for(target_info, target_shares)
    if acquirer_net_income is None or target_net_income is None:
        raise ValueError("Missing net income data for acquirer or target — cannot run accretion/dilution.")

    combined_net_income = (
        acquirer_net_income + target_net_income
        + aftertax_synergies
        - aftertax_interest_on_new_debt
        - aftertax_foregone_income_on_cash
        - aftertax_incremental_da_amort
    )

    pro_forma_diluted_shares = acquirer_shares + new_shares_issued
    pro_forma_eps = combined_net_income / pro_forma_diluted_shares if pro_forma_diluted_shares else None
    acquirer_standalone_eps = acquirer_info.get("trailingEps") or (
        acquirer_net_income / acquirer_shares if acquirer_shares else None
    )

    accretion_dilution_pct = None
    if pro_forma_eps is not None and acquirer_standalone_eps:
        accretion_dilution_pct = (pro_forma_eps - acquirer_standalone_eps) / abs(acquirer_standalone_eps)

    # --- Premium vs. independent valuation anchors ---
    premium_vs_dcf = None
    if target_dcf and target_dcf["implied_share_price"]:
        premium_vs_dcf = (offer_price_per_share - target_dcf["implied_share_price"]) / target_dcf["implied_share_price"]

    premium_vs_comps = None
    implied_comps_price = None
    if target_comps["implied_ev_from_revenue"] and target_shares:
        target_net_debt = (target_info.get("totalDebt", 0) or 0) - (target_info.get("totalCash", 0) or 0)
        implied_comps_equity = target_comps["implied_ev_from_revenue"] - target_net_debt
        implied_comps_price = implied_comps_equity / target_shares
        premium_vs_comps = (offer_price_per_share - implied_comps_price) / implied_comps_price

    return {
        "acquirer_ticker": acquirer_ticker,
        "target_ticker": target_ticker,
        "acquirer_name": acquirer["company_name"],
        "target_name": target["company_name"],
        "acquirer_currency_symbol": get_currency_symbol(acquirer["financial_currency"]),
        "target_currency_symbol": get_currency_symbol(target["financial_currency"]),
        "target_current_price": target_current_price,
        "acquirer_current_price": acquirer_current_price,
        "offer_premium": offer_premium,
        "pct_cash": pct_cash,
        "pct_stock": pct_stock,
        "target_shares": target_shares,
        "acquirer_shares": acquirer_shares,
        "acquirer_cash_on_hand": acquirer_cash_on_hand,
        "target_revenue": target_revenue,
        "target_book_equity": target_book_equity,
        "excess_purchase_price": excess_purchase_price,
        "intangible_step_up": intangible_step_up,
        "ppe_step_up": ppe_step_up,
        "goodwill": goodwill,
        "intangible_useful_life_years": intangible_useful_life_years,
        "ppe_useful_life_years": ppe_useful_life_years,
        "incremental_amortization": incremental_amortization,
        "incremental_da_from_stepup": incremental_da_from_stepup,
        "total_incremental_da_amort": total_incremental_da_amort,
        "aftertax_incremental_da_amort": aftertax_incremental_da_amort,
        "deferred_tax_liability": deferred_tax_liability,
        "synergy_pct_of_target_revenue": synergy_pct_of_target_revenue,
        "cost_of_new_debt": cost_of_new_debt,
        "tax_rate": tax_rate,
        "foregone_cash_return": foregone_cash_return,
        "offer_price_per_share": offer_price_per_share,
        "total_consideration": total_consideration,
        "cash_consideration": cash_consideration,
        "stock_consideration": stock_consideration,
        "cash_from_balance_sheet": cash_from_balance_sheet,
        "new_debt_raised": new_debt_raised,
        "new_shares_issued": new_shares_issued,
        "pretax_synergies": pretax_synergies,
        "aftertax_synergies": aftertax_synergies,
        "aftertax_interest_on_new_debt": aftertax_interest_on_new_debt,
        "aftertax_foregone_income_on_cash": aftertax_foregone_income_on_cash,
        "acquirer_net_income": acquirer_net_income,
        "target_net_income": target_net_income,
        "combined_net_income": combined_net_income,
        "pro_forma_diluted_shares": pro_forma_diluted_shares,
        "pro_forma_eps": pro_forma_eps,
        "acquirer_standalone_eps": acquirer_standalone_eps,
        "accretion_dilution_pct": accretion_dilution_pct,
        "acquirer_dcf": acquirer_dcf,
        "target_dcf": target_dcf,
        "target_comps": target_comps,
        "implied_comps_price": implied_comps_price,
        "premium_vs_dcf": premium_vs_dcf,
        "premium_vs_comps": premium_vs_comps,
    }


def print_ma_summary(deal):
    divider = "─" * 60
    tcs, acs = deal["target_currency_symbol"], deal["acquirer_currency_symbol"]

    print(divider)
    print(f"{deal['acquirer_name']} ({deal['acquirer_ticker']}) acquiring {deal['target_name']} ({deal['target_ticker']})")
    print(divider)
    print(f"Offer Price:        {tcs}{deal['offer_price_per_share']:,.2f}  ({deal['offer_premium']:+.1%} premium)")
    print(f"Total Consideration:{tcs}{deal['total_consideration']:,.0f}")
    print(f"  Cash:             {tcs}{deal['cash_consideration']:,.0f}  (of which {tcs}{deal['new_debt_raised']:,.0f} new debt)")
    print(f"  Stock:            {tcs}{deal['stock_consideration']:,.0f}  ({deal['new_shares_issued']:,.0f} new shares)")
    print()
    print(f"Run-Rate Synergies: {tcs}{deal['pretax_synergies']:,.0f} pretax / {tcs}{deal['aftertax_synergies']:,.0f} after-tax")
    print()
    print("Accretion / Dilution")
    if deal["accretion_dilution_pct"] is not None:
        direction = "ACCRETIVE" if deal["accretion_dilution_pct"] > 0 else "DILUTIVE"
        print(f"  Standalone EPS:   {acs}{deal['acquirer_standalone_eps']:.2f}")
        print(f"  Pro Forma EPS:    {acs}{deal['pro_forma_eps']:.2f}")
        print(f"  {direction} by {deal['accretion_dilution_pct']:+.1%}")
    else:
        print("  Unavailable — missing EPS data.")
    print()
    print("Premium vs. Independent Valuation")
    if deal["premium_vs_dcf"] is not None:
        print(f"  vs. DCF implied value:   {deal['premium_vs_dcf']:+.1%}")
    if deal["premium_vs_comps"] is not None:
        print(f"  vs. comps implied value: {deal['premium_vs_comps']:+.1%}")
    print(divider)


def _copy_sheet_into(dest_wb, source_sheet, new_title):
    """
    Copies cell values, formatting, and column widths from a worksheet built
    by another export_*_to_excel() call into this workbook — so a deal
    workbook can embed the standalone DCF/comps tabs without duplicating
    their layout code. Charts aren't copied (openpyxl can't move chart
    objects across workbooks), so embedded tabs carry data/format only.
    """
    dest = dest_wb.create_sheet(new_title)
    dest.sheet_view.showGridLines = False
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = dest.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter, dim in source_sheet.column_dimensions.items():
        if dim.width:
            dest.column_dimensions[col_letter].width = dim.width
    return dest


def _add_dcf_snapshot_tab(wb, dcf_result, sheet_name):
    """
    A static-value snapshot of a DCF result, for embedding inside another
    workbook (e.g. the M&A model). Deliberately NOT live formulas — the
    Python-computed dcf_result values are used directly, since the source
    DCF export's own cross-sheet formulas (Assumptions!/DCF!) wouldn't
    resolve without their source tabs also being copied over. For the
    live, flexible version, generate that ticker's standalone DCF export.
    """
    fmt = currency_formats(dcf_result["currency_symbol"])
    tab = wb.create_sheet(sheet_name)
    tab.sheet_view.showGridLines = False
    tab["A1"] = f"{dcf_result['company_name']} ({dcf_result['ticker']}) — DCF Snapshot"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:B1")
    tab["A2"] = "Static snapshot for reference — see the standalone DCF export for a live, editable version."
    tab["A2"].font = Font(italic=True, size=9, color="666666")

    a = dcf_result["assumptions_used"]
    rows = [
        ("Implied Share Price", dcf_result["implied_share_price"], fmt["price"]),
        ("Current Price", dcf_result["current_price"], fmt["price"]),
        ("Upside / Downside", dcf_result["upside"], "0.0%"),
        ("Enterprise Value", dcf_result["enterprise_value"], fmt["currency"]),
        ("Equity Value", dcf_result["equity_value"], fmt["currency"]),
        ("Net Debt", dcf_result["net_debt"], fmt["currency"]),
        ("WACC", a["wacc"], "0.00%"),
        ("Terminal Growth Rate", a["terminal_growth_rate"], "0.00%"),
    ]
    for i, (label, value, number_fmt) in enumerate(rows, start=4):
        tab.cell(row=i, column=1, value=label).font = LABEL_FONT
        cell = tab.cell(row=i, column=2, value=value)
        cell.number_format = number_fmt
    tab.column_dimensions["A"].width = 26
    tab.column_dimensions["B"].width = 18


def export_ma_model_to_excel(deal):
    """
    LIVE accretion/dilution model — matching standard merger-model
    conventions (see e.g. Breaking Into Wall Street's Accretion/Dilution
    and Sources & Uses guides): every deal term, financing split, and
    synergy assumption is a blue input; offer price, consideration mix,
    sources & uses, pro forma net income, and accretion/dilution are all
    Excel formulas chained off those inputs on an Assumptions tab.
    """
    wb = Workbook()
    fmt = currency_formats(deal["target_currency_symbol"])
    acs = deal["acquirer_currency_symbol"]

    # ============================================
    # ASSUMPTIONS TAB
    # ============================================
    assump = wb.active
    assump.title = "Assumptions"
    assump.sheet_view.showGridLines = False
    assump["A1"] = f"{deal['acquirer_name']} ({deal['acquirer_ticker']}) acquiring {deal['target_name']} ({deal['target_ticker']})"
    assump["A1"].font = TITLE_FONT
    assump.merge_cells("A1:C1")
    assump["A2"] = "Blue cells are hardcoded inputs — change these to flex the deal. Black cells are formulas."
    assump["A2"].font = Font(italic=True, size=9, color="666666")

    input_rows = [
        ("Target Current Price", deal["target_current_price"], fmt["price"]),
        ("Acquirer Current Price", deal["acquirer_current_price"], fmt["price"]),
        ("Offer Premium", deal["offer_premium"], "0.0%"),
        ("% Consideration in Cash", deal["pct_cash"], "0.0%"),
        ("% Consideration in Stock", deal["pct_stock"], "0.0%"),
        ("Target Shares Outstanding", deal["target_shares"], "#,##0"),
        ("Acquirer Shares Outstanding", deal["acquirer_shares"], "#,##0"),
        ("Acquirer Cash on Hand", deal["acquirer_cash_on_hand"], fmt["currency"]),
        ("Target Revenue", deal["target_revenue"], fmt["currency"]),
        ("Run-Rate Synergies (% of Target Revenue)", deal["synergy_pct_of_target_revenue"], "0.0%"),
        ("Cost of New Debt", deal["cost_of_new_debt"], "0.0%"),
        ("Tax Rate", deal["tax_rate"], "0.0%"),
        ("Foregone Return on Cash Used", deal["foregone_cash_return"], "0.0%"),
        ("Acquirer Standalone Net Income", deal["acquirer_net_income"], fmt["currency"]),
        ("Target Standalone Net Income", deal["target_net_income"], fmt["currency"]),
        ("Acquirer Standalone EPS", deal["acquirer_standalone_eps"], fmt["price"]),
        ("Target Book Equity", deal["target_book_equity"], fmt["currency"]),
        ("Intangible Step-Up (% of Excess Purchase Price)", 0.40, "0.0%"),
        ("PP&E Step-Up (% of Excess Purchase Price)", 0.10, "0.0%"),
        ("Intangible Useful Life (Years)", deal["intangible_useful_life_years"], "0"),
        ("PP&E Step-Up Useful Life (Years)", deal["ppe_useful_life_years"], "0"),
    ]
    for i, (label, value, number_fmt) in enumerate(input_rows, start=4):
        assump.cell(row=i, column=1, value=label).font = LABEL_FONT
        _write_input(assump, f"B{i}", value, number_fmt)
    ROW = {label: i for i, (label, *_) in enumerate(input_rows, start=4)}
    assump.column_dimensions["A"].width = 36
    assump.column_dimensions["B"].width = 20

    def A(label):
        return f"Assumptions!$B${ROW[label]}"

    # ============================================
    # PURCHASE PRICE ALLOCATION TAB — live formulas throughout
    # ============================================
    ppa = wb.create_sheet("Purchase Price Allocation")
    ppa.sheet_view.showGridLines = False
    ppa["A1"] = f"{deal['target_name']} — Purchase Price Allocation"
    ppa["A1"].font = TITLE_FONT
    ppa.merge_cells("A1:B1")
    ppa["A2"] = "The excess of purchase price over book value is allocated to intangibles and a PP&E step-up; " \
                "goodwill is the RESIDUAL, not a separate estimate. Step-ups create deferred tax and incremental " \
                "non-cash D&A/amortization that reduces pro forma net income below."
    ppa["A2"].font = Font(italic=True, size=9, color="666666")
    ppa.merge_cells("A2:D2")

    r_ppa = 4
    def ppa_row(label, formula, number_fmt, bold=False):
        nonlocal r_ppa
        ppa.cell(row=r_ppa, column=1, value=label).font = Font(bold=True, color="1F4E78") if bold else LABEL_FONT
        _write_formula(ppa, f"B{r_ppa}", formula, number_fmt, bold=bold)
        this_row = r_ppa
        r_ppa += 1
        return this_row

    offer_price_ref = f"({A('Target Current Price')}*(1+{A('Offer Premium')}))"
    total_consid_ppa_row = ppa_row("Total Consideration (recomputed here for self-containment)",
                                     f"={offer_price_ref}*{A('Target Shares Outstanding')}", fmt["currency"])
    excess_row = ppa_row("Excess Purchase Price (over Book Equity)",
                          f"=MAX(B{total_consid_ppa_row}-{A('Target Book Equity')},0)", fmt["currency"])
    intangible_row = ppa_row("Intangible Step-Up", f"=B{excess_row}*{A('Intangible Step-Up (% of Excess Purchase Price)')}", fmt["currency"])
    ppe_row = ppa_row("PP&E Step-Up", f"=B{excess_row}*{A('PP&E Step-Up (% of Excess Purchase Price)')}", fmt["currency"])
    goodwill_row = ppa_row("Goodwill (residual)", f"=B{excess_row}-B{intangible_row}-B{ppe_row}", fmt["currency"], bold=True)
    r_ppa += 1
    dtl_row = ppa_row("Deferred Tax Liability", f"=(B{intangible_row}+B{ppe_row})*{A('Tax Rate')}", fmt["currency"])
    r_ppa += 1
    incr_amort_row = ppa_row("Incremental Amortization (Intangibles)", f"=B{intangible_row}/{A('Intangible Useful Life (Years)')}", fmt["currency"])
    incr_da_row = ppa_row("Incremental D&A (PP&E Step-Up)", f"=B{ppe_row}/{A('PP&E Step-Up Useful Life (Years)')}", fmt["currency"])
    total_incr_row = ppa_row("Total Incremental D&A / Amortization", f"=B{incr_amort_row}+B{incr_da_row}", fmt["currency"], bold=True)
    aftertax_incr_row = ppa_row("After-Tax Impact on Net Income", f"=-B{total_incr_row}*(1-{A('Tax Rate')})", fmt["currency"], bold=True)

    ppa.column_dimensions["A"].width = 42
    ppa.column_dimensions["B"].width = 20

    # ============================================
    # DEAL SUMMARY TAB — the live waterfall
    # ============================================
    summary = wb.create_sheet("Deal Summary", 0)
    summary.sheet_view.showGridLines = False
    summary["A1"] = "Deal Terms & Accretion / Dilution"
    summary["A1"].font = TITLE_FONT
    summary.merge_cells("A1:C1")

    r = 3
    def add_row(label, formula_or_value, number_fmt, bold=False, is_formula=True):
        nonlocal r
        summary.cell(row=r, column=1, value=label).font = Font(bold=True, size=12, color="1F4E78") if bold else LABEL_FONT
        if is_formula:
            _write_formula(summary, f"B{r}", formula_or_value, number_fmt, bold=bold)
        else:
            summary.cell(row=r, column=2, value=formula_or_value).number_format = number_fmt
        summary[f"A{r}"].border = THIN_BORDER
        summary[f"B{r}"].border = THIN_BORDER
        this_row = r
        r += 1
        return this_row

    offer_price_row = add_row("Offer Price / Share", f"={A('Target Current Price')}*(1+{A('Offer Premium')})", fmt["price"], bold=True)
    total_consid_row = add_row("Total Consideration", f"=B{offer_price_row}*{A('Target Shares Outstanding')}", fmt["currency"], bold=True)
    cash_consid_row = add_row("  Cash Consideration", f"=B{total_consid_row}*{A('% Consideration in Cash')}", fmt["currency"])
    stock_consid_row = add_row("  Stock Consideration", f"=B{total_consid_row}*{A('% Consideration in Stock')}", fmt["currency"])
    new_shares_row = add_row("New Shares Issued", f"=B{stock_consid_row}/{A('Acquirer Current Price')}", "#,##0")
    r += 1

    cash_from_bs_row = add_row("Cash from Balance Sheet", f"=MIN(B{cash_consid_row},{A('Acquirer Cash on Hand')})", fmt["currency"])
    new_debt_row = add_row("New Debt Raised", f"=MAX(B{cash_consid_row}-{A('Acquirer Cash on Hand')},0)", fmt["currency"])
    r += 1

    pretax_syn_row = add_row("Pretax Run-Rate Synergies", f"={A('Target Revenue')}*{A('Run-Rate Synergies (% of Target Revenue)')}", fmt["currency"])
    aftertax_syn_row = add_row("After-Tax Synergies", f"=B{pretax_syn_row}*(1-{A('Tax Rate')})", fmt["currency"])
    aftertax_int_row = add_row("Less: After-Tax Interest on New Debt", f"=-B{new_debt_row}*{A('Cost of New Debt')}*(1-{A('Tax Rate')})", fmt["currency"])
    aftertax_forgone_row = add_row("Less: After-Tax Foregone Cash Income", f"=-B{cash_from_bs_row}*{A('Foregone Return on Cash Used')}*(1-{A('Tax Rate')})", fmt["currency"])
    r += 1

    combined_ni_row = add_row("Pro Forma Combined Net Income",
                               f"={A('Acquirer Standalone Net Income')}+{A('Target Standalone Net Income')}+B{aftertax_syn_row}+B{aftertax_int_row}+B{aftertax_forgone_row}"
                               f"+'Purchase Price Allocation'!B{aftertax_incr_row}", fmt["currency"], bold=True)
    pf_shares_row = add_row("Pro Forma Diluted Shares", f"={A('Acquirer Shares Outstanding')}+B{new_shares_row}", "#,##0")
    pf_eps_row = add_row("Pro Forma EPS", f"=B{combined_ni_row}/B{pf_shares_row}", fmt["price"], bold=True)
    standalone_eps_row = add_row("Acquirer Standalone EPS", f"={A('Acquirer Standalone EPS')}", fmt["price"])
    accretion_row = add_row("Accretion / (Dilution)", f"=(B{pf_eps_row}-B{standalone_eps_row})/ABS(B{standalone_eps_row})", "0.0%", bold=True)

    if deal["premium_vs_dcf"] is not None or deal["premium_vs_comps"] is not None:
        r += 1
        if deal["premium_vs_dcf"] is not None:
            add_row("Premium vs. DCF Implied Value", deal["premium_vs_dcf"], fmt["percent"], is_formula=False)
        if deal["premium_vs_comps"] is not None:
            add_row("Premium vs. Comps Implied Value", deal["premium_vs_comps"], fmt["percent"], is_formula=False)

    summary.column_dimensions["A"].width = 38
    summary.column_dimensions["B"].width = 20

    chart = BarChart()
    chart.title = "Standalone vs. Pro Forma EPS"
    chart.style = 12
    chart.height, chart.width = 7, 12
    chart.y_axis.number_format = fmt["price"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart_row = r + 2
    summary.cell(row=chart_row, column=1, value="Standalone EPS")
    _write_formula(summary, f"B{chart_row}", f"=B{standalone_eps_row}", fmt["price"])
    summary.cell(row=chart_row + 1, column=1, value="Pro Forma EPS")
    _write_formula(summary, f"B{chart_row + 1}", f"=B{pf_eps_row}", fmt["price"])
    chart.add_data(Reference(summary, min_col=2, min_row=chart_row, max_row=chart_row + 1))
    chart.set_categories(Reference(summary, min_col=1, min_row=chart_row, max_row=chart_row + 1))
    summary.add_chart(chart, "D3")

    # ============================================
    # SOURCES & USES TAB — live, must balance by construction
    # ============================================
    su = wb.create_sheet("Sources & Uses")
    su.sheet_view.showGridLines = False
    su["A1"] = "Sources & Uses"
    su["A1"].font = TITLE_FONT
    su.merge_cells("A1:E1")

    for col, header in enumerate(["Sources", "Amount", "", "Uses", "Amount"], start=1):
        if header:
            cell = su.cell(row=3, column=col, value=header)
            cell.font, cell.fill = HEADER_FONT, HEADER_FILL

    su.cell(row=4, column=1, value="Cash from Balance Sheet")
    _write_formula(su, "B4", f"='Deal Summary'!B{cash_from_bs_row}", fmt["currency"])
    su.cell(row=5, column=1, value="New Debt Raised")
    _write_formula(su, "B5", f"='Deal Summary'!B{new_debt_row}", fmt["currency"])
    su.cell(row=6, column=1, value="Stock Issued to Target")
    _write_formula(su, "B6", f"='Deal Summary'!B{stock_consid_row}", fmt["currency"])
    su.cell(row=7, column=1, value="Total Sources").font = LABEL_FONT
    _write_formula(su, "B7", "=SUM(B4:B6)", fmt["currency"], bold=True)

    su.cell(row=4, column=4, value="Purchase of Target Equity")
    _write_formula(su, "E4", f"='Deal Summary'!B{total_consid_row}", fmt["currency"])
    su.cell(row=7, column=4, value="Total Uses").font = LABEL_FONT
    _write_formula(su, "E7", "=E4", fmt["currency"], bold=True)

    su.cell(row=9, column=1, value="Balance Check (Sources - Uses)").font = Font(bold=True, size=10)
    check_cell = _write_formula(su, "B9", "=B7-E7", fmt["currency"], bold=True)
    check_cell.font = Font(bold=True, color="1B7A1B")

    for col_letter, width in zip("ABCDE", [26, 16, 4, 26, 16]):
        su.column_dimensions[col_letter].width = width

    # --- Embed standalone DCF snapshots (static values, not the live formulas
    # from export_dcf_to_excel — that Summary tab now cross-references its
    # own Assumptions/DCF tabs, which don't exist here. This snapshot is
    # reference context; open the standalone DCF export for the live version). ---
    if deal["acquirer_dcf"]:
        _add_dcf_snapshot_tab(wb, deal["acquirer_dcf"], f"{deal['acquirer_ticker']} DCF"[:31])
    if deal["target_dcf"]:
        _add_dcf_snapshot_tab(wb, deal["target_dcf"], f"{deal['target_ticker']} DCF"[:31])

    # Comps tab is safe to copy directly — its formulas are all same-sheet
    # references (peer averages, implied EV), so they stay valid after copying.
    comps_wb = export_comps_to_excel(deal["target_comps"], deal["target_ticker"])
    _copy_sheet_into(wb, comps_wb["Comps"], "Target Comps")

    return wb


# ============================================
# IPO VALUATION (comps-driven — no trading history to anchor to)
# ============================================

def run_sotp_valuation(company_name, segments, net_debt=0, shares_outstanding=None, currency_symbol="$"):
    """
    Sum-of-the-parts, built on segment data YOU supply (from a 10-K, investor
    presentation, or your own research) — not auto-fetched. yfinance doesn't
    reliably expose segment-level financials for most companies, so rather
    than fake a split, this takes real segment data as input, exactly the
    way run_ipo_valuation() takes a private company's financials directly.

    segments: list of dicts, each with:
      name (str), metric_value (revenue or EBITDA for that segment),
      metric_type ("revenue" or "ebitda"), multiple (the peer/segment
      multiple to apply — different segments legitimately warrant
      different multiples, which is the whole point of doing this
      instead of one blended multiple for the whole company).
    """
    if not segments:
        raise ValueError("Need at least one segment to run a sum-of-the-parts valuation.")

    for seg in segments:
        if seg["metric_type"] not in ("revenue", "ebitda"):
            raise ValueError(f"Segment '{seg['name']}': metric_type must be 'revenue' or 'ebitda'.")
        seg["implied_ev"] = seg["metric_value"] * seg["multiple"]

    total_ev = sum(seg["implied_ev"] for seg in segments)
    equity_value = total_ev - net_debt
    implied_price_per_share = equity_value / shares_outstanding if shares_outstanding else None

    for seg in segments:
        seg["pct_of_total_ev"] = seg["implied_ev"] / total_ev if total_ev else None

    return {
        "company_name": company_name,
        "currency_symbol": currency_symbol,
        "segments": segments,
        "total_ev": total_ev,
        "net_debt": net_debt,
        "equity_value": equity_value,
        "shares_outstanding": shares_outstanding,
        "implied_price_per_share": implied_price_per_share,
    }


def print_sotp_valuation(result):
    divider = "─" * 60
    sym = result["currency_symbol"]
    print(divider)
    print(f"{result['company_name']} — Sum-of-the-Parts Valuation")
    print(divider)
    for seg in result["segments"]:
        metric_label = "Revenue" if seg["metric_type"] == "revenue" else "EBITDA"
        print(f"{seg['name']:<24} {metric_label} {sym}{seg['metric_value']:,.0f} × {seg['multiple']:.1f}x "
              f"= {sym}{seg['implied_ev']:,.0f}  ({seg['pct_of_total_ev']:.0%} of total)")
    print(divider)
    print(f"Total Enterprise Value:  {sym}{result['total_ev']:,.0f}")
    print(f"Less: Net Debt:          {sym}{result['net_debt']:,.0f}")
    print(f"Equity Value:            {sym}{result['equity_value']:,.0f}")
    if result["implied_price_per_share"] is not None:
        print(f"Implied Price / Share:   {sym}{result['implied_price_per_share']:,.2f}")
    print(divider)


def export_sotp_to_excel(result):
    """LIVE model: each segment's raw metric and multiple are blue inputs;
    implied EV per segment, total EV, and the equity value bridge are all
    Excel formulas."""
    wb = Workbook()
    fmt = currency_formats(result["currency_symbol"])
    tab = wb.active
    tab.title = "Sum-of-the-Parts"
    tab.sheet_view.showGridLines = False

    tab["A1"] = f"{result['company_name']} — Sum-of-the-Parts Valuation"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:F1")
    tab["A2"] = "Blue cells are raw inputs (segment financials and multiples you supply). Implied EV, totals, and the equity bridge are live formulas."
    tab["A2"].font = Font(italic=True, size=9, color="666666")

    headers = ["Segment", "Metric", "Metric Value", "Multiple", "Implied EV", "% of Total EV"]
    for col, h in enumerate(headers, start=1):
        cell = tab.cell(row=4, column=col, value=h)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    first_row = 5
    for i, seg in enumerate(result["segments"], start=first_row):
        tab.cell(row=i, column=1, value=seg["name"])
        tab.cell(row=i, column=2, value="Revenue" if seg["metric_type"] == "revenue" else "EBITDA")
        _write_input(tab, f"C{i}", seg["metric_value"], fmt["currency"])
        _write_input(tab, f"D{i}", seg["multiple"], '0.00"x"')
        _write_formula(tab, f"E{i}", f"=C{i}*D{i}", fmt["currency"], bold=True)
        for col in (1, 2):
            tab.cell(row=i, column=col).alignment = Alignment(horizontal="left")
        for col in range(3, 6):
            tab.cell(row=i, column=col).alignment = Alignment(horizontal="center")
    last_row = first_row + len(result["segments"]) - 1

    total_row = last_row + 1
    tab.cell(row=total_row, column=1, value="Total Enterprise Value").font = Font(bold=True, color="1F4E78")
    _write_formula(tab, f"E{total_row}", f"=SUM(E{first_row}:E{last_row})", fmt["currency"], bold=True)

    for i in range(first_row, last_row + 1):
        _write_formula(tab, f"F{i}", f"=IFERROR(E{i}/$E${total_row},\"N/A\")", "0.0%")
    tab.cell(row=total_row, column=6, value=1.0).number_format = "0.0%"
    tab.cell(row=total_row, column=6).font = Font(bold=True)

    bridge_row = total_row + 2
    tab.cell(row=bridge_row, column=1, value="Total Enterprise Value").font = LABEL_FONT
    _write_formula(tab, f"B{bridge_row}", f"=E{total_row}", fmt["currency"], bold=True)
    tab.cell(row=bridge_row + 1, column=1, value="Net Debt").font = LABEL_FONT
    _write_input(tab, f"B{bridge_row + 1}", result["net_debt"], fmt["currency"])
    tab.cell(row=bridge_row + 2, column=1, value="Equity Value").font = Font(bold=True, color="1F4E78")
    _write_formula(tab, f"B{bridge_row + 2}", f"=B{bridge_row}-B{bridge_row + 1}", fmt["currency"], bold=True)
    tab.cell(row=bridge_row + 3, column=1, value="Shares Outstanding").font = LABEL_FONT
    _write_input(tab, f"B{bridge_row + 3}", result["shares_outstanding"], "#,##0")
    tab.cell(row=bridge_row + 4, column=1, value="Implied Price / Share").font = Font(bold=True, size=12, color="1F4E78")
    _write_formula(tab, f"B{bridge_row + 4}", f'=IFERROR(B{bridge_row + 2}/B{bridge_row + 3},"N/A")', fmt["price"], bold=True)

    chart = BarChart()
    chart.title = "Implied Enterprise Value by Segment"
    chart.style = 10
    chart.y_axis.number_format = fmt["currency"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16
    chart.add_data(Reference(tab, min_col=5, min_row=4, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(tab, min_col=1, min_row=first_row, max_row=last_row))
    tab.add_chart(chart, f"H4")

    for col_letter, width in zip("ABCDEF", [22, 12, 16, 12, 16, 14]):
        tab.column_dimensions[col_letter].width = width

    return wb


def build_pitch_book_sotp(result, output_path):
    prs = _new_pitch_book()
    logo_info = {"shortName": result["company_name"]}
    sym = result["currency_symbol"]; today = datetime.now().strftime("%B %d, %Y")

    page = 1
    _add_title_slide(prs, result["company_name"], "Sum-of-the-Parts Valuation", logo_info, None, today)
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["Executive Summary", "Segment Valuation", "Value Composition", "Equity Bridge", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "Executive Summary", "Segment-level valuation, summed to a total enterprise value", logo_info, None, page)
    cards = [("Total EV", f"{sym}{result['total_ev']:,.0f}", "neutral"), ("Equity Value", f"{sym}{result['equity_value']:,.0f}", "neutral"),
             ("Segments", str(len(result["segments"])), "neutral")]
    if result["implied_price_per_share"] is not None:
        cards.append(("Implied Price / Share", f"{sym}{result['implied_price_per_share']:,.2f}", "positive"))
    _add_kpi_cards(s, cards, top=1.35, card_width=2.75, height=1.3)
    _add_bullet_box(s, .7, 3.1, 11.9, 2.7, "Methodology",
                     ["Each segment is valued independently at its own multiple, reflecting that different lines of business "
                      "warrant different valuation approaches — a growth segment and a mature cash-generative segment should "
                      "rarely be valued at the same multiple.",
                      "Segment financials and multiples are supplied directly (from disclosed segment data or your own "
                      "research), not automatically sourced — automated data providers don't reliably expose true segment-level "
                      "financials.", "Total EV is the sum of each segment's implied value; equity value follows the standard "
                      "EV less net debt bridge."], PPT_NAVY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Segment Valuation", "Implied enterprise value by segment", logo_info, None, page)
    rows = [[seg["name"], "Revenue" if seg["metric_type"] == "revenue" else "EBITDA", f"{sym}{seg['metric_value']:,.0f}",
             f"{seg['multiple']:.1f}x", f"{sym}{seg['implied_ev']:,.0f}", f"{seg['pct_of_total_ev']:.0%}"] for seg in result["segments"]]
    _add_table(s, .6, 1.5, 12.1, 3.5, ["Segment", "Metric", "Value", "Multiple", "Implied EV", "% of Total"], rows,
               col_widths=[2.2, 1.2, 1.6, 1.2, 1.8, 1.2], font_size=9.3, first_col_bold=True)
    _add_banner(s, f"Total Enterprise Value: {sym}{result['total_ev']:,.0f}", 5.3, fill=PPT_NAVY, color=PPT_WHITE)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Value Composition", "Where the total enterprise value comes from", logo_info, None, page)
    _add_column_chart(s, .8, 1.55, 11.8, 4.7, [seg["name"] for seg in result["segments"]], [("Implied EV", [seg["implied_ev"] for seg in result["segments"]])],
                       title="Implied EV by Segment", number_format='#,##0', colors=[PPT_NAVY])

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Equity Value Bridge", "From total enterprise value to implied price per share", logo_info, None, page)
    _add_waterfall(s, .8, 1.55, 11.6, 4.6, ["Total EV", "Net Debt", "Equity Value"],
                    [result["total_ev"], -result["net_debt"], result["equity_value"]], title="EV to Equity Value Bridge", number_format='#,##0')
    if result["implied_price_per_share"] is not None:
        _add_banner(s, f"Implied Price / Share: {sym}{result['implied_price_per_share']:,.2f}  ({result['shares_outstanding']:,.0f} shares outstanding)", 6.35, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Segment Detail & Assumptions", "Every input behind this valuation", logo_info, None, page)
    rows = [[seg["name"], seg["metric_type"], f"{sym}{seg['metric_value']:,.0f}", f"{seg['multiple']:.2f}x"] for seg in result["segments"]]
    _add_table(s, .8, 1.5, 6.6, 3.5, ["Segment", "Metric Type", "Metric Value", "Multiple Applied"], rows, col_widths=[2.2, 1.4, 1.8, 1.4], font_size=9.3, first_col_bold=True)
    _add_bullet_box(s, 7.65, 1.5, 4.85, 3.5, "Caveats",
                     ["Segment financials/multiples are user-supplied, not automatically fetched — verify against the latest disclosed figures.",
                      "Different segments may warrant different capital structures; this model uses one consolidated net debt figure.",
                      "No inter-segment eliminations are modeled."], PPT_GOLD)

    prs.save(output_path)
    return output_path


def run_ipo_valuation(company_name, industry, revenue, ebitda=None, net_income=None,
                       shares_outstanding_post_ipo=None, net_debt=0, peer_tickers=None,
                       max_peers=5, illiquidity_discount=0.15):
    """
    For a pre-IPO company there's no current price to check a DCF against,
    so peer multiples carry the valuation instead. Peers come from
    INDUSTRY_UNIVERSE by industry name (same lookup the comps platform uses)
    unless an explicit peer_tickers list is passed — useful when the company
    doesn't map cleanly to a built-in industry bucket, or a hand-picked peer
    set is wanted.

    revenue/ebitda/net_income/shares_outstanding_post_ipo are supplied
    directly (from the S-1 or management projections) rather than pulled
    from yfinance, since a private company has no yfinance ticker.

    illiquidity_discount is a standard IPO-pricing convention (newly public
    stock trades at a discount to seasoned peers) applied to the blended
    comps-implied price — pass 0 to see the raw comps-implied price.
    """
    if peer_tickers is None:
        peer_tickers = INDUSTRY_UNIVERSE.get(industry, [])
        if not peer_tickers:
            raise ValueError(
                f"No built-in peer list for industry '{industry}'. Pass peer_tickers explicitly."
            )

    peer_data = [p for p in (get_company_multiples(t) for t in peer_tickers[:max_peers]) if p is not None]
    if not peer_data:
        raise ValueError("Could not pull multiples for any peer ticker — check the peer list.")

    def average(values, min_val=None, max_val=None):
        clean = [v for v in values if v is not None]
        if min_val is not None:
            clean = [v for v in clean if v > min_val]
        if max_val is not None:
            clean = [v for v in clean if v < max_val]
        return sum(clean) / len(clean) if clean else None

    avg_ev_revenue = average([p["ev_revenue"] for p in peer_data])
    avg_ev_ebitda = average([p["ev_ebitda"] for p in peer_data], min_val=0, max_val=60)
    avg_pe = average([p["pe_ratio"] for p in peer_data], min_val=0, max_val=100)

    implied_ev_from_revenue = avg_ev_revenue * revenue if (avg_ev_revenue and revenue) else None
    implied_ev_from_ebitda = (
        avg_ev_ebitda * ebitda if (avg_ev_ebitda and ebitda and ebitda > 0) else None
    )
    implied_equity_from_pe = avg_pe * net_income if (avg_pe and net_income and net_income > 0) else None

    implied_price_from_revenue = None
    implied_price_from_ebitda = None
    implied_price_from_pe = None
    if shares_outstanding_post_ipo:
        if implied_ev_from_revenue is not None:
            implied_price_from_revenue = (implied_ev_from_revenue - net_debt) / shares_outstanding_post_ipo
        if implied_ev_from_ebitda is not None:
            implied_price_from_ebitda = (implied_ev_from_ebitda - net_debt) / shares_outstanding_post_ipo
        if implied_equity_from_pe is not None:
            implied_price_from_pe = implied_equity_from_pe / shares_outstanding_post_ipo

    # Blend whichever per-share anchors are available (equal-weight across the ones that exist)
    per_share_estimates = [
        v for v in [implied_price_from_revenue, implied_price_from_ebitda, implied_price_from_pe] if v is not None
    ]
    implied_price_per_share = sum(per_share_estimates) / len(per_share_estimates) if per_share_estimates else None
    discounted_price_per_share = (
        implied_price_per_share * (1 - illiquidity_discount) if implied_price_per_share is not None else None
    )

    return {
        "company_name": company_name,
        "industry": industry,
        "revenue": revenue,
        "ebitda": ebitda,
        "net_income": net_income,
        "net_debt": net_debt,
        "shares_outstanding_post_ipo": shares_outstanding_post_ipo,
        "peers": peer_data,
        "avg_ev_revenue": avg_ev_revenue,
        "avg_ev_ebitda": avg_ev_ebitda,
        "avg_pe": avg_pe,
        "implied_ev_from_revenue": implied_ev_from_revenue,
        "implied_ev_from_ebitda": implied_ev_from_ebitda,
        "implied_equity_from_pe": implied_equity_from_pe,
        "implied_price_from_revenue": implied_price_from_revenue,
        "implied_price_from_ebitda": implied_price_from_ebitda,
        "implied_price_from_pe": implied_price_from_pe,
        "illiquidity_discount": illiquidity_discount,
        "implied_price_per_share": implied_price_per_share,
        "discounted_price_per_share": discounted_price_per_share,
    }


def print_ipo_valuation(ipo, currency_symbol="$"):
    divider = "─" * 60
    print(divider)
    print(f"{ipo['company_name']} — IPO Valuation  ·  {ipo['industry']}")
    print(divider)
    print(f"Revenue:            {currency_symbol}{ipo['revenue']:,.0f}")
    if ipo["ebitda"]:
        print(f"EBITDA:             {currency_symbol}{ipo['ebitda']:,.0f}")
    if ipo["net_income"]:
        print(f"Net Income:         {currency_symbol}{ipo['net_income']:,.0f}")
    print(f"Peers used:         {len(ipo['peers'])}  ({', '.join(p['ticker'] for p in ipo['peers'])})")
    print()
    print("Implied Price / Share")
    for label, key in [("EV/Revenue", "implied_price_from_revenue"),
                        ("EV/EBITDA", "implied_price_from_ebitda"),
                        ("P/E", "implied_price_from_pe")]:
        val = ipo.get(key)
        print(f"  {label:<12}{currency_symbol}{val:,.2f}" if val is not None else f"  {label:<12}N/A")
    if ipo["implied_price_per_share"] is not None:
        print(f"  {'Blended':<12}{currency_symbol}{ipo['implied_price_per_share']:,.2f}")
        print(f"  {'w/ discount':<12}{currency_symbol}{ipo['discounted_price_per_share']:,.2f}  "
              f"({ipo['illiquidity_discount']:.0%} illiquidity discount)")
    else:
        print("  Blended:      N/A — insufficient peer data")
    print(divider)


def export_ipo_valuation_to_excel(ipo, currency_symbol="$"):
    """
    LIVE model, same philosophy as comps: raw EV/Revenue/EBITDA per peer
    are blue inputs; every multiple, peer average, and implied price is
    an Excel formula. Change a peer's raw figures, or the company's own
    financials, and the whole valuation recalculates.
    """
    wb = Workbook()
    fmt = currency_formats(currency_symbol)
    tab = wb.active
    tab.title = "IPO Valuation"
    tab.sheet_view.showGridLines = False

    tab["A1"] = f"{ipo['company_name']} — IPO Valuation ({ipo['industry']})"
    tab["A1"].font = TITLE_FONT
    tab.merge_cells("A1:H1")
    tab["A2"] = "Blue cells are raw inputs; multiples, averages, and implied prices below are live formulas."
    tab["A2"].font = Font(italic=True, size=9, color="666666")

    headers = ["Ticker", "Company", "Enterprise Value", "Revenue", "EBITDA", "EV / Revenue", "EV / EBITDA", "P/E"]
    for col, header in enumerate(headers, start=1):
        cell = tab.cell(row=4, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for i, peer in enumerate(ipo["peers"], start=5):
        tab.cell(row=i, column=1, value=peer["ticker"])
        tab.cell(row=i, column=2, value=peer["company_name"])
        _write_input(tab, f"C{i}", peer.get("enterprise_value"), fmt["currency"])
        _write_input(tab, f"D{i}", peer.get("revenue"), fmt["currency"])
        _write_input(tab, f"E{i}", peer.get("ebitda"), fmt["currency"])
        _write_formula(tab, f"F{i}", f'=IFERROR(C{i}/D{i},"N/A")', '0.00"x"')
        _write_formula(tab, f"G{i}", f'=IFERROR(C{i}/E{i},"N/A")', '0.00"x"')
        _write_input(tab, f"H{i}", peer.get("pe_ratio"), '0.00"x"')
        for col in (1, 2):
            tab.cell(row=i, column=col).alignment = Alignment(horizontal="left")
        for col in range(3, 9):
            tab.cell(row=i, column=col).alignment = Alignment(horizontal="center")

    first_peer_row, last_peer_row = 5, 4 + len(ipo["peers"])
    avg_row = last_peer_row + 1
    tab.cell(row=avg_row, column=2, value="Peer Average").font = LABEL_FONT
    for col_letter in ("F", "G", "H"):
        _write_formula(tab, f"{col_letter}{avg_row}",
                        f'=IFERROR(AVERAGE({col_letter}{first_peer_row}:{col_letter}{last_peer_row}),"N/A")',
                        '0.00"x"', bold=True)
    for col in (1, 2):
        tab.cell(row=avg_row, column=col).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    # --- Target company's own financials, as inputs ---
    fin_row = avg_row + 2
    tab.cell(row=fin_row, column=1, value=f"{ipo['company_name']} Financials").font = LABEL_FONT
    fin_inputs = [
        ("Revenue", ipo["revenue"], fmt["currency"]),
        ("EBITDA", ipo["ebitda"], fmt["currency"]),
        ("Net Income", ipo["net_income"], fmt["currency"]),
        ("Net Debt", ipo["net_debt"], fmt["currency"]),
        ("Shares Outstanding (Post-IPO)", ipo["shares_outstanding_post_ipo"], "#,##0"),
        ("Illiquidity Discount", ipo["illiquidity_discount"], "0.0%"),
    ]
    r = fin_row + 1
    fin_cell_row = {}
    for label, value, number_fmt in fin_inputs:
        tab.cell(row=r, column=1, value=label)
        _write_input(tab, f"B{r}", value, number_fmt)
        fin_cell_row[label] = r
        r += 1

    revenue_ref = f"B{fin_cell_row['Revenue']}"
    ebitda_ref = f"B{fin_cell_row['EBITDA']}"
    net_income_ref = f"B{fin_cell_row['Net Income']}"
    net_debt_ref = f"B{fin_cell_row['Net Debt']}"
    shares_ref = f"B{fin_cell_row['Shares Outstanding (Post-IPO)']}"
    discount_ref = f"B{fin_cell_row['Illiquidity Discount']}"

    # --- Implied price per share, each method a live formula ---
    implied_row = r + 1
    tab.cell(row=implied_row, column=1, value="Implied Price / Share").font = LABEL_FONT
    rev_price_row = implied_row + 1
    ebitda_price_row = implied_row + 2
    pe_price_row = implied_row + 3
    blended_row = implied_row + 4
    blended_discounted_row = implied_row + 5

    tab.cell(row=rev_price_row, column=1, value="From EV/Revenue")
    _write_formula(tab, f"B{rev_price_row}",
                    f'=IFERROR((F{avg_row}*{revenue_ref}-{net_debt_ref})/{shares_ref},"N/A")', fmt["price"])

    tab.cell(row=ebitda_price_row, column=1, value="From EV/EBITDA")
    _write_formula(tab, f"B{ebitda_price_row}",
                    f'=IFERROR((G{avg_row}*{ebitda_ref}-{net_debt_ref})/{shares_ref},"N/A")', fmt["price"])

    tab.cell(row=pe_price_row, column=1, value="From P/E")
    _write_formula(tab, f"B{pe_price_row}",
                    f'=IFERROR(H{avg_row}*{net_income_ref}/{shares_ref},"N/A")', fmt["price"])

    tab.cell(row=blended_row, column=1, value="Blended (pre-discount)").font = LABEL_FONT
    _write_formula(tab, f"B{blended_row}",
                    f'=IFERROR(AVERAGE(B{rev_price_row}:B{pe_price_row}),"N/A")', fmt["price"], bold=True)

    tab.cell(row=blended_discounted_row, column=1, value="Blended, After Illiquidity Discount").font = Font(bold=True, size=12, color="1F4E78")
    _write_formula(tab, f"B{blended_discounted_row}",
                    f'=IFERROR(B{blended_row}*(1-{discount_ref}),"N/A")', fmt["price"], bold=True)

    # --- Chart: implied price by method, reading live from the formula cells ---
    chart = BarChart()
    chart.title = "Implied Price / Share by Method"
    chart.style = 10
    chart.height, chart.width = 8, 16
    chart.y_axis.number_format = fmt["price"]
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.add_data(Reference(tab, min_col=2, min_row=rev_price_row, max_row=blended_discounted_row))
    chart.set_categories(Reference(tab, min_col=1, min_row=rev_price_row, max_row=blended_discounted_row))
    tab.add_chart(chart, "D4")

    for col_letter, width in zip("ABCDEFGH", [30, 20, 16, 16, 14, 12, 12, 10]):
        tab.column_dimensions[col_letter].width = width

    return wb

# ============================================
# PITCH BOOK GENERATOR
# Presentation layer over research / DCF / comps / M&A / IPO outputs — it
# never re-derives a valuation, it packages results already produced by
# run_research(), run_ma_model(), or run_ipo_valuation().
# ============================================

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from urllib.request import Request, urlopen
from io import BytesIO
import math

# Same palette as the Excel exports (HEADER_FILL / TITLE_FONT use 1F4E78)
# so a pitch book and its supporting workbook read as one product.
PPT_NAVY = RGBColor(0x1F, 0x4E, 0x78)
PPT_NAVY_2 = RGBColor(0x1F, 0x4E, 0x78)
PPT_ACCENT = RGBColor(0x2E, 0x75, 0xB6)
PPT_TEAL = RGBColor(0x00, 0x8C, 0x95)
PPT_GREEN = RGBColor(0x35, 0x8A, 0x3C)
PPT_RED = RGBColor(0xB5, 0x2B, 0x35)
PPT_GOLD = RGBColor(0xB7, 0x8A, 0x2B)
PPT_DARK_TEXT = RGBColor(0x33, 0x33, 0x33)
PPT_MUTED_TEXT = RGBColor(0x66, 0x66, 0x66)
PPT_LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
PPT_MID_GRAY = RGBColor(0xD9, 0xDE, 0xE5)
PPT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PPT_BLACK = RGBColor(0x00, 0x00, 0x00)
PPT_WARM_GRAY = RGBColor(0xE8, 0xE3, 0xDA)

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5


def _new_pitch_book():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    return prs


def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank layout — full control over placement


def _add_textbox(slide, left, top, width, height, text, size=18, bold=False,
                  color=PPT_DARK_TEXT, align=PP_ALIGN.LEFT, italic=False, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def _add_top_rule(slide, color=PPT_NAVY, y=0.0, h=0.10):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(y), Inches(SLIDE_WIDTH_IN), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = color; sh.line.fill.background()
    return sh


def _add_section_header(slide, title, subtitle=None, info=None, ticker=None, page_no=None, source=None):
    _add_top_rule(slide, PPT_NAVY, 0, 0.08)
    _add_textbox(slide, 0.58, 0.28, 10.9, 0.42, title, size=25, bold=False, color=PPT_BLACK)
    if subtitle:
        _add_textbox(slide, 0.60, 0.72, 11.0, 0.34, subtitle, size=13.5, color=PPT_MUTED_TEXT)
    if info is not None:
        _add_logo(slide, info, ticker=ticker)
    elif ticker:
        _add_logo(slide, ticker, ticker=ticker)
    if page_no is not None:
        _add_footer(slide, page_no, source=source)


def _add_title_slide(prs, title, subtitle, info=None, ticker=None, date_text=None):
    slide = _blank_slide(prs)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_WIDTH_IN), Inches(SLIDE_HEIGHT_IN))
    bg.fill.solid()
    bg.fill.fore_color.rgb = PPT_NAVY
    bg.line.fill.background()
    bg.shadow.inherit = False

    _add_textbox(slide, 1, 2.8, 11.3, 1.4, title, size=40, bold=True, color=PPT_WHITE)
    _add_textbox(slide, 1, 4.1, 11.3, 0.8, subtitle, size=18, color=RGBColor(0xD9, 0xE2, 0xF0))
    if date_text:
        _add_textbox(slide, 1, 4.85, 5.5, 0.35, date_text, size=11, color=RGBColor(0xC6, 0xD5, 0xE7))
    if info is not None:
        _add_logo(slide, info, ticker=ticker, left=10.9, top=2.45, size=1.15)
    return slide


def _add_bullet_slide(prs, title, bullets, subtitle=None):
    slide = _blank_slide(prs)
    _add_section_header(slide, title)
    top = 1.3
    if subtitle:
        _add_textbox(slide, 0.6, top, 12.1, 0.5, subtitle, size=14, color=PPT_MUTED_TEXT)
        top += 0.6

    box = slide.shapes.add_textbox(Inches(0.6), Inches(top), Inches(12.1), Inches(5.6))
    tf = box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {bullet}"
        p.font.size = Pt(16)
        p.font.color.rgb = PPT_DARK_TEXT
        p.space_after = Pt(10)
    return slide


def _add_stat_cards_slide(prs, title, stats):
    """stats: list of (label, value_str) tuples — up to 4 cards across the slide."""
    slide = _blank_slide(prs)
    _add_section_header(slide, title)

    stats = stats[:4]
    n = len(stats)
    if n == 0:
        return slide
    card_width, gap = 2.7, 0.35
    total_width = n * card_width + (n - 1) * gap
    start_left = (SLIDE_WIDTH_IN - total_width) / 2
    top = 2.6

    for i, (label, value) in enumerate(stats):
        left = start_left + i * (card_width + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(card_width), Inches(2.0))
        card.fill.solid()
        card.fill.fore_color.rgb = PPT_LIGHT_GRAY
        card.line.fill.background()
        card.shadow.inherit = False

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = str(value)
        r1.font.size = Pt(24)
        r1.font.bold = True
        r1.font.color.rgb = PPT_NAVY

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = label
        r2.font.size = Pt(12)
        r2.font.color.rgb = PPT_MUTED_TEXT

    return slide


def _add_bar_chart_slide(prs, title, categories, series_name, values, number_format='#,##0.00', subtitle=None):
    slide = _blank_slide(prs)
    _add_section_header(slide, title)
    top = 1.3
    if subtitle:
        _add_textbox(slide, 0.6, top, 12.1, 0.5, subtitle, size=14, color=PPT_MUTED_TEXT)
        top += 0.6

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series_name, values)

    gframe = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1.0), Inches(top + 0.3), Inches(11.3), Inches(5.1), chart_data
    )
    chart = gframe.chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = number_format
    plot.data_labels.number_format_is_linked = False
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = PPT_ACCENT
    return slide


def _add_disclaimer_slide(prs):
    """
    A generic 'Notice to Recipient' slide, written in our own words — NOT
    copied from any real bank's actual disclaimer text, which is that
    firm's own drafted legal language. Every real pitch book has a slide
    like this; the specific wording here is original.
    """
    text = (
        "These materials are provided for discussion purposes only and do not constitute investment, legal, tax, "
        "or accounting advice, nor an offer or solicitation to buy or sell any security. Figures are estimates "
        "based on available data as of the date of preparation, derived from a combination of public market data "
        "and modeling assumptions, and are subject to change without notice. Actual results may differ materially. "
        "Recipients should conduct their own independent analysis and consult qualified professional advisors "
        "before relying on any information contained herein."
    )
    slide = _blank_slide(prs)
    _add_textbox(slide, 0.8, 0.6, 11.7, 0.6, "Notice to Recipient", size=24, bold=True, color=PPT_NAVY)
    _add_textbox(slide, 0.8, 1.5, 11.7, 4.5, text, size=13, color=PPT_DARK_TEXT)
    return slide


def _add_toc_slide(prs, items):
    """items: an ordered list of section title strings."""
    slide = _blank_slide(prs)
    _add_section_header(slide, "Table of Contents")
    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(9.5), Inches(5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{i + 1}.  {item}"
        p.font.size = Pt(18)
        p.font.color.rgb = PPT_DARK_TEXT
        p.space_after = Pt(14)
    return slide


def _add_table_slide(prs, title, headers, rows, subtitle=None, col_widths=None):
    """
    Generic styled data-table slide — used for transaction rationale,
    sensitivity/scenario analysis, sources & uses, and process timelines.
    `rows` is a list of lists/tuples, each matching len(headers).
    col_widths (optional) are relative weights, e.g. [2, 1, 1, 1].
    """
    slide = _blank_slide(prs)
    _add_section_header(slide, title)
    top = 1.3
    if subtitle:
        _add_textbox(slide, 0.6, top, 12.1, 0.5, subtitle, size=13, color=PPT_MUTED_TEXT)
        top += 0.55

    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_height = Inches(min(0.45 * n_rows + 0.3, 5.6))

    graphic_frame = slide.shapes.add_table(n_rows, n_cols, Inches(0.6), Inches(top + 0.2), Inches(12.1), table_height)
    table = graphic_frame.table

    if col_widths:
        total = sum(col_widths)
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(12.1 * w / total)

    for c, header in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = str(header)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PPT_NAVY
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size = Pt(11)
                run.font.bold = True
                run.font.color.rgb = PPT_WHITE

    for r, row_data in enumerate(rows, start=1):
        row_fill = PPT_LIGHT_GRAY if r % 2 == 0 else PPT_WHITE
        for c, value in enumerate(row_data):
            cell = table.cell(r, c)
            cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(10.5)
                    run.font.color.rgb = PPT_DARK_TEXT
    return slide


def _add_table(slide, left, top, width, height, headers, rows, col_widths=None, font_size=9.2, first_col_bold=False):
    n_rows = len(rows) + 1; n_cols = len(headers)
    table = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height)).table
    if col_widths:
        total = sum(col_widths)
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(width * w / total)
    for c, h in enumerate(headers):
        cell = table.cell(0, c); cell.text = str(h)
        cell.fill.solid(); cell.fill.fore_color.rgb = PPT_NAVY
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(font_size); r.font.bold = True; r.font.color.rgb = PPT_WHITE; r.font.name = 'Aptos'
    for r_idx, row in enumerate(rows, 1):
        fill = PPT_WHITE if r_idx % 2 else PPT_LIGHT_GRAY
        for c, val in enumerate(row):
            cell = table.cell(r_idx, c); cell.text = str(val)
            cell.fill.solid(); cell.fill.fore_color.rgb = fill
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                for rr in p.runs:
                    rr.font.size = Pt(font_size); rr.font.color.rgb = PPT_DARK_TEXT; rr.font.name = 'Aptos'
                    if c == 0 and first_col_bold:
                        rr.font.bold = True
    return table


def _add_logo(slide, info_or_name, ticker=None, left=11.95, top=0.25, size=0.75):
    """Real logo, tried in order: yfinance -> Wikipedia (free, no key) ->
    Finnhub (needs a key) -> a clean monogram. Best-effort throughout, so
    a blocked image request or missing source never breaks deck generation."""
    info = info_or_name if isinstance(info_or_name, dict) else {}
    name = info.get('shortName') or info.get('longName') or info_or_name or ticker or 'Co.'
    logo_url = fetch_best_available_logo(info, ticker or '', str(name))
    if logo_url:
        try:
            req = Request(logo_url, headers={'User-Agent': 'Mozilla/5.0'})
            raw = urlopen(req, timeout=5).read()
            slide.shapes.add_picture(BytesIO(raw), Inches(left), Inches(top), height=Inches(size))
            return
        except Exception:
            pass
    initials = ''.join(w[0] for w in re.findall(r'[A-Za-z0-9]+', str(name))[:2]).upper()
    initials = initials or (str(ticker or 'CO')[:2].upper())
    c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(size), Inches(size))
    c.fill.solid(); c.fill.fore_color.rgb = PPT_NAVY; c.line.fill.background()
    tf = c.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = initials; r.font.size = Pt(17); r.font.bold = True
    r.font.color.rgb = PPT_WHITE; r.font.name = 'Aptos'


def _add_footer(slide, page_no, source=None, confidential=True):
    _add_textbox(slide, 0.55, 7.13, 9.7, 0.20,
                 source or "Source: Company filings, public market data and internal calculations.",
                 size=7.5, color=PPT_MUTED_TEXT)
    if confidential:
        _add_textbox(slide, 10.0, 7.13, 2.15, 0.20, "CONFIDENTIAL", size=7.5, bold=True,
                     color=PPT_MUTED_TEXT, align=PP_ALIGN.RIGHT)
    _add_textbox(slide, 12.25, 7.10, 0.55, 0.22, str(page_no), size=8, color=PPT_MUTED_TEXT, align=PP_ALIGN.RIGHT)


def _add_banner(slide, text, top=1.02, fill=PPT_WARM_GRAY, color=PPT_DARK_TEXT, height=0.40):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.58), Inches(top), Inches(12.15), Inches(height))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill; sh.line.fill.background()
    _add_textbox(slide, 0.72, top + 0.04, 11.9, height - 0.06, text, size=10.5, bold=True,
                 color=color, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def _add_divider_slide(prs, title, subtitle=None, info=None, ticker=None, page_no=None):
    slide = _blank_slide(prs)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SLIDE_WIDTH_IN), Inches(0.10))
    band.fill.solid(); band.fill.fore_color.rgb = PPT_NAVY; band.line.fill.background()
    _add_textbox(slide, 0.95, 2.70, 11.4, 0.65, title, size=31, color=PPT_BLACK, align=PP_ALIGN.CENTER)
    if subtitle:
        _add_textbox(slide, 1.2, 3.42, 10.9, 0.45, subtitle, size=15, color=PPT_MUTED_TEXT, align=PP_ALIGN.CENTER)
    band2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6.35), Inches(SLIDE_WIDTH_IN), Inches(0.65))
    band2.fill.solid(); band2.fill.fore_color.rgb = PPT_NAVY; band2.line.fill.background()
    if info is not None:
        _add_logo(slide, info, ticker=ticker, left=0.65, top=6.72, size=0.48)
    if page_no is not None:
        _add_footer(slide, page_no, source=None)
    return slide


def _add_kpi_cards(slide, stats, top=1.55, card_width=2.82, height=1.45):
    stats = stats[:4]
    gap = 0.23
    total = len(stats) * card_width + (len(stats) - 1) * gap
    left = (SLIDE_WIDTH_IN - total) / 2
    for i, (label, value, kind) in enumerate(stats):
        x = left + i * (card_width + gap)
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(top), Inches(card_width), Inches(height))
        sh.fill.solid(); sh.fill.fore_color.rgb = PPT_LIGHT_GRAY; sh.line.color.rgb = PPT_MID_GRAY
        tf = sh.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(value); r.font.size = Pt(21); r.font.bold = True
        r.font.color.rgb = PPT_NAVY; r.font.name = 'Aptos'
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run(); r2.text = label; r2.font.size = Pt(9.5); r2.font.color.rgb = PPT_MUTED_TEXT; r2.font.name = 'Aptos'
        if kind == 'positive':
            sh.line.color.rgb = PPT_GREEN
        elif kind == 'negative':
            sh.line.color.rgb = PPT_RED


def _add_bullet_box(slide, left, top, width, height, title, bullets, accent=PPT_NAVY):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    sh.fill.solid(); sh.fill.fore_color.rgb = PPT_WHITE; sh.line.color.rgb = PPT_MID_GRAY
    tag = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.12), Inches(height))
    tag.fill.solid(); tag.fill.fore_color.rgb = accent; tag.line.fill.background()
    _add_textbox(slide, left + 0.25, top + 0.16, width - 0.45, 0.3, title, size=12, bold=True, color=accent)
    box = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.55), Inches(width - 0.45), Inches(height - 0.65))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = '• ' + str(b); p.font.size = Pt(10.3); p.font.color.rgb = PPT_DARK_TEXT; p.space_after = Pt(7)
    return sh


def _sanitize_series_values(values):
    """
    Replaces NaN/Inf with None so chart-writing libraries don't choke on
    them (xlsxwriter's write_number rejects NaN/Inf outright) — real,
    fairly common data gaps from yfinance, especially for non-US tickers
    with different trading calendars. None becomes a genuine gap in the
    chart, which is honest; substituting 0 would misleadingly imply the
    value was actually zero that day.
    """
    clean = []
    for v in values:
        if v is None or isinstance(v, str):
            clean.append(v)
        elif isinstance(v, (int, float)) and (math.isnan(v) or math.isinf(v)):
            clean.append(None)
        else:
            clean.append(v)
    return clean


def _add_column_chart(slide, left, top, width, height, categories, series, title=None, number_format='#,##0', colors=None):
    data = CategoryChartData(); data.categories = categories
    for name, vals in series:
        data.add_series(name, _sanitize_series_values(vals))
    frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(left), Inches(top), Inches(width), Inches(height), data)
    chart = frame.chart; chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    chart.has_legend = len(series) > 1
    if len(series) > 1:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.tick_labels.number_format = number_format
    chart.category_axis.tick_labels.font.size = Pt(9)
    for i, ser in enumerate(chart.series):
        ser.has_data_labels = True
        ser.data_labels.number_format = number_format
        ser.data_labels.font.size = Pt(8)
        if colors and i < len(colors):
            ser.format.fill.solid(); ser.format.fill.fore_color.rgb = colors[i]
    return chart


def _add_line_chart(slide, left, top, width, height, categories, series, title=None, number_format='0.0%'):
    data = CategoryChartData(); data.categories = categories
    for name, vals in series:
        data.add_series(name, _sanitize_series_values(vals))
    frame = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(left), Inches(top), Inches(width), Inches(height), data)
    chart = frame.chart; chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    chart.has_legend = len(series) > 1
    if len(series) > 1:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.value_axis.tick_labels.number_format = number_format
    chart.value_axis.has_major_gridlines = True
    for ser in chart.series:
        ser.format.line.width = Pt(2.2)
    return chart


def _add_waterfall(slide, left, top, width, height, labels, values, title=None, number_format='#,##0'):
    """Bank-style bridge drawn with shapes, so it doesn't depend on a
    third-party waterfall-chart library."""
    max_abs = max([abs(v) for v in values] + [1.0])
    baseline = top + height * 0.78
    chart_h = height * 0.62
    x0 = left; step = width / max(len(values), 1); bar_w = step * 0.55; cumulative = 0
    _add_textbox(slide, left, top - 0.30, width, 0.3, title or '', size=12, bold=True, color=PPT_NAVY)
    for i, (lab, val) in enumerate(zip(labels, values)):
        if i == 0:
            start, end = 0, val
        else:
            start, end = cumulative, cumulative + val
        lo, hi = min(start, end), max(start, end)
        h = chart_h * abs(hi - lo) / max_abs
        y = baseline - chart_h * hi / max_abs if hi >= 0 else baseline
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x0 + i * step + step * 0.22), Inches(y), Inches(bar_w), Inches(max(h, 0.02)))
        sh.fill.solid(); sh.fill.fore_color.rgb = PPT_GREEN if val >= 0 else PPT_RED; sh.line.fill.background()
        _add_textbox(slide, x0 + i * step, baseline + 0.10, step, 0.42, lab, size=8.5, color=PPT_DARK_TEXT, align=PP_ALIGN.CENTER)
        _add_textbox(slide, x0 + i * step + step * 0.02, y - 0.26, step * 0.96, 0.22, f"{val:+,.1f}", size=8.5, bold=True, color=PPT_DARK_TEXT, align=PP_ALIGN.CENTER)
        cumulative = end
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(baseline), Inches(width), Inches(0.015))
    ln.fill.solid(); ln.fill.fore_color.rgb = PPT_MID_GRAY; ln.line.fill.background()


def _price_fmt(v, symbol='$'):
    return 'N/A' if v is None else f'{symbol}{v:,.2f}'


def _safe_hist(ticker_obj, period='5y'):
    try:
        h = ticker_obj.history(period=period, auto_adjust=False)
        return None if (h is None or h.empty) else h
    except Exception:
        return None


def _get_financial_series(ticker_obj, row_names, max_years=5):
    try:
        df = ticker_obj.financials
    except Exception:
        return []
    for name in row_names:
        if df is not None and not df.empty and name in df.index:
            return [float(v) for v in reversed(df.loc[name].dropna().tolist()[:max_years])]
    return []


def _add_football_field(slide, left, top, width, height, ranges, current=None, title='Valuation Football Field'):
    """ranges: list of (label, low, high, color)."""
    valid = [x for x in ranges if x[1] is not None and x[2] is not None]
    if not valid:
        return
    lo = min(x[1] for x in valid); hi = max(x[2] for x in valid)
    if current is not None:
        lo = min(lo, current); hi = max(hi, current)
    span = max(hi - lo, 1e-9)
    _add_textbox(slide, left, top - 0.38, width, 0.3, title, size=13, bold=True, color=PPT_NAVY)
    y = top + 0.18; row_h = 0.58
    for label, a, b, color in valid:
        _add_textbox(slide, left, y, 1.75, 0.28, label, size=9.5, color=PPT_DARK_TEXT)
        x = left + 1.85; w = width - 2.25
        axis = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y + 0.13), Inches(w), Inches(0.015))
        axis.fill.solid(); axis.fill.fore_color.rgb = PPT_MID_GRAY; axis.line.fill.background()
        sx = x + w * (a - lo) / span; ex = x + w * (b - lo) / span
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(y + 0.03), Inches(max(ex - sx, 0.08)), Inches(0.22))
        bar.fill.solid(); bar.fill.fore_color.rgb = color; bar.line.fill.background()
        _add_textbox(slide, max(x, sx - 0.25), y + 0.28, 0.7, 0.18, f'{a:,.1f}', size=7.5, color=PPT_MUTED_TEXT)
        _add_textbox(slide, max(x, ex - 0.25), y + 0.28, 0.7, 0.18, f'{b:,.1f}', size=7.5, color=PPT_MUTED_TEXT, align=PP_ALIGN.RIGHT)
        y += row_h
    if current is not None:
        x = left + 1.85 + (current - lo) / span * (width - 2.25)
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(top), Inches(0.025), Inches(row_h * len(valid) - 0.08))
        line.fill.solid(); line.fill.fore_color.rgb = PPT_RED; line.line.fill.background()
        _add_textbox(slide, max(left + 1.7, x - 0.35), top - 0.02, 0.75, 0.20, f'Current {current:,.2f}', size=8, bold=True, color=PPT_RED, align=PP_ALIGN.CENTER)


def _add_heatmap(slide, left, top, width, height, row_values, col_values, matrix, row_label, col_label,
                  number_fmt='{:.1f}', row_fmt=None, col_fmt=None):
    row_fmt = row_fmt or (lambda v: f'{v:.1%}')  # default preserves existing percentage-based callers
    col_fmt = col_fmt or (lambda v: f'{v:.1%}')
    n_r, n_c = len(row_values), len(col_values)
    cell_w = width / (n_c + 1.4); cell_h = height / (n_r + 1.4)
    _add_textbox(slide, left, top - 0.38, width, 0.25, f'{row_label} vs. {col_label}', size=11, bold=True, color=PPT_NAVY)
    _add_textbox(slide, left, top - 0.08, 1.15, 0.25, row_label, size=8.5, bold=True, color=PPT_MUTED_TEXT, align=PP_ALIGN.CENTER)
    for c, v in enumerate(col_values):
        _add_textbox(slide, left + 1.2 + c * cell_w, top - 0.08, cell_w, 0.25, col_fmt(v), size=8.5, bold=True, color=PPT_NAVY, align=PP_ALIGN.CENTER)
    vals = [x for row in matrix for x in row if x is not None]
    mn = min(vals) if vals else 0; mx = max(vals) if vals else 1; span = max(mx - mn, 1e-9)
    for r, rv in enumerate(row_values):
        _add_textbox(slide, left, top + 0.30 + r * cell_h, 1.15, cell_h, row_fmt(rv), size=8.5, bold=True, color=PPT_NAVY, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        for c, cv in enumerate(col_values):
            val = matrix[r][c]
            t = (val - mn) / span if val is not None else 0.5
            rr = int(245 - 90 * t); gg = int(248 - 60 * t); bb = int(250 - 15 * t)
            sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left + 1.2 + c * cell_w), Inches(top + 0.25 + r * cell_h), Inches(cell_w - 0.03), Inches(cell_h - 0.03))
            sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(rr, gg, bb); sh.line.color.rgb = PPT_WHITE
            display = 'N/A' if val is None else (number_fmt(val) if callable(number_fmt) else number_fmt.format(val))
            _add_textbox(slide, left + 1.2 + c * cell_w, top + 0.25 + r * cell_h, cell_w - 0.03, cell_h - 0.03, display, size=8.5, bold=True, color=PPT_DARK_TEXT, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def _add_process_timeline(slide, left, top, width, phases):
    """phases: [(label, duration, detail), ...]"""
    n = len(phases); step = width / n; y = top + 1.1
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(left), Inches(y), Inches(width), Inches(0.42))
    arrow.fill.solid(); arrow.fill.fore_color.rgb = PPT_NAVY; arrow.line.fill.background()
    for i, (label, duration, detail) in enumerate(phases):
        x = left + i * step
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.05), Inches(top), Inches(step - 0.10), Inches(1.0))
        card.fill.solid(); card.fill.fore_color.rgb = PPT_LIGHT_GRAY; card.line.color.rgb = PPT_MID_GRAY
        _add_textbox(slide, x + 0.12, top + 0.10, step - 0.24, 0.27, label, size=10, bold=True, color=PPT_NAVY, align=PP_ALIGN.CENTER)
        _add_textbox(slide, x + 0.12, top + 0.39, step - 0.24, 0.20, duration, size=8.5, bold=True, color=PPT_ACCENT, align=PP_ALIGN.CENTER)
        _add_textbox(slide, x + 0.12, top + 1.72, step - 0.24, 1.0, detail, size=8.5, color=PPT_DARK_TEXT, align=PP_ALIGN.CENTER)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + step / 2 - 0.08), Inches(y + 0.10), Inches(0.16), Inches(0.16))
        circ.fill.solid(); circ.fill.fore_color.rgb = PPT_WHITE; circ.line.color.rgb = PPT_NAVY


def _dcf_valuation_range(dcf_result, ticker_symbol):
    """Real sensitivity endpoints from the actual DCF, not an arbitrary +/- %."""
    a = dcf_result['assumptions_used']; w = a['wacc']; g = a['terminal_growth_rate']
    vals = []
    for ww, gg in [(w + 0.02, g - 0.01), (w - 0.02, g + 0.01)]:
        try:
            r = run_dcf(ticker_symbol, starting_growth_rate=a['starting_growth_rate'], ending_growth_rate=a['ending_growth_rate'],
                        starting_fcf_margin=a['starting_fcf_margin'], ending_fcf_margin=a['ending_fcf_margin'],
                        wacc=ww, terminal_growth_rate=gg)
            vals.append(r['implied_share_price'])
        except Exception:
            pass
    if not vals:
        return dcf_result['implied_share_price'], dcf_result['implied_share_price']
    return min(vals), max(vals)


def _comps_implied_price_range(comps, target_shares):
    if not comps or not target_shares:
        return None
    t = comps.get('target') or {}
    net_debt = (t.get('enterprise_value') or 0) - (t.get('market_cap') or 0)
    ranges = []
    for key, field in [('ev_revenue', 'revenue'), ('ev_ebitda', 'ebitda')]:
        vals = [p.get(key) for p in comps.get('peers', []) if p.get(key) is not None and (key != 'ev_ebitda' or p.get(key) > 0)]
        if t.get(field) and vals:
            lowm, highm = _percentile(sorted(vals), .25), _percentile(sorted(vals), .75)
            ranges.append(((lowm * t[field] - net_debt) / target_shares, (highm * t[field] - net_debt) / target_shares))
    if not ranges:
        return None
    return min(x[0] for x in ranges), max(x[1] for x in ranges)


def _research_history_arrays(research):
    t = research['company']['ticker_obj']
    rev = _get_financial_series(t, REVENUE_ROW_NAMES, 5)
    gp = _get_financial_series(t, ["Gross Profit"], 5)
    ni = _get_financial_series(t, ['Net Income', 'NetIncome', 'Net Income Common Stockholders'], 5)
    years = []
    if rev:
        try:
            years = [str(x.year) for x in t.financials.columns[:len(rev)][::-1]]
        except Exception:
            years = []
        if len(years) != len(rev):
            # Real year labels weren't extractable (e.g. an unusual column
            # format) — fall back to generic FY labels rather than leaving
            # years shorter than rev, which would crash chart rendering
            # with a categories/data-length mismatch.
            years = [f"FY-{len(rev) - i - 1}" if i < len(rev) - 1 else "FY (Latest)" for i in range(len(rev))]
    return years, rev, gp, ni

def build_pitch_book_research(research, output_path):
    prs = _new_pitch_book(); c = research["company"]; info = c["info"]; ticker = c["ticker"]; snap = research["snapshot"]; dcf = research.get("dcf")
    sym = get_currency_symbol(c["quote_currency"]); today = datetime.now().strftime("%B %d, %Y")

    # Bug fix: comps and scenarios were previously expected on `research`
    # but never actually placed there by run_research(), silently leaving
    # the Comps and Scenario slides empty every time. Computed here instead.
    try:
        comps = build_comps_table(ticker, max_peers=6)
    except Exception:
        comps = None
    try:
        scenarios = run_scenarios(ticker)
    except Exception:
        scenarios = {}

    page = 1
    _add_title_slide(prs, c["company_name"], f"{ticker} | Equity Research & Valuation Materials", info, ticker, today)
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["Executive Summary", "Business & Industry", "Financial Performance", "Market Performance", "Valuation", "Investment Case", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "Investment Summary", "Headline valuation, operating profile and key debate", info, ticker, page)
    cards = [("Share Price", _price_fmt(research.get("current_price"), sym), "neutral")]
    if dcf:
        cards += [("DCF Value", _price_fmt(dcf.get("implied_share_price"), dcf.get("currency_symbol", sym)), "positive" if dcf.get("upside", 0) >= 0 else "negative"),
                  ("DCF Upside", f"{dcf.get('upside', 0):+.1%}", "positive" if dcf.get("upside", 0) >= 0 else "negative")]
    if snap.get("market_cap") is not None:
        cards.append(("Market Cap", f"{sym}{snap['market_cap']:,.0f}", "neutral"))
    _add_kpi_cards(s, cards, top=1.35, height=1.3, card_width=2.75)
    _add_bullet_box(s, .7, 3.05, 5.85, 2.7, "Investment Thesis",
                     [f"{c['company_name']} operates in {c['industry']} within {c['sector']}.",
                      "The valuation case should be underwritten through growth durability, margin trajectory, cash conversion and cost of capital.",
                      "DCF and trading comparables provide complementary intrinsic and market-based anchors."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.05, 5.85, 2.7, "Key Debate",
                     ["Sustainable growth versus normalization", "Margin durability", "Terminal value / WACC sensitivity",
                      "Whether the current multiple reflects growth and risk"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Dashboard", "Triangulation of intrinsic and market-based valuation", info, ticker, page)
    labels, vals = [], []
    if dcf:
        labels.append("DCF"); vals.append(dcf.get("implied_share_price"))
    if comps and comps.get("implied_ev_from_revenue") and comps.get("target", {}).get("revenue"):
        td = comps["target"]; nd = (td.get("total_debt", 0) or 0) - (td.get("cash", 0) or 0)
        sh = info.get("sharesOutstanding")
        if sh:
            labels.append("Trading Comps"); vals.append((comps["implied_ev_from_revenue"] - nd) / sh)
    if snap.get("analyst_target"):
        labels.append("Street Target"); vals.append(snap["analyst_target"])
    clean = [(l, v) for l, v in zip(labels, vals) if isinstance(v, (int, float))]
    if clean:
        _add_column_chart(s, .8, 1.55, 11.8, 4.7, [x[0] for x in clean], [("Implied Value", [x[1] for x in clean])],
                           title="Valuation Reference Points", number_format='#,##0', colors=[PPT_NAVY])
    if dcf:
        _add_banner(s, f"Current price: {sym}{research.get('current_price', 0):,.2f}  |  DCF upside / (downside): {dcf.get('upside', 0):+.1%}", 6.25, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Company Overview", "Business profile, market position and operating footprint", info, ticker, page)
    desc = info.get("longBusinessSummary") or info.get("description") or "Public-company business description was not available from the market-data source."
    _add_textbox(s, .75, 1.35, 7.5, 2.05, desc, size=11, color=PPT_DARK_TEXT)
    from_rows = [("Exchange", c.get("exchange")), ("Country", c.get("country")), ("Sector", c.get("sector")), ("Industry", c.get("industry")), ("Employees", info.get("fullTimeEmployees"))]
    _add_table(s, 8.55, 1.35, 3.9, 3.3, ["Metric", "Value"], [[str(a), str(b)] for a, b in from_rows], col_widths=[1.6, 2.3], font_size=8.8, first_col_bold=True)
    _add_bullet_box(s, .75, 4.05, 5.75, 1.85, "Operating Model", ["Evaluate revenue drivers, unit economics, gross-margin structure, operating leverage and FCF conversion."], PPT_NAVY)
    _add_bullet_box(s, 6.75, 4.05, 5.75, 1.85, "Data Quality", ["Descriptions and statistics depend on the available data feed; missing fields are left unavailable, not fabricated."], PPT_GOLD)

    years, rev, gp, ni = _research_history_arrays(research)
    page += 1; s = _blank_slide(prs); _add_section_header(s, "Historical Financial Performance", "Reported financial trajectory where available", info, ticker, page)
    if rev:
        _add_column_chart(s, .7, 1.4, 7.6, 4.9, years, [("Revenue", rev), ("Gross Profit", gp or [0] * len(rev)), ("Net Income", ni or [0] * len(rev))], title="Historical Income Statement", number_format='#,##0')
        rows = [[y, f"{rev[i]:,.0f}", f"{gp[i]:,.0f}" if gp else "N/A", f"{ni[i]:,.0f}" if ni else "N/A"] for i, y in enumerate(years)]
        _add_table(s, 8.55, 1.45, 3.85, 4.75, ["Year", "Revenue", "Gross Profit", "Net Income"], rows, col_widths=[.75, 1.1, 1.25, 1.15], font_size=7.8)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Historical data unavailable", ["The selected public-data feed did not provide sufficient historical financials for this chart."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Growth & Profitability", "Revenue growth, gross-margin and earnings trajectory", info, ticker, page)
    if rev and len(rev) >= 2:
        growth = [None] + [(rev[i] / rev[i - 1] - 1) if rev[i - 1] else None for i in range(1, len(rev))]
        gm = [(gp[i] / rev[i]) if gp and rev[i] else None for i in range(len(rev))]
        _add_line_chart(s, .75, 1.45, 6.0, 4.8, years, [("Revenue Growth", [x if x is not None else 0 for x in growth])], title="Revenue Growth", number_format='0.0%')
        _add_line_chart(s, 6.95, 1.45, 5.55, 4.8, years, [("Gross Margin", [x if x is not None else 0 for x in gm])], title="Gross Margin", number_format='0.0%')
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Margin history unavailable", ["Insufficient historical line items were returned for a robust margin series."], PPT_RED)

    hist = _safe_hist(c["ticker_obj"], "5y")
    page += 1; s = _blank_slide(prs); _add_section_header(s, "Share Price Performance", "Five-year market performance and valuation context", info, ticker, page)
    if hist is not None and not hist.empty:
        cats = [str(x.date()) for x in hist.index][::max(len(hist) // 80, 1)]
        prices = hist["Close"].tolist()[::max(len(hist) // 80, 1)]
        _add_line_chart(s, .7, 1.45, 8.0, 4.8, cats, [(ticker, prices)], title="Share Price", number_format='0.00')
        rows = [("Current", _price_fmt(research.get("current_price"), sym)), ("52W Low", _price_fmt(snap.get("fifty_two_week_low"), sym)),
                ("52W High", _price_fmt(snap.get("fifty_two_week_high"), sym)), ("Beta", f"{snap.get('beta'):.2f}" if isinstance(snap.get('beta'), (int, float)) else "N/A")]
        _add_table(s, 8.95, 1.55, 3.4, 3.2, ["Metric", "Value"], [[a, b] for a, b in rows], col_widths=[1.6, 1.8], font_size=9.2, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Price history unavailable", ["Historical market prices could not be retrieved from the selected feed."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Trading Comparables", "Peer benchmarking across valuation metrics", info, ticker, page)
    if comps:
        rows = []
        for p in [comps.get("target")] + comps.get("peers", []):
            if not p:
                continue
            rows.append([p.get("ticker", ""), p.get("company_name", "")[:24], f"{p.get('ev_revenue'):.1f}x" if isinstance(p.get('ev_revenue'), (int, float)) else "NM",
                         f"{p.get('ev_ebitda'):.1f}x" if isinstance(p.get('ev_ebitda'), (int, float)) else "NM", f"{p.get('pe_ratio'):.1f}x" if isinstance(p.get('pe_ratio'), (int, float)) else "NM"])
        _add_table(s, .65, 1.45, 12.05, 4.95, ["Ticker", "Company", "EV / Rev.", "EV / EBITDA", "P / E"], rows, col_widths=[.75, 3.5, 1.5, 1.6, 1.3], font_size=8.4, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Trading comps unavailable", ["No peer set was returned for this company."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "DCF Forecast", "Explicit forecast and free-cash-flow build", info, ticker, page)
    if dcf:
        cats = [f"FY+{i+1}" for i in range(len(dcf.get("projected_revenue", [])))]
        _add_column_chart(s, .7, 1.45, 7.0, 4.8, cats, [("Revenue", dcf.get("projected_revenue", [])), ("FCF", dcf.get("projected_fcf", []))], title="Revenue and FCF Forecast", number_format='#,##0')
        a = dcf["assumptions_used"]
        rows = [("Starting Growth", f"{a.get('starting_growth_rate', 0):.1%}"), ("Ending Growth", f"{a.get('ending_growth_rate', 0):.1%}"),
                ("Starting FCF Margin", f"{a.get('starting_fcf_margin', 0):.1%}"), ("Ending FCF Margin", f"{a.get('ending_fcf_margin', 0):.1%}"),
                ("WACC", f"{a.get('wacc', 0):.2%}"), ("Terminal Growth", f"{a.get('terminal_growth_rate', 0):.2%}")]
        _add_table(s, 8.0, 1.55, 4.4, 3.9, ["Assumption", "Value"], [[a2, b] for a2, b in rows], col_widths=[2.4, 1.6], font_size=9.3, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "DCF unavailable", [research.get("dcf_error") or "Insufficient data to run a DCF for this company."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "DCF Sensitivity", "Implied share price sensitivity to WACC and terminal growth", info, ticker, page)
    if dcf:
        waccs = [.07, .08, .09, .10, .11]; tgs = [.01, .02, .03, .04, .05]
        matrix = []
        for w in waccs:
            row = []
            for g in tgs:
                try:
                    row.append(run_dcf(ticker, wacc=w, terminal_growth_rate=g)["implied_share_price"])
                except Exception:
                    row.append(None)
            matrix.append(row)
        _add_heatmap(s, 1.0, 1.45, 11.3, 4.8, waccs, tgs, matrix, "WACC", "Terminal Growth", number_fmt=lambda x: f"{sym}{x:,.2f}")

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Football Field", "Triangulation of valuation methodologies", info, ticker, page)
    ranges = []
    if dcf and isinstance(dcf.get("implied_share_price"), (int, float)):
        lo, hi = _dcf_valuation_range(dcf, ticker); ranges.append(("DCF", lo, hi, PPT_NAVY))
    if comps:
        cr = _comps_implied_price_range(comps, info.get("sharesOutstanding"))
        if cr:
            ranges.append(("Trading Comps", cr[0], cr[1], PPT_ACCENT))
    if snap.get("analyst_target"):
        p = snap["analyst_target"]; ranges.append(("Street Target", p * .9, p * 1.1, PPT_TEAL))
    _add_football_field(s, .7, 1.55, 11.9, 4.7, ranges, current=research.get("current_price"))

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Bull / Base / Bear", "Scenario analysis across growth, margins and discount rate", info, ticker, page)
    rows = []
    for name in ("Bear", "Base", "Bull"):
        d = scenarios.get(name)
        if d:
            rows.append([name, _price_fmt(d.get("implied_share_price"), sym), f"{d.get('upside', 0):+.1%}",
                         f"{d.get('assumptions_used', {}).get('wacc', 0):.2%}", f"{d.get('assumptions_used', {}).get('starting_growth_rate', 0):.1%}"])
    if rows:
        _add_table(s, .8, 1.55, 7.0, 3.3, ["Scenario", "Value", "Upside", "WACC", "Starting Growth"], rows, col_widths=[1.1, 1.4, 1.3, 1.3, 1.5], font_size=9, first_col_bold=True)
        cats = ['Bear', 'Base', 'Bull']; vals2 = [scenarios[x]['implied_share_price'] for x in cats]
        _add_column_chart(s, .8, 5.05, 7.0, 1.9, cats, [("Implied Price", vals2)], number_format='0.00', colors=[PPT_RED, PPT_NAVY, PPT_GREEN])
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Scenarios unavailable", ["Could not compute Bear/Base/Bull scenarios for this company."], PPT_RED)
    _add_bullet_box(s, 8.15, 1.55, 4.15, 3.3, "Scenario Interpretation",
                     ["Bear tests lower growth / weaker margins and a higher discount rate.", "Bull tests stronger growth / margin realization and a lower discount rate.",
                      "A risk framework, not a forecast guarantee."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Investment Catalysts", "Potential factors that could change the valuation debate", info, ticker, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Fundamental", ["Revenue re-acceleration", "Margin expansion", "Better cash conversion", "New products / geographies"], PPT_GREEN)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Market", ["Multiple re-rating", "Positive earnings revisions", "Sector rotation", "Lower rates / WACC"], PPT_ACCENT)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Corporate", ["M&A / strategic actions", "Capital returns", "Guidance changes", "New partnerships"], PPT_TEAL)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Key Risks", "Downside cases investors should underwrite", info, ticker, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Operating", ["Growth deceleration", "Competitive share loss", "Margin pressure", "Customer concentration"], PPT_RED)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Financial", ["Weak FCF conversion", "Higher CapEx / NWC", "Leverage", "Refinancing risk"], PPT_RED)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Valuation", ["Multiple compression", "Higher WACC", "Lower terminal growth", "Estimate misses"], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Bridge", "From enterprise value to equity value", info, ticker, page)
    if dcf:
        _add_waterfall(s, .8, 1.45, 7.2, 4.7, ["PV of FCF", "PV of TV", "Enterprise Value", "Net Debt", "Equity Value"],
                        [sum(dcf.get("discounted_fcf", [])), dcf.get("enterprise_value", 0) - sum(dcf.get("discounted_fcf", [])), dcf.get("enterprise_value", 0), -dcf.get("net_debt", 0), dcf.get("equity_value", 0)],
                        title="DCF Enterprise-to-Equity Bridge", number_format='#,##0')
        _add_bullet_box(s, 8.25, 1.55, 4.1, 3.8, "Key Takeaway",
                         [f"DCF implies {sym}{dcf.get('implied_share_price', 0):,.2f} per share versus {sym}{dcf.get('current_price', 0):,.2f} current, "
                          f"{dcf.get('upside', 0):+.1%} upside / downside."], PPT_NAVY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Recent News & Information Set", "Public-information monitoring relevant to the debate", info, ticker, page)
    news = research.get("news") or []
    bullets = [n.get("title", "")[:160] for n in news[:6] if n.get("title")]
    if bullets:
        _add_bullet_box(s, .8, 1.45, 11.6, 4.9, "Recent Items", bullets, PPT_ACCENT)
    else:
        _add_bullet_box(s, .8, 1.45, 11.6, 4.9, "No recent news available", ["The connected data source did not return a usable news set."], PPT_GOLD)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Methodology", "Calculation framework and limitations", info, ticker, page)
    _add_bullet_box(s, .8, 1.4, 5.7, 4.9, "DCF", ["Explicit forecast discounted at WACC", "Terminal value via perpetual growth", "EV less net debt = equity value", "Price = equity value / diluted shares"], PPT_NAVY)
    _add_bullet_box(s, 6.75, 1.4, 5.7, 4.9, "Trading Comps", ["Peers selected by industry / sector mapping", "EV/Revenue, EV/EBITDA, P/E", "Quartiles used to triangulate value", "Non-meaningful multiples excluded"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Final Valuation Summary", "Selected valuation outputs", info, ticker, page)
    final_rows = []
    if dcf:
        final_rows.append(["DCF", _price_fmt(dcf.get("implied_share_price"), sym), f"{dcf.get('upside', 0):+.1%}"])
    if snap.get("analyst_target") and research.get("current_price"):
        final_rows.append(["Street Target", _price_fmt(snap.get("analyst_target"), sym), f"{(snap.get('analyst_target') - research.get('current_price')) / research.get('current_price'):+.1%}"])
    _add_table(s, .9, 1.5, 8.0, 3.6, ["Method", "Value / Share", "Upside / (Downside)"], final_rows, col_widths=[2.4, 2.2, 2.4], font_size=10, first_col_bold=True)
    _add_banner(s, "Investment conclusion should combine valuation triangulation, scenario analysis, fundamental diligence and current market conditions.", 5.55, fill=PPT_NAVY, color=PPT_WHITE, height=.65)

    prs.save(output_path)
    return output_path


def build_ma_sensitivity_rows(acquirer_ticker, target_ticker, premiums, pct_cash=0.5, pct_stock=0.5):
    """
    Re-runs the actual M&A model at each premium in `premiums` — this is
    genuinely equivalent to a real bank's "Analysis at Various Prices"
    slide, not a static/interpolated table. Each row's accretion/dilution
    and DCF/comps premium figures are freshly computed, not approximated.
    Slower than most pitch book slides since it means several full model
    runs (each involving its own DCF and comps calls) — expected, not a bug.
    """
    print(f"Building premium sensitivity analysis across {len(premiums)} scenarios "
          f"(re-runs the full model each time — may take a minute)...")
    rows = []
    for premium in premiums:
        try:
            scenario = run_ma_model(acquirer_ticker, target_ticker, offer_premium=premium,
                                     pct_cash=pct_cash, pct_stock=pct_stock)
        except Exception:
            continue
        accretion_str = (f"{scenario['accretion_dilution_pct']:+.1%}"
                          if scenario["accretion_dilution_pct"] is not None else "N/A")
        dcf_str = f"{scenario['premium_vs_dcf']:+.1%}" if scenario["premium_vs_dcf"] is not None else "N/A"
        comps_str = f"{scenario['premium_vs_comps']:+.1%}" if scenario["premium_vs_comps"] is not None else "N/A"
        rows.append([
            f"{premium:+.0%}" if premium != 0 else "0% (Current)",
            f"{scenario['target_currency_symbol']}{scenario['offer_price_per_share']:,.2f}",
            f"{scenario['target_currency_symbol']}{scenario['total_consideration']:,.0f}",
            accretion_str, dcf_str, comps_str,
        ])
    return rows


def build_pitch_book_ma(deal, output_path):
    prs = _new_pitch_book()
    tcs, acs = deal["target_currency_symbol"], deal["acquirer_currency_symbol"]
    acq, tgt = deal["acquirer_ticker"], deal["target_ticker"]
    today = datetime.now().strftime("%B %d, %Y")

    try:
        ai = resolve_company(acq)["info"]
    except Exception:
        ai = {}
    try:
        ti = resolve_company(tgt)["info"]
    except Exception:
        ti = {}

    page = 1
    _add_title_slide(prs, f"{deal['acquirer_name']} / {deal['target_name']}",
                      f"{acq} acquisition of {tgt}  |  M&A Transaction Materials", ti, tgt, today)
    try:
        _add_logo(prs.slides[0], ai, ticker=acq, left=9.45, top=2.45, size=1.05)
    except Exception:
        pass
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["Executive Summary", "Transaction Rationale", "Acquirer & Target Overview",
                                     "Valuation Analysis", "Deal Mechanics", "Synergies & Accretion / Dilution",
                                     "Sensitivities", "Execution & Risk", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "Executive Summary", "Transaction terms, valuation and headline economics", ti, tgt, page)
    _add_kpi_cards(s, [
        ("Offer Price", _price_fmt(deal.get("offer_price_per_share"), tcs), "neutral"),
        ("Premium", f"{deal.get('offer_premium', 0):.1%}", "negative"),
        ("Purchase Price", f"{tcs}{deal.get('total_consideration', 0):,.0f}", "neutral"),
        ("Accretion / Dilution", f"{deal.get('accretion_dilution_pct', 0):+.1%}" if deal.get('accretion_dilution_pct') is not None else "N/A",
         "positive" if (deal.get('accretion_dilution_pct') or 0) >= 0 else "negative"),
    ], top=1.35, card_width=2.75, height=1.3)
    _add_bullet_box(s, .7, 3.1, 5.8, 2.7, "Headline Takeaways",
                     [f"Offer represents a {deal['offer_premium']:.0%} premium to the target's undisturbed price.",
                      f"After-tax run-rate synergies estimated at {tcs}{deal['aftertax_synergies']:,.0f}.",
                      f"Transaction is {'accretive' if (deal['accretion_dilution_pct'] or 0) > 0 else 'dilutive'} to acquirer EPS under modeled assumptions."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.1, 5.8, 2.7, "Valuation Context",
                     [f"Premium vs. target DCF: {deal['premium_vs_dcf']:+.1%}" if deal['premium_vs_dcf'] is not None else "DCF comparison unavailable",
                      f"Premium vs. target comps: {deal['premium_vs_comps']:+.1%}" if deal['premium_vs_comps'] is not None else "Comps comparison unavailable",
                      "Deal value should be considered against standalone value and strategic synergies."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Transaction Rationale", "Strategic benefits and considerations by stakeholder", ti, tgt, page)
    rows = [
        ["For Acquirer", "Adds target earnings and cash flows", "Integration and execution risk"],
        ["Synergies", f"{tcs}{deal['pretax_synergies']:,.0f} pretax run-rate", "Realization timing / one-time costs"],
        ["Financing", f"{deal['pct_cash']:.0%} cash / {deal['pct_stock']:.0%} stock", f"New debt: {acs}{deal['new_debt_raised']:,.0f}"],
        ["For Target", f"{deal['offer_premium']:.0%} premium to current price", "Foregoes future standalone upside"],
    ]
    _add_table(s, .75, 1.45, 11.75, 3.1, ["Dimension", "Benefit / Rationale", "Key Consideration"], rows, col_widths=[1.5, 4, 3], font_size=9.5, first_col_bold=True)
    _add_banner(s, "A compelling transaction balances strategic value, price discipline, financing capacity and execution certainty.", 5.0, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Acquirer & Target Overview", "Relative scale, valuation and operating context", ti, tgt, page)
    rows = []
    for label, d, cur in [("Target", deal.get("target_dcf") or {}, tcs), ("Acquirer", deal.get("acquirer_dcf") or {}, acs)]:
        rows.append([label, _price_fmt(d.get("current_price"), cur), f"{d.get('market_cap', 0):,.0f}" if d.get('market_cap') else "N/A",
                     f"{d.get('enterprise_value', 0):,.0f}" if d.get('enterprise_value') else "N/A",
                     f"{d.get('implied_share_price', 0):,.2f}" if d.get('implied_share_price') else "N/A"])
    _add_table(s, .75, 1.5, 11.75, 2.0, ["Company", "Current Price", "Market Cap", "Enterprise Value", "DCF Value / Share"], rows, col_widths=[1.5, 1.4, 1.6, 1.8, 1.6], font_size=9.5, first_col_bold=True)
    _add_bullet_box(s, .75, 4.0, 5.75, 1.8, "Target", [f"{deal['target_name']} ({tgt})", "Standalone valuation should anchor the purchase-price discussion.", "Target shareholders receive immediate liquidity at a negotiated premium."], PPT_NAVY)
    _add_bullet_box(s, 6.75, 4.0, 5.75, 1.8, "Acquirer", [f"{deal['acquirer_name']} ({acq})", "Funding combines cash, new debt and equity issuance.", "Accretion depends on financing cost, synergy capture and purchase price."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Analysis", "Standalone valuation, offer price and premium context", ti, tgt, page)
    ranges = []
    if deal.get("target_dcf") and isinstance(deal["target_dcf"].get("implied_share_price"), (int, float)):
        lo, hi = _dcf_valuation_range(deal["target_dcf"], tgt); ranges.append(("Target DCF", lo, hi, PPT_NAVY))
    if deal.get("target_comps"):
        cr = _comps_implied_price_range(deal["target_comps"], deal.get("target_shares"))
        if cr:
            ranges.append(("Trading Comps", cr[0], cr[1], PPT_ACCENT))
    ranges.append(("Offer Price", deal["offer_price_per_share"], deal["offer_price_per_share"], PPT_RED))
    _add_football_field(s, .75, 1.55, 11.75, 3.6, ranges, current=deal["target_current_price"])
    rows = [["Undisturbed Price", _price_fmt(deal["target_current_price"], tcs)], ["Offer Price", _price_fmt(deal["offer_price_per_share"], tcs)],
            ["Premium", f"{deal['offer_premium']:.1%}"], ["DCF Premium", f"{deal['premium_vs_dcf']:.1%}" if deal['premium_vs_dcf'] is not None else "N/A"],
            ["Comps Premium", f"{deal['premium_vs_comps']:.1%}" if deal['premium_vs_comps'] is not None else "N/A"]]
    _add_table(s, .75, 5.45, 5.3, 1.25, ["Metric", "Value"], rows, col_widths=[2.4, 1.2], font_size=8.6, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Target Trading Comparables", "Peer multiples used to triangulate standalone value", ti, tgt, page)
    comps = deal.get("target_comps")
    if comps:
        rows = []
        for p in [comps.get("target")] + comps.get("peers", []):
            if not p:
                continue
            rows.append([p.get("ticker", ""), p.get("company_name", "")[:22], f"{p['ev_revenue']:.1f}x" if p.get('ev_revenue') else "N/A",
                         f"{p['ev_ebitda']:.1f}x" if p.get('ev_ebitda') else "N/A", f"{p['pe_ratio']:.1f}x" if p.get('pe_ratio') else "N/A"])
        _add_table(s, .65, 1.45, 12.05, 3.45, ["Ticker", "Company", "EV / Revenue", "EV / EBITDA", "P / E"], rows, col_widths=[.9, 3, 1.5, 1.5, 1.1], font_size=8.8, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Trading comps unavailable", ["No peer set was returned."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Sources & Uses", "Illustrative financing and purchase consideration", ti, tgt, page)
    rows = [["Cash on Balance Sheet", f"{acs}{deal['cash_from_balance_sheet']:,.0f}", "Purchase of Target Equity", f"{tcs}{deal['total_consideration']:,.0f}"],
            ["New Debt", f"{acs}{deal['new_debt_raised']:,.0f}", "", ""], ["Stock Issued", f"{acs}{deal['stock_consideration']:,.0f}", "", ""],
            ["Total Sources", f"{acs}{deal['cash_from_balance_sheet']+deal['new_debt_raised']+deal['stock_consideration']:,.0f}", "Total Uses", f"{tcs}{deal['total_consideration']:,.0f}"]]
    _add_table(s, .75, 1.5, 11.75, 2.55, ["Sources", "Amount", "Uses", "Amount"], rows, col_widths=[2.7, 1.5, 2.7, 1.5], font_size=9.2, first_col_bold=True)
    _add_bullet_box(s, .75, 4.45, 5.75, 1.6, "Cash Funding", [f"Cash used: {acs}{deal['cash_from_balance_sheet']:,.0f}", f"Foregone after-tax cash income: {acs}{deal['aftertax_foregone_income_on_cash']:,.0f}"], PPT_NAVY)
    _add_bullet_box(s, 6.75, 4.45, 5.75, 1.6, "Debt / Equity Funding", [f"New debt: {acs}{deal['new_debt_raised']:,.0f}", f"New shares issued: {deal['new_shares_issued']:,.0f}", f"After-tax interest: {acs}{deal['aftertax_interest_on_new_debt']:,.0f}"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Synergy Bridge", "Illustrative run-rate synergy build", ti, tgt, page)
    _add_waterfall(s, .7, 1.55, 11.4, 4.6, ["Target NI", "Synergies", "New Debt Interest", "Cash Opp. Cost", "Purchase Acct.", "Pro Forma NI"],
                    [deal['target_net_income'], deal['aftertax_synergies'], -deal['aftertax_interest_on_new_debt'], -deal['aftertax_foregone_income_on_cash'], -deal['aftertax_incremental_da_amort'], deal['combined_net_income']],
                    title="Illustrative Net Income Bridge", number_format='#,##0')
    _add_banner(s, f"Modeled after-tax synergies: {tcs}{deal['aftertax_synergies']:,.0f}; timing/one-time costs not separately modeled.", 6.3, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Purchase Accounting", "Purchase price allocation and its impact on pro forma earnings", ti, tgt, page)
    rows = [["Total Consideration", f"{tcs}{deal['total_consideration']:,.0f}"], ["Target Book Equity", f"{tcs}{deal['target_book_equity']:,.0f}"],
            ["Excess Purchase Price", f"{tcs}{deal['excess_purchase_price']:,.0f}"], ["Intangible Step-Up", f"{tcs}{deal['intangible_step_up']:,.0f}"],
            ["PP&E Step-Up", f"{tcs}{deal['ppe_step_up']:,.0f}"], ["Goodwill (residual)", f"{tcs}{deal['goodwill']:,.0f}"],
            ["Deferred Tax Liability", f"{tcs}{deal['deferred_tax_liability']:,.0f}"]]
    _add_table(s, .75, 1.5, 6.6, 3.5, ["Item", "Value"], rows, col_widths=[2.6, 1.6], font_size=9.3, first_col_bold=True)
    _add_bullet_box(s, 7.65, 1.5, 4.85, 3.5, "Earnings Impact",
                     [f"Incremental amortization: {tcs}{deal['incremental_amortization']:,.0f}/yr over {deal['intangible_useful_life_years']} years",
                      f"Incremental D&A: {tcs}{deal['incremental_da_from_stepup']:,.0f}/yr over {deal['ppe_useful_life_years']} years",
                      f"After-tax drag on pro forma net income: {tcs}{deal['aftertax_incremental_da_amort']:,.0f}/yr"], PPT_NAVY)
    _add_banner(s, "Goodwill is the RESIDUAL of the excess purchase price after allocating to identifiable intangibles and a PP&E fair-value step-up — not a separately estimated figure.", 5.35, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Pro Forma Accretion / Dilution", "Standalone versus pro forma EPS", ti, tgt, page)
    _add_column_chart(s, .85, 1.55, 5.6, 4.5, ["Acquirer Standalone", "Pro Forma"], [("EPS", [deal['acquirer_standalone_eps'], deal['pro_forma_eps']])], number_format='0.00', colors=[PPT_NAVY, PPT_TEAL])
    rows = [["Standalone EPS", f"{acs}{deal['acquirer_standalone_eps']:.2f}"], ["Pro Forma EPS", f"{acs}{deal['pro_forma_eps']:.2f}"],
            ["Accretion / (Dilution)", f"{deal['accretion_dilution_pct']:+.1%}" if deal['accretion_dilution_pct'] is not None else "N/A"],
            ["Pro Forma Shares", f"{deal['pro_forma_diluted_shares']:,.0f}"]]
    _add_table(s, 7.0, 1.7, 5.0, 2.25, ["Metric", "Value"], rows, col_widths=[2.7, 1.5], font_size=9.2, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Analysis at Various Prices", "Accretion / dilution across offer premiums — each row a real model re-run", ti, tgt, page)
    premiums_to_test = sorted(set([0.0, 0.05, 0.10, 0.15, 0.20, 0.25, deal["offer_premium"]]))
    sensitivity_rows = build_ma_sensitivity_rows(acq, tgt, premiums_to_test,
                                                   pct_cash=deal["cash_consideration"] / deal["total_consideration"],
                                                   pct_stock=deal["stock_consideration"] / deal["total_consideration"])
    if sensitivity_rows:
        _add_table(s, .6, 1.5, 12.1, 3.7, ["Premium", "Offer Price", "Total Consideration", "Accretion/(Dilution)", "Prem. vs. DCF", "Prem. vs. Comps"],
                   sensitivity_rows, col_widths=[.9, 1.3, 2, 1.6, 1.3, 1.3], font_size=8.4)
    _add_banner(s, "Each row re-runs the full transaction model rather than interpolating between scenarios.", 5.45, fill=PPT_LIGHT_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Premium Sensitivity", "Offer price at each premium level", ti, tgt, page)
    base = deal.get("target_current_price", 0) or 0
    offers = [base * (1 + p) for p in [.10, .20, .30, .40, .50]]
    _add_column_chart(s, .8, 1.45, 7.1, 4.7, [f"{p:.0%}" for p in [.10, .20, .30, .40, .50]], [("Offer Price", offers)], title="Offer Price by Premium", number_format='0.00')

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Illustrative Process Timeline", "Typical workstreams from evaluation through closing", ti, tgt, page)
    phases = [("Initial Evaluation", "2-4 wks", "Board approval, advisor engagement"), ("Negotiation", "4-6 wks", "Diligence, valuation, definitive terms"),
              ("Signing / Announcement", "~2 wks", "Documentation, signing, disclosure"), ("Regulatory / Vote", "8-14 wks", "Regulatory review, shareholder process"),
              ("Closing", "—", "Conditions satisfied, consideration paid")]
    _add_process_timeline(s, .75, 1.55, 11.8, phases)
    _add_banner(s, "Illustrative only — actual timing depends on deal structure, jurisdiction and negotiation dynamics.", 5.9, fill=PPT_LIGHT_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Transaction Considerations & Risks", "Key diligence topics before execution", ti, tgt, page)
    _add_bullet_box(s, .75, 1.4, 5.8, 4.9, "Execution Considerations", ["Synergy validation and timing", "Integration planning and one-time costs", "Financing availability", "Regulatory / shareholder approvals"], PPT_NAVY)
    _add_bullet_box(s, 6.78, 1.4, 5.8, 4.9, "Downside Risks", ["Overpayment relative to standalone value", "Lower synergy capture than modeled", "Higher funding costs / leverage", "EPS dilution from stock consideration"], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Standalone DCF", "Target DCF reference outputs", ti, tgt, page)
    if deal.get("target_dcf"):
        d = deal["target_dcf"]; a = d["assumptions_used"]
        rows = [["Current Price", _price_fmt(d['current_price'], tcs)], ["Implied Share Price", _price_fmt(d['implied_share_price'], tcs)],
                ["Upside / Downside", f"{d['upside']:+.1%}"], ["WACC", f"{a['wacc']:.1%}"], ["Terminal Growth", f"{a['terminal_growth_rate']:.1%}"]]
        _add_table(s, .8, 1.45, 5.3, 2.8, ["Metric", "Value"], rows, col_widths=[2.7, 1.5], font_size=9.3, first_col_bold=True)
        _add_column_chart(s, 6.55, 1.45, 5.6, 3.5, [f'Year {i}' for i in range(1, len(d['projected_revenue']) + 1)], [("Revenue", d['projected_revenue']), ("FCF", d['projected_fcf'])], number_format='#,##0', colors=[PPT_ACCENT, PPT_TEAL])
    _add_banner(s, "See the standalone Excel DCF for a fully editable live model with WACC build-up and Checks tab.", 5.35, fill=PPT_LIGHT_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Transaction Summary", "Key terms at a glance", ti, tgt, page)
    rows = [["Target", f"{deal['target_name']} ({tgt})"], ["Acquirer", f"{deal['acquirer_name']} ({acq})"],
            ["Offer Price", f"{tcs}{deal['offer_price_per_share']:,.2f}"], ["Total Consideration", f"{tcs}{deal['total_consideration']:,.0f}"],
            ["Cash / Stock", f"{deal['pct_cash']:.0%} / {deal['pct_stock']:.0%}"], ["Run-Rate Synergies", f"{tcs}{deal['pretax_synergies']:,.0f} pre-tax"],
            ["Accretion / Dilution", f"{deal['accretion_dilution_pct']:+.1%}" if deal['accretion_dilution_pct'] is not None else "N/A"]]
    _add_table(s, 1.6, 1.55, 10.1, 3.8, ["Term", "Illustrative Value"], rows, col_widths=[2.8, 4.2], font_size=10, first_col_bold=True)

    prs.save(output_path)
    return output_path


def build_pitch_book_ipo(ipo, currency_symbol, output_path):
    prs = _new_pitch_book(); logo_info = {"shortName": ipo["company_name"]}
    sym = currency_symbol; today = datetime.now().strftime("%B %d, %Y")
    base = ipo.get("discounted_price_per_share") or ipo.get("implied_price_per_share") or 0

    page = 1
    _add_title_slide(prs, ipo["company_name"], f"IPO Valuation & Investor Presentation Materials  |  {ipo['industry']}", logo_info, None, today)
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["IPO Overview", "Business & Financial Profile", "Peer Benchmarking", "Valuation", "Pricing", "Diligence & Risks", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "IPO Executive Summary", "Indicative valuation, capitalization and key considerations", logo_info, None, page)
    stats = [("Revenue", f"{sym}{ipo.get('revenue', 0):,.0f}", "neutral")]
    if ipo.get('ebitda') is not None:
        stats.append(("EBITDA", f"{sym}{ipo.get('ebitda', 0):,.0f}", "neutral"))
    stats.append(("Blended Value", f"{sym}{ipo.get('implied_price_per_share', 0):,.2f}", "neutral"))
    stats.append(("After Discount", f"{sym}{base:,.2f}", "positive"))
    _add_kpi_cards(s, stats, top=1.35, card_width=2.75, height=1.3)
    _add_bullet_box(s, .7, 3.05, 5.85, 2.75, "Equity Story", ["IPO valuation should be supported by durable growth, attractive unit economics, credible cash conversion and a differentiated market position."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.05, 5.85, 2.75, "Pricing Objective", ["Balance valuation maximization with aftermarket performance, demand depth, float, dilution and market conditions."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "IPO Snapshot", "Capitalization, valuation and peer framework", logo_info, None, page)
    rows = [("Industry", ipo.get("industry")), ("Revenue", f"{sym}{ipo.get('revenue', 0):,.0f}"),
            ("EBITDA", f"{sym}{ipo.get('ebitda', 0):,.0f}" if ipo.get('ebitda') is not None else "N/A"),
            ("Net Debt", f"{sym}{ipo.get('net_debt', 0):,.0f}"),
            ("Post-IPO Shares", f"{ipo.get('shares_outstanding_post_ipo', 0):,.0f}" if ipo.get('shares_outstanding_post_ipo') else "Not supplied")]
    _add_table(s, .8, 1.45, 5.3, 4.0, ["Metric", "Value"], [[a, b] for a, b in rows], col_widths=[2, 3], font_size=9.5, first_col_bold=True)
    _add_bullet_box(s, 6.55, 1.45, 5.5, 4.0, "Valuation Framework", ["Public-market peers provide the primary external anchor. EV/Revenue, EV/EBITDA and P/E are blended where the relevant metric is available."], PPT_NAVY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Business & Investment Highlights", "IPO investor lens", logo_info, None, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Growth", ["TAM expansion", "Customer growth", "Pricing / mix", "New markets / products"], PPT_GREEN)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Economics", ["Gross margin", "EBITDA margin", "FCF conversion", "Capital intensity"], PPT_NAVY)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Public Market", ["Peer positioning", "Multiple support", "Float / liquidity", "Post-IPO catalysts"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Financial Profile", "Provided financial inputs", logo_info, None, page)
    labels = ["Revenue", "EBITDA", "Net Income"]; vals = [ipo.get("revenue"), ipo.get("ebitda"), ipo.get("net_income")]
    clean = [(l, v) for l, v in zip(labels, vals) if isinstance(v, (int, float))]
    if clean:
        _add_column_chart(s, .8, 1.5, 11.8, 4.7, [x[0] for x in clean], [("Provided Financials", [x[1] for x in clean])], title="Financial Profile", number_format='#,##0', colors=[PPT_NAVY])

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Trading Comparables", "Selected public-company peer benchmarking", logo_info, None, page)
    if ipo.get("peers"):
        rows = [[p["ticker"], p["company_name"][:24], f"{p['ev_revenue']:.1f}x" if p.get('ev_revenue') else "N/A",
                 f"{p['ev_ebitda']:.1f}x" if p.get('ev_ebitda') else "N/A", f"{p['pe_ratio']:.1f}x" if p.get('pe_ratio') else "N/A"] for p in ipo["peers"]]
        rows.append(["", "Peer Average", f"{ipo['avg_ev_revenue']:.1f}x" if ipo.get('avg_ev_revenue') else "N/A",
                      f"{ipo['avg_ev_ebitda']:.1f}x" if ipo.get('avg_ev_ebitda') else "N/A", f"{ipo['avg_pe']:.1f}x" if ipo.get('avg_pe') else "N/A"])
        _add_table(s, .6, 1.4, 12.1, 3.7, ["Ticker", "Company", "EV / Revenue", "EV / EBITDA", "P / E"], rows, col_widths=[.9, 3.3, 1.5, 1.5, 1.1], font_size=8.6, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Peer set unavailable", ["No public-company peer data was returned."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Framework", "Implied share price by methodology", logo_info, None, page)
    cats, vals2 = [], []
    for label, key in [("EV / Revenue", "implied_price_from_revenue"), ("EV / EBITDA", "implied_price_from_ebitda"),
                        ("P / E", "implied_price_from_pe"), ("Blended", "implied_price_per_share"), ("After Discount", "discounted_price_per_share")]:
        if ipo.get(key) is not None:
            cats.append(label); vals2.append(ipo[key])
    if cats:
        _add_column_chart(s, .75, 1.45, 7.0, 4.55, cats, [("Implied Price / Share", vals2)], number_format='0.00', colors=[PPT_ACCENT])
    rows = [[label, f"{sym}{ipo[key]:,.2f}"] for label, key in
            [("EV / Revenue", "implied_price_from_revenue"), ("EV / EBITDA", "implied_price_from_ebitda"),
             ("P / E", "implied_price_from_pe"), ("Blended", "implied_price_per_share"), ("After Discount", "discounted_price_per_share")] if ipo.get(key) is not None]
    _add_table(s, 8.2, 1.55, 4.2, 3.1, ["Method", "Price / Share"], rows, col_widths=[2, 1.2], font_size=9.3, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "IPO Valuation Football Field", "Valuation range by method", logo_info, None, page)
    ranges = []
    for label, key, color in [("EV / Revenue", "implied_price_from_revenue", PPT_NAVY), ("EV / EBITDA", "implied_price_from_ebitda", PPT_ACCENT), ("P / E", "implied_price_from_pe", PPT_TEAL)]:
        p = ipo.get(key)
        if isinstance(p, (int, float)):
            ranges.append((label, p * .9, p * 1.1, color))
    if ipo.get("discounted_price_per_share"):
        p = ipo["discounted_price_per_share"]; ranges.append(("After Discount", p * .95, p * 1.05, PPT_RED))
    _add_football_field(s, .7, 1.55, 11.9, 4.6, ranges)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Indicative IPO Pricing Range", "Low / base / high pricing framework", logo_info, None, page)
    _add_kpi_cards(s, [("Low", f"{sym}{base*.90:,.2f}", "neutral"), ("Base", f"{sym}{base:,.2f}", "positive"), ("High", f"{sym}{base*1.10:,.2f}", "neutral")], top=1.55, card_width=3.4, height=1.5)
    _add_bullet_box(s, .8, 3.7, 5.5, 2.1, "Issuer Lens", ["Maximize proceeds while retaining sufficient investor demand and healthy aftermarket performance."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.7, 5.5, 2.1, "Investor Lens", ["Require an appropriate discount / valuation premium relative to public peers and growth quality."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Capitalization & Equity Value", "Post-IPO share count and enterprise value", logo_info, None, page)
    shares = ipo.get("shares_outstanding_post_ipo"); equity = shares * base if shares else None
    if equity is not None:
        _add_waterfall(s, .8, 1.5, 7.0, 4.6, ["Equity Value", "Net Debt", "Enterprise Value"], [equity, ipo.get("net_debt", 0), equity + (ipo.get("net_debt", 0) or 0)],
                        title="Equity-to-Enterprise Value Bridge", number_format='#,##0')
    rows = [("Post-IPO Shares", f"{shares:,.0f}" if shares else "N/A"), ("Base Price", f"{sym}{base:,.2f}"),
            ("Equity Value", f"{sym}{equity:,.0f}" if equity is not None else "N/A"), ("Net Debt", f"{sym}{ipo.get('net_debt', 0):,.0f}")]
    _add_table(s, 8.15, 1.55, 4.0, 3.7, ["Metric", "Value"], [[a, b] for a, b in rows], col_widths=[1.8, 1.6], font_size=9.2, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "IPO Risks & Diligence Topics", "Items investors should diligence", logo_info, None, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Business", ["Competitive intensity", "Customer concentration", "Market size and growth"], PPT_RED)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Financial", ["Margin durability", "Cash conversion", "Net debt / leverage"], PPT_RED)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "IPO / Market", ["Lock-up expiry and float", "Valuation compression", "Execution risk"], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Methodology", "How the IPO value is calculated", logo_info, None, page)
    _add_bullet_box(s, .8, 1.4, 5.7, 4.9, "Core Methods", ["EV/Revenue = peer multiple × revenue, less net debt, / shares", "EV/EBITDA = peer multiple × EBITDA, less net debt, / shares",
                     "P/E = peer multiple × net income / shares", "Blended = equal-weighted average of available methods"], PPT_NAVY)
    _add_bullet_box(s, 6.75, 1.4, 5.7, 4.9, "IPO Adjustment", [f"Illustrative illiquidity discount ({ipo.get('illiquidity_discount', 0):.0%})",
                     "An analytical convention, not a market guarantee", "Final price must reflect investor demand and market conditions"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Final Valuation Summary", "Selected point estimates", logo_info, None, page)
    rows = [[label, f"{sym}{ipo[key]:,.2f}" if ipo.get(key) else "N/A"] for label, key in
            [("EV / Revenue", "implied_price_from_revenue"), ("EV / EBITDA", "implied_price_from_ebitda"), ("P / E", "implied_price_from_pe"),
             ("Blended", "implied_price_per_share"), ("After Discount", "discounted_price_per_share")]]
    _add_table(s, 1.1, 1.5, 8.8, 3.8, ["Method", "Price / Share"], rows, col_widths=[4.0, 2.3], font_size=10, first_col_bold=True)
    _add_banner(s, f"Illustrative base pricing: {sym}{base:,.2f} per share", 5.65, fill=PPT_NAVY, color=PPT_WHITE, height=.65)

    prs.save(output_path)
    return output_path


def _lbo_quick_returns(result, entry_multiple, exit_multiple, leverage_multiple=None,
                        fees_pct=0.02, interest_rate=0.08, mandatory_amort_pct=0.05, cash_sweep_pct=1.0):
    """
    Pure-Python recompute of LBO MOIC/IRR for sensitivity grids — mirrors
    the Excel model's debt schedule exactly (same iterative solve for the
    interest/debt circularity used in export_lbo_to_excel), without
    needing to re-fetch company data for every grid cell. Verified
    separately to reproduce the Excel model's own MOIC/IRR to the same
    precision on the base case.
    """
    a = result["assumptions_used"]
    leverage_multiple = leverage_multiple if leverage_multiple is not None else result["leverage_multiple"]
    entry_ebitda = result["entry_ebitda"]
    n = result["hold_years"]
    entry_ev = entry_ebitda * entry_multiple
    new_debt = entry_ebitda * leverage_multiple
    fees = entry_ev * fees_pct
    sponsor_equity = (entry_ev + fees) - new_debt
    if sponsor_equity <= 0:
        return None, None

    revenue = result["most_recent_revenue"]
    debt = new_debt
    ebitda = entry_ebitda
    for year in range(1, n + 1):
        progress = (year - 1) / (n - 1) if n > 1 else 1
        growth = a["starting_growth_rate"] + (a["ending_growth_rate"] - a["starting_growth_rate"]) * progress
        revenue = revenue * (1 + growth)
        margin = a["starting_fcf_margin"] + (a["ending_fcf_margin"] - a["starting_fcf_margin"]) * progress
        ebitda = revenue * margin
        da = revenue * a["da_pct_revenue"]
        ebit = ebitda - da
        capex = revenue * a["capex_pct_revenue"]
        beginning_debt = debt
        mandatory_amort = min(beginning_debt * mandatory_amort_pct, beginning_debt)
        ending_debt_guess = beginning_debt - mandatory_amort
        for _ in range(20):
            avg_debt = (beginning_debt + ending_debt_guess) / 2
            interest = avg_debt * interest_rate
            pretax = ebit - interest
            tax = max(pretax, 0) * a["tax_rate"]
            ni = pretax - tax
            fcf = ni + da - capex
            available = fcf - mandatory_amort
            sweep = max(min(available * cash_sweep_pct, beginning_debt - mandatory_amort), 0)
            new_guess = beginning_debt - mandatory_amort - sweep
            if abs(new_guess - ending_debt_guess) < 1:
                ending_debt_guess = new_guess
                break
            ending_debt_guess = new_guess
        debt = ending_debt_guess

    exit_ev = ebitda * exit_multiple
    exit_equity = exit_ev - debt
    moic = exit_equity / sponsor_equity
    irr = moic ** (1 / n) - 1 if moic > 0 else None
    return moic, irr


def build_pitch_book_lbo(result, output_path):
    prs = _new_pitch_book()
    logo_info = {"shortName": result["company_name"]}
    sym = result["currency_symbol"]; today = datetime.now().strftime("%B %d, %Y")
    a = result["assumptions_used"]; n = result["hold_years"]

    entry_ev = result["entry_ebitda"] * result["entry_multiple"]
    new_debt = result["entry_ebitda"] * result["leverage_multiple"]
    fees = entry_ev * 0.02
    sponsor_equity = (entry_ev + fees) - new_debt
    base_moic, base_irr = _lbo_quick_returns(result, result["entry_multiple"], result["exit_multiple"])

    page = 1
    _add_title_slide(prs, f"{result['company_name']} ({result['ticker']})", f"Leveraged Buyout Analysis  |  {n}-Year Hold Period", logo_info, result["ticker"], today)
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["Executive Summary", "Transaction Overview", "Operating Plan", "Debt Paydown",
                                     "Returns Analysis", "Returns Sensitivity", "Value Creation", "Risks", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "Executive Summary", "Entry economics and headline returns", logo_info, result["ticker"], page)
    _add_kpi_cards(s, [
        ("Entry EV", f"{sym}{entry_ev:,.0f}", "neutral"),
        ("Leverage", f"{result['leverage_multiple']:.2f}x", "neutral"),
        ("MOIC", f"{base_moic:.2f}x" if base_moic else "N/A", "positive" if (base_moic or 0) >= 2.0 else "neutral"),
        ("IRR", f"{base_irr:.1%}" if base_irr else "N/A", "positive" if (base_irr or 0) >= 0.20 else "neutral"),
    ], top=1.35, card_width=2.75, height=1.3)
    _add_bullet_box(s, .7, 3.1, 5.8, 2.7, "Transaction Highlights",
                     [f"Entry at {result['entry_multiple']:.1f}x EBITDA, {'the company\u2019s own current market multiple' if result.get('market_ev_ebitda') else 'an assumed multiple'}.",
                      f"Financed with {result['leverage_multiple']:.1f}x EBITDA of new debt, {sym}{sponsor_equity:,.0f} of sponsor equity.",
                      f"Exit assumed at the same {result['exit_multiple']:.1f}x multiple — no multiple expansion baked in."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.1, 5.8, 2.7, "Return Drivers",
                     ["EBITDA growth from revenue growth and margin expansion.", "Debt paydown via a 100% free-cash-flow sweep after mandatory amortization.",
                      "No assumed multiple expansion — returns are driven by operating performance and deleveraging, not re-rating."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Transaction Overview", "Entry sources & uses", logo_info, result["ticker"], page)
    rows = [["New Debt", f"{sym}{new_debt:,.0f}", "Purchase of Enterprise Value", f"{sym}{entry_ev:,.0f}"],
            ["Sponsor Equity", f"{sym}{sponsor_equity:,.0f}", "Transaction Fees", f"{sym}{fees:,.0f}"],
            ["Total Sources", f"{sym}{new_debt + sponsor_equity:,.0f}", "Total Uses", f"{sym}{entry_ev + fees:,.0f}"]]
    _add_table(s, .75, 1.5, 11.75, 2.0, ["Sources", "Amount", "Uses", "Amount"], rows, col_widths=[2.7, 1.5, 2.7, 1.5], font_size=9.5, first_col_bold=True)
    _add_bullet_box(s, .75, 3.9, 11.75, 1.9, "Financing Structure",
                     [f"Debt sized at {result['leverage_multiple']:.1f}x entry EBITDA — a standard leverage level for a first-pass LBO screen.",
                      "Transaction fees estimated at 2% of enterprise value.", "Sponsor equity is the plug: whatever isn't covered by debt."], PPT_NAVY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Operating Plan", "Revenue and EBITDA build over the hold period", logo_info, result["ticker"], page)
    revenue, ebitda_series, revs = result["most_recent_revenue"], [], []
    rev_track = revenue
    for year in range(1, n + 1):
        progress = (year - 1) / (n - 1) if n > 1 else 1
        growth = a["starting_growth_rate"] + (a["ending_growth_rate"] - a["starting_growth_rate"]) * progress
        rev_track = rev_track * (1 + growth)
        margin = a["starting_fcf_margin"] + (a["ending_fcf_margin"] - a["starting_fcf_margin"]) * progress
        revs.append(rev_track); ebitda_series.append(rev_track * margin)
    cats = [f"Year {i}" for i in range(1, n + 1)]
    _add_column_chart(s, .75, 1.5, 11.75, 4.7, cats, [("Revenue", revs), ("EBITDA", ebitda_series)], title="Revenue & EBITDA Forecast", number_format='#,##0', colors=[PPT_NAVY, PPT_ACCENT])
    _add_banner(s, f"Revenue growth fades from {a['starting_growth_rate']:.1%} to {a['ending_growth_rate']:.1%}; EBITDA margin moves from {a['starting_fcf_margin']:.1%} to {a['ending_fcf_margin']:.1%}.", 6.35, fill=PPT_WARM_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Debt Paydown", "Deleveraging via mandatory amortization and 100% FCF cash sweep", logo_info, result["ticker"], page)
    debt_track = new_debt; debt_series = []
    for i, year in enumerate(range(1, n + 1)):
        rev_y = revs[i]; ebitda_y = ebitda_series[i]
        da = rev_y * a["da_pct_revenue"]; ebit = ebitda_y - da; capex = rev_y * a["capex_pct_revenue"]
        beginning_debt = debt_track
        mandatory_amort = min(beginning_debt * 0.05, beginning_debt)
        ending_debt_guess = beginning_debt - mandatory_amort
        for _ in range(20):
            avg_debt = (beginning_debt + ending_debt_guess) / 2
            interest = avg_debt * 0.08
            pretax = ebit - interest; tax = max(pretax, 0) * a["tax_rate"]; ni = pretax - tax
            fcf = ni + da - capex; available = fcf - mandatory_amort
            sweep = max(min(available, beginning_debt - mandatory_amort), 0)
            new_guess = beginning_debt - mandatory_amort - sweep
            if abs(new_guess - ending_debt_guess) < 1:
                ending_debt_guess = new_guess; break
            ending_debt_guess = new_guess
        debt_track = ending_debt_guess
        debt_series.append(debt_track)
    _add_line_chart(s, .75, 1.5, 11.75, 4.7, cats, [("Ending Debt Balance", debt_series)], title="Debt Balance Over the Hold Period", number_format='#,##0')
    _add_banner(s, f"Debt reduces from {sym}{new_debt:,.0f} at entry to {sym}{debt_series[-1]:,.0f} at exit — a {(1 - debt_series[-1]/new_debt):.0%} reduction.", 6.35, fill=PPT_LIGHT_GRAY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Returns Analysis", "Entry to exit equity value bridge", logo_info, result["ticker"], page)
    exit_ev = ebitda_series[-1] * result["exit_multiple"]
    exit_equity = exit_ev - debt_series[-1]
    _add_waterfall(s, .8, 1.55, 11.6, 4.6, ["Sponsor Equity (Entry)", "EBITDA Growth Value", "Debt Paydown Value", "Exit Equity Value"],
                    [sponsor_equity, (exit_ev - entry_ev), (new_debt - debt_series[-1]), exit_equity],
                    title="Illustrative Value Creation Bridge", number_format='#,##0')
    _add_bullet_box(s, .8, 6.35, 11.6, .8, "", [f"Exit Enterprise Value: {sym}{exit_ev:,.0f}  |  Exit Equity Value: {sym}{exit_equity:,.0f}  |  MOIC: {base_moic:.2f}x  |  IRR: {base_irr:.1%}" if base_moic else "Returns unavailable"], PPT_NAVY)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Returns Sensitivity", "MOIC across entry and exit multiple assumptions — each cell independently recomputed", logo_info, result["ticker"], page)
    entry_mults = [result["entry_multiple"] - 1.0, result["entry_multiple"] - 0.5, result["entry_multiple"], result["entry_multiple"] + 0.5, result["entry_multiple"] + 1.0]
    exit_mults = [result["exit_multiple"] - 1.0, result["exit_multiple"] - 0.5, result["exit_multiple"], result["exit_multiple"] + 0.5, result["exit_multiple"] + 1.0]
    matrix = []
    for em in entry_mults:
        row = []
        for xm in exit_mults:
            moic, _ = _lbo_quick_returns(result, em, xm)
            row.append(moic)
        matrix.append(row)
    _add_heatmap(s, 1.0, 1.55, 11.3, 4.8, entry_mults, exit_mults, matrix,
                 "Entry Multiple", "Exit Multiple", number_fmt=lambda v: f"{v:.2f}x",
                 row_fmt=lambda v: f"{v:.2f}x", col_fmt=lambda v: f"{v:.2f}x")

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Value Creation Drivers", "Where returns come from in this deal", logo_info, result["ticker"], page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "EBITDA Growth", ["Revenue growth", "Margin expansion", "Operating leverage over the hold period"], PPT_GREEN)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Deleveraging", ["Mandatory amortization", "100% free-cash-flow sweep", "No assumed refinancing"], PPT_NAVY)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Multiple", ["No expansion assumed here", "Exit at the same multiple as entry", "A conservative, defensible base case"], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Key Risks", "What could break the return case", logo_info, result["ticker"], page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Operating", ["Slower growth than modeled", "Margin compression", "Execution risk under leverage"], PPT_RED)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Financial", ["Higher interest rates", "Covenant pressure", "Inability to sweep cash as modeled"], PPT_RED)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Exit", ["Multiple compression at exit", "Fewer buyers / weaker M&A market", "Longer hold than planned"], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Model Assumptions", "Every driver behind this analysis", logo_info, result["ticker"], page)
    rows = [("Entry Multiple", f"{result['entry_multiple']:.2f}x"), ("Exit Multiple", f"{result['exit_multiple']:.2f}x"),
            ("Leverage", f"{result['leverage_multiple']:.2f}x"), ("Hold Period", f"{n} years"),
            ("Revenue Growth (Start/End)", f"{a['starting_growth_rate']:.1%} / {a['ending_growth_rate']:.1%}"),
            ("EBITDA Margin (Start/End)", f"{a['starting_fcf_margin']:.1%} / {a['ending_fcf_margin']:.1%}"),
            ("D&A % of Revenue", f"{a['da_pct_revenue']:.1%}"), ("CapEx % of Revenue", f"{a['capex_pct_revenue']:.1%}"),
            ("Tax Rate", f"{a['tax_rate']:.1%}"), ("Interest Rate on Debt", "8.0%"), ("Mandatory Amortization", "5.0%"), ("Cash Sweep", "100%")]
    _add_table(s, 1.0, 1.5, 11.3, 4.8, ["Assumption", "Value"], [[a2, b] for a2, b in rows], col_widths=[3.0, 2.0], font_size=9.5, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Final Returns Summary", "Base-case outcome", logo_info, result["ticker"], page)
    rows = [("Sponsor Equity (Entry)", f"{sym}{sponsor_equity:,.0f}"), ("Exit Equity Value", f"{sym}{exit_equity:,.0f}"),
            ("MOIC", f"{base_moic:.2f}x" if base_moic else "N/A"), ("IRR", f"{base_irr:.1%}" if base_irr else "N/A")]
    _add_table(s, 1.6, 1.55, 10.1, 2.6, ["Metric", "Value"], [[a2, b] for a2, b in rows], col_widths=[2.8, 2.2], font_size=10, first_col_bold=True)
    _add_banner(s, "See the standalone Excel LBO model for a fully editable live model with a real circular debt schedule and Checks tab.", 4.9, fill=PPT_NAVY, color=PPT_WHITE, height=.65)

    prs.save(output_path)
    return output_path


def build_pitch_book_public_company(result, output_path):
    prs = _new_pitch_book()
    logo_info = {"shortName": result["company_name"]}
    sym = result["currency_symbol"]; today = datetime.now().strftime("%B %d, %Y")
    ticker = result["ticker"]

    page = 1
    _add_title_slide(prs, f"{result['company_name']} ({ticker})", "Public Company Valuation", logo_info, ticker, today)
    page += 1; _add_disclaimer_slide(prs)
    page += 1; _add_toc_slide(prs, ["Executive Summary", "Company Overview", "Trading Performance", "Share Price History",
                                     "Ownership & Float", "Trading Multiples", "Analyst Consensus",
                                     "DCF Valuation", "DCF Sensitivity", "Trading Comparables", "Historical Financials",
                                     "Growth & Profitability", "Valuation Synthesis", "Investment Catalysts", "Key Risks", "Appendix"])
    page += 1

    s = _blank_slide(prs); _add_section_header(s, "Executive Summary", "Headline valuation and market position", logo_info, ticker, page)
    cards = [("Current Price", _price_fmt(result["current_price"], sym), "neutral")]
    if result["market_cap"]:
        cards.append(("Market Cap", f"{sym}{result['market_cap']:,.0f}", "neutral"))
    if result["dcf"]:
        cards.append(("DCF Value", _price_fmt(result["dcf"]["implied_share_price"], sym), "positive" if result["dcf"]["upside"] >= 0 else "negative"))
    if result["target_mean"]:
        cards.append(("Analyst Target", _price_fmt(result["target_mean"], sym), "neutral"))
    _add_kpi_cards(s, cards, top=1.35, card_width=2.75, height=1.3)
    _add_bullet_box(s, .7, 3.1, 5.85, 2.7, "Positioning",
                     [f"{result['company_name']} trades in {result['sector']} / {result['industry']} on {result['exchange']}.",
                      f"52-week range: {sym}{result['fifty_two_week_low']:,.2f} to {sym}{result['fifty_two_week_high']:,.2f}."], PPT_NAVY)
    _add_bullet_box(s, 6.8, 3.1, 5.85, 2.7, "Valuation Approach",
                     ["Synthesized from DCF, trading comparables, and analyst consensus where each is available.",
                      "Each method's range is shown independently in the Valuation Synthesis section — a triangulation, not a single number."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Company Overview", "Business profile and operating footprint", logo_info, ticker, page)
    desc = result.get("business_description") or "A business description was not available from the market-data source."
    _add_textbox(s, .75, 1.35, 7.5, 3.2, desc, size=11, color=PPT_DARK_TEXT)
    rows = [("Exchange", result["exchange"]), ("Sector", result["sector"]), ("Industry", result["industry"]),
            ("Employees", f"{result['employees']:,}" if result.get("employees") else "N/A")]
    _add_table(s, 8.55, 1.35, 3.9, 2.6, ["Metric", "Value"], [[a, str(b)] for a, b in rows], col_widths=[1.6, 2.3], font_size=9.0, first_col_bold=True)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Trading Performance", "Historical returns and risk profile", logo_info, ticker, page)
    rows2 = [("1-Year Return", f"{result['returns_1y']:+.1%}" if result["returns_1y"] is not None else "N/A"),
             ("3-Year Return", f"{result['returns_3y']:+.1%}" if result["returns_3y"] is not None else "N/A"),
             ("Annualized Volatility", f"{result['volatility_annualized']:.1%}" if result["volatility_annualized"] is not None else "N/A"),
             ("Beta", f"{result['beta']:.2f}" if result["beta"] is not None else "N/A"),
             ("52-Week High", _price_fmt(result["fifty_two_week_high"], sym)),
             ("52-Week Low", _price_fmt(result["fifty_two_week_low"], sym))]
    _add_table(s, .8, 1.5, 5.4, 3.6, ["Metric", "Value"], rows2, col_widths=[2.6, 1.6], font_size=9.5, first_col_bold=True)
    _add_bullet_box(s, 6.55, 1.5, 5.5, 3.6, "Risk Context",
                     ["Higher beta implies more sensitivity to broad market moves.", "Annualized volatility reflects realized price variability, not a forecast.",
                      "Returns are historical and not indicative of future performance."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Share Price History", "Five-year trading history", logo_info, ticker, page)
    hist = _safe_hist(result["ticker_obj"], "5y")
    if hist is not None and not hist.empty:
        cats = [str(x.date()) for x in hist.index][::max(len(hist) // 80, 1)]
        prices = hist["Close"].tolist()[::max(len(hist) // 80, 1)]
        _add_line_chart(s, .7, 1.45, 11.8, 4.9, cats, [(ticker, prices)], title="Share Price", number_format='0.00')
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Price history unavailable", ["Historical market prices could not be retrieved from the selected feed."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Ownership & Float", "Who owns the shares, and how liquid is the stock", logo_info, ticker, page)
    rows3 = [("Shares Outstanding", f"{result['shares_outstanding']:,.0f}" if result["shares_outstanding"] else "N/A"),
             ("Float (Shares)", f"{result['float_shares']:,.0f}" if result["float_shares"] else "N/A"),
             ("Insider Ownership", f"{result['insider_pct']:.1%}" if result["insider_pct"] is not None else "N/A"),
             ("Institutional Ownership", f"{result['institution_pct']:.1%}" if result["institution_pct"] is not None else "N/A"),
             ("Average Daily Volume", f"{result['avg_volume']:,.0f}" if result["avg_volume"] else "N/A")]
    _add_table(s, .8, 1.5, 6.6, 3.6, ["Metric", "Value"], rows3, col_widths=[3.0, 2.0], font_size=9.5, first_col_bold=True)
    if result["insider_pct"] is not None and result["institution_pct"] is not None:
        public_pct = max(1 - result["insider_pct"] - result["institution_pct"], 0)
        _add_column_chart(s, 7.6, 1.55, 4.7, 3.6, ["Insider", "Institutional", "Public/Other"],
                           [("Ownership", [result["insider_pct"], result["institution_pct"], public_pct])],
                           title="Ownership Split", number_format='0.0%', colors=[PPT_NAVY])

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Trading Multiples", "Current market-implied valuation multiples", logo_info, ticker, page)
    mult_rows = [("EV / Revenue", f"{result['ev_revenue']:.2f}x" if result["ev_revenue"] else "N/A"),
                 ("EV / EBITDA", f"{result['ev_ebitda']:.2f}x" if result["ev_ebitda"] else "N/A"),
                 ("Trailing P/E", f"{result['trailing_pe']:.2f}x" if result["trailing_pe"] else "N/A"),
                 ("Forward P/E", f"{result['forward_pe']:.2f}x" if result["forward_pe"] else "N/A"),
                 ("Price / Book", f"{result['price_to_book']:.2f}x" if result["price_to_book"] else "N/A"),
                 ("Dividend Yield", f"{result['dividend_yield']:.1%}" if result["dividend_yield"] else "N/A")]
    _add_table(s, .8, 1.5, 6.6, 3.9, ["Multiple", "Value"], mult_rows, col_widths=[3.0, 2.0], font_size=9.7, first_col_bold=True)
    _add_bullet_box(s, 7.6, 1.5, 4.7, 3.9, "Interpretation",
                     ["Compare against the peer set on the Trading Comparables slide for context.", "A premium or discount to peers should be explained by growth, margins, or risk — not assumed."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Analyst Consensus", "Street ratings and price targets", logo_info, ticker, page)
    if result["target_mean"]:
        _add_kpi_cards(s, [
            ("Rating", (result["recommendation_key"] or "N/A").replace("_", " ").title(), "neutral"),
            ("Analysts", str(result["num_analysts"] or "N/A"), "neutral"),
            ("Target Mean", _price_fmt(result["target_mean"], sym), "neutral"),
            ("Target Range", f"{sym}{result['target_low']:,.0f}-{sym}{result['target_high']:,.0f}" if result["target_low"] else "N/A", "neutral"),
        ], top=1.5, card_width=2.75, height=1.3)
        if result["target_low"] and result["target_high"]:
            _add_football_field(s, .8, 3.4, 11.6, 1.6, [("Analyst Range", result["target_low"], result["target_high"], PPT_TEAL)], current=result["current_price"])
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Analyst consensus unavailable", ["No analyst coverage data was returned for this company."], PPT_GOLD)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "DCF Valuation", "Explicit forecast and free-cash-flow build", logo_info, ticker, page)
    dcf = result["dcf"]
    if dcf:
        cats2 = [f"FY+{i+1}" for i in range(len(dcf.get("projected_revenue", [])))]
        _add_column_chart(s, .7, 1.45, 7.0, 4.8, cats2, [("Revenue", dcf.get("projected_revenue", [])), ("FCF", dcf.get("projected_fcf", []))], title="Revenue and FCF Forecast", number_format='#,##0')
        a = dcf["assumptions_used"]
        rows4 = [("Starting Growth", f"{a.get('starting_growth_rate', 0):.1%}"), ("Ending Growth", f"{a.get('ending_growth_rate', 0):.1%}"),
                 ("WACC", f"{a.get('wacc', 0):.2%}"), ("Terminal Growth", f"{a.get('terminal_growth_rate', 0):.2%}"),
                 ("Implied Price", _price_fmt(dcf["implied_share_price"], sym)), ("Upside / (Downside)", f"{dcf['upside']:+.1%}")]
        _add_table(s, 8.0, 1.55, 4.4, 3.9, ["Assumption", "Value"], rows4, col_widths=[2.4, 1.6], font_size=9.2, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "DCF unavailable", [result.get("dcf_error") or "Insufficient data to run a DCF for this company."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "DCF Sensitivity", "Implied share price across WACC and terminal growth", logo_info, ticker, page)
    if dcf:
        waccs = [.07, .08, .09, .10, .11]; tgs = [.01, .02, .03, .04, .05]
        matrix = []
        for w in waccs:
            row = []
            for g in tgs:
                try:
                    row.append(run_dcf(ticker, wacc=w, terminal_growth_rate=g)["implied_share_price"])
                except Exception:
                    row.append(None)
            matrix.append(row)
        _add_heatmap(s, 1.0, 1.45, 11.3, 4.8, waccs, tgs, matrix, "WACC", "Terminal Growth", number_fmt=lambda x: f"{sym}{x:,.2f}")
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Sensitivity unavailable", ["No DCF was available to sensitize."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Trading Comparables", "Peer benchmarking across valuation metrics", logo_info, ticker, page)
    comps = result["comps"]
    if comps:
        rows5 = []
        for p in [comps.get("target")] + comps.get("peers", []):
            if not p:
                continue
            rows5.append([p.get("ticker", ""), p.get("company_name", "")[:24], f"{p['ev_revenue']:.1f}x" if isinstance(p.get('ev_revenue'), (int, float)) else "NM",
                          f"{p['ev_ebitda']:.1f}x" if isinstance(p.get('ev_ebitda'), (int, float)) else "NM", f"{p['pe_ratio']:.1f}x" if isinstance(p.get('pe_ratio'), (int, float)) else "NM"])
        _add_table(s, .65, 1.45, 12.05, 4.95, ["Ticker", "Company", "EV / Rev.", "EV / EBITDA", "P / E"], rows5, col_widths=[.9, 3.3, 1.5, 1.5, 1.1], font_size=8.6, first_col_bold=True)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Trading comps unavailable", ["No peer set was returned for this company."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Peer Multiple Comparison", "EV/EBITDA across the target and its peer set", logo_info, ticker, page)
    if comps:
        peer_entries = [p for p in [comps.get("target")] + comps.get("peers", []) if p and isinstance(p.get("ev_ebitda"), (int, float))]
        if peer_entries:
            _add_column_chart(s, .8, 1.55, 11.6, 4.7, [p.get("ticker", "") for p in peer_entries],
                               [("EV / EBITDA", [p["ev_ebitda"] for p in peer_entries])], title="EV / EBITDA by Company",
                               number_format='0.0"x"', colors=[PPT_ACCENT])
            _add_banner(s, f"{ticker} shown alongside its peer set — no multiple is meaningful in isolation.", 6.35, fill=PPT_WARM_GRAY)
        else:
            _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "No comparable EV/EBITDA data", ["Insufficient peer data to chart this multiple."], PPT_GOLD)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Trading comps unavailable", ["No peer set was returned for this company."], PPT_RED)

    years, rev, gp, ni = _research_history_arrays({"company": {"ticker_obj": result["ticker_obj"]}})
    page += 1; s = _blank_slide(prs); _add_section_header(s, "Historical Financials", "Reported financial trajectory where available", logo_info, ticker, page)
    if rev:
        _add_column_chart(s, .7, 1.4, 7.6, 4.9, years, [("Revenue", rev), ("Gross Profit", gp or [0] * len(rev)), ("Net Income", ni or [0] * len(rev))], title="Historical Income Statement", number_format='#,##0')
        rows6 = [[y, f"{rev[i]:,.0f}", f"{gp[i]:,.0f}" if gp else "N/A", f"{ni[i]:,.0f}" if ni else "N/A"] for i, y in enumerate(years)]
        _add_table(s, 8.55, 1.45, 3.85, 4.75, ["Year", "Revenue", "Gross Profit", "Net Income"], rows6, col_widths=[.75, 1.1, 1.25, 1.15], font_size=7.8)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Historical data unavailable", ["The selected public-data feed did not provide sufficient historical financials for this chart."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Growth & Profitability", "Revenue growth and gross-margin trajectory", logo_info, ticker, page)
    if rev and len(rev) >= 2:
        growth = [None] + [(rev[i] / rev[i - 1] - 1) if rev[i - 1] else None for i in range(1, len(rev))]
        gm = [(gp[i] / rev[i]) if gp and rev[i] else None for i in range(len(rev))]
        _add_line_chart(s, .75, 1.45, 6.0, 4.8, years, [("Revenue Growth", [x if x is not None else 0 for x in growth])], title="Revenue Growth", number_format='0.0%')
        _add_line_chart(s, 6.95, 1.45, 5.55, 4.8, years, [("Gross Margin", [x if x is not None else 0 for x in gm])], title="Gross Margin", number_format='0.0%')
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "Margin history unavailable", ["Insufficient historical line items were returned for a robust margin series."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Valuation Synthesis", "Triangulating value across every available method", logo_info, ticker, page)
    ranges = []
    if dcf and isinstance(dcf.get("implied_share_price"), (int, float)):
        lo, hi = _dcf_valuation_range(dcf, ticker); ranges.append(("DCF", lo, hi, PPT_NAVY))
    if comps:
        cr = _comps_implied_price_range(comps, result["shares_outstanding"])
        if cr:
            ranges.append(("Trading Comps", cr[0], cr[1], PPT_ACCENT))
    if result["target_low"] and result["target_high"]:
        ranges.append(("Analyst Targets", result["target_low"], result["target_high"], PPT_TEAL))
    _add_football_field(s, .7, 1.55, 11.9, 4.7, ranges, current=result["current_price"])

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Investment Catalysts", "Potential factors that could change the valuation debate", logo_info, ticker, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Fundamental", ["Revenue re-acceleration", "Margin expansion", "Better cash conversion", "New products / geographies"], PPT_GREEN)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Market", ["Multiple re-rating", "Positive earnings revisions", "Sector rotation", "Lower rates / WACC"], PPT_ACCENT)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Corporate", ["M&A / strategic actions", "Capital returns", "Guidance changes", "New partnerships"], PPT_TEAL)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Key Risks", "Downside cases investors should underwrite", logo_info, ticker, page)
    _add_bullet_box(s, .7, 1.45, 3.8, 4.8, "Operating", ["Growth deceleration", "Competitive share loss", "Margin pressure", "Customer concentration"], PPT_RED)
    _add_bullet_box(s, 4.75, 1.45, 3.8, 4.8, "Financial", ["Weak FCF conversion", "Higher CapEx / NWC", "Leverage", "Refinancing risk"], PPT_RED)
    _add_bullet_box(s, 8.8, 1.45, 3.8, 4.8, "Valuation", ["Multiple compression", "Higher WACC", "Lower terminal growth", "Estimate misses"], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — DCF Year-by-Year Detail", "Full explicit forecast, undiscounted and discounted", logo_info, ticker, page)
    if dcf:
        proj_rev = dcf.get("projected_revenue", []); proj_fcf = dcf.get("projected_fcf", []); disc_fcf = dcf.get("discounted_fcf", [])
        rows_dcf = [[f"Year {i+1}", f"{sym}{proj_rev[i]:,.0f}", f"{sym}{proj_fcf[i]:,.0f}", f"{sym}{disc_fcf[i]:,.0f}"] for i in range(len(proj_rev))]
        _add_table(s, 1.3, 1.5, 10.4, 3.6, ["Year", "Revenue", "Free Cash Flow", "PV of FCF"], rows_dcf, col_widths=[1.2, 2.4, 2.4, 2.4], font_size=9.5, first_col_bold=True)
        _add_banner(s, f"Sum of PV of explicit FCF: {sym}{sum(disc_fcf):,.0f}  |  Enterprise Value: {sym}{dcf.get('enterprise_value', 0):,.0f}", 5.5, fill=PPT_WARM_GRAY)
    else:
        _add_bullet_box(s, .8, 1.6, 11.5, 4.5, "DCF unavailable", ["No DCF was available for this company."], PPT_RED)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Methodology", "How this package is assembled", logo_info, ticker, page)
    _add_bullet_box(s, .8, 1.4, 5.7, 4.9, "Data Sources", ["Trading data, ownership, multiples, and analyst consensus from live market data.",
                     "Historical returns and volatility computed directly from price history.", "DCF and trading comps reuse this tool's existing, independently-tested models."], PPT_NAVY)
    _add_bullet_box(s, 6.75, 1.4, 5.7, 4.9, "Valuation Synthesis", ["Each method's range is shown independently — this is a triangulation exercise, not a single answer.",
                     "See the standalone Excel workbook for a live, editable version of every figure here."], PPT_ACCENT)

    page += 1; s = _blank_slide(prs); _add_section_header(s, "Appendix — Final Valuation Summary", "Selected valuation outputs", logo_info, ticker, page)
    final_rows = []
    if dcf:
        final_rows.append(["DCF", _price_fmt(dcf["implied_share_price"], sym), f"{dcf['upside']:+.1%}"])
    if result["target_mean"]:
        upside_target = (result["target_mean"] - result["current_price"]) / result["current_price"] if result["current_price"] else None
        final_rows.append(["Street Target", _price_fmt(result["target_mean"], sym), f"{upside_target:+.1%}" if upside_target is not None else "N/A"])
    _add_table(s, .9, 1.5, 8.0, 3.6, ["Method", "Value / Share", "Upside / (Downside)"], final_rows, col_widths=[2.4, 2.2, 2.4], font_size=10, first_col_bold=True)
    _add_banner(s, "Investment conclusion should combine valuation triangulation, fundamental diligence, and current market conditions.", 5.55, fill=PPT_NAVY, color=PPT_WHITE, height=.65)

    prs.save(output_path)
    return output_path


def build_pitch_book(mode, output_path, **kwargs):
    """
    mode="research" -> kwargs: research=<dict from run_research()>
    mode="ma"        -> kwargs: deal=<dict from run_ma_model()>
    mode="ipo"        -> kwargs: ipo=<dict from run_ipo_valuation()>, currency_symbol="$"
    mode="lbo"        -> kwargs: result=<dict from run_lbo_model()>
    mode="pubco"      -> kwargs: result=<dict from run_public_company_valuation()>
    """
    if mode == "research":
        return build_pitch_book_research(kwargs["research"], output_path)
    elif mode == "pubco":
        return build_pitch_book_public_company(kwargs["result"], output_path)
    elif mode == "ma":
        return build_pitch_book_ma(kwargs["deal"], output_path)
    elif mode == "ipo":
        return build_pitch_book_ipo(kwargs["ipo"], kwargs.get("currency_symbol", "$"), output_path)
    elif mode == "lbo":
        return build_pitch_book_lbo(kwargs["result"], output_path)
    else:
        raise ValueError(f"Unknown pitch book mode: '{mode}'. Use 'research', 'ma', 'ipo', 'lbo', or 'pubco'.")

# ============================================
# WATCHLIST
# Persists to a simple JSON file in the current folder — no database, so it
# survives between runs but is specific to wherever you run the tool from.
# ============================================

WATCHLIST_FILE = os.path.join(APP_DATA_DIR, "watchlist.json")


def load_watchlist():
    if not os.path.exists(WATCHLIST_FILE):
        return []
    try:
        with open(WATCHLIST_FILE, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []


def save_watchlist(tickers):
    with open(WATCHLIST_FILE, "w") as f:
        json.dump(tickers, f, indent=2)


def add_to_watchlist(ticker_symbol):
    """Verifies the ticker is real (via resolve_company) before saving it —
    a typo shouldn't silently sit in the watchlist forever."""
    ticker_symbol = ticker_symbol.upper()
    resolve_company(ticker_symbol)  # raises ValueError if invalid, caller should catch

    watchlist = load_watchlist()
    if ticker_symbol in watchlist:
        print(f"{ticker_symbol} is already on your watchlist.")
    else:
        watchlist.append(ticker_symbol)
        save_watchlist(watchlist)
        print(f"Added {ticker_symbol} to your watchlist ({len(watchlist)} tickers total).")
    return watchlist


def remove_from_watchlist(ticker_symbol):
    ticker_symbol = ticker_symbol.upper()
    watchlist = load_watchlist()
    if ticker_symbol not in watchlist:
        print(f"{ticker_symbol} is not on your watchlist.")
    else:
        watchlist.remove(ticker_symbol)
        save_watchlist(watchlist)
        print(f"Removed {ticker_symbol} from your watchlist ({len(watchlist)} tickers remaining).")
    return watchlist


def print_watchlist_overview():
    """
    A fast scan across the whole watchlist — current price, today's move,
    day's trading range, market cap — rather than running a full DCF for
    every ticker (which would be much slower). Use 'research TICKER' on
    any individual name for the deeper view.
    """
    watchlist = load_watchlist()
    if not watchlist:
        print("Your watchlist is empty. Add one with: watchlist add TICKER")
        return

    divider = "─" * 96
    print(divider)
    print(f"{'Ticker':<8}{'Company':<24}{'Price':>12}{'Day %':>9}{'Day Low':>12}{'Day High':>12}{'Market Cap':>18}")
    print(divider)
    for ticker_symbol in watchlist:
        try:
            company = resolve_company(ticker_symbol)
            info = company["info"]
            price_divisor = company["price_divisor"]
            symbol = get_currency_symbol(company["quote_currency"])

            price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / price_divisor
            prev_close = (info.get("regularMarketPreviousClose") or price) / price_divisor
            day_change_pct = ((price - prev_close) / prev_close) if prev_close else 0

            day_low = (info.get("dayLow") or info.get("regularMarketDayLow") or 0) / price_divisor
            day_high = (info.get("dayHigh") or info.get("regularMarketDayHigh") or 0) / price_divisor
            day_low_str = f"{symbol}{day_low:,.2f}" if day_low else "N/A"
            day_high_str = f"{symbol}{day_high:,.2f}" if day_high else "N/A"

            market_cap = info.get("marketCap")
            market_cap_str = f"{symbol}{market_cap:,.0f}" if market_cap else "N/A"
            price_str = f"{symbol}{price:,.2f}"

            print(f"{ticker_symbol:<8}{company['company_name'][:22]:<24}{price_str:>12}{day_change_pct:>8.1%}"
                  f"{day_low_str:>12}{day_high_str:>12}{market_cap_str:>18}")
        except ValueError as e:
            print(f"{ticker_symbol:<8}Error resolving this ticker — {e}")
    print(divider)


# ============================================
# GLOBAL MARKET NEWS
# Reuses get_recent_news() (already built for research mode) but points it
# at major market INDEXES rather than individual companies — Yahoo Finance
# tags broad market-moving headlines to these symbols, giving a rough
# cross-market news scan without needing a separate news API.
# ============================================

MAJOR_MARKET_INDEXES = {
    "United States (S&P 500)": "^GSPC",
    "United Kingdom (FTSE 100)": "^FTSE",
    "Germany (DAX)": "^GDAXI",
    "Japan (Nikkei 225)": "^N225",
    "Hong Kong (Hang Seng)": "^HSI",
    "Australia (ASX 200)": "^AXJO",
    "Canada (TSX)": "^GSPTSE",
    "India (Nifty 50)": "^NSEI",
}


def get_global_market_news(max_per_market=3):
    results = {}
    for market_name, index_symbol in MAJOR_MARKET_INDEXES.items():
        try:
            results[market_name] = get_recent_news(index_symbol, max_items=max_per_market)
        except Exception:
            results[market_name] = []
    return results


def print_global_market_news(news_by_market):
    divider = "─" * 60
    print(divider)
    print("Global Market News")
    print(divider)
    for market_name, items in news_by_market.items():
        print(f"\n{market_name}")
        if not items:
            print("  No headlines found for this market right now.")
        for item in items:
            print(f"  • {item['title']}  ({item['publisher']})")
    print(f"\n{divider}")


# ============================================
# DAILY WATCHLIST SUMMARY + MACOS NOTIFICATION
# Designed to be triggered by a scheduled macOS LaunchAgent (see setup
# instructions), not run interactively — so it writes a dated log file
# (reviewable even if the notification is missed) in addition to firing a
# native notification banner with the headline numbers.
# ============================================
# PERSONAL INVESTMENT TRACKER (Australian CGT)
# Tracks YOUR actual buy/sell transactions and computes realized/unrealized
# capital gains using FIFO lot matching and the Australian 50% CGT discount
# for parcels held over 12 months. Deliberately NOT exposed through the
# natural-language parser — for tax-relevant numbers, explicit commands
# beat guessing which number was meant to be quantity vs. price vs. fees.
#
# IMPORTANT: this is a planning/estimation tool, not tax advice. It doesn't
# model prior-year loss carry-forwards, wash sale-style rules, corporate
# actions (splits/mergers adjusting cost base), or anything ATO-specific
# beyond the core FIFO + 12-month-discount mechanics. Talk to a registered
# tax agent before lodging a return.
# ============================================

PORTFOLIO_FILE = os.path.join(APP_DATA_DIR, "portfolio.json")
CGT_DISCOUNT_ELIGIBILITY_DAYS = 365  # "more than 12 months" — a reasonable
                                       # approximation; the ATO's actual rule
                                       # is date-exact (day after the 1-year
                                       # anniversary), which this doesn't
                                       # account for in leap years.

TAX_SETTINGS_FILE = os.path.join(APP_DATA_DIR, "tax_settings.json")


def load_tax_settings():
    if not os.path.exists(TAX_SETTINGS_FILE):
        return {}
    try:
        with open(TAX_SETTINGS_FILE, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def save_tax_settings(settings):
    with open(TAX_SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=2)


def set_marginal_tax_rate(rate):
    """Persists your marginal tax rate so future tax reports apply it
    automatically — update it any time your income/bracket changes."""
    if rate < 0 or rate > 100:
        raise ValueError("Marginal tax rate must be between 0 and 100.")
    settings = load_tax_settings()
    settings["marginal_tax_rate"] = rate
    save_tax_settings(settings)
    print(f"Saved marginal tax rate: {rate:.1f}% — this will be used automatically in tax reports "
          f"until you update it again with: portfolio settaxrate RATE")


def get_marginal_tax_rate():
    return load_tax_settings().get("marginal_tax_rate")


def load_portfolio():
    if not os.path.exists(PORTFOLIO_FILE):
        return []
    try:
        with open(PORTFOLIO_FILE, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []


def save_portfolio(transactions):
    with open(PORTFOLIO_FILE, "w") as f:
        json.dump(transactions, f, indent=2)


# --- nabtrade's online brokerage schedule (identical tiers for domestic and
# international shares; "over the phone" pricing is deliberately not modeled) ---
DOMESTIC_BROKERAGE_TIERS = [
    (1000, 9.95),
    (5000, 14.95),
    (20000, 19.95),
]
BROKERAGE_PCT_ABOVE_20K = 0.0011  # 0.11% of trade value, for trades over $20,000

# International (non-AUD) trades also incur a currency conversion spread.
# nabtrade states a range of 0.50%-0.80%; 0.80% (the conservative/higher
# end) is used so the estimate doesn't understate your true cost base.
INTERNATIONAL_FX_SPREAD_PCT = 0.008


def calculate_brokerage(trade_value_aud):
    """trade_value_aud MUST already be in AUD — nabtrade's tiers are AUD
    thresholds, regardless of what currency the underlying stock trades in."""
    for threshold, flat_fee in DOMESTIC_BROKERAGE_TIERS:
        if trade_value_aud <= threshold:
            return flat_fee
    return trade_value_aud * BROKERAGE_PCT_ABOVE_20K


def get_historical_fx_rate(date_str, pair="AUDUSD=X"):
    """
    Returns the FX rate for the given pair on the given date (e.g.
    "AUDUSD=X" returns USD per 1 AUD) — used both for correctly converting
    nabtrade's AUD-denominated brokerage fee, and for the ATO-correct CGT
    approach of converting each leg of a foreign trade using the exchange
    rate on ITS OWN transaction date. Looks back up to a week if the exact
    date has no data (weekends/holidays), using the most recent prior close.
    """
    fx_ticker = get_ticker(pair)
    target_date = datetime.strptime(date_str, "%Y-%m-%d")
    start = (target_date - timedelta(days=7)).strftime("%Y-%m-%d")
    end = (target_date + timedelta(days=1)).strftime("%Y-%m-%d")
    history = fx_ticker.history(start=start, end=end)
    if history.empty:
        raise ValueError(f"Could not fetch an exchange rate for {pair} near {date_str}.")
    return float(history["Close"].iloc[-1])


def get_ticker_currency(ticker_symbol, currency_cache):
    """Cached currency lookup — avoids repeated network calls for the same
    ticker when processing many transactions in one report."""
    if ticker_symbol not in currency_cache:
        company = resolve_company(ticker_symbol)
        currency_cache[ticker_symbol] = company["quote_currency"]
    return currency_cache[ticker_symbol]


def convert_to_aud(amount, currency, date_str, fx_cache):
    """
    Converts `amount` (in `currency`) to AUD using the actual exchange
    rate ON date_str — the ATO-correct method for foreign shares: convert
    at the rate on the specific transaction date, not today's rate.
    Returns (aud_amount, fx_rate_used, conversion_succeeded). On failure,
    aud_amount falls back to the raw (unconverted) figure so a report
    doesn't crash outright — conversion_succeeded=False flags this clearly
    to the caller instead of silently mislabeling a native amount as AUD.
    fx_cache avoids re-fetching the same (currency, date) pair repeatedly.
    """
    if currency == "AUD":
        return amount, 1.0, True

    cache_key = (currency, date_str)
    if cache_key not in fx_cache:
        pair = f"AUD{currency}=X"
        try:
            fx_cache[cache_key] = get_historical_fx_rate(date_str, pair=pair)
        except Exception:
            fx_cache[cache_key] = None

    rate = fx_cache[cache_key]
    if rate is None:
        return amount, None, False
    return amount / rate, rate, True


def calculate_trade_fees(is_domestic, trade_value, native_currency=None, trade_date=None):
    """
    Returns (total_fees, brokerage, fx_cost) so callers can show the
    breakdown, both expressed in the trade's own currency (native_currency)
    so they combine correctly with a price_per_share you're keeping in
    that same currency.

    For domestic (AUD) trades, trade_value already IS the AUD figure the
    tiers expect, so no conversion is needed.

    For international trades, nabtrade still bills brokerage in AUD
    regardless of the stock's currency — so this converts your (e.g. USD)
    trade value to AUD using the actual rate on the trade date, picks the
    correct AUD tier, then converts that AUD fee back into your currency.
    Only AUD<->USD is handled directly; other currencies fall back to
    applying the tiers to the raw (unconverted) value, which is less
    accurate — a printed note flags this when it happens.
    """
    if is_domestic:
        brokerage = calculate_brokerage(trade_value)
        fx_cost = trade_value * 0  # explicit for clarity: no FX cost on a domestic trade
        return brokerage, brokerage, fx_cost

    fx_cost = trade_value * INTERNATIONAL_FX_SPREAD_PCT

    if native_currency == "USD" and trade_date:
        try:
            fx_rate = get_historical_fx_rate(trade_date)  # USD per 1 AUD
            trade_value_aud = trade_value / fx_rate
            brokerage_aud = calculate_brokerage(trade_value_aud)
            brokerage = brokerage_aud * fx_rate  # convert the AUD fee back into USD
            return brokerage + fx_cost, brokerage, fx_cost
        except Exception:
            print("  (Couldn't fetch a historical AUD/USD rate — brokerage tier estimated "
                  "from the raw USD trade value instead, which may be slightly off.)")

    # Fallback: no conversion available (currency isn't USD, or the FX
    # lookup failed) — apply tiers to the raw value as a rough estimate.
    brokerage = calculate_brokerage(trade_value)
    return brokerage + fx_cost, brokerage, fx_cost


def add_transaction(ticker_symbol, txn_type, quantity, price_per_share, fees=None, date=None):
    """
    Records a buy or sell. Verifies the ticker resolves to something real
    before saving — same philosophy as the watchlist, a typo shouldn't sit
    in your tax records. date defaults to today if not given, format
    YYYY-MM-DD.

    fees is calculated AUTOMATICALLY from nabtrade's brokerage schedule
    (plus an FX spread estimate for international shares) unless you pass
    an explicit value yourself — e.g. if you actually traded over the
    phone, or got a different rate.
    """
    ticker_symbol = ticker_symbol.upper()
    txn_type = txn_type.lower()
    if txn_type not in ("buy", "sell"):
        raise ValueError(f"Transaction type must be 'buy' or 'sell', got '{txn_type}'.")
    if quantity <= 0:
        raise ValueError("Quantity must be positive.")
    if price_per_share < 0:
        raise ValueError("Price per share cannot be negative.")

    company = resolve_company(ticker_symbol)  # raises ValueError if not a real ticker

    if date is None:
        date = datetime.now().strftime("%Y-%m-%d")
    else:
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"Date must be in YYYY-MM-DD format, got '{date}'.")

    if txn_type == "sell":
        holdings = compute_current_holdings(load_portfolio())
        available = holdings.get(ticker_symbol, {}).get("quantity", 0)
        if quantity > available:
            raise ValueError(
                f"Cannot sell {quantity} shares of {ticker_symbol} — only {available} currently held. "
                f"Check your transaction history with: portfolio holdings"
            )

    trade_value = quantity * price_per_share
    is_domestic = company["quote_currency"] == "AUD"

    if fees is None:
        fees, brokerage, fx_cost = calculate_trade_fees(
            is_domestic, trade_value, native_currency=company["quote_currency"], trade_date=date
        )
        if fx_cost:
            fee_breakdown = f"brokerage ${brokerage:,.2f} + est. FX spread ${fx_cost:,.2f}"
        else:
            fee_breakdown = f"brokerage ${brokerage:,.2f}"
    else:
        fee_breakdown = "manually specified"

    transactions = load_portfolio()
    transactions.append({
        "date": date, "ticker": ticker_symbol, "type": txn_type,
        "quantity": quantity, "price_per_share": price_per_share, "fees": fees,
    })
    save_portfolio(transactions)
    print(f"Recorded: {txn_type.upper()} {quantity} {ticker_symbol} @ {format_price(price_per_share)} "
          f"(+ ${fees:,.2f} fees — {fee_breakdown}) on {date}")
    return transactions


def financial_year_for_date(date_str):
    """
    Australian financial year runs July 1 - June 30. A sale on 2025-08-15
    falls in FY 2025-26; a sale on 2025-03-15 falls in FY 2024-25.
    Returns a label like '2025-26'.
    """
    d = datetime.strptime(date_str, "%Y-%m-%d")
    if d.month >= 7:
        start_year = d.year
    else:
        start_year = d.year - 1
    return f"{start_year}-{str(start_year + 1)[-2:]}"


def compute_realized_gains(transactions):
    """
    FIFO lot matching: each sell consumes the OLDEST remaining buy lots
    first for that ticker. Fees on both the original buy and the sell are
    incorporated into cost base / proceeds respectively, prorated across
    however many lots a single sell consumes.

    For non-AUD tickers, ALSO computes the ATO-correct AUD-converted
    figures: cost basis is converted using the exchange rate on the BUY
    date, proceeds using the rate on the SELL date — each leg at its own
    date, not today's rate. Native-currency figures are kept alongside
    for reference/display; the AUD figures are what actually feed the tax
    summary math.

    Returns a list of matched-lot records, one per (sell, lot-portion)
    pair — a single sell can produce multiple records if it draws from
    more than one buy lot.
    """
    # Sort chronologically so FIFO matching processes buys before the sells
    # that consume them, regardless of the order they were entered.
    sorted_txns = sorted(transactions, key=lambda t: t["date"])

    open_lots = {}  # ticker -> list of {quantity, date, cost_per_share}
    realized = []
    currency_cache = {}
    fx_cache = {}

    for txn in sorted_txns:
        ticker = txn["ticker"]
        if txn["type"] == "buy":
            cost_per_share = txn["price_per_share"] + (txn["fees"] / txn["quantity"] if txn["quantity"] else 0)
            open_lots.setdefault(ticker, []).append({
                "quantity": txn["quantity"], "date": txn["date"], "cost_per_share": cost_per_share,
            })
        elif txn["type"] == "sell":
            remaining_to_sell = txn["quantity"]
            sell_fee_per_share = txn["fees"] / txn["quantity"] if txn["quantity"] else 0
            lots = open_lots.get(ticker, [])
            currency = get_ticker_currency(ticker, currency_cache)

            while remaining_to_sell > 1e-9 and lots:
                lot = lots[0]
                qty_from_this_lot = min(lot["quantity"], remaining_to_sell)

                proceeds_per_share = txn["price_per_share"] - sell_fee_per_share
                cost_basis = qty_from_this_lot * lot["cost_per_share"]
                proceeds = qty_from_this_lot * proceeds_per_share
                gain_loss = proceeds - cost_basis

                held_days = (datetime.strptime(txn["date"], "%Y-%m-%d")
                             - datetime.strptime(lot["date"], "%Y-%m-%d")).days
                discount_eligible = held_days > CGT_DISCOUNT_ELIGIBILITY_DAYS

                # --- AUD conversion: cost basis at the BUY date's rate,
                # proceeds at the SELL date's rate — the correct ATO method,
                # NOT converting the final native gain at a single rate. ---
                cost_basis_aud, buy_fx_rate, buy_fx_ok = convert_to_aud(cost_basis, currency, lot["date"], fx_cache)
                proceeds_aud, sell_fx_rate, sell_fx_ok = convert_to_aud(proceeds, currency, txn["date"], fx_cache)
                fx_conversion_ok = buy_fx_ok and sell_fx_ok
                gain_loss_aud = proceeds_aud - cost_basis_aud

                # Taxable amount for THIS trade considered alone (in AUD —
                # the figure that actually matters for tax): the 50%
                # discount applies to eligible gains, losses are never
                # discounted. This is per-trade — the actual final taxable
                # figure for the year (in summarize_financial_year) can
                # differ, since losses from OTHER trades in the same year
                # get offset against this one first.
                taxable_gain_aud = gain_loss_aud * 0.5 if (discount_eligible and gain_loss_aud > 0) else gain_loss_aud

                realized.append({
                    "ticker": ticker, "currency": currency, "buy_date": lot["date"], "sell_date": txn["date"],
                    "quantity": qty_from_this_lot,
                    "cost_basis": cost_basis, "proceeds": proceeds, "gain_loss": gain_loss,  # native currency
                    "cost_basis_aud": cost_basis_aud, "proceeds_aud": proceeds_aud,
                    "gain_loss_aud": gain_loss_aud, "taxable_gain_aud": taxable_gain_aud,
                    "fx_conversion_ok": fx_conversion_ok,
                    "held_days": held_days, "discount_eligible": discount_eligible,
                    "financial_year": financial_year_for_date(txn["date"]),
                })

                lot["quantity"] -= qty_from_this_lot
                remaining_to_sell -= qty_from_this_lot
                if lot["quantity"] <= 1e-9:
                    lots.pop(0)

            if remaining_to_sell > 1e-9:
                # Shouldn't happen if add_transaction()'s pre-check ran, but
                # don't silently misreport tax figures if it somehow does.
                raise ValueError(
                    f"Data inconsistency: tried to sell {txn['quantity']} {ticker} on {txn['date']} "
                    f"but only enough lots for {txn['quantity'] - remaining_to_sell} were found."
                )

    return realized


def summarize_financial_year(realized_gains, financial_year):
    """
    Applies the ATO's optimal ordering: net capital losses against
    non-discount-eligible (short-term) gains FIRST, since that preserves
    the 50% discount on long-term gains — then apply the discount to
    whatever long-term gain remains. This is a legitimate, commonly-used
    approach, not the only possible one; a tax agent may order things
    differently based on your full position (e.g. prior-year loss
    carry-forwards, which this tool doesn't track).

    Runs on the AUD-converted figures (gain_loss_aud), not the native
    currency amounts — this IS the actual taxable calculation.
    """
    year_events = [r for r in realized_gains if r["financial_year"] == financial_year]
    any_conversion_failed = any(not e["fx_conversion_ok"] for e in year_events)

    short_term_gain = sum(r["gain_loss_aud"] for r in year_events if not r["discount_eligible"] and r["gain_loss_aud"] > 0)
    long_term_gain = sum(r["gain_loss_aud"] for r in year_events if r["discount_eligible"] and r["gain_loss_aud"] > 0)
    total_losses = sum(-r["gain_loss_aud"] for r in year_events if r["gain_loss_aud"] < 0)

    losses_remaining = total_losses
    short_term_after_losses = max(short_term_gain - losses_remaining, 0)
    losses_remaining = max(losses_remaining - short_term_gain, 0)
    long_term_after_losses = max(long_term_gain - losses_remaining, 0)
    losses_remaining = max(losses_remaining - long_term_gain, 0)

    discounted_long_term = long_term_after_losses * 0.5
    net_capital_gain = short_term_after_losses + discounted_long_term

    return {
        "financial_year": financial_year,
        "events": year_events,
        "any_conversion_failed": any_conversion_failed,
        "short_term_gain_gross": short_term_gain,
        "long_term_gain_gross": long_term_gain,
        "total_losses": total_losses,
        "unused_losses": losses_remaining,  # would carry forward to next FY — not automatically applied
        "short_term_gain_after_losses": short_term_after_losses,
        "long_term_gain_after_losses": long_term_after_losses,
        "discounted_long_term_gain": discounted_long_term,
        "net_capital_gain": net_capital_gain,
    }


def compute_current_holdings(transactions):
    """
    What you actually still own right now, computed by replaying every
    buy/sell in order — NOT a running balance stored separately, so it's
    always internally consistent with the transaction history.
    """
    sorted_txns = sorted(transactions, key=lambda t: t["date"])
    open_lots = {}

    for txn in sorted_txns:
        ticker = txn["ticker"]
        if txn["type"] == "buy":
            cost_per_share = txn["price_per_share"] + (txn["fees"] / txn["quantity"] if txn["quantity"] else 0)
            open_lots.setdefault(ticker, []).append({
                "quantity": txn["quantity"], "date": txn["date"], "cost_per_share": cost_per_share,
            })
        elif txn["type"] == "sell":
            remaining = txn["quantity"]
            lots = open_lots.get(ticker, [])
            while remaining > 1e-9 and lots:
                lot = lots[0]
                take = min(lot["quantity"], remaining)
                lot["quantity"] -= take
                remaining -= take
                if lot["quantity"] <= 1e-9:
                    lots.pop(0)

    holdings = {}
    for ticker, lots in open_lots.items():
        total_qty = sum(lot["quantity"] for lot in lots if lot["quantity"] > 1e-9)
        if total_qty <= 1e-9:
            continue
        total_cost = sum(lot["quantity"] * lot["cost_per_share"] for lot in lots if lot["quantity"] > 1e-9)
        holdings[ticker] = {
            "quantity": total_qty,
            "avg_cost_per_share": total_cost / total_qty,
            "total_cost_base": total_cost,
            "lots": [lot for lot in lots if lot["quantity"] > 1e-9],
        }
    return holdings


def print_holdings_overview():
    transactions = load_portfolio()
    holdings = compute_current_holdings(transactions)

    if not holdings:
        print("No current holdings — add a transaction with: portfolio buy")
        return

    divider = "─" * 96
    print(divider)
    print(f"{'Ticker':<8}{'Qty':>10}{'Avg Cost':>14}{'Current':>14}{'Mkt Value':>16}{'Unrealized':>18}{'%':>9}")
    print(divider)

    total_cost_base = 0
    total_market_value = 0

    for ticker, position in holdings.items():
        try:
            company = resolve_company(ticker)
            info = company["info"]
            symbol = get_currency_symbol(company["quote_currency"])
            current_price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / company["price_divisor"]
            market_value = current_price * position["quantity"]
            cost_base = position["total_cost_base"]
            unrealized = market_value - cost_base
            unrealized_pct = (unrealized / cost_base) if cost_base else 0

            total_cost_base += cost_base
            total_market_value += market_value

            avg_cost_str = f"{symbol}{position['avg_cost_per_share']:,.2f}"
            current_str = f"{symbol}{current_price:,.2f}"
            market_value_str = f"{symbol}{market_value:,.2f}"
            unrealized_str = f"{symbol}{unrealized:,.2f}"

            print(f"{ticker:<8}{position['quantity']:>10,.2f}{avg_cost_str:>14}{current_str:>14}"
                  f"{market_value_str:>16}{unrealized_str:>18}{unrealized_pct:>9.1%}")
        except ValueError as e:
            print(f"{ticker:<8}Error resolving this ticker — {e}")

    print(divider)
    total_unrealized = total_market_value - total_cost_base
    total_unrealized_pct = (total_unrealized / total_cost_base) if total_cost_base else 0
    print(f"{'TOTAL':<8}{'':>10}{'':>14}{'':>14}"
          f"{'$' + format(total_market_value, ',.2f'):>16}"
          f"{'$' + format(total_unrealized, ',.2f'):>18}{total_unrealized_pct:>9.1%}")
    print(divider)


def print_tax_report(financial_year=None, marginal_tax_rate=None):
    transactions = load_portfolio()
    realized = compute_realized_gains(transactions)

    if not realized:
        print("No sell transactions recorded yet — nothing to report.")
        return

    if financial_year is None:
        financial_year = max(r["financial_year"] for r in realized)  # most recent FY with activity

    summary = summarize_financial_year(realized, financial_year)

    divider = "─" * 106
    print(divider)
    print(f"Capital Gains Tax Estimate — FY {financial_year}  "
          f"(1 Jul {financial_year[:4]} – 30 Jun 20{financial_year[-2:]})")
    print(divider)
    print("Planning estimate only — not tax advice. Confirm with a registered tax agent before lodging,")
    print("especially around prior-year loss carry-forwards, which this tool doesn't track.")
    print("Foreign trades are converted to AUD using the exchange rate on each transaction's own date")
    print("(buy date for cost, sell date for proceeds) — the ATO-correct method, not today's rate.")
    print(divider)

    if not summary["events"]:
        print(f"No disposals recorded in FY {financial_year}.")
        print(divider)
        return

    if summary["any_conversion_failed"]:
        print("⚠ WARNING: at least one trade's exchange rate couldn't be fetched — its AUD figures below")
        print("  are a fallback estimate, not a real conversion. Check that trade's currency manually.")
        print(divider)

    print(f"\n{'Ticker':<7}{'Ccy':<5}{'Bought':<12}{'Sold':<12}{'Qty':>8}"
          f"{'Gain (native)':>15}{'Gain (AUD)':>13}{'Held':>6}{'Disc?':>6}{'Taxable AUD*':>14}")
    for e in summary["events"]:
        native_gain_str = f"{e['gain_loss']:,.2f} {e['currency']}"
        conversion_flag = "" if e["fx_conversion_ok"] else "⚠"
        print(f"{e['ticker']:<7}{e['currency']:<5}{e['buy_date']:<12}{e['sell_date']:<12}{e['quantity']:>8,.2f}"
              f"{native_gain_str:>15}"
              f"{'$' + format(e['gain_loss_aud'], ',.2f') + conversion_flag:>13}"
              f"{e['held_days']:>6}{'Yes' if e['discount_eligible'] else 'No':>6}"
              f"{'$' + format(e['taxable_gain_aud'], ',.2f'):>14}")
    print("*Taxable amount (AUD) for that trade considered alone — see final total below for the actual")
    print(" figure after offsetting losses across all trades in the year. AUD figures are what the ATO")
    print(" cares about; native-currency figures are shown for your own reference only.")

    print(f"\n{divider}")
    print(f"Short-term gains (held ≤12mo, no discount):        ${summary['short_term_gain_gross']:,.2f} AUD")
    print(f"Long-term gains (held >12mo, discount-eligible):    ${summary['long_term_gain_gross']:,.2f} AUD")
    print(f"Total capital losses:                               ${summary['total_losses']:,.2f} AUD")
    print(f"  → applied against short-term gains first, then long-term")
    if summary["unused_losses"] > 0:
        print(f"  → ${summary['unused_losses']:,.2f} in losses unused this year "
              f"(would carry forward — track manually)")
    print(f"Long-term gain after losses, before discount:       ${summary['long_term_gain_after_losses']:,.2f} AUD")
    print(f"Long-term gain after 50% CGT discount:              ${summary['discounted_long_term_gain']:,.2f} AUD")
    print(divider)
    print(f"ESTIMATED NET CAPITAL GAIN for FY {financial_year}:       ${summary['net_capital_gain']:,.2f} AUD")

    if marginal_tax_rate is not None:
        tax_owed = summary["net_capital_gain"] * (marginal_tax_rate / 100)
        print(divider)
        print(f"At your marginal rate of {marginal_tax_rate:.1f}%:")
        print(f"ESTIMATED TAX OWED ON THIS GAIN:                       ${tax_owed:,.2f} AUD")
        print("(Include the Medicare Levy, ~2%, in this rate if it applies to you — not added automatically.)")

    print(divider)


PORTFOLIO_TAX_REPORT_FILE = os.path.join(APP_DATA_DIR, "Portfolio_CGT_Report.xlsx")


def export_tax_report_to_excel(financial_year=None, marginal_tax_rate=None):
    transactions = load_portfolio()
    realized = compute_realized_gains(transactions)
    if not realized:
        raise ValueError("No sell transactions recorded yet — nothing to export.")

    if financial_year is None:
        financial_year = max(r["financial_year"] for r in realized)
    summary = summarize_financial_year(realized, financial_year)

    # Load the existing consolidated workbook if one exists (preserving every
    # other financial year's tabs), otherwise start a fresh one.
    if os.path.exists(PORTFOLIO_TAX_REPORT_FILE):
        wb = load_workbook(PORTFOLIO_TAX_REPORT_FILE)
    else:
        wb = Workbook()
        wb.remove(wb.active)  # drop openpyxl's default blank "Sheet" — we only want named FY tabs

    trades_sheet_name = f"{financial_year} Trades"
    summary_sheet_name = f"{financial_year} Summary"

    # If this financial year was exported before, remove its old tabs first —
    # re-running the report UPDATES that year in place, rather than piling
    # up duplicate tabs every time you run it again.
    for existing_name in (trades_sheet_name, summary_sheet_name):
        if existing_name in wb.sheetnames:
            del wb[existing_name]

    gain_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   # light green
    loss_fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")   # light red
    price_fmt = '$#,##0.00'

    # ============================================
    # TAB 1: TRADES — every matched lot, one row each
    # ============================================
    trades_tab = wb.create_sheet(trades_sheet_name)
    trades_tab.sheet_view.showGridLines = False


    trades_tab["A1"] = f"Capital Gains Tax Estimate — FY {financial_year}"
    trades_tab["A1"].font = TITLE_FONT
    trades_tab.merge_cells("A1:O1")
    trades_tab["A2"] = ("Planning estimate only — not tax advice. Confirm with a registered tax agent before lodging. "
                         "Foreign trades converted to AUD using the exchange rate on each transaction's own date.")
    trades_tab["A2"].font = Font(italic=True, size=9, color="9C0000")
    trades_tab.merge_cells("A2:O2")

    headers = ["Ticker", "Ccy", "Bought", "Sold", "Quantity",
               "Cost (native)", "Proceeds (native)", "Gain/Loss (native)",
               "Cost (AUD)", "Proceeds (AUD)", "Gain/Loss (AUD)",
               "Held (days)", "12mo+ Discount", "Taxable AUD (this trade)", "Tax to Withhold (AUD)"]
    for col, header in enumerate(headers, start=1):
        cell = trades_tab.cell(row=4, column=col, value=header)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    withhold_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # light amber
    total_withhold = 0.0

    for i, e in enumerate(summary["events"], start=5):
        row_fill = gain_fill if e["gain_loss_aud"] >= 0 else loss_fill
        # Only profitable trades need anything withheld — a loss owes no
        # tax by itself. Uses THIS trade's taxable amount alone (not
        # offset against other trades' losses), which is the safer,
        # conservative figure to set aside as you go, since you can't
        # know in advance whether a later loss will reduce your final bill.
        withhold_amount = None
        if marginal_tax_rate is not None:
            withhold_amount = max(e["taxable_gain_aud"], 0) * (marginal_tax_rate / 100)
            total_withhold += withhold_amount

        values = [e["ticker"], e["currency"], e["buy_date"], e["sell_date"], e["quantity"],
                  e["cost_basis"], e["proceeds"], e["gain_loss"],
                  e["cost_basis_aud"], e["proceeds_aud"], e["gain_loss_aud"],
                  e["held_days"], "Yes" if e["discount_eligible"] else "No", e["taxable_gain_aud"],
                  withhold_amount if withhold_amount is not None else "Set a rate: portfolio settaxrate"]
        for col, val in enumerate(values, start=1):
            cell = trades_tab.cell(row=i, column=col, value=val)
            cell.fill = withhold_fill if col == 15 else row_fill
            if col in (6, 7, 8, 9, 10, 11, 14) or (col == 15 and withhold_amount is not None):
                cell.number_format = price_fmt
            cell.alignment = Alignment(horizontal="center" if col not in (1, 2, 3, 4) else "left")

    total_row = 5 + len(summary["events"])
    first_event_row, last_event_row = 5, total_row - 1

    trades_tab.cell(row=total_row, column=9,
                     value="TOTAL PROFIT (GAIN/LOSS) FOR THE YEAR (AUD):").font = Font(bold=True, size=11, color="1F4E78")
    total_profit_cell = trades_tab.cell(row=total_row, column=11, value=f"=SUM(K{first_event_row}:K{last_event_row})")
    total_profit_cell.number_format = price_fmt
    total_profit_cell.font = Font(bold=True, size=11, color="1F4E78")
    total_profit_cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

    if marginal_tax_rate is not None:
        trades_tab.cell(row=total_row, column=13, value="TOTAL TO WITHHOLD THIS YEAR:").font = Font(bold=True, size=11)
        total_cell = trades_tab.cell(row=total_row, column=15, value=total_withhold)
        total_cell.number_format = price_fmt
        total_cell.font = Font(bold=True, size=11, color="1F4E78")
        total_cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

    note_row = total_row + 2
    trades_tab.cell(row=note_row, column=1,
                     value="\"Native\" = the currency the stock actually trades in. \"AUD\" columns are the "
                           "actual taxable figures. \"Taxable AUD (this trade)\" and \"Tax to Withhold\" are "
                           "that trade considered alone (a conservative, safe-to-set-aside figure) — see the "
                           "Summary tab for the real final total after offsetting losses across all trades.")
    trades_tab.cell(row=note_row, column=1).font = Font(italic=True, size=9, color="666666")
    trades_tab.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=15)

    if any(not e["fx_conversion_ok"] for e in summary["events"]):
        warn_row = note_row + 1
        trades_tab.cell(row=warn_row, column=1,
                         value="⚠ At least one trade's exchange rate could not be fetched — its AUD figures "
                               "are an unconverted fallback, not a real conversion. Verify manually.")
        trades_tab.cell(row=warn_row, column=1).font = Font(italic=True, size=9, color="9C0000", bold=True)
        trades_tab.merge_cells(start_row=warn_row, start_column=1, end_row=warn_row, end_column=15)

    col_widths = [8, 6, 12, 12, 10, 13, 13, 13, 13, 13, 13, 11, 13, 16, 18]
    for col_letter, width in zip("ABCDEFGHIJKLMNO", col_widths):
        trades_tab.column_dimensions[col_letter].width = width

    # Chart: gain/loss (AUD) per trade — the actual taxable figure
    chart = BarChart()
    chart.title = "Gain / Loss (AUD) by Trade"
    chart.style = 10
    chart.y_axis.number_format = price_fmt
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.height, chart.width = 8, 16

    n_events = len(summary["events"])
    chart.add_data(Reference(trades_tab, min_col=11, min_row=4, max_row=4 + n_events), titles_from_data=True)
    chart.set_categories(Reference(trades_tab, min_col=1, min_row=5, max_row=4 + n_events))
    trades_tab.add_chart(chart, "P4")

    # ============================================
    # TAB 2: SUMMARY — the FY-level breakdown, with loss offsetting applied
    # ============================================
    summary_tab = wb.create_sheet(summary_sheet_name)
    summary_tab.sheet_view.showGridLines = False

    summary_tab["A1"] = f"FY {financial_year} Summary — Net Capital Gain Calculation"
    summary_tab["A1"].font = TITLE_FONT
    summary_tab.merge_cells("A1:B1")

    summary_rows = [
        ("Short-term gains (held ≤12mo, no discount)", summary["short_term_gain_gross"], False),
        ("Long-term gains (held >12mo, discount-eligible)", summary["long_term_gain_gross"], False),
        ("Total capital losses", summary["total_losses"], False),
        ("Unused losses (carry forward — track manually)", summary["unused_losses"], False),
        ("Short-term gain after losses", summary["short_term_gain_after_losses"], False),
        ("Long-term gain after losses (pre-discount)", summary["long_term_gain_after_losses"], False),
        ("Long-term gain after 50% CGT discount", summary["discounted_long_term_gain"], False),
        ("ESTIMATED NET CAPITAL GAIN", summary["net_capital_gain"], True),
    ]
    if marginal_tax_rate is not None:
        tax_owed = summary["net_capital_gain"] * (marginal_tax_rate / 100)
        summary_rows.append((f"Estimated tax owed at {marginal_tax_rate:.1f}% marginal rate", tax_owed, True))

    for i, (label, value, is_final) in enumerate(summary_rows, start=3):
        label_cell = summary_tab.cell(row=i, column=1, value=label)
        value_cell = summary_tab.cell(row=i, column=2, value=value)
        value_cell.number_format = price_fmt
        if is_final:
            label_cell.font = Font(bold=True, size=13, color="1F4E78")
            value_cell.font = Font(bold=True, size=13, color="1F4E78")
            label_cell.fill = value_cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        else:
            label_cell.font = LABEL_FONT

    note_row_2 = 3 + len(summary_rows)
    if marginal_tax_rate is not None:
        summary_tab.cell(row=note_row_2, column=1,
                          value="Include the Medicare Levy (~2%) in your rate if it applies to you — not added automatically.")
        summary_tab.cell(row=note_row_2, column=1).font = Font(italic=True, size=9, color="666666")
        summary_tab.merge_cells(start_row=note_row_2, start_column=1, end_row=note_row_2, end_column=2)

    summary_tab["A2"] = "Planning estimate only — not tax advice."
    summary_tab["A2"].font = Font(italic=True, size=9, color="9C0000")

    summary_tab.column_dimensions["A"].width = 42
    summary_tab.column_dimensions["B"].width = 18

    # Chart: the three headline figures side by side
    chart2 = BarChart()
    chart2.title = "Short-Term vs Long-Term vs Net Capital Gain"
    chart2.style = 12
    chart2.y_axis.number_format = price_fmt
    chart2.y_axis.delete = False
    chart2.x_axis.delete = False
    chart2.height, chart2.width = 8, 14

    chart_data_row = len(summary_rows) + 5
    summary_tab.cell(row=chart_data_row, column=1, value="Short-Term Gain")
    summary_tab.cell(row=chart_data_row, column=2, value=summary["short_term_gain_after_losses"])
    summary_tab.cell(row=chart_data_row + 1, column=1, value="Long-Term Gain (Discounted)")
    summary_tab.cell(row=chart_data_row + 1, column=2, value=summary["discounted_long_term_gain"])
    summary_tab.cell(row=chart_data_row + 2, column=1, value="Net Capital Gain")
    summary_tab.cell(row=chart_data_row + 2, column=2, value=summary["net_capital_gain"])

    chart2.add_data(Reference(summary_tab, min_col=2, min_row=chart_data_row, max_row=chart_data_row + 2))
    chart2.set_categories(Reference(summary_tab, min_col=1, min_row=chart_data_row, max_row=chart_data_row + 2))
    summary_tab.add_chart(chart2, "D3")

    wb.save(PORTFOLIO_TAX_REPORT_FILE)
    return PORTFOLIO_TAX_REPORT_FILE


# ============================================

DAILY_SUMMARY_DIR = os.path.join(APP_DATA_DIR, "daily_summaries")


def send_mac_notification(title, message):
    """
    Fires a native macOS notification banner via AppleScript. Fails
    silently (does nothing) on non-Mac systems or if osascript isn't
    available — the saved log file is the real record either way, this is
    just a nice-to-have heads-up.
    """
    try:
        safe_title = title.replace('"', "'")
        safe_message = message.replace('"', "'")
        subprocess.run(
            ["osascript", "-e", f'display notification "{safe_message}" with title "{safe_title}"'],
            check=False, timeout=5,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass


def send_email_summary(subject, html_body, plain_text_body):
    """
    Sends a proper HTML email with a plain-text fallback part (best
    practice for HTML email — some clients/screen readers use the plain
    part instead). Requires the same EMAIL_ADDRESS/EMAIL_APP_PASSWORD
    environment variables as before; returns True/False rather than
    raising, so a missing/broken email setup doesn't take down the rest
    of the daily summary.
    """
    sender_email = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipient_email = os.environ.get("EMAIL_RECIPIENT", sender_email)

    if not sender_email or not app_password:
        print("Email not sent — set EMAIL_ADDRESS and EMAIL_APP_PASSWORD environment "
              "variables to enable email summaries (see setup instructions).")
        return False

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg.attach(MIMEText(plain_text_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))  # email clients render the LAST part that they support -- HTML last

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=15) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)
        print(f"Email sent to {recipient_email}")
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False


def send_email_with_attachments(subject, body, attachment_paths):
    """
    Sends a plain-text email with one or more files attached — used by
    the mobile-trigger workflow to deliver a generated .xlsx/.pptx
    straight to your inbox. Same EMAIL_ADDRESS/EMAIL_APP_PASSWORD
    environment variables as everything else; returns True/False rather
    than raising. Skips any path that doesn't actually exist rather than
    failing the whole email over one missing file.
    """
    sender_email = os.environ.get("EMAIL_ADDRESS")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipient_email = os.environ.get("EMAIL_RECIPIENT", sender_email)

    if not sender_email or not app_password:
        print("Email not sent — set EMAIL_ADDRESS and EMAIL_APP_PASSWORD environment variables.")
        return False

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg.attach(MIMEText(body, "plain"))

    attached_count = 0
    for path in attachment_paths:
        if not os.path.exists(path):
            continue
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
        msg.attach(part)
        attached_count += 1

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)
        print(f"Email with {attached_count} attachment(s) sent to {recipient_email}")
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False


def fetch_daily_summary_data():
    """
    Fetches every watchlist ticker's data ONCE — shared by both the
    plain-text log/notification and the HTML email, so nothing gets
    fetched twice and the two versions can never disagree.
    """
    watchlist = load_watchlist()
    ticker_details = []
    for ticker_symbol in watchlist:
        try:
            company = resolve_company(ticker_symbol)
            info = company["info"]
            price_divisor = company["price_divisor"]
            symbol = get_currency_symbol(company["quote_currency"])

            open_price = (info.get("regularMarketOpen") or info.get("open") or 0) / price_divisor
            day_low = (info.get("dayLow") or info.get("regularMarketDayLow") or 0) / price_divisor
            day_high = (info.get("dayHigh") or info.get("regularMarketDayHigh") or 0) / price_divisor
            close_price = (info.get("currentPrice") or info.get("regularMarketPrice") or 0) / price_divisor
            prev_close = (info.get("regularMarketPreviousClose") or close_price) / price_divisor
            day_change_pct = ((close_price - prev_close) / prev_close) if prev_close else 0

            ticker_details.append({
                "ticker": ticker_symbol, "company_name": company["company_name"], "symbol": symbol,
                "open": open_price, "low": day_low, "high": day_high, "close": close_price,
                "change_pct": day_change_pct, "news": get_recent_news(ticker_symbol, max_items=2), "error": None,
            })
        except ValueError as e:
            ticker_details.append({"ticker": ticker_symbol, "error": str(e)})
    return ticker_details


def build_daily_summary_text(ticker_details):
    """Plain-text version — used for the saved log file, terminal print,
    and as the HTML email's plain-text fallback part."""
    lines = [f"Daily Watchlist Summary — {datetime.now().strftime('%Y-%m-%d %H:%M')}", "=" * 70]
    if not ticker_details:
        lines.append("Watchlist is empty — nothing to summarize.")
        return "\n".join(lines)

    for t in ticker_details:
        if t.get("error"):
            lines.append(f"\n{t['ticker']}: Error resolving this ticker — {t['error']}")
            continue
        lines.append(f"\n{t['ticker']} — {t['company_name']}")
        lines.append(f"  Open: {t['symbol']}{t['open']:,.2f}   Low: {t['symbol']}{t['low']:,.2f}   "
                     f"High: {t['symbol']}{t['high']:,.2f}   Close: {t['symbol']}{t['close']:,.2f}   "
                     f"({t['change_pct']:+.1%})")
        if t["news"]:
            for item in t["news"]:
                lines.append(f"  • {item['title']} ({item['publisher']})")
        else:
            lines.append("  No recent headlines.")

    movers = [(t["ticker"], t["change_pct"]) for t in ticker_details if not t.get("error")]
    up_count = sum(1 for _, pct in movers if pct > 0)
    down_count = sum(1 for _, pct in movers if pct < 0)
    lines.append("\n" + "=" * 70)
    lines.append(f"{up_count} up, {down_count} down, {len(movers) - up_count - down_count} unchanged")
    return "\n".join(lines)


def build_daily_summary_html(ticker_details):
    """
    A clean, card-based HTML email — one card per ticker, color-coded
    moves with up/down arrows, real spacing and typography instead of a
    monospace ASCII-divider text dump. Uses a system font stack so it
    renders natively on whatever device it's opened on, and inline styles
    throughout since email clients don't reliably support <style> blocks.
    """
    today_str = datetime.now().strftime("%A, %B %d")
    font = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif"
    valid = [t for t in ticker_details if not t.get("error")]
    errored = [t for t in ticker_details if t.get("error")]

    if not valid and not errored:
        body_inner = f'<p style="font-family:{font}; font-size:14px; color:#555;">Your watchlist is empty right now — nothing to report today.</p>'
    else:
        up = sum(1 for t in valid if t["change_pct"] > 0)
        down = sum(1 for t in valid if t["change_pct"] < 0)
        flat = len(valid) - up - down
        flat_text = f", {flat} unchanged" if flat else ""
        intro = (f'<p style="font-family:{font}; font-size:15px; color:#333; margin:0 0 20px 0;">'
                 f"Here&rsquo;s how your watchlist did &mdash; <strong>{up} up</strong>, <strong>{down} down</strong>{flat_text}.</p>")

        cards = []
        for t in valid:
            if t["change_pct"] > 0:
                color, arrow = "#1a7f37", "&#9650;"
            elif t["change_pct"] < 0:
                color, arrow = "#c0362c", "&#9660;"
            else:
                color, arrow = "#888", "&#8212;"  # flat: neutral gray dash, not a misleading colored arrow
            news_html = ""
            if t["news"]:
                news_html = "".join(
                    f'<div style="font-size:13px; color:#555; margin-top:4px; font-family:{font};">'
                    f'&bull; {n["title"]} <span style="color:#999;">({n["publisher"]})</span></div>'
                    for n in t["news"])
            cards.append(f'''
            <div style="border:1px solid #e6e6e6; border-radius:10px; padding:16px 18px; margin-bottom:12px; font-family:{font};">
              <table width="100%" style="border-collapse:collapse;"><tr>
                <td style="text-align:left;">
                  <span style="font-size:16px; font-weight:600; color:#111;">{t['ticker']}</span>
                  <span style="font-size:13px; color:#888; margin-left:6px;">{t['company_name']}</span>
                </td>
                <td style="text-align:right; font-size:16px; font-weight:600; color:{color}; white-space:nowrap;">
                  {t['symbol']}{t['close']:,.2f} &nbsp; {arrow} {abs(t['change_pct']):.1%}
                </td>
              </tr></table>
              <div style="font-size:12.5px; color:#777; margin-top:6px;">
                Open {t['symbol']}{t['open']:,.2f} &nbsp;&middot;&nbsp; Low {t['symbol']}{t['low']:,.2f} &nbsp;&middot;&nbsp; High {t['symbol']}{t['high']:,.2f}
              </div>
              {news_html}
            </div>''')

        error_html = ""
        if errored:
            err_lines = "".join(f'<div style="font-size:13px; color:#c0362c; font-family:{font};">{t["ticker"]}: {t["error"]}</div>' for t in errored)
            error_html = f'<div style="margin-top:16px;">{err_lines}</div>'

        body_inner = intro + "".join(cards) + error_html

    return f'''<!DOCTYPE html>
<html>
<body style="margin:0; padding:0; background-color:#f7f7f8;">
  <div style="max-width:560px; margin:0 auto; padding:28px 20px;">
    <p style="font-family:{font}; font-size:13px; color:#999; margin:0 0 4px 0;">{today_str}</p>
    <h2 style="font-family:{font}; font-size:20px; color:#111; margin:0 0 18px 0; font-weight:600;">Your watchlist today</h2>
    {body_inner}
    <p style="font-family:{font}; font-size:11px; color:#bbb; margin-top:28px;">Sent automatically each day from your watchlist tracker.</p>
  </div>
</body>
</html>'''


def build_daily_summary():
    """Backward-compatible wrapper — returns the same (text, movers) shape
    as before, now built from the shared fetch_daily_summary_data()."""
    ticker_details = fetch_daily_summary_data()
    text = build_daily_summary_text(ticker_details)
    movers = [(t["ticker"], t["change_pct"]) for t in ticker_details if not t.get("error")]
    return text, movers


def run_daily_summary_and_notify():
    """
    The function a scheduled task actually calls. Fetches every watchlist
    ticker's data once, builds both a plain-text log entry and a clean
    HTML email from that same data, fires a short native Mac notification,
    and emails the HTML version (with the plain text as a fallback part)
    if EMAIL_ADDRESS/EMAIL_APP_PASSWORD are set.
    """
    ticker_details = fetch_daily_summary_data()
    summary_text = build_daily_summary_text(ticker_details)
    summary_html = build_daily_summary_html(ticker_details)
    movers = [(t["ticker"], t["change_pct"]) for t in ticker_details if not t.get("error")]

    os.makedirs(DAILY_SUMMARY_DIR, exist_ok=True)
    filename = os.path.join(DAILY_SUMMARY_DIR, f"{datetime.now().strftime('%Y-%m-%d')}.txt")
    with open(filename, "w") as f:
        f.write(summary_text)

    print(summary_text)
    print(f"\nSaved: {filename}")

    if movers:
        top_movers = sorted(movers, key=lambda x: abs(x[1]), reverse=True)[:3]
        digest = ", ".join(f"{t} {pct:+.1%}" for t, pct in top_movers)
    else:
        digest = "Watchlist is empty — nothing to report."

    send_mac_notification("Watchlist Daily Summary", digest)
    email_subject = f"Watchlist Daily Summary — {datetime.now().strftime('%Y-%m-%d')} ({digest})"
    send_email_summary(email_subject, summary_html, summary_text)



# Rule-based (regex/keyword) parsing, not a second AI call — deliberately,
# so this layer can't break the way the earnings-call feature did when a
# model provider changed its API. It scans plain English for intent
# keywords and ticker-looking words, verifies each candidate ticker against
# real data before using it, and falls through to the exact argv commands
# below if it can't confidently parse a request.
# ============================================

# Common words that would otherwise look like tickers (1-5 letters). Not
# exhaustive — just enough to keep everyday phrasing from being misread as
# a stock symbol. Real disambiguation happens by actually resolving each
# candidate against yfinance, not by this list alone.
COMMAND_STOPWORDS = {
    "a", "an", "the", "for", "on", "of", "and", "or", "to", "me", "please",
    "can", "you", "i", "is", "does", "build", "buy", "buying", "buys",
    "research", "dcf", "model", "models", "comps", "comp", "comparable",
    "companies", "company", "ma", "acquisition", "acquiring", "acquire",
    "acquires", "merger", "ipo", "pitch", "book", "pitchbook", "deck",
    "valuation", "want", "create", "make", "show", "get", "give", "stock",
    "share", "shares", "price", "current", "run", "do", "with", "at",
    "premium", "using", "using", "into", "about", "vs", "versus", "in",
    # Single-letter/contraction fragments that regex splitting produces
    # (e.g. "m&a" -> stray "m"; "what's" -> stray "s") and that happen to
    # also be real tickers (M = Macy's), which would otherwise get
    # mistaken for an intended ticker.
    "m", "s", "t", "d", "re", "ve", "ll", "new", "going",
    "add", "watchlist", "list", "my", "news", "global", "world", "markets",
}


def _extract_ticker_candidates(text):
    """
    Pulls out short, alphabetic, non-stopword tokens — the pool of
    "might be a ticker" guesses. Order is preserved (first mention first)
    since word order usually carries meaning (e.g. "X acquiring Y").
    """
    raw_tokens = re.findall(r"\b[A-Za-z]{1,5}\b", text)
    seen = set()
    candidates = []
    for token in raw_tokens:
        lower = token.lower()
        if lower in COMMAND_STOPWORDS:
            continue
        upper = token.upper()
        if upper not in seen:
            seen.add(upper)
            candidates.append(upper)
    return candidates


def _resolve_first_valid_tickers(candidates, count):
    """
    Tries each candidate against real data (via resolve_company) in order,
    keeping the first `count` that actually resolve to a tradeable ticker.
    This is what separates "looks like a ticker" from "is one" — e.g. in
    "build a dcf for zs", stopword-filtering leaves just "ZS", and this
    confirms it's real before running anything.
    """
    valid = []
    for candidate in candidates:
        if len(valid) >= count:
            break
        try:
            resolve_company(candidate)
            valid.append(candidate)
        except ValueError:
            continue
    return valid


def parse_natural_language_command(text):
    """
    Returns a dict describing what to do: {"mode": ..., plus whatever that
    mode needs}, or {"mode": "unknown", "reason": "..."} if nothing
    confident could be parsed. Checked in order from most to least specific,
    since e.g. "pitch book for a merger" should hit pitchbook-ma, not
    fall through to plain research.
    """
    lower_text = text.lower()
    candidates = _extract_ticker_candidates(text)

    is_ma_language = any(word in lower_text for word in
                          ["acqui", "merger", "buying", "buys", "takeover"])

    # --- Global market news (check very early — distinct from single-ticker research) ---
    if any(phrase in lower_text for phrase in
           ["global market news", "global news", "world markets", "major markets",
            "market news", "news across markets", "news from us", "news from uk"]):
        return {"mode": "global_news"}

    # --- Daily watchlist summary ---
    if any(phrase in lower_text for phrase in
           ["daily summary", "end of day", "how did my watchlist", "watchlist summary",
            "how did my stocks do", "today's summary"]):
        return {"mode": "daily_summary"}

    # --- Watchlist (also checked early, since "watchlist" isn't ambiguous with anything else) ---
    if "watchlist" in lower_text or "my list" in lower_text:
        if any(word in lower_text for word in ["add", "put", "save"]):
            tickers = _resolve_first_valid_tickers(candidates, 1)
            if not tickers:
                return {"mode": "unknown", "reason": "Couldn't find a valid ticker to add to the watchlist."}
            return {"mode": "watchlist_add", "ticker": tickers[0]}
        elif any(word in lower_text for word in ["remove", "delete", "drop"]):
            tickers = _resolve_first_valid_tickers(candidates, 1)
            if not tickers:
                return {"mode": "unknown", "reason": "Couldn't find a valid ticker to remove from the watchlist."}
            return {"mode": "watchlist_remove", "ticker": tickers[0]}
        else:
            return {"mode": "watchlist_show"}

    # --- Pitch book (check before plain research/dcf, since it mentions those words too) ---
    if "pitch book" in lower_text or "pitchbook" in lower_text or "deck" in lower_text:
        if is_ma_language:
            tickers = _resolve_first_valid_tickers(candidates, 2)
            if len(tickers) < 2:
                return {"mode": "unknown", "reason": "Needed two valid tickers (acquirer and target) for a pitch book M&A deck, but couldn't confirm two."}
            premium_match = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
            premium = float(premium_match.group(1)) / 100 if premium_match else 0.30
            return {"mode": "pitchbook_ma", "acquirer": tickers[0], "target": tickers[1], "premium": premium}
        else:
            tickers = _resolve_first_valid_tickers(candidates, 1)
            if not tickers:
                return {"mode": "unknown", "reason": "Couldn't find a valid ticker for the pitch book."}
            return {"mode": "pitchbook_research", "ticker": tickers[0]}

    # --- M&A ---
    if is_ma_language:
        tickers = _resolve_first_valid_tickers(candidates, 2)
        if len(tickers) < 2:
            return {"mode": "unknown", "reason": "M&A needs two valid tickers (acquirer and target), but couldn't confirm two."}
        premium_match = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
        premium = float(premium_match.group(1)) / 100 if premium_match else 0.30
        return {"mode": "ma", "acquirer": tickers[0], "target": tickers[1], "premium": premium}

    # --- IPO (needs manual financials, so this just triggers interactive follow-up questions) ---
    if "ipo" in lower_text:
        return {"mode": "ipo"}

    # --- Comps ---
    if "comp" in lower_text:
        tickers = _resolve_first_valid_tickers(candidates, 1)
        if not tickers:
            return {"mode": "unknown", "reason": "Couldn't find a valid ticker for comps."}
        return {"mode": "comps", "ticker": tickers[0]}

    # --- 3-statement model (check before plain DCF, since it's more specific) ---
    if any(phrase in lower_text for phrase in
           ["3 statement", "3-statement", "three statement", "three-statement",
            "income statement", "balance sheet", "full financial model"]):
        tickers = _resolve_first_valid_tickers(candidates, 1)
        if not tickers:
            return {"mode": "unknown", "reason": "Couldn't find a valid ticker for the 3-statement model."}
        return {"mode": "3statement", "ticker": tickers[0]}

    # --- DCF / valuation model ---
    if any(word in lower_text for word in ["dcf", "valuation model", "discounted cash flow"]):
        tickers = _resolve_first_valid_tickers(candidates, 1)
        if not tickers:
            return {"mode": "unknown", "reason": "Couldn't find a valid ticker for the DCF."}
        return {"mode": "dcf", "ticker": tickers[0]}

    # --- Default: research (the most common, lowest-commitment ask) ---
    tickers = _resolve_first_valid_tickers(candidates, 1)
    if tickers:
        return {"mode": "research", "ticker": tickers[0]}

    return {"mode": "unknown", "reason": "Couldn't find a recognizable ticker or command in that request."}


def _ask(prompt_text):
    """Small wrapper around input() so IPO's follow-up questions read consistently."""
    return input(f"  {prompt_text}: ").strip()


def _ask_float(prompt_text, allow_blank=False, blank_default=0.0):
    """Keeps re-prompting until a valid number is entered — important here
    since these numbers feed directly into tax calculations."""
    while True:
        raw = _ask(prompt_text)
        if allow_blank and not raw:
            return blank_default
        try:
            return float(raw)
        except ValueError:
            print("  Please enter a valid number.")


def interactive_log_transaction(txn_type):
    """
    Asks for each field one at a time instead of requiring exact
    positional command-line arguments — for tax-relevant numbers, a
    guided prompt is safer than hoping nothing got mixed up on one long
    command line. Brokerage/FX fees are calculated automatically (see
    calculate_trade_fees) rather than asked for.
    """
    print(f"\nLet's log a {txn_type}.")
    ticker_symbol = _ask("Ticker symbol (e.g. ZS)").upper()
    quantity = _ask_float("Number of shares")
    price = _ask_float("Price per share")
    date_input = _ask("Date of trade, YYYY-MM-DD (press Enter for today)")
    date = date_input if date_input else None

    add_transaction(ticker_symbol, txn_type, quantity, price, date=date)  # fees auto-calculated


def print_transactions_list():
    """Shows every logged transaction with a #-index, so edit/delete have
    something stable to reference. Shown in the order they're stored."""
    transactions = load_portfolio()
    if not transactions:
        print("No transactions recorded yet.")
        return

    divider = "─" * 78
    print(divider)
    print(f"{'#':<4}{'Date':<12}{'Type':<6}{'Ticker':<8}{'Qty':>10}{'Price':>14}{'Fees':>10}")
    print(divider)
    for i, t in enumerate(transactions, start=1):
        print(f"{i:<4}{t['date']:<12}{t['type'].upper():<6}{t['ticker']:<8}"
              f"{t['quantity']:>10,.2f}{format_price(t['price_per_share']):>14}"
              f"{'$' + format(t['fees'], ',.2f'):>10}")
    print(divider)


def delete_transaction(index_1_based):
    """index_1_based matches what print_transactions_list() displays —
    the #1 shown to the user, not a 0-based list index."""
    transactions = load_portfolio()
    if index_1_based < 1 or index_1_based > len(transactions):
        raise ValueError(f"No transaction #{index_1_based} — you have {len(transactions)} recorded. "
                          f"Check numbers with: portfolio list")

    removed = transactions.pop(index_1_based - 1)
    save_portfolio(transactions)
    print(f"Deleted #{index_1_based}: {removed['type'].upper()} {removed['quantity']} {removed['ticker']} "
          f"@ {format_price(removed['price_per_share'])} on {removed['date']}")
    return transactions


def edit_transaction(index_1_based):
    """
    Walks through each field with the current value shown, so you can
    just press Enter to keep anything unchanged. Fees are always
    recalculated automatically for the new details (consistent with how
    add_transaction works) — a manually-specified fee from long ago isn't
    preserved through an edit.
    """
    transactions = load_portfolio()
    if index_1_based < 1 or index_1_based > len(transactions):
        raise ValueError(f"No transaction #{index_1_based} — you have {len(transactions)} recorded. "
                          f"Check numbers with: portfolio list")

    old = transactions[index_1_based - 1]
    print(f"\nEditing #{index_1_based}: {old['type'].upper()} {old['quantity']} {old['ticker']} "
          f"@ {format_price(old['price_per_share'])} on {old['date']}")
    print("Press Enter on any question to keep its current value.\n")

    ticker_input = _ask(f"Ticker [{old['ticker']}]")
    ticker_symbol = ticker_input.upper() if ticker_input else old["ticker"]

    type_input = _ask(f"Type, buy or sell [{old['type']}]").strip().lower()
    txn_type = type_input if type_input in ("buy", "sell") else old["type"]

    qty_input = _ask(f"Quantity [{old['quantity']}]")
    quantity = float(qty_input) if qty_input else old["quantity"]

    price_input = _ask(f"Price per share [{old['price_per_share']}]")
    price = float(price_input) if price_input else old["price_per_share"]

    date_input = _ask(f"Date [{old['date']}]")
    date = date_input if date_input else old["date"]

    # Remove the old entry BEFORE re-adding, so a sell-quantity check
    # against current holdings doesn't get confused by counting the very
    # entry being replaced. Rolls back if the new details fail validation.
    transactions.pop(index_1_based - 1)
    save_portfolio(transactions)
    try:
        add_transaction(ticker_symbol, txn_type, quantity, price, date=date)
        print(f"Updated #{index_1_based}.")
    except ValueError as e:
        transactions.insert(index_1_based - 1, old)
        save_portfolio(transactions)
        raise ValueError(f"Edit failed and was rolled back — original entry restored: {e}")


def run_natural_language_command(text):
    """
    Parses one line of plain English and actually runs the matching mode —
    the bridge between parse_natural_language_command()'s guess and the
    real functions (run_research, run_dcf, etc.) built earlier.
    """
    parsed = parse_natural_language_command(text)
    mode = parsed["mode"]

    try:
        if mode == "global_news":
            print("\n(Understood: global market news scan)")
            print_global_market_news(get_global_market_news())

        elif mode == "daily_summary":
            print("\n(Understood: daily watchlist summary)")
            run_daily_summary_and_notify()

        elif mode == "watchlist_add":
            print(f"\n(Understood: add {parsed['ticker']} to watchlist)")
            add_to_watchlist(parsed["ticker"])

        elif mode == "watchlist_remove":
            print(f"\n(Understood: remove {parsed['ticker']} from watchlist)")
            remove_from_watchlist(parsed["ticker"])

        elif mode == "watchlist_show":
            print("\n(Understood: show watchlist)")
            print_watchlist_overview()

        elif mode == "research":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: research mode for {ticker_symbol})")
            print_research_report(run_research(ticker_symbol))

        elif mode == "dcf":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: DCF model for {ticker_symbol})")
            dcf_result = run_dcf(ticker_symbol)
            scenarios = run_scenarios(ticker_symbol)
            wb = export_dcf_to_excel(dcf_result)
            add_scenarios_tab(wb, scenarios)
            wb.save(f"{ticker_symbol}_DCF.xlsx")
            print(f"Saved: {ticker_symbol}_DCF.xlsx")

        elif mode == "comps":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: comps platform for {ticker_symbol})")
            comps = build_comps_table(ticker_symbol)
            export_comps_to_excel(comps, ticker_symbol).save(f"{ticker_symbol}_Comps.xlsx")
            print(f"Saved: {ticker_symbol}_Comps.xlsx")

        elif mode == "footballfield":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: football field valuation summary for {ticker_symbol})")
            field = build_football_field(ticker_symbol)
            for m in field["methods"]:
                print(f"  {m['method']:<28} {field['currency_symbol']}{m['low']:,.2f} - "
                      f"{field['currency_symbol']}{m['high']:,.2f}  (mid {field['currency_symbol']}{m['mid']:,.2f})")
            filename = f"{ticker_symbol}_FootballField.xlsx"
            export_football_field_to_excel(field).save(filename)
            print(f"Saved: {filename}")

        elif mode == "lbo":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: LBO model for {ticker_symbol})")
            lbo_result = run_lbo_model(ticker_symbol)
            filename = f"{ticker_symbol}_LBO.xlsx"
            export_lbo_to_excel(lbo_result).save(filename)
            print(f"Saved: {filename}")
            pitch_path = build_pitch_book("lbo", f"{ticker_symbol}_LBO_PitchBook.pptx", result=lbo_result)
            print(f"Saved: {pitch_path}")

        elif mode == "pubco":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: public company valuation for {ticker_symbol})")
            pubco_result = run_public_company_valuation(ticker_symbol)
            print_public_company_valuation(pubco_result)
            filename = f"{ticker_symbol}_PublicCompanyValuation.xlsx"
            export_public_company_valuation_to_excel(pubco_result).save(filename)
            print(f"Saved: {filename}")
            pitch_path = build_pitch_book("pubco", f"{ticker_symbol}_PublicCompanyValuation_PitchBook.pptx", result=pubco_result)
            print(f"Saved: {pitch_path}")

        elif mode == "3statement":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: 3-statement model for {ticker_symbol})")
            result = run_three_statement_model(ticker_symbol)
            print_three_statement_summary(result)
            filename = f"{ticker_symbol}_3Statement.xlsx"
            export_three_statement_to_excel(result).save(filename)
            print(f"Saved: {filename}")

        elif mode == "ma":
            print(f"\n(Understood: M&A model — {parsed['acquirer']} acquiring {parsed['target']} "
                  f"at a {parsed['premium']:.0%} premium)")
            deal = run_ma_model(parsed["acquirer"], parsed["target"], offer_premium=parsed["premium"])
            print_ma_summary(deal)
            filename = f"{parsed['acquirer']}_{parsed['target']}_MA.xlsx"
            export_ma_model_to_excel(deal).save(filename)
            print(f"Saved: {filename}")

        elif mode == "pitchbook_research":
            ticker_symbol = parsed["ticker"]
            print(f"\n(Understood: pitch book, research mode, for {ticker_symbol})")
            research = run_research(ticker_symbol)
            path = build_pitch_book("research", f"{ticker_symbol}_PitchBook.pptx", research=research)
            print(f"Saved: {path}")

        elif mode == "pitchbook_ma":
            print(f"\n(Understood: pitch book, M&A — {parsed['acquirer']} acquiring {parsed['target']})")
            deal = run_ma_model(parsed["acquirer"], parsed["target"], offer_premium=parsed["premium"])
            path = build_pitch_book("ma", f"{parsed['acquirer']}_{parsed['target']}_PitchBook.pptx", deal=deal)
            print(f"Saved: {path}")

        elif mode == "ipo":
            print("\n(Understood: IPO valuation — a pre-IPO company has no ticker, so a few quick details are needed.)")
            company_name = _ask("Company name")
            print(f"  Available industries: {', '.join(INDUSTRY_UNIVERSE.keys())}")
            industry = _ask("Industry (must match one above)")
            revenue = float(_ask("Revenue (e.g. 500000000 for $500M)"))
            ebitda_input = _ask("EBITDA (optional, press Enter to skip)")
            ebitda = float(ebitda_input) if ebitda_input else None
            shares_input = _ask("Shares outstanding post-IPO (optional, press Enter to skip)")
            shares = float(shares_input) if shares_input else None

            ipo = run_ipo_valuation(company_name, industry, revenue, ebitda=ebitda,
                                     shares_outstanding_post_ipo=shares)
            print_ipo_valuation(ipo)
            filename = f"{company_name.replace(' ', '_')}_IPO.xlsx"
            export_ipo_valuation_to_excel(ipo).save(filename)
            print(f"Saved: {filename}")

        else:
            print(f"\nCouldn't confidently understand that request: {parsed['reason']}")
            print("Try being explicit, e.g. 'research ZS', 'build a DCF for AAPL', "
                  "'build an M&A model PANW acquiring ZS at 25% premium'.")

    except (ValueError, RuntimeError) as e:
        print(f"\nError: {e}")


def run_interactive_mode():
    """
    The natural-language REPL: keeps asking for requests until the person
    types 'exit' or 'quit'. This is what makes 'python3 deal_tool.py' with
    no arguments feel like talking to the tool instead of memorizing exact
    command syntax — the argv-based commands below still work too, for
    scripting or when precision matters more than convenience.
    """
    print("Deal Tool — tell me what you want in plain English.")
    print("Examples: 'research ZS', 'build a DCF for AAPL', 'build comps for ZS',")
    print("          'build an M&A model PANW acquiring ZS at 25% premium', 'ipo valuation'")
    print("Type 'exit' or 'quit' to leave.\n")

    while True:
        try:
            text = input("> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break

        if not text:
            continue
        if text.lower() in ("exit", "quit"):
            print("Exiting.")
            break

        run_natural_language_command(text)
        print()


# ============================================
# ENTRY POINT
# argv-based exact commands (unchanged, for scripting/precision) PLUS the
# natural-language front door above. No arguments -> interactive NL mode.
# ============================================

def _usage():
    print("Usage:")
    print("  python3 deal_tool.py research TICKER")
    print("  python3 deal_tool.py dcf TICKER")
    print("  python3 deal_tool.py comps TICKER")
    print("  python3 deal_tool.py 3statement TICKER")
    print("  python3 deal_tool.py watchlist add TICKER")
    print("  python3 deal_tool.py watchlist remove TICKER")
    print("  python3 deal_tool.py watchlist show")
    print("  python3 deal_tool.py globalnews")
    print("  python3 deal_tool.py dailysummary")
    print("  python3 deal_tool.py portfolio buy TICKER QUANTITY PRICE [DATE=YYYY-MM-DD]   (fees auto-calculated)")
    print("  python3 deal_tool.py portfolio sell TICKER QUANTITY PRICE [DATE=YYYY-MM-DD]  (fees auto-calculated)")
    print("  python3 deal_tool.py portfolio holdings")
    print("  python3 deal_tool.py portfolio list")
    print("  python3 deal_tool.py portfolio edit [#]")
    print("  python3 deal_tool.py portfolio delete [#]")
    print("  python3 deal_tool.py portfolio taxreport                          (covers every year automatically)")
    print("  python3 deal_tool.py portfolio taxreport [FINANCIAL_YEAR] [RATE]   (focus on just one year)")
    print("  python3 deal_tool.py portfolio settaxrate RATE   (saved and reused automatically)")
    print("  python3 deal_tool.py portfolio taxrate           (show currently saved rate)")
    print("  python3 deal_tool.py ma ACQUIRER_TICKER TARGET_TICKER [premium]")
    print("  python3 deal_tool.py ipo COMPANY_NAME INDUSTRY REVENUE [EBITDA] [SHARES_OUTSTANDING]")
    print("  python3 deal_tool.py pitchbook research TICKER")
    print("  python3 deal_tool.py pitchbook ma ACQUIRER_TICKER TARGET_TICKER [premium]")


KNOWN_EXACT_COMMANDS = {"research", "dcf", "comps", "3statement", "ma", "ipo", "pitchbook", "watchlist", "globalnews", "dailysummary", "portfolio", "lbo", "footballfield", "sotp", "pubco", "mobiletrigger"}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        # No arguments at all -> drop into the natural-language REPL.
        run_interactive_mode()
        sys.exit(0)

    command = sys.argv[1].strip().lower()

    if command not in KNOWN_EXACT_COMMANDS:
        # Not one of the exact commands -> treat the whole argv as one
        # natural-language request, e.g.:
        #   python3 deal_tool.py research ZS for me please
        #   python3 deal_tool.py "build a dcf for aapl"
        run_natural_language_command(" ".join(sys.argv[1:]))
        sys.exit(0)

    try:
        if command == "research":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== RESEARCH MODE ===")
            print_research_report(run_research(ticker_symbol))

        elif command == "dcf":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== DCF MODEL ===")
            dcf_result = run_dcf(ticker_symbol)
            scenarios = run_scenarios(ticker_symbol)
            sensitivity = build_sensitivity_table(ticker_symbol, row_range=[0.08, 0.09, 0.10], col_range=[0.02, 0.03, 0.04])
            sensitivity_growth = build_sensitivity_table(
                ticker_symbol, row_range=[0.08, 0.09, 0.10], col_range=[0.15, 0.20, 0.23, 0.26, 0.30],
                vary="wacc_vs_growth",
            )
            wb = export_dcf_to_excel(dcf_result)
            add_scenarios_tab(wb, scenarios)
            add_sensitivity_tab(wb, "Sensitivity", f"{ticker_symbol} — WACC vs Terminal Growth",
                                 sensitivity, "WACC", "Terminal Growth", currency_symbol=dcf_result["currency_symbol"])
            add_sensitivity_tab(wb, "Sensitivity (Growth)", f"{ticker_symbol} — WACC vs Revenue Growth",
                                 sensitivity_growth, "WACC", "Starting Growth", currency_symbol=dcf_result["currency_symbol"])
            wb.save(f"{ticker_symbol}_DCF.xlsx")
            print(f"Saved: {ticker_symbol}_DCF.xlsx")

        elif command == "comps":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== COMPS PLATFORM ===")
            comps = build_comps_table(ticker_symbol)
            export_comps_to_excel(comps, ticker_symbol).save(f"{ticker_symbol}_Comps.xlsx")
            print(f"Saved: {ticker_symbol}_Comps.xlsx")

        elif command == "footballfield":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== FOOTBALL FIELD VALUATION SUMMARY ===")
            field = build_football_field(ticker_symbol)
            for m in field["methods"]:
                print(f"  {m['method']:<28} {field['currency_symbol']}{m['low']:,.2f} - "
                      f"{field['currency_symbol']}{m['high']:,.2f}  (mid {field['currency_symbol']}{m['mid']:,.2f})")
            filename = f"{ticker_symbol}_FootballField.xlsx"
            export_football_field_to_excel(field).save(filename)
            print(f"Saved: {filename}")

        elif command == "lbo":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== LBO MODEL ===")
            lbo_result = run_lbo_model(ticker_symbol)
            filename = f"{ticker_symbol}_LBO.xlsx"
            export_lbo_to_excel(lbo_result).save(filename)
            print(f"Saved: {filename}")
            pitch_path = build_pitch_book("lbo", f"{ticker_symbol}_LBO_PitchBook.pptx", result=lbo_result)
            print(f"Saved: {pitch_path}")

        elif command == "pubco":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== PUBLIC COMPANY VALUATION ===")
            pubco_result = run_public_company_valuation(ticker_symbol)
            print_public_company_valuation(pubco_result)
            filename = f"{ticker_symbol}_PublicCompanyValuation.xlsx"
            export_public_company_valuation_to_excel(pubco_result).save(filename)
            print(f"Saved: {filename}")
            pitch_path = build_pitch_book("pubco", f"{ticker_symbol}_PublicCompanyValuation_PitchBook.pptx", result=pubco_result)
            print(f"Saved: {pitch_path}")

        elif command == "mobiletrigger":
            # Usage: python3 deal_tool.py mobiletrigger COMMAND [ARGS...]
            # Runs the given command exactly as if typed directly (as a
            # subprocess, reusing every existing code path unchanged),
            # then emails every file it created as an attachment. Built
            # for the GitHub Actions workflow_dispatch mobile trigger —
            # not something you'd normally type by hand.
            real_args = sys.argv[2:]
            if not real_args:
                raise ValueError("Usage: mobiletrigger COMMAND [ARGS...], e.g. mobiletrigger dcf AAPL")

            before_files = set(os.listdir("."))
            proc = subprocess.run([sys.executable, sys.argv[0]] + real_args)
            after_files = set(os.listdir("."))
            new_files = sorted(f for f in (after_files - before_files) if f.lower().endswith((".xlsx", ".pptx", ".pdf")))

            command_str = " ".join(real_args)
            subject = f"DealTool Mobile — {command_str}"
            if new_files:
                body = f"Ran: {command_str}\n\nAttached {len(new_files)} file(s):\n" + "\n".join(new_files)
                send_email_with_attachments(subject, body, new_files)
            else:
                body = (f"Ran: {command_str}\n\nNo files were generated (this command may only print output, "
                        f"or it may have failed — exit code {proc.returncode}).")
                send_email_summary(subject, f"<pre>{body}</pre>", body)
            sys.exit(proc.returncode)

        elif command == "3statement":
            ticker_symbol = sys.argv[2].strip().upper()
            print("\n=== 3-STATEMENT MODEL ===")
            result = run_three_statement_model(ticker_symbol)
            print_three_statement_summary(result)
            filename = f"{ticker_symbol}_3Statement.xlsx"
            export_three_statement_to_excel(result).save(filename)
            print(f"Saved: {filename}")

        elif command == "watchlist":
            sub_action = sys.argv[2].strip().lower() if len(sys.argv) > 2 else "show"
            if sub_action == "add":
                add_to_watchlist(sys.argv[3].strip().upper())
            elif sub_action == "remove":
                remove_from_watchlist(sys.argv[3].strip().upper())
            elif sub_action == "show":
                print_watchlist_overview()
            else:
                print(f"Unknown watchlist action: '{sub_action}'. Use 'add', 'remove', or 'show'.")

        elif command == "globalnews":
            print("\n=== GLOBAL MARKET NEWS ===")
            print_global_market_news(get_global_market_news())

        elif command == "dailysummary":
            run_daily_summary_and_notify()

        elif command == "portfolio":
            sub_action = sys.argv[2].strip().lower() if len(sys.argv) > 2 else None
            if sub_action == "buy" or sub_action == "sell":
                if len(sys.argv) > 5:
                    # Full details given on the command line — use them directly.
                    # Fees are calculated automatically, so no [FEES] slot anymore.
                    ticker_symbol = sys.argv[3].strip().upper()
                    quantity = float(sys.argv[4])
                    price = float(sys.argv[5])
                    txn_date = sys.argv[6].strip() if len(sys.argv) > 6 else None
                    add_transaction(ticker_symbol, sub_action, quantity, price, date=txn_date)
                else:
                    # Not enough given upfront — ask for each field instead.
                    interactive_log_transaction(sub_action)
            elif sub_action == "holdings":
                print_holdings_overview()
            elif sub_action == "list":
                print_transactions_list()
            elif sub_action == "delete":
                if len(sys.argv) > 3:
                    index_to_delete = int(sys.argv[3])
                else:
                    print_transactions_list()
                    index_to_delete = int(_ask("Enter the # of the transaction to delete"))
                confirm = _ask(f"Delete transaction #{index_to_delete}? This can't be undone (y/n)").strip().lower()
                if confirm == "y":
                    delete_transaction(index_to_delete)
                else:
                    print("Cancelled — nothing deleted.")
            elif sub_action == "edit":
                if len(sys.argv) > 3:
                    index_to_edit = int(sys.argv[3])
                else:
                    print_transactions_list()
                    index_to_edit = int(_ask("Enter the # of the transaction to edit"))
                edit_transaction(index_to_edit)
            elif sub_action == "taxreport":
                fy = sys.argv[3].strip() if len(sys.argv) > 3 else None
                if len(sys.argv) > 4:
                    # Explicit override for this one report only — doesn't change the saved setting.
                    marginal_rate = float(sys.argv[4])
                else:
                    marginal_rate = get_marginal_tax_rate()
                    if marginal_rate is not None:
                        print(f"(Using saved marginal tax rate: {marginal_rate:.1f}% — "
                              f"update anytime with: portfolio settaxrate RATE)")
                    else:
                        rate_input = _ask("No saved tax rate yet. Enter your marginal tax rate as a percent "
                                           "(e.g. 32.5, include Medicare Levy if it applies) — or press Enter to skip")
                        if rate_input:
                            try:
                                marginal_rate = float(rate_input)
                                save_it = _ask("Save this rate for future reports too? (y/n)").strip().lower()
                                if save_it == "y":
                                    set_marginal_tax_rate(marginal_rate)
                            except ValueError:
                                print("  Couldn't read that as a number — showing the gain amount only.")

                if fy:
                    # A specific year was named — just do that one.
                    print_tax_report(fy, marginal_tax_rate=marginal_rate)
                    filename = export_tax_report_to_excel(fy, marginal_tax_rate=marginal_rate)
                    print(f"\nSaved: {filename}")
                else:
                    # No year given — cover every financial year that has any
                    # sell activity, so you never have to look up or type one.
                    all_transactions = load_portfolio()
                    realized = compute_realized_gains(all_transactions)
                    if not realized:
                        print("No sell transactions recorded yet — nothing to report.")
                    else:
                        years = sorted(set(r["financial_year"] for r in realized))
                        filename = None
                        for year in years:
                            print_tax_report(year, marginal_tax_rate=marginal_rate)
                            filename = export_tax_report_to_excel(year, marginal_tax_rate=marginal_rate)
                            print()
                        print(f"Updated {len(years)} financial year(s) in: {filename}")
            elif sub_action == "settaxrate":
                if len(sys.argv) > 3:
                    set_marginal_tax_rate(float(sys.argv[3]))
                else:
                    rate_input = _ask("Your marginal tax rate as a percent (e.g. 32.5, include Medicare Levy if applicable)")
                    set_marginal_tax_rate(float(rate_input))
            elif sub_action == "taxrate":
                rate = get_marginal_tax_rate()
                if rate is None:
                    print("No tax rate saved yet. Set one with: portfolio settaxrate RATE")
                else:
                    print(f"Saved marginal tax rate: {rate:.1f}%")
            else:
                print("Usage: portfolio buy|sell   (asks questions interactively)")
                print("   or: portfolio buy|sell TICKER QUANTITY PRICE [FEES] [DATE]  (all in one line)")
                print("       portfolio holdings")
                print("       portfolio list")
                print("       portfolio edit [#]")
                print("       portfolio delete [#]")
                print("       portfolio taxreport                          (covers every year automatically)")
                print("       portfolio taxreport [FINANCIAL_YEAR] [RATE]   (focus on just one year)")
                print("       portfolio settaxrate RATE   (saved and reused automatically)")
                print("       portfolio taxrate           (show currently saved rate)")

        elif command == "ma":
            acquirer_ticker = sys.argv[2].strip().upper()
            target_ticker = sys.argv[3].strip().upper()
            premium = float(sys.argv[4]) if len(sys.argv) > 4 else 0.30
            print("\n=== M&A DEAL MODEL ===")
            deal = run_ma_model(acquirer_ticker, target_ticker, offer_premium=premium)
            print_ma_summary(deal)
            filename = f"{acquirer_ticker}_{target_ticker}_MA.xlsx"
            export_ma_model_to_excel(deal).save(filename)
            print(f"Saved: {filename}")

        elif command == "ipo":
            company_name = sys.argv[2]
            industry = sys.argv[3]
            revenue = float(sys.argv[4])
            ebitda = float(sys.argv[5]) if len(sys.argv) > 5 else None
            shares = float(sys.argv[6]) if len(sys.argv) > 6 else None
            print("\n=== IPO VALUATION ===")
            ipo = run_ipo_valuation(company_name, industry, revenue, ebitda=ebitda,
                                     shares_outstanding_post_ipo=shares)
            print_ipo_valuation(ipo)
            filename = f"{company_name.replace(' ', '_')}_IPO.xlsx"
            export_ipo_valuation_to_excel(ipo).save(filename)
            print(f"Saved: {filename}")

        elif command == "sotp":
            company_name = sys.argv[2] if len(sys.argv) > 2 else _ask("Company name")
            print("\n=== SUM-OF-THE-PARTS VALUATION ===")
            print("Enter each segment's data. Segment financials and multiples aren't auto-fetched — "
                  "supply them from the company's disclosed segment data or your own research.\n")
            segments = []
            while True:
                seg_name = _ask(f"Segment {len(segments) + 1} name (blank to finish)")
                if not seg_name:
                    break
                metric_type = _ask("  Metric type (revenue/ebitda)").strip().lower()
                if metric_type not in ("revenue", "ebitda"):
                    print("  Must be 'revenue' or 'ebitda' — skipping this segment.")
                    continue
                metric_value = _ask_float(f"  {metric_type.title()} for this segment")
                multiple = _ask_float("  Multiple to apply (e.g. 8.5 for 8.5x)")
                segments.append({"name": seg_name, "metric_type": metric_type, "metric_value": metric_value, "multiple": multiple})
            if not segments:
                raise ValueError("Need at least one segment to run a sum-of-the-parts valuation.")
            net_debt = _ask_float("Consolidated net debt", allow_blank=True, blank_default=0.0)
            shares_input = _ask("Shares outstanding (optional, press Enter to skip)")
            shares = float(shares_input) if shares_input else None

            sotp_result = run_sotp_valuation(company_name, segments, net_debt=net_debt, shares_outstanding=shares)
            print_sotp_valuation(sotp_result)
            filename = f"{company_name.replace(' ', '_')}_SOTP.xlsx"
            export_sotp_to_excel(sotp_result).save(filename)
            print(f"Saved: {filename}")
            pitch_path = build_pitch_book_sotp(sotp_result, f"{company_name.replace(' ', '_')}_SOTP_PitchBook.pptx")
            print(f"Saved: {pitch_path}")

        elif command == "pitchbook":
            sub_mode = sys.argv[2].strip().lower()
            if sub_mode == "research":
                ticker_symbol = sys.argv[3].strip().upper()
                research = run_research(ticker_symbol)
                path = build_pitch_book("research", f"{ticker_symbol}_PitchBook.pptx", research=research)
            elif sub_mode == "ma":
                acquirer_ticker = sys.argv[3].strip().upper()
                target_ticker = sys.argv[4].strip().upper()
                premium = float(sys.argv[5]) if len(sys.argv) > 5 else 0.30
                deal = run_ma_model(acquirer_ticker, target_ticker, offer_premium=premium)
                path = build_pitch_book("ma", f"{acquirer_ticker}_{target_ticker}_PitchBook.pptx", deal=deal)
            elif sub_mode == "ipo":
                if len(sys.argv) < 6:
                    raise ValueError("IPO pitchbook requires COMPANY_NAME INDUSTRY REVENUE [EBITDA] [SHARES_OUTSTANDING].")
                company_name = sys.argv[3]
                industry = sys.argv[4]
                revenue = float(sys.argv[5])
                ebitda = float(sys.argv[6]) if len(sys.argv) > 6 else None
                shares = float(sys.argv[7]) if len(sys.argv) > 7 else None
                ipo = run_ipo_valuation(company_name, industry, revenue, ebitda=ebitda, shares_outstanding_post_ipo=shares)
                path = build_pitch_book("ipo", f"{company_name.replace(' ', '_')}_IPO_PitchBook.pptx", ipo=ipo, currency_symbol="$")
            elif sub_mode == "lbo":
                ticker_symbol = sys.argv[3].strip().upper()
                lbo_result = run_lbo_model(ticker_symbol)
                path = build_pitch_book("lbo", f"{ticker_symbol}_LBO_PitchBook.pptx", result=lbo_result)
            else:
                raise ValueError(f"Unknown pitchbook sub-mode: '{sub_mode}'. Use 'research', 'ma', 'ipo', or 'lbo'.")
            print(f"Saved: {path}")

        else:
            _usage()
            sys.exit(1)

    except (ValueError, RuntimeError, NotImplementedError) as e:
        print(f"\nError: {e}")