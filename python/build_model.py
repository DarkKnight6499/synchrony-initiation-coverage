"""Build SYF_Model.xlsx (live formulas) from the raw snapshots in data/."""
import datetime as dt
import json
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ============================ CONFIG (edit here) ============================
TICKER = "SYF"
COMPANY = "Synchrony Financial"
AUTHOR = "Yazad Madan"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR.parent
RAW_XLSX = BASE_DIR / "_build" / "SYF_Model_raw.xlsx"
FINAL_XLSX = OUT_DIR / "SYF_Model.xlsx"
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
HISTORY_FILE = DATA_DIR / "syf_history.json"
MARKET_FILE = DATA_DIR / "market_snapshot.json"
PEER_FILE = DATA_DIR / "peer_xbrl.json"
PRICES_FILE = DATA_DIR / "prices_1y.csv"
PEERS = ["COF", "AXP", "BFH", "OMF", "SOFI"]
PEER_NAMES = {"SYF": "Synchrony Financial", "COF": "Capital One", "AXP": "American Express", "BFH": "Bread Financial", "OMF": "OneMain Holdings", "SOFI": "SoFi Technologies"}
# peer tangible book components: XBRL tags (latest instant fact of each is used)
PEER_TAGS = {
    "COF": {"equity": "StockholdersEquity", "pref": "PreferredStockIncludingAdditionalPaidInCapitalNetOfDiscount", "intang": ["IntangibleAssetsNetIncludingGoodwill"]},
    "AXP": {"equity": "StockholdersEquity", "pref": None, "intang": ["Goodwill", "FiniteLivedIntangibleAssetsNet"]},
    "BFH": {"equity": "StockholdersEquity", "pref": "PreferredStockLiquidationPreferenceValue", "intang": ["IntangibleAssetsNetIncludingGoodwill"]},
    "OMF": {"equity": "StockholdersEquity", "pref": None, "intang": ["Goodwill", "IntangibleAssetsNetExcludingGoodwill"]},
    "SOFI": {"equity": "StockholdersEquity", "pref": None, "intang": ["Goodwill", "IntangibleAssetsNetExcludingGoodwill"]},
}
# Comps sheet layout: named columns (letter, header)
COMPS_HEADS = {
    "ticker": ("A", "Ticker"), "name": ("B", "Company"), "price": ("C", "Price ($)"), "shares": ("D", "Shares out (mm)"), "mcap": ("E", "Market cap ($mm)"),
    "eps_ttm": ("F", "TTM EPS ($)"), "eps26": ("G", "EPS 2026E ($)"), "eps27": ("H", "EPS 2027E ($)"), "pe_ttm": ("I", "P/E TTM"), "pe26": ("J", "P/E 2026E"), "pe27": ("K", "P/E 2027E"),
    "equity": ("L", "Equity ($mm)"), "pref": ("M", "Preferred ($mm)"), "gwi": ("N", "Goodwill and intangibles ($mm)"), "tce": ("O", "TCE ($mm)"), "tbvps": ("P", "TBV per share ($)"),
    "ptbv": ("Q", "P/TBV"), "rotce": ("R", "ROTCE (TTM EPS / TBVPS)"), "bsdate": ("S", "Balance sheet date"), "roe": ("T", "Yahoo ROE (memo)"), "pdate": ("U", "Price date"),
}
COMPS_COL = {k: v[0] for k, v in COMPS_HEADS.items()}
COMPS_VALUE_COL = "H"  # column holding the cross-check outputs
FY_LIST = [2020, 2021, 2022, 2023, 2024, 2025]
Q_LIST = ["Q1_2025", "Q2_2025", "Q1_2026", "Q2_2026"]
Q_LABELS = {"Q1_2025": "Q1 2025", "Q2_2025": "Q2 2025", "Q1_2026": "Q1 2026", "Q2_2026": "Q2 2026"}
Q_DAYS = {"Q1_2025": 90, "Q2_2025": 91, "Q1_2026": 90, "Q2_2026": 91}
BALANCE_DATE = dt.date(2026, 6, 30)  # latest reported balance sheet (Q2 2026 10-Q)

# Scenario drivers: FY2026E, FY2027E, FY2028E for bear / base / bull
SCENARIO_DRIVERS = [
    ("loan_g", "Loan receivables growth, year-end", "pct1", {"bear": [0.000, 0.010, 0.020], "base": [0.025, 0.040, 0.040], "bull": [0.040, 0.060, 0.060]},
     "FY2025A -0.9%; Q2 2026 period-end +2.4% YoY (10-Q)"),
    ("pv_g", "Purchase volume growth", "pct1", {"bear": [0.030, 0.010, 0.020], "base": [0.065, 0.045, 0.040], "bull": [0.080, 0.065, 0.060]},
     "H1 2026 +6.9% YoY (10-Q); FY2025A +0.1%"),
    ("yield", "Loan yield (interest and fees on loans / average loans)", "pct2", {"bear": [0.2130, 0.2100, 0.2100], "base": [0.2155, 0.2150, 0.2150], "bull": [0.2170, 0.2170, 0.2170]},
     "FY2025A 21.64%; Q2 2026 21.43%; bear reflects late-fee and pricing pressure"),
    ("cof", "Cost of funds (interest expense / average interest-bearing liabilities)", "pct2", {"bear": [0.0405, 0.0410, 0.0405], "base": [0.0395, 0.0390, 0.0390], "bull": [0.0390, 0.0375, 0.0370]},
     "FY2025A 4.28%; Q2 2026 3.96%"),
    ("nco", "Net charge-off rate (% of average loans)", "pct2", {"bear": [0.0560, 0.0700, 0.0650], "base": [0.0525, 0.0510, 0.0500], "bull": [0.0500, 0.0470, 0.0460]},
     "FY2025A 5.65%; H1 2026 5.43%; bear is a 2027 credit downturn"),
    ("cov", "Allowance coverage ratio (ACL / year-end loans)", "pct2", {"bear": [0.1020, 0.1120, 0.1080], "base": [0.0990, 0.0985, 0.0985], "bull": [0.0985, 0.0970, 0.0960]},
     "Dec 2025 10.06%; Jun 2026 10.09%; CECL builds ahead of losses in bear"),
    ("opex_g", "Other expense growth", "pct1", {"bear": [0.050, 0.030, 0.030], "base": [0.055, 0.045, 0.040], "bull": [0.060, 0.050, 0.045]},
     "FY2025A +6.1%; H1 2026 +6.4% YoY (10-Q)"),
    ("bb", "Share repurchases ($mm)", "num", {"bear": [3000, 1200, 1600], "base": [3200, 2600, 2700], "bull": [3400, 3200, 3400]},
     "H1 2026A $1,750mm; $5.7bn authorization remaining at Jun 30 2026 (10-Q)"),
]
# Assumptions common to all scenarios (FY2026E, FY2027E, FY2028E)
COMMON_ASSUMPTIONS = [
    ("avgf", "Average loans / simple average of opening and closing loans", "x3", [0.9650, 0.9700, 0.9700], "5-yr mean 0.972; FY2025A 0.962; 2026 set to reconcile to H1 2026A average loans"),
    ("oth_iea", "Other interest-earning assets as % of average loans", "pct1", [0.215, 0.210, 0.210], "FY2025A 20.8%; Q2 2026 21.7%"),
    ("ibl", "Interest-bearing liabilities as % of average interest-earning assets", "pct1", [0.800, 0.800, 0.800], "FY2025A 79.8%; Q2 2026 80.5%"),
    ("oth_y", "Yield on other interest-earning assets (cash, securities)", "pct2", [0.0375, 0.0360, 0.0350], "FY2025A 4.33%; Q2 2026 cash 3.71%, securities 3.83%"),
    ("rsa_res", "RSA structural residual vs credit regression (% of average loans)", "pct2", [0.0025, 0.0025, 0.0025], "FY2025A residual +0.22pp; H1 2026 about +0.3pp"),
    ("oi", "Other income as % of average loans", "pct2", [0.0054, 0.0054, 0.0054], "H1 2026A annualized 0.54%; FY2025A 0.52%"),
    ("tax", "Effective tax rate", "pct1", [0.240, 0.240, 0.240], "FY2025A 23.1%; H1 2026 24.3%"),
    ("pref", "Preferred dividends ($mm)", "num", [100, 120, 120], "FY2025A $83mm plus new $500mm 7.25% Series C (Jun 2026)"),
    ("dps", "Common dividend per share ($)", "usd", [1.28, 1.40, 1.50], "2026: 2 x $0.30 paid plus 2 x $0.34 announced; later years are author assumptions"),
    ("px_g", "Assumed annual share price drift for buyback pricing", "pct1", [0.00, 0.09, 0.09], "FY2026 H2 repurchases at spot; author assumption thereafter"),
    ("iss", "Net share issuance (mm; FY2026 is H2 only)", "num1", [0.5, 2.5, 2.5], "FY2025A: shares fell 41mm vs 43.7mm repurchased (about 2.7mm net issuance)"),
    ("dil", "Dilution added to average basic shares (mm)", "num1", [4.0, 4.0, 4.0], "FY2025A diluted 373.9mm vs basic 369.9mm"),
    ("rwa", "Risk-weighted assets / year-end loans", "pct1", [1.025, 1.025, 1.025], "Dec 2025 102.0%; Jun 2026 102.7%"),
    ("oth_tce", "Other TCE movements ($mm; OCI, share comp)", "num", [0, 0, 0], "Set to zero to keep clean surplus"),
]
# Terminal ROTCE by scenario (through-cycle return on tangible common equity after FY2028E)
SCENARIO_ROT = {"bear": 0.12, "base": 0.16, "bull": 0.20}
ROT_NOTE = "Base is about the midpoint of cost of equity and the FY2023 trough ROTCE (20.3%); FY2021-25 average was 28.8% but is flattered by post-COVID reserve releases"
# Valuation and market assumptions: key -> (label, value, format, note)
VAL_ASSUMPTIONS = [
    ("erp", "Equity risk premium", 0.050, "pct1", "Author assumption"),
    ("g", "Terminal growth rate of tangible equity", 0.030, "pct1", "Author assumption"),
    ("w_ri", "Weight on residual income value", 0.5, "pct0", "Author assumption"),
    ("w_pt", "Weight on justified P/TBV value", 0.5, "pct0", "Author assumption"),
    ("dps_int", "Dividend per share paid between balance sheet date and price date ($)", 0.34, "usd", "Q3 2026 dividend of $0.34 (10-Q Q2 2026); ex-date Aug 2026 per Yahoo"),
    ("dps_fwd", "Dividends per share in the 12 months after the price date ($)", 1.36, "usd", "4 x $0.34 quarterly rate (10-Q Q2 2026)"),
    ("ow", "Overweight if expected total return at or above", 0.15, "pct0", "Rating rule"),
    ("uw", "Underweight if expected total return at or below", -0.10, "pct0", "Rating rule"),
]
DEFAULT_SCENARIO = 2  # 1 bear, 2 base, 3 bull
# ==========================================================================

# field-name constants for the history JSON
F_LOANS_END, F_AVG_LOANS, F_IEA, F_IBL = "loans_end", "avg_loans", "iea_avg_bal", "ibl_avg_bal"
F_PV, F_ACTIVE = "purchase_volume", "avg_active_accounts"
F_INT_LOANS, F_INT_TOTAL, F_INT_EXP = "interest_and_fees_on_loans", "total_interest_income", "total_interest_expense"
F_NII, F_RSA, F_PROV = "net_interest_income", "rsa", "provision"
F_OI, F_OE, F_PRETAX, F_TAX = "other_income", "other_expense", "pretax_income", "income_tax"
F_NI, F_NIC, F_EPS, F_DSH = "net_earnings", "net_earnings_common", "eps_diluted", "diluted_shares_m"
F_NCO, F_ACL, F_DQ30, F_DQ90 = "net_chargeoffs", "acl_end", "dq30_reported", "dq90_reported"
F_EQ, F_PREF, F_GW, F_INT = "total_equity", "preferred_equity", "goodwill", "intangibles"
F_SH, F_CET1A, F_RWA, F_CET1R = "shares_outstanding_m", "cet1_amount", "rwa", "cet1_ratio"
F_NIM_REP, F_NCO_REP, F_EFF_REP = "nim_reported", "nco_rate_reported", "efficiency_reported"
F_LOAN_Y_REP, F_COF_REP = "loan_yield_reported", "cost_of_funds_reported"

FMT = {"num": "#,##0;(#,##0)", "num1": "#,##0.0", "usd": "$#,##0.00", "pct0": "0%", "pct1": "0.0%", "pct2": "0.00%", "x3": "0.000", "x2": '0.00"x"', "date": "yyyy-mm-dd", "txt": "@"}
BOLD = Font(bold=True)


def col(i):
    return get_column_letter(i)


def q(sheet):
    return f"'{sheet}'!"


class Sheet:
    def __init__(self, wb, name, widths=None):
        self.ws = wb.create_sheet(name)
        self.name = name
        if widths:
            for k, w in widths.items():
                self.ws.column_dimensions[k].width = w

    def put(self, cell, value, fmt=None, bold=False, wrap=False):
        c = self.ws[cell]
        c.value = value
        if fmt:
            c.number_format = FMT.get(fmt, fmt)
        if bold:
            c.font = BOLD
        if wrap:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        return c


def load_all():
    hist = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    market = json.loads(MARKET_FILE.read_text(encoding="utf-8"))
    peers = json.loads(PEER_FILE.read_text(encoding="utf-8"))
    return hist, market, peers


def close_before(ticker, cutoff):
    """Last completed daily close strictly before the retrieval date (the retrieval-day session may be incomplete)."""
    import csv
    best = None
    with PRICES_FILE.open(newline="") as f:
        for row in csv.DictReader(f):
            d = dt.date.fromisoformat(row["date"])
            if row["ticker"] == ticker and d < cutoff and (best is None or d > best[0]):
                best = (d, float(row["close"]))
    return best


def raw(d, k):
    return (d.get(k) or {}).get("raw") if isinstance(d.get(k), dict) else d.get(k)


# ---------------------------------------------------------------- Historical
HIST_COL = {}  # period label -> column letter
for _i, _y in enumerate(FY_LIST):
    HIST_COL[f"FY{_y}"] = col(2 + _i)
for _i, _q in enumerate(Q_LIST):
    HIST_COL[_q] = col(2 + len(FY_LIST) + _i)
COL_H1, COL_LTM = col(2 + len(FY_LIST) + len(Q_LIST)), col(3 + len(FY_LIST) + len(Q_LIST))
COL_SRC = col(5 + len(FY_LIST) + len(Q_LIST))
PERIODS = [f"FY{y}" for y in FY_LIST] + Q_LIST
PREV = {}  # column of the comparable prior period
for _i, _y in enumerate(FY_LIST[1:], start=1):
    PREV[HIST_COL[f"FY{_y}"]] = HIST_COL[f"FY{FY_LIST[_i - 1]}"]
PREV[HIST_COL["Q1_2026"]] = HIST_COL["Q1_2025"]
PREV[HIST_COL["Q2_2026"]] = HIST_COL["Q2_2025"]

SRC_TEXT = {
    "S_STAT": "10-K and 10-Q MD&A, Other Financial and Statistical Data table",
    "S_AVG": "10-K and 10-Q MD&A, Average Balance Sheet table",
    "S_IS": "Consolidated Statements of Earnings (XBRL cross-checked)",
    "S_EPSN": "Note: Earnings per share, weighted average common shares outstanding, dilutive",
    "S_BS": "Consolidated Statements of Financial Position (XBRL cross-checked for equity)",
    "S_CAP": "10-K and 10-Q MD&A, Capital, Basel III table (Synchrony Financial)",
    "S_RECAST": "10-Q Q2 2026 comparative column for Dec-25 (software reclassification recast)",
    "S_XBRL": "SEC XBRL companyfacts (us-gaap tag in row label on Inputs_Sources) for fiscal years; MD&A Dividend and Share Repurchases tables for quarters",
    "S_DAYS": "Calendar",
}
# input rows: key, label, json field, transform, fmt, source-key (for the register)
H_INPUT_SECTIONS = [
    ("Balances and volumes ($mm unless stated)", [
        ("loans_end", "Loan receivables, period end", F_LOANS_END, 1, "num", "S_STAT"),
        ("avg_loans", "Average loan receivables, including held for sale", F_AVG_LOANS, 1, "num", "S_STAT"),
        ("iea", "Average interest-earning assets", F_IEA, 1, "num", "S_AVG"),
        ("ibl", "Average interest-bearing liabilities", F_IBL, 1, "num", "S_AVG"),
        ("pv", "Purchase volume", F_PV, 1, "num", "S_STAT"),
        ("active", "Average active accounts (thousands)", F_ACTIVE, 1, "num", "S_STAT"),
    ]),
    ("Income statement ($mm)", [
        ("int_loans", "Interest and fees on loans", F_INT_LOANS, 1, "num", "S_IS"),
        ("int_total", "Total interest income", F_INT_TOTAL, 1, "num", "S_IS"),
        ("int_exp", "Interest expense", F_INT_EXP, 1, "num", "S_IS"),
        ("nii", "Net interest income", F_NII, 1, "num", "S_IS"),
        ("rsa", "Retailer share arrangements (RSA), shown positive", F_RSA, -1, "num", "S_IS"),
        ("prov", "Provision for credit losses", F_PROV, 1, "num", "S_IS"),
        ("oi", "Other income", F_OI, 1, "num", "S_IS"),
        ("oe", "Other expense", F_OE, 1, "num", "S_IS"),
        ("pretax", "Earnings before income taxes (reported)", F_PRETAX, 1, "num", "S_IS"),
        ("tax", "Provision for income taxes", F_TAX, 1, "num", "S_IS"),
        ("ni", "Net earnings", F_NI, 1, "num", "S_IS"),
        ("nic", "Net earnings available to common stockholders", F_NIC, 1, "num", "S_IS"),
        ("eps", "Diluted EPS ($)", F_EPS, 1, "usd", "S_IS"),
        ("dsh", "Diluted weighted average shares (mm)", F_DSH, 1, "num1", "S_EPSN"),
    ]),
    ("Credit quality ($mm)", [
        ("nco", "Net charge-offs", F_NCO, 1, "num", "S_STAT"),
        ("acl", "Allowance for credit losses, period end", F_ACL, 1, "num", "S_STAT"),
        ("dq30", "30+ days past due, % of period-end loans", F_DQ30, 0.01, "pct2", "S_STAT"),
        ("dq90", "90+ days past due, % of period-end loans", F_DQ90, 0.01, "pct2", "S_STAT"),
    ]),
    ("Equity and capital ($mm unless stated)", [
        ("equity", "Total equity", F_EQ, 1, "num", "S_BS"),
        ("pref", "Preferred stock", F_PREF, 1, "num", "S_BS"),
        ("gw", "Goodwill", F_GW, 1, "num", "S_BS"),
        ("intang", "Intangible assets, net (as originally reported)", F_INT, 1, "num", "S_BS"),
        ("sh", "Common shares outstanding, period end (mm)", F_SH, 1, "num1", "S_BS"),
        ("cet1a", "Common equity Tier 1 capital", F_CET1A, 1, "num", "S_CAP"),
        ("rwa", "Risk-weighted assets", F_RWA, 1, "num", "S_CAP"),
        ("cet1r", "CET1 ratio (as originally reported)", F_CET1R, 0.01, "pct1", "S_CAP"),
    ]),
]

H_REPORTED_ROWS = [  # reported ratios kept for reconciliation
    ("nim_rep", "Net interest margin, reported", F_NIM_REP, 0.01, "pct2", "S_STAT"),
    ("nco_rep", "Net charge-off rate, reported", F_NCO_REP, 0.01, "pct2", "S_STAT"),
    ("eff_rep", "Efficiency ratio, reported", F_EFF_REP, 0.01, "pct1", "S_STAT"),
    ("yld_rep", "Loan yield, reported", F_LOAN_Y_REP, 0.01, "pct2", "S_AVG"),
    ("cof_rep", "Cost of funds, reported", F_COF_REP, 0.01, "pct2", "S_AVG"),
]

# derived rows: key, label, template, scope, fmt. Tokens {k} same column, {k@p} comparable prior column
H_DERIVED = [
    ("Derived ratios (formulas)", None),
    ("ann", "Annualization factor (365 / days; 1 for fiscal years)", None, None, "x3"),
    ("loan_gr", "Loan receivables growth, YoY", "={loans_end}/{loans_end@p}-1", "all", "pct1"),
    ("pv_gr", "Purchase volume growth, YoY", "={pv}/{pv@p}-1", "all", "pct1"),
    ("pv_int", "Purchase volume / average loans (annualized)", "={pv}*{ann}/{avg_loans}", "all", "pct1"),
    ("act_gr", "Average active accounts growth, YoY", "={active}/{active@p}-1", "all", "pct1"),
    ("yield", "Loan yield", "={int_loans}*{ann}/{avg_loans}", "all", "pct2"),
    ("oth_y", "Yield on other interest-earning assets", "=({int_total}-{int_loans})*{ann}/({iea}-{avg_loans})", "all", "pct2"),
    ("cof", "Cost of funds", "={int_exp}*{ann}/{ibl}", "all", "pct2"),
    ("nim", "Net interest margin (recomputed)", "={nii}*{ann}/{iea}", "all", "pct2"),
    ("nim_d", "NIM: recomputed less reported", "={nim}-{nim_rep}", "all", "0.00%"),
    ("rsa_pct", "RSA as % of average loans", "={rsa}*{ann}/{avg_loans}", "all", "pct2"),
    ("nco_rate", "Net charge-off rate (recomputed)", "={nco}*{ann}/{avg_loans}", "all", "pct2"),
    ("nco_d", "NCO rate: recomputed less reported", "={nco_rate}-{nco_rep}", "all", "0.00%"),
    ("cov", "Allowance coverage ratio", "={acl}/{loans_end}", "all", "pct2"),
    ("res_b", "Reserve build (ACL change)", "={acl}-{acl@p}", "fy", "num"),
    ("acl_oth", "Other ACL movements (plug to roll-forward)", "={acl}-({acl@p}+{prov}-{nco})", "fy", "num"),
    ("oi_pct", "Other income as % of average loans", "={oi}*{ann}/{avg_loans}", "all", "pct2"),
    ("oe_gr", "Other expense growth, YoY", "={oe}/{oe@p}-1", "all", "pct1"),
    ("oe_pct", "Other expense as % of average loans", "={oe}*{ann}/{avg_loans}", "all", "pct2"),
    ("eff", "Efficiency ratio (recomputed)", "={oe}/({nii}-{rsa}+{oi})", "all", "pct1"),
    ("eff_d", "Efficiency: recomputed less reported", "={eff}-{eff_rep}", "all", "0.0%"),
    ("taxr", "Effective tax rate", "={tax}/{pretax}", "all", "pct1"),
    ("pretax_c", "Pretax check: NII - RSA - provision + other income - other expense", "={nii}-{rsa}-{prov}+{oi}-{oe}", "all", "num"),
    ("tce", "Tangible common equity (equity - preferred - goodwill - intangibles)", "={equity}-{pref}-{gw}-{intang}", "bs", "num"),
    ("rotce", "ROTCE (net earnings to common / average TCE; annualized)", "={nic}*{ann}/AVERAGE({tce},{tce@p})", "fy", "pct1"),
    ("tbvps", "Tangible book value per share ($)", "={tce}/{sh}", "bs", "usd"),
    ("avgf", "Average loans / simple average of opening and closing loans", "={avg_loans}/AVERAGE({loans_end},{loans_end@p})", "fy", "x3"),
    ("oth_iea", "Other interest-earning assets as % of average loans", "=({iea}-{avg_loans})/{avg_loans}", "all", "pct1"),
    ("ibl_pct", "Interest-bearing liabilities as % of average interest-earning assets", "={ibl}/{iea}", "all", "pct1"),
    ("rwa_pct", "Risk-weighted assets / loans", "={rwa}/{loans_end}", "cap", "pct1"),
    ("pay_bb", "Buybacks as % of net earnings to common", "={bb}/{nic}", "fy", "pct0"),
]


def hist_value(hist, market, period, field, scale):
    if period.startswith("FY"):
        v = hist["annual"][period[2:]].get(field)
    else:
        v = hist["quarterly"][period].get(field)
    if v is None:
        return None
    return v * scale


def build_historical(wb, hist):
    sh = Sheet(wb, "Historical", {"A": 62})
    ws = sh.ws
    for c in range(2, 16):
        ws.column_dimensions[col(c)].width = 12
    sh.put("A1", f"{COMPANY} ({TICKER}): historical financials, $ in millions unless stated", bold=True)
    sh.put("A2", "Inputs are as reported (see Inputs_Sources for the filing and table behind each row). Ratio rows are formulas. Quarterly ratios are annualized on 365/days, as the company does.")
    hdr = 4
    sh.put(f"A{hdr}", "Period", bold=True)
    for p in PERIODS:
        sh.put(f"{HIST_COL[p]}{hdr}", p.replace("_", " ").replace("FY", "FY ") if p.startswith("FY") else Q_LABELS[p], bold=True)
    sh.put(f"{COL_H1}{hdr}", "H1 2026", bold=True)
    sh.put(f"{COL_LTM}{hdr}", "LTM Jun-26", bold=True)
    sh.put(f"{COL_SRC}{hdr}", "Source (filing and table; see also Inputs_Sources)", bold=True)
    ws.column_dimensions[COL_SRC].width = 70
    R = {}
    row = hdr + 1
    written = set()

    def set_input(key, label, field, scale, fmt, src=None):
        nonlocal row
        R[key] = row
        sh.put(f"A{row}", label)
        if src:
            sh.put(f"{COL_SRC}{row}", SRC_TEXT[src])
        for p in PERIODS:
            v = hist_value(hist, None, p, field, scale)
            if v is not None:
                sh.put(f"{HIST_COL[p]}{row}", v, fmt)
                written.add((key, HIST_COL[p]))
        row += 1

    R["days"] = row
    sh.put(f"A{row}", "Days in period")
    sh.put(f"{COL_SRC}{row}", SRC_TEXT["S_DAYS"])
    for p in PERIODS:
        if p.startswith("FY"):
            y = int(p[2:])
            sh.put(f"{HIST_COL[p]}{row}", 366 if y % 4 == 0 else 365, "num")
        else:
            sh.put(f"{HIST_COL[p]}{row}", Q_DAYS[p], "num")
        written.add(("days", HIST_COL[p]))
    row += 1
    for title, items in H_INPUT_SECTIONS:
        row += 1
        sh.put(f"A{row}", title, bold=True)
        row += 1
        for it in items:
            set_input(*it)
        if title.startswith("Equity"):
            # recast rows and XBRL/text-only capital return inputs
            for key, label, val, fmt in [
                ("intang_r", "Intangible assets, net, recast for reclassified internal-use software (Q2 2026 10-Q, Dec-25 comparative)", hist["quarterly"]["Q2_2026"]["_recast_prior_ye"]["intangibles"], "num"),
                ("cet1a_r", "CET1 capital, recast (Q2 2026 10-Q, Dec-25 comparative)", hist["quarterly"]["Q2_2026"]["_recast_prior_ye"]["cet1_amount"], "num"),
                ("rwa_r", "Risk-weighted assets, recast (Q2 2026 10-Q, Dec-25 comparative)", hist["quarterly"]["Q2_2026"]["_recast_prior_ye"]["rwa"], "num"),
                ("cet1r_r", "CET1 ratio, recast (Q2 2026 10-Q, Dec-25 comparative)", hist["quarterly"]["Q2_2026"]["_recast_prior_ye"]["cet1_ratio"] / 100, "pct1"),
            ]:
                R[key] = row
                sh.put(f"A{row}", label)
                sh.put(f"{COL_SRC}{row}", SRC_TEXT["S_RECAST"])
                sh.put(f"{HIST_COL['FY2025']}{row}", val, fmt)
                written.add((key, HIST_COL["FY2025"]))
                row += 1
            R["tce_r"] = row
            sh.put(f"A{row}", "TCE, recast basis (used as forecast starting point)")
            sh.put(f"{COL_SRC}{row}", "Formula: equity - preferred - goodwill - recast intangibles")
            g = HIST_COL["FY2025"]
            sh.put(f"{g}{row}", f"={g}{R['equity']}-{g}{R['pref']}-{g}{R['gw']}-{g}{R['intang_r']}", "num")
            k = HIST_COL["Q2_2026"]
            sh.put(f"{k}{row}", f"={k}{R['equity']}-{k}{R['pref']}-{k}{R['gw']}-{k}{R['intang']}", "num")
            written.update({("tce_r", g), ("tce_r", k)})
            row += 1
            # capital return
            xb = hist["xbrl_extra"]
            R["dps"], R["bb"], R["div"], R["bb_sh"] = row, row + 1, row + 2, row + 3
            sh.put(f"A{row}", "Common dividends declared per share ($)")
            sh.put(f"A{row + 1}", "Share repurchases ($mm; cash flow basis for fiscal years, program purchases for quarters)")
            sh.put(f"A{row + 2}", "Common dividends ($mm)")
            sh.put(f"A{row + 3}", "Shares repurchased under programs (mm)")
            for off, txt in enumerate(["XBRL CommonStockDividendsPerShareDeclared (FY); MD&A Dividend and Share Repurchases table (quarters)", "XBRL PaymentsForRepurchaseOfCommonStock, cash flow basis (FY); MD&A program purchases table (quarters)", "XBRL DividendsCommonStockCash (FY); MD&A dividends table (quarters)", "MD&A Common Shares Repurchased Under Publicly Announced Programs table"]):
                sh.put(f"{COL_SRC}{row + off}", txt)
            for y in FY_LIST:
                c = HIST_COL[f"FY{y}"]
                sh.put(f"{c}{row}", xb["dividends_per_share"][str(y)]["val"], "usd")
                sh.put(f"{c}{row + 1}", xb["buyback_cash_mm"][str(y)]["val"], "num")
                sh.put(f"{c}{row + 2}", xb["common_dividends_mm"][str(y)]["val"], "num")
                written.update({("dps", c), ("bb", c), ("div", c)})
            bbq = hist["buybacks_quarterly_text"]
            dvq = hist["misc"]["dividends_quarterly_text"]
            c25 = HIST_COL["FY2025"]
            sh.put(f"{c25}{row + 3}", sum(bbq[f"Q{i}_2025"]["shares_m"] for i in range(1, 5)), "num1")
            written.add(("bb_sh", c25))
            for qk in Q_LIST:
                c = HIST_COL[qk]
                sh.put(f"{c}{row}", dvq[qk]["dps"], "usd")
                sh.put(f"{c}{row + 1}", bbq[qk]["dollars_mm"], "num")
                sh.put(f"{c}{row + 2}", dvq[qk]["dollars_mm"], "num")
                sh.put(f"{c}{row + 3}", bbq[qk]["shares_m"], "num1")
                written.update({("dps", c), ("bb", c), ("div", c), ("bb_sh", c)})
            row += 4
    row += 1
    sh.put(f"A{row}", "Reported ratios (kept to reconcile the recomputed ratios below)", bold=True)
    row += 1
    for it in H_REPORTED_ROWS:
        set_input(*it)

    # derived
    def tok(m, c, key, at_prev):
        target = PREV.get(c) if at_prev else c
        if target is None or (key, target) not in written:
            return None
        return f"{target}{R[key]}"

    import re
    pat = re.compile(r"\{(\w+)(@p)?\}")
    for entry in H_DERIVED:
        if entry[1] is None:
            row += 1
            sh.put(f"A{row}", entry[0], bold=True)
            row += 1
            continue
        key, label, tmpl, scope, fmt = entry
        R[key] = row
        sh.put(f"A{row}", label)
        for p in PERIODS:
            c = HIST_COL[p]
            if key == "ann":
                sh.put(f"{c}{row}", 1 if p.startswith("FY") else f"=365/{c}{R['days']}", fmt)
                written.add((key, c))
                continue
            if scope == "fy" and not p.startswith("FY"):
                continue
            if scope == "bs" and (key, c) is None:
                pass
            ok = True

            def repl(m):
                nonlocal ok
                ref = tok(m, c, m.group(1), bool(m.group(2)))
                if ref is None:
                    ok = False
                    return "0"
                return ref

            f = pat.sub(repl, tmpl)
            if scope == "bs" and (("equity", c) not in written or ("intang", c) not in written):
                ok = False
            if scope == "cap" and ("rwa", c) not in written:
                ok = False
            if ok:
                sh.put(f"{c}{row}", f, fmt)
                written.add((key, c))
        row += 1
    # H1 2026 and LTM columns for flow rows
    j, k, h_, i_ = HIST_COL["Q1_2026"], HIST_COL["Q2_2026"], HIST_COL["Q1_2025"], HIST_COL["Q2_2025"]
    g = HIST_COL["FY2025"]
    for key in ["pv", "int_loans", "int_total", "int_exp", "nii", "rsa", "prov", "oi", "oe", "pretax", "tax", "ni", "nic", "eps", "nco", "bb", "div", "bb_sh"]:
        r = R[key]
        fmt = "usd" if key == "eps" else ("num1" if key == "bb_sh" else "num")
        sh.put(f"{COL_H1}{r}", f"={j}{r}+{k}{r}", fmt)
        if key != "bb_sh":
            sh.put(f"{COL_LTM}{r}", f"={g}{r}+{COL_H1}{r}-({h_}{r}+{i_}{r})", fmt)
    sh.put(f"{COL_H1}{R['dps']}", f"={j}{R['dps']}+{k}{R['dps']}", "usd")
    sh.put(f"{COL_H1}{R['dsh']}", f"=AVERAGE({j}{R['dsh']},{k}{R['dsh']})", "num1")
    sh.put(f"{COL_H1}{R['ann']}", f"=365/({j}{R['days']}+{k}{R['days']})", "x3")
    # regression block
    row += 1
    sh.put(f"A{row}", "RSA versus credit: FY2020-FY2025 regression of RSA % of average loans on the net charge-off rate", bold=True)
    row += 1
    b, gcol = HIST_COL["FY2020"], HIST_COL["FY2025"]
    rsa_rng = f"{b}{R['rsa_pct']}:{gcol}{R['rsa_pct']}"
    nco_rng = f"{b}{R['nco_rate']}:{gcol}{R['nco_rate']}"
    R["slope"], R["intercept"], R["correl"] = row, row + 1, row + 2
    sh.put(f"A{row}", "Slope (RSA % change per 1.0 change in NCO rate)")
    sh.put(f"B{row}", f"=SLOPE({rsa_rng},{nco_rng})", "0.000")
    sh.put(f"A{row + 1}", "Intercept")
    sh.put(f"B{row + 1}", f"=INTERCEPT({rsa_rng},{nco_rng})", "0.0000")
    sh.put(f"A{row + 2}", "Correlation")
    sh.put(f"B{row + 2}", f"=CORREL({rsa_rng},{nco_rng})", "0.000")
    row += 3
    R["fit_res"] = row
    sh.put(f"A{row}", "Regression residual (actual less fitted RSA %)")
    for y in FY_LIST:
        c = HIST_COL[f"FY{y}"]
        sh.put(f"{c}{row}", f"={c}{R['rsa_pct']}-($B${R['intercept']}+$B${R['slope']}*{c}{R['nco_rate']})", "pct2")
    row += 1
    R["rotce_avg"] = row
    sh.put(f"A{row}", "Average ROTCE FY2021-FY2025")
    sh.put(f"B{row}", f"=AVERAGE({HIST_COL['FY2021']}{R['rotce']}:{gcol}{R['rotce']})", "pct1")
    row += 1
    R["avgf_mean"] = row
    sh.put(f"A{row}", "Mean average-loan factor FY2021-FY2025")
    sh.put(f"B{row}", f"=AVERAGE({HIST_COL['FY2021']}{R['avgf']}:{gcol}{R['avgf']})", "x3")
    ws.freeze_panes = "B5"
    return sh, R


# ---------------------------------------------------------------- Inputs_Sources
def build_inputs(wb, hist, market, peers):
    sh = Sheet(wb, "Inputs_Sources", {"A": 60, "B": 16, "C": 18, "D": 70, "E": 60})
    I = {}
    syf = market["tickers"][TICKER]
    retrieved = dt.date.fromisoformat(market["retrieved_date"])
    price_date, price = close_before(TICKER, retrieved)
    sh.put("A1", f"{COMPANY} ({TICKER}): market inputs, valuation assumptions and source register", bold=True)
    sh.put("A2", "All historical figures are from SEC EDGAR filings by Synchrony Financial (CIK 0001601712). Market data is from Yahoo Finance. Forecast and valuation assumptions are the author's illustrative inputs.")
    r = 4
    sh.put(f"A{r}", "A. Market data (Yahoo Finance)", bold=True)
    sh.put(f"B{r}", "Value", bold=True)
    sh.put(f"C{r}", "As of", bold=True)
    sh.put(f"D{r}", "Source", bold=True)
    r += 1
    trend = {t["period"]: t for t in syf["earningsTrend"]["trend"]}
    rows = [
        ("price", "SYF share price ($, last completed daily close)", price, "usd", price_date, "Yahoo Finance chart API daily close, symbol SYF; data retrieved " + market["retrieved_date"] + " (last completed session before retrieval)"),
        ("shares", "Common shares outstanding (mm), 10-Q cover", hist["misc"]["cover_shares_outstanding_m"], "num1", dt.date.fromisoformat(hist["misc"]["cover_shares_date"]), "SEC XBRL dei:EntityCommonStockSharesOutstanding, 10-Q Q2 2026 (accession " + hist["misc"]["cover_shares_accn"] + ")"),
        ("beta", "Beta (Yahoo, 5-year monthly)", raw(syf["summaryDetail"], "beta"), "0.00", dt.date.fromisoformat(market["retrieved_date"]), "Yahoo Finance quoteSummary summaryDetail.beta"),
        ("rf", "Risk-free rate: 10-year US Treasury yield", market["risk_free"]["yield_pct"] / 100, "pct2", dt.date.fromisoformat(market["retrieved_date"]), "Yahoo Finance chart API, symbol ^TNX (percent), retrieved " + market["retrieved_date"]),
        ("cons26", "Consensus diluted EPS FY2026E ($)", raw(trend["0y"]["earningsEstimate"], "avg"), "usd", dt.date.fromisoformat(market["retrieved_date"]), f"Yahoo Finance earningsTrend 0y, {raw(trend['0y']['earningsEstimate'], 'numberOfAnalysts')} analysts"),
        ("cons27", "Consensus diluted EPS FY2027E ($)", raw(trend["+1y"]["earningsEstimate"], "avg"), "usd", dt.date.fromisoformat(market["retrieved_date"]), f"Yahoo Finance earningsTrend +1y, {raw(trend['+1y']['earningsEstimate'], 'numberOfAnalysts')} analysts"),
        ("ttm_eps", "Trailing-twelve-month diluted EPS per Yahoo ($)", raw(syf["defaultKeyStatistics"], "trailingEps"), "usd", dt.date.fromisoformat(market["retrieved_date"]), "Yahoo Finance defaultKeyStatistics.trailingEps (cross-check to Historical LTM EPS)"),
    ]
    for key, label, val, fmt, asof, src in rows:
        I[key] = f"Inputs_Sources!$B${r}"
        sh.put(f"A{r}", label)
        sh.put(f"B{r}", val, fmt)
        sh.put(f"C{r}", asof, "date")
        sh.put(f"D{r}", src)
        r += 1
    sh.put(f"A{r}", "Balance sheet date used as valuation base")
    sh.put(f"B{r}", BALANCE_DATE, "date")
    I["bal_date"] = f"Inputs_Sources!$B${r}"
    sh.put(f"D{r}", "Latest reported balance sheet: Form 10-Q for the quarter ended 2026-06-30")
    r += 2
    sh.put(f"A{r}", "B. Valuation and rating assumptions (author's inputs)", bold=True)
    sh.put(f"B{r}", "Value", bold=True)
    sh.put(f"D{r}", "Basis", bold=True)
    r += 1
    for key, label, val, fmt, note in VAL_ASSUMPTIONS:
        I[key] = f"Inputs_Sources!$B${r}"
        sh.put(f"A{r}", label)
        sh.put(f"B{r}", val, fmt)
        sh.put(f"D{r}", note)
        r += 1
    sh.put(f"A{r}", "Cost of equity (CAPM: risk-free + beta x equity risk premium)")
    sh.put(f"B{r}", f"={I['rf'].split('!')[1]}+{I['beta'].split('!')[1]}*{I['erp'].split('!')[1]}", "pct2")
    I["ke"] = f"Inputs_Sources!$B${r}"
    r += 2
    I["_next_row"] = r
    return sh, I


def source_register(sh, I, hist, market, peers):
    r = I["_next_row"]
    sh.put(f"A{r}", "C. Source register for historical inputs (Historical sheet rows)", bold=True)
    r += 1
    for c, t in zip("ABCD", ["Item (Historical row)", "Periods", "Filing", "Table or XBRL tag and cross-check"]):
        sh.put(f"{c}{r}", t, bold=True)
    r += 1
    cks = hist["xbrl_crosscheck"]
    def ck(field):
        xs = [x for x in cks if x["field"] == field]
        return f"XBRL {xs[0]['tag']}: {sum(1 for x in xs if x['match'])} of {len(xs)} fiscal years match" if xs else "no XBRL cross-check"
    reg = [
        ("Loan receivables, average loans, purchase volume, average active accounts, NIM, NCO, ACL, delinquency, efficiency, ROE", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "MD&A, Other Financial and Statistical Data table; " + ck("acl_end").replace("XBRL", "ACL")),
        ("Average interest-earning assets, interest-bearing liabilities, yields", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "MD&A, Average Balance Sheet table"),
        ("Interest and fees on loans, total interest income, interest expense, NII, RSA, provision, other income, other expense, pretax, tax, net earnings, EPS", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "Consolidated Statements of Earnings; " + ck("net_interest_income") + "; " + ck("net_earnings") + "; " + ck("eps_diluted")),
        ("Diluted weighted average shares", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "Note: Earnings per share, weighted average common shares outstanding, dilutive"),
        ("Total equity, preferred stock, goodwill, intangibles, shares outstanding", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "Consolidated Statements of Financial Position; " + ck("total_equity")),
        ("CET1 capital, ratio, risk-weighted assets", "FY2020-FY2025", "10-K FY2025, FY2023, FY2021", "MD&A, Capital, Basel III table for Synchrony Financial (ratios as originally reported)"),
        ("Dividends per share, common dividends, share repurchases (cash)", "FY2020-FY2025", "10-K FY2025 (FY2023 to FY2025), earlier 10-Ks", "XBRL CommonStockDividendsPerShareDeclared, DividendsCommonStockCash, PaymentsForRepurchaseOfCommonStock"),
        ("Quarterly income statement, credit, balance sheet and capital data", "Q1 2025, Q2 2025, Q1 2026, Q2 2026", "10-Q Q1 2026, 10-Q Q2 2026", "Same tables as the 10-K; three-month columns; prior-year quarters are the comparative columns"),
        ("Quarterly dividends and share repurchases", "Q1 2025 to Q2 2026", "10-K FY2025, 10-Q Q2 2026", "MD&A, Capital, Dividend and Share Repurchases tables"),
        ("Recast intangibles, CET1 and RWA at 2025-12-31", "Dec 2025", "10-Q Q2 2026", "Note 2 (internal-use capitalized software reclassified out of Intangible assets to Other assets); Capital table footnote (1)"),
        ("Series C preferred issuance ($500mm, 7.250%) and dividend increase to $0.34 quarterly", "Q2 2026", "10-Q Q2 2026", "MD&A, Dividend and Share Repurchases"),
        ("Full filing-text versus XBRL cross-check", "FY2020-FY2025", "SEC companyfacts API", f"{sum(1 for x in cks if x['match'])} of {len(cks)} checks match within $0.6mm or $0.005 per share (see python/pull_data.py)"),
    ]
    for it in reg:
        for c, t in zip("ABCD", it):
            sh.put(f"{c}{r}", t, wrap=(c in "AD"))
        r += 1
    r += 1
    sh.put(f"A{r}", "D. Filings used", bold=True)
    r += 1
    for c, t in zip("ABCD", ["Form and period", "Accession", "Document", "URL"]):
        sh.put(f"{c}{r}", t, bold=True)
    r += 1
    for f in hist["filings"]:
        sh.put(f"A{r}", f"{f['form']} {f['period']}")
        sh.put(f"B{r}", f["accession"])
        sh.put(f"C{r}", f["document"])
        sh.put(f"D{r}", f["url"])
        r += 1
    r += 1
    sh.put(f"A{r}", "E. Definitions, simplifications and known limitations", bold=True)
    r += 1
    notes = [
        "Annual forecast granularity (FY2026E to FY2028E). FY2026E is anchored to reported H1 2026 for share count, TCE and buybacks; income statement lines are full-year driver-based.",
        "No full balance sheet is forecast. Funding is implied through cost of funds on interest-bearing liabilities set as a % of interest-earning assets. Loans, allowance, TCE, CET1 and shares are forecast.",
        "TCE = total equity less preferred stock, goodwill and intangible assets. In the Q2 2026 10-Q the company reclassified internal-use software ($1,151mm at Dec 2025) out of intangibles, which lifts TCE and CET1 versus earlier periods. History uses the as-originally-reported basis; the forecast uses the recast basis.",
        "Reserve build: provision = net charge-offs + change in allowance (other ACL movements assumed zero). History shows the small plug for portfolio sales, acquisitions and off-balance-sheet reserves.",
        "RSA is modelled from the FY2020-FY2025 regression of RSA % of average loans on the net charge-off rate, plus a structural residual (product, pricing and policy changes).",
        "Quarterly ratios use 365/days annualization, matching the company's disclosed NIM and NCO rate.",
        "Comparable-company data: Discover is no longer listed (acquired by Capital One), so Bread Financial is used. Peer tangible book uses the latest XBRL balance sheet per peer (dates on the Comps sheet).",
        "Forecasts, price target and rating are the author's illustrative analysis for portfolio purposes. Not investment advice. No affiliation with any bank or research department.",
    ]
    for n in notes:
        sh.put(f"A{r}", n)
        r += 1
    return sh


# ---------------------------------------------------------------- Scenarios
def build_scenarios(wb):
    sh = Sheet(wb, "Scenarios", {"A": 64, "B": 12, "C": 12, "D": 12, "E": 90})
    S = {"switch": "Scenarios!$B$3", "name": "Scenarios!$B$4"}
    sh.put("A1", "Scenarios and forecast assumptions", bold=True)
    sh.put("A3", "Scenario switch: enter 1 = Bear, 2 = Base, 3 = Bull", bold=True)
    sh.put("B3", DEFAULT_SCENARIO, "0")
    sh.put("A4", "Selected scenario")
    sh.put("B4", '=CHOOSE(B3,"Bear","Base","Bull")')
    r = 6
    sh.put(f"A{r}", "Scenario drivers", bold=True)
    for c, t in zip("BCD", ["FY2026E", "FY2027E", "FY2028E"]):
        sh.put(f"{c}{r}", t, bold=True)
    sh.put(f"E{r}", "Reference points (filings)", bold=True)
    r += 1
    for key, label, fmt, vals, note in SCENARIO_DRIVERS:
        S[key] = {}
        sh.put(f"A{r}", label, bold=True)
        sh.put(f"E{r}", note)
        r += 1
        for scen in ("bear", "base", "bull"):
            S[key][scen] = r
            sh.put(f"A{r}", "   " + scen.capitalize())
            for i, v in enumerate(vals[scen]):
                sh.put(f"{col(2 + i)}{r}", v, fmt)
            r += 1
        S[key]["live"] = r
        sh.put(f"A{r}", "   Live (selected scenario)")
        for i in range(3):
            c = col(2 + i)
            sh.put(f"{c}{r}", f"=CHOOSE($B$3,{c}{S[key]['bear']},{c}{S[key]['base']},{c}{S[key]['bull']})", fmt)
        r += 2
    sh.put(f"A{r}", "Terminal ROTCE by scenario", bold=True)
    sh.put(f"E{r}", ROT_NOTE)
    r += 1
    S["rot"] = {}
    for scen in ("bear", "base", "bull"):
        S["rot"][scen] = r
        sh.put(f"A{r}", "   " + scen.capitalize())
        sh.put(f"B{r}", SCENARIO_ROT[scen], "pct1")
        r += 1
    S["rot"]["live"] = r
    sh.put(f"A{r}", "   Live (selected scenario)")
    sh.put(f"B{r}", f"=CHOOSE($B$3,B{S['rot']['bear']},B{S['rot']['base']},B{S['rot']['bull']})", "pct1")
    r += 2
    sh.put(f"A{r}", "Assumptions common to all scenarios", bold=True)
    for c, t in zip("BCD", ["FY2026E", "FY2027E", "FY2028E"]):
        sh.put(f"{c}{r}", t, bold=True)
    sh.put(f"E{r}", "Reference points (filings)", bold=True)
    r += 1
    for key, label, fmt, vals, note in COMMON_ASSUMPTIONS:
        S["c_" + key] = r
        sh.put(f"A{r}", label)
        for i, v in enumerate(vals):
            sh.put(f"{col(2 + i)}{r}", v, fmt)
        sh.put(f"E{r}", note)
        r += 1
    return sh, S


# ---------------------------------------------------------------- Forecast + Valuation engine
FC = {"prev": {"C": "B", "D": "C", "E": "D"}}
SEC = None


def build_engine(wb, fname, vname, S, I, HR, drv):
    """Write forecast to sheet fname and valuation to vname (may be the same sheet). drv: live|bear|base|bull."""
    same = fname == vname
    fs = Sheet(wb, fname, {"A": 66, "B": 13, "C": 13, "D": 13, "E": 13, "F": 3, "G": 13, "H": 13})
    vs = fs if same else Sheet(wb, vname, {"A": 70, "B": 15, "C": 15, "D": 15, "E": 15})
    Hs = "Historical!"
    hc = lambda key, period: f"{Hs}${HIST_COL[period]}${HR[key]}"
    hL = lambda key: f"{Hs}${COL_H1}${HR[key]}"
    hM = lambda key: f"{Hs}${COL_LTM}${HR[key]}"

    def drvref(key, yi):
        return f"Scenarios!{col(2 + yi)}{S[key][drv]}"

    def comref(key, yi):
        return f"Scenarios!{col(2 + yi)}{S['c_' + key]}"

    # ---- row plan
    order = [
        ("H", "Drivers"),
        ("d_loan_g", "Loan receivables growth, year-end", "pct1"), ("d_pv_g", "Purchase volume growth", "pct1"),
        ("d_yield", "Loan yield", "pct2"), ("d_cof", "Cost of funds", "pct2"), ("d_nco", "Net charge-off rate", "pct2"),
        ("d_cov", "Allowance coverage ratio (reserve build under CECL)", "pct2"), ("d_opex_g", "Other expense growth", "pct1"),
        ("d_bb", "Share repurchases ($mm)", "num"), ("d_avgf", "Average loans / simple average of opening and closing loans", "x3"),
        ("d_oth_iea", "Other interest-earning assets as % of average loans", "pct1"), ("d_ibl", "Interest-bearing liabilities as % of average interest-earning assets", "pct1"),
        ("d_oth_y", "Yield on other interest-earning assets", "pct2"), ("d_rsa_res", "RSA structural residual (% of average loans)", "pct2"),
        ("d_oi", "Other income as % of average loans", "pct2"), ("d_tax", "Effective tax rate", "pct1"), ("d_pref", "Preferred dividends ($mm)", "num"),
        ("d_dps", "Common dividend per share ($)", "usd"), ("d_px", "Assumed average repurchase price ($; H2 for FY2026E)", "usd"),
        ("d_iss", "Net share issuance (mm; H2 only for FY2026E)", "num1"), ("d_dil", "Dilution added to average basic shares (mm)", "num1"),
        ("d_rwa", "Risk-weighted assets / year-end loans", "pct1"),
        ("H", "Balances and volumes ($mm)"),
        ("loans_end", "Loan receivables, year end", "num"), ("avg_loans", "Average loan receivables", "num"), ("avg_iea", "Average interest-earning assets", "num"),
        ("oth_iea", "  of which other interest-earning assets", "num"), ("avg_ibl", "Average interest-bearing liabilities", "num"),
        ("pv", "Purchase volume", "num"), ("pv_int", "Purchase volume / average loans", "x2"),
        ("H", "Income statement ($mm)"),
        ("int_loans", "Interest and fees on loans", "num"), ("int_oth", "Interest on cash and securities", "num"), ("int_total", "Total interest income", "num"),
        ("int_exp", "Interest expense", "num"), ("nii", "Net interest income", "num"), ("nim", "Net interest margin", "pct2"),
        ("rsa_pct", "RSA as % of average loans (regression-fitted plus residual)", "pct2"), ("rsa", "Retailer share arrangements", "num"),
        ("prov", "Provision for credit losses", "num"), ("oi", "Other income", "num"), ("oe", "Other expense", "num"),
        ("eff", "Efficiency ratio", "pct1"), ("pretax", "Earnings before income taxes", "num"), ("tax", "Provision for income taxes", "num"),
        ("ni", "Net earnings", "num"), ("pref", "Preferred dividends", "num"), ("nic", "Net earnings available to common", "num"),
        ("H", "Allowance for credit losses roll-forward ($mm)"),
        ("acl_beg", "Beginning allowance", "num"), ("acl_prov", "Plus: provision for credit losses", "num"), ("acl_nco", "Less: net charge-offs", "num"),
        ("acl_oth", "Other movements", "num"), ("acl_end", "Ending allowance", "num"), ("res_b", "Reserve build (ending less beginning)", "num"),
        ("H", "Shares and capital return"),
        ("sh_beg", "Common shares, beginning (mm)", "num1"), ("sh_bb", "Less: shares repurchased (mm)", "num1"), ("sh_iss", "Plus: net issuance (mm)", "num1"),
        ("sh_end", "Common shares, ending (mm)", "num1"), ("sh_avg", "Average basic shares (mm)", "num1"), ("sh_dil", "Diluted shares (mm)", "num1"),
        ("eps", "Diluted EPS ($)", "usd"), ("dps", "Dividends per share ($)", "usd"), ("div", "Common dividends ($mm)", "num"),
        ("bb", "Share repurchases ($mm)", "num"), ("payout", "Total payout (dividends + buybacks) / net earnings to common", "pct0"),
        ("H", "Tangible common equity and capital ($mm)"),
        ("tce_beg", "TCE, beginning", "num"), ("tce_nic", "Plus: net earnings to common", "num"), ("tce_div", "Less: common dividends", "num"),
        ("tce_bb", "Less: share repurchases", "num"), ("tce_oth", "Plus: other movements", "num"), ("tce_rec", "Plus: reconciliation to reported Jun-26 TCE (FY2026E only)", "num"),
        ("tce_end", "TCE, ending", "num"), ("rotce", "ROTCE (net earnings to common / average TCE)", "pct1"), ("tbvps", "Tangible book value per share ($)", "usd"),
        ("cet1_beg", "CET1 capital, beginning", "num"), ("cet1_end", "CET1 capital, ending", "num"), ("rwa", "Risk-weighted assets", "num"), ("cet1r", "CET1 ratio", "pct1"),
    ]
    R = {}
    row = 3
    hdr_rows = []
    for it in order:
        if it[0] == "H":
            row += 1
            hdr_rows.append((row, it[1]))
            row += 1
        else:
            R[it[0]] = row
            row += 1
    fs.put("A1", f"{COMPANY} ({TICKER}): annual forecast, $ in millions unless stated" if not same else f"{COMPANY} ({TICKER}): {drv.capitalize()} scenario forecast and valuation, $ in millions unless stated", bold=True)
    fs.put("A2", "Scenario:")
    fs.put("B2", "=Scenarios!$B$4" if drv == "live" else drv.capitalize())
    for c, t in zip("BCDE", ["FY2025A", "FY2026E", "FY2027E", "FY2028E"]):
        fs.put(f"{c}3", t, bold=True)
    fs.put("G3", "H1 2026A", bold=True)
    fs.put("H3", "H2 2026E", bold=True)
    for rr, t in hdr_rows:
        fs.put(f"A{rr}", t, bold=True)
    for it in order:
        if it[0] != "H":
            fs.put(f"A{R[it[0]]}", it[1])
    FMTS = {it[0]: it[2] for it in order if it[0] != "H"}
    A = lambda key, c: f"{c}{R[key]}"
    F = {}  # (key, col) -> formula

    def w(key, c, formula):
        fs.put(f"{c}{R[key]}", formula, FMTS[key])

    g25 = "FY2025"
    # base year (FY2025A) column B
    base = {
        "d_loan_g": f"={hc('loan_gr', g25)}", "d_pv_g": f"={hc('pv_gr', g25)}", "d_yield": f"={hc('yield', g25)}", "d_cof": f"={hc('cof', g25)}",
        "d_nco": f"={hc('nco_rate', g25)}", "d_cov": f"={hc('cov', g25)}", "d_opex_g": f"={hc('oe_gr', g25)}", "d_bb": f"={hc('bb', g25)}",
        "d_avgf": f"={hc('avgf', g25)}", "d_oth_iea": f"={hc('oth_iea', g25)}", "d_ibl": f"={hc('ibl_pct', g25)}", "d_oth_y": f"={hc('oth_y', g25)}",
        "d_rsa_res": f"={Hs}{HIST_COL[g25]}{HR['fit_res']}", "d_oi": f"={hc('oi_pct', g25)}", "d_tax": f"={hc('taxr', g25)}", "d_pref": f"={hc('nic', g25)}*0+({hc('ni', g25)}-{hc('nic', g25)})",
        "d_dps": f"={hc('dps', g25)}", "d_rwa": f"={Hs}{HIST_COL[g25]}{HR['rwa_r']}/{hc('loans_end', g25)}",
        "loans_end": f"={hc('loans_end', g25)}", "avg_loans": f"={hc('avg_loans', g25)}", "avg_iea": f"={hc('iea', g25)}", "oth_iea": f"={hc('iea', g25)}-{hc('avg_loans', g25)}",
        "avg_ibl": f"={hc('ibl', g25)}", "pv": f"={hc('pv', g25)}", "int_loans": f"={hc('int_loans', g25)}", "int_oth": f"={hc('int_total', g25)}-{hc('int_loans', g25)}",
        "int_total": f"={hc('int_total', g25)}", "int_exp": f"={hc('int_exp', g25)}", "nii": f"={hc('nii', g25)}", "rsa_pct": f"={hc('rsa_pct', g25)}", "rsa": f"={hc('rsa', g25)}",
        "prov": f"={hc('prov', g25)}", "oi": f"={hc('oi', g25)}", "oe": f"={hc('oe', g25)}", "pretax": f"={hc('pretax', g25)}", "tax": f"={hc('tax', g25)}", "ni": f"={hc('ni', g25)}",
        "pref": f"={hc('ni', g25)}-{hc('nic', g25)}", "nic": f"={hc('nic', g25)}", "acl_beg": f"={Hs}{HIST_COL['FY2024']}{HR['acl']}", "acl_prov": f"={hc('prov', g25)}",
        "acl_nco": f"={hc('nco', g25)}", "acl_oth": f"={hc('acl_oth', g25)}", "acl_end": f"={hc('acl', g25)}", "res_b": f"={hc('res_b', g25)}",
        "sh_end": f"={hc('sh', g25)}", "eps": f"={hc('eps', g25)}", "dps": f"={hc('dps', g25)}", "div": f"={hc('div', g25)}", "bb": f"={hc('bb', g25)}",
        "sh_dil": f"={hc('dsh', g25)}", "tce_end": f"={Hs}{HIST_COL[g25]}{HR['tce_r']}", "tbvps": f"={Hs}{HIST_COL[g25]}{HR['tce_r']}/{hc('sh', g25)}",
        "cet1_end": f"={Hs}{HIST_COL[g25]}{HR['cet1a_r']}", "rwa": f"={Hs}{HIST_COL[g25]}{HR['rwa_r']}", "cet1r": f"={Hs}{HIST_COL[g25]}{HR['cet1a_r']}/{Hs}{HIST_COL[g25]}{HR['rwa_r']}",
        "eff": f"={hc('eff', g25)}", "nim": f"={hc('nim', g25)}",
    }
    base["d_pref"] = f"={hc('ni', g25)}-{hc('nic', g25)}"
    for k, v in base.items():
        w(k, "B", v)
    w("pv_int", "B", f"={A('pv', 'B')}/{A('avg_loans', 'B')}")
    w("payout", "B", f"=({A('div', 'B')}+{A('bb', 'B')})/{A('nic', 'B')}")
    w("tce_beg", "B", "n/a")
    w("rotce", "B", f"={hc('rotce', g25)}")
    # forecast columns
    for yi, c in enumerate("CDE"):
        p = FC["prev"][c]
        d = lambda key: A(key, c)
        w("d_loan_g", c, f"={drvref('loan_g', yi)}")
        w("d_pv_g", c, f"={drvref('pv_g', yi)}")
        w("d_yield", c, f"={drvref('yield', yi)}")
        w("d_cof", c, f"={drvref('cof', yi)}")
        w("d_nco", c, f"={drvref('nco', yi)}")
        w("d_cov", c, f"={drvref('cov', yi)}")
        w("d_opex_g", c, f"={drvref('opex_g', yi)}")
        w("d_bb", c, f"={drvref('bb', yi)}")
        for k_src, k_dst in [("avgf", "d_avgf"), ("oth_iea", "d_oth_iea"), ("ibl", "d_ibl"), ("oth_y", "d_oth_y"), ("rsa_res", "d_rsa_res"), ("oi", "d_oi"), ("tax", "d_tax"), ("pref", "d_pref"), ("dps", "d_dps"), ("iss", "d_iss"), ("dil", "d_dil"), ("rwa", "d_rwa")]:
            w(k_dst, c, f"={comref(k_src, yi)}")
        w("d_px", c, f"={I['price']}*(1+{comref('px_g', yi)})" if yi == 0 else f"={A('d_px', p)}*(1+{comref('px_g', yi)})")
        # volumes
        w("loans_end", c, f"={A('loans_end', p)}*(1+{d('d_loan_g')})")
        w("avg_loans", c, f"=AVERAGE({A('loans_end', p)},{d('loans_end')})*{d('d_avgf')}")
        w("oth_iea", c, f"={d('avg_loans')}*{d('d_oth_iea')}")
        w("avg_iea", c, f"={d('avg_loans')}+{d('oth_iea')}")
        w("avg_ibl", c, f"={d('avg_iea')}*{d('d_ibl')}")
        w("pv", c, f"={A('pv', p)}*(1+{d('d_pv_g')})")
        w("pv_int", c, f"={d('pv')}/{d('avg_loans')}")
        # income
        w("int_loans", c, f"={d('avg_loans')}*{d('d_yield')}")
        w("int_oth", c, f"={d('oth_iea')}*{d('d_oth_y')}")
        w("int_total", c, f"={d('int_loans')}+{d('int_oth')}")
        w("int_exp", c, f"={d('avg_ibl')}*{d('d_cof')}")
        w("nii", c, f"={d('int_total')}-{d('int_exp')}")
        w("nim", c, f"={d('nii')}/{d('avg_iea')}")
        w("rsa_pct", c, f"=Historical!$B${HR['intercept']}+Historical!$B${HR['slope']}*{d('d_nco')}+{d('d_rsa_res')}")
        w("rsa", c, f"={d('rsa_pct')}*{d('avg_loans')}")
        w("prov", c, f"={d('acl_nco')}+{d('res_b')}")
        w("oi", c, f"={d('avg_loans')}*{d('d_oi')}")
        w("oe", c, f"={A('oe', p)}*(1+{d('d_opex_g')})")
        w("eff", c, f"={d('oe')}/({d('nii')}-{d('rsa')}+{d('oi')})")
        w("pretax", c, f"={d('nii')}-{d('rsa')}-{d('prov')}+{d('oi')}-{d('oe')}")
        w("tax", c, f"={d('pretax')}*{d('d_tax')}")
        w("ni", c, f"={d('pretax')}-{d('tax')}")
        w("pref", c, f"={d('d_pref')}")
        w("nic", c, f"={d('ni')}-{d('pref')}")
        # allowance
        w("acl_beg", c, f"={A('acl_end', p)}")
        w("acl_nco", c, f"={d('d_nco')}*{d('avg_loans')}")
        w("acl_prov", c, f"={d('prov')}")
        w("acl_oth", c, "=0")
        w("acl_end", c, f"={d('d_cov')}*{d('loans_end')}")
        w("res_b", c, f"={d('acl_end')}-{d('acl_beg')}")
        # shares
        junsh = f"{Hs}${HIST_COL['Q2_2026']}${HR['sh']}"
        # note: cover-page count is used for valuation; the balance-sheet count (rounded) anchors the roll
        w("sh_beg", c, f"={A('sh_end', p)}")
        if yi == 0:
            w("sh_bb", c, f"={hL('bb_sh')}+({d('d_bb')}-{hL('bb')})/{d('d_px')}")
            w("sh_iss", c, f"=({I['shares']}-{d('sh_beg')}+{hL('bb_sh')})+{d('d_iss')}")
            w("sh_end", c, f"={d('sh_beg')}-{d('sh_bb')}+{d('sh_iss')}")
            w("sh_avg", c, f"=AVERAGE({d('sh_beg')},{d('sh_end')})")
            w("sh_dil", c, f"=AVERAGE({hL('dsh')},AVERAGE({I['shares']},{d('sh_end')})+{d('d_dil')})")
            w("div", c, f"={hL('div')}+({d('d_dps')}-{hL('dps')})*AVERAGE({I['shares']},{d('sh_end')})")
        else:
            w("sh_bb", c, f"={d('d_bb')}/{d('d_px')}")
            w("sh_iss", c, f"={d('d_iss')}")
            w("sh_end", c, f"={d('sh_beg')}-{d('sh_bb')}+{d('sh_iss')}")
            w("sh_avg", c, f"=AVERAGE({d('sh_beg')},{d('sh_end')})")
            w("sh_dil", c, f"={d('sh_avg')}+{d('d_dil')}")
            w("div", c, f"={d('d_dps')}*{d('sh_avg')}")
        w("eps", c, f"={d('nic')}/{d('sh_dil')}")
        w("dps", c, f"={d('d_dps')}")
        w("bb", c, f"={d('d_bb')}")
        w("payout", c, f"=({d('div')}+{d('bb')})/{d('nic')}")
        # tce
        w("tce_beg", c, f"={A('tce_end', p)}")
        w("tce_nic", c, f"={d('nic')}")
        w("tce_div", c, f"=-{d('div')}")
        w("tce_bb", c, f"=-{d('bb')}")
        w("tce_oth", c, f"={comref('oth_tce', yi)}")
        if yi == 0:
            w("tce_rec", c, f"=({Hs}${HIST_COL['Q2_2026']}${HR['tce_r']})-({d('tce_beg')}+{hL('nic')}-{hL('div')}-{hL('bb')})")
        else:
            w("tce_rec", c, "=0")
        w("tce_end", c, f"={d('tce_beg')}+{d('tce_nic')}+{d('tce_div')}+{d('tce_bb')}+{d('tce_oth')}+{d('tce_rec')}")
        w("rotce", c, f"={d('nic')}/AVERAGE({d('tce_beg')},{d('tce_end')})")
        w("tbvps", c, f"={d('tce_end')}/{d('sh_end')}")
        w("cet1_beg", c, f"={A('cet1_end', p)}")
        w("cet1_end", c, f"={d('cet1_beg')}+{d('nic')}-{d('div')}-{d('bb')}+{d('tce_oth')}")
        w("rwa", c, f"={d('loans_end')}*{d('d_rwa')}")
        w("cet1r", c, f"={d('cet1_end')}/{d('rwa')}")
    # H1 actual and implied H2 memo columns
    for key, hk in [("pv", "pv"), ("int_loans", "int_loans"), ("int_exp", "int_exp"), ("nii", "nii"), ("rsa", "rsa"), ("prov", "prov"), ("oi", "oi"), ("oe", "oe"), ("pretax", "pretax"), ("tax", "tax"), ("ni", "ni"), ("nic", "nic"), ("eps", "eps"), ("acl_nco", "nco"), ("div", "div"), ("bb", "bb")]:
        fs.put(f"G{R[key]}", f"={hL(hk)}", FMTS[key])
        fs.put(f"H{R[key]}", f"=C{R[key]}-G{R[key]}", FMTS[key])
    fs.put(f"G{R['tce_end']}", f"={Hs}${HIST_COL['Q2_2026']}${HR['tce_r']}", "num")
    fs.put(f"G{R['loans_end']}", f"={Hs}${HIST_COL['Q2_2026']}${HR['loans_end']}", "num")
    fs.put(f"G{R['acl_end']}", f"={Hs}${HIST_COL['Q2_2026']}${HR['acl']}", "num")
    fs.put(f"G{R['sh_end']}", f"={I['shares']}", "num1")
    fs.put(f"G{R['cet1r']}", f"={Hs}${HIST_COL['Q2_2026']}${HR['cet1r']}", "pct1")
    fs.put(f"G{R['eff']}", f"=G{R['oe']}/(G{R['nii']}-G{R['rsa']}+G{R['oi']})", "pct1")
    fs.put(f"H{R['eff']}", f"=H{R['oe']}/(H{R['nii']}-H{R['rsa']}+H{R['oi']})", "pct1")
    fs.put(f"A{max(R.values()) + 2}", "Notes: H1 2026A is reported; H2 2026E is FY2026E less H1 2026A. FY2025A ROTCE is on the as-originally-reported TCE basis; forecast ROTCE uses the recast basis. Loans, allowance, TCE, CET1 and shares tie to their roll-forwards on the Checks sheet.")
    FR = R
    fpre = "" if same else q(fname)

    # ================= valuation
    if same:
        vrow = max(R.values()) + 5
    else:
        vrow = 3
    V = {}

    def vput(key, label, formula, fmt, c="B", bold=False):
        nonlocal vrow
        V[key] = vrow
        vs.put(f"A{vrow}", label, bold=bold)
        if formula is not None:
            vs.put(f"{c}{vrow}", formula, fmt)
        vrow += 1

    if same:
        vs.put(f"A{vrow - 2}", f"Valuation ({drv} scenario)", bold=True)
    else:
        vs.put("A1", f"{COMPANY} ({TICKER}): valuation, price target and rating (live scenario)", bold=True)
        vs.put("A2", "Scenario:")
        vs.put("B2", "=Scenarios!$B$4")
    fr = lambda key, c: f"{fpre}{c}{FR[key]}"
    vs.put(f"A{vrow}", "Cost of equity and terminal assumptions", bold=True)
    vrow += 1
    vput("price", "Share price ($)", f"={I['price']}", "usd")
    vput("rf", "Risk-free rate", f"={I['rf']}", "pct2")
    vput("beta", "Beta", f"={I['beta']}", "0.00")
    vput("erp", "Equity risk premium", f"={I['erp']}", "pct1")
    vput("ke", "Cost of equity (CAPM)", f"={I['ke']}", "pct2")
    vput("g", "Terminal growth", f"={I['g']}", "pct1")
    vput("rot", "Terminal ROTCE (scenario-dependent)", f"=Scenarios!$B${S['rot'][drv]}", "pct1")
    vput("tau", "Years from balance sheet date to price date", f"=(Inputs_Sources!$C${int(I['price'].split('$')[-1])}-{I['bal_date']})/365", "0.000")
    vrow += 1
    vs.put(f"A{vrow}", "Residual income model (equity value of the common franchise, base = tangible common equity)", bold=True)
    for c, t in zip("BCD", ["H2 2026E", "FY2027E", "FY2028E"]):
        vs.put(f"{c}{vrow}", t, bold=True)
    vrow += 1
    vput("len", "Period length (years)", None, "0.0")
    vs.put(f"B{V['len']}", 0.5, "0.0")
    vs.put(f"C{V['len']}", 1, "0.0")
    vs.put(f"D{V['len']}", 1, "0.0")
    vput("t", "Time from balance sheet date to end of period (years)", None, "0.0")
    vs.put(f"B{V['t']}", f"=B{V['len']}", "0.0")
    vs.put(f"C{V['t']}", f"=B{V['t']}+C{V['len']}", "0.0")
    vs.put(f"D{V['t']}", f"=C{V['t']}+D{V['len']}", "0.0")
    vput("nic", "Net earnings to common", None, "num")
    vs.put(f"B{V['nic']}", f"={fr('nic', 'C')}-{fr('nic', 'G')}", "num")
    vs.put(f"C{V['nic']}", f"={fr('nic', 'D')}", "num")
    vs.put(f"D{V['nic']}", f"={fr('nic', 'E')}", "num")
    vput("tceb", "TCE, beginning of period", None, "num")
    vs.put(f"B{V['tceb']}", f"={fr('tce_end', 'G')}", "num")
    vs.put(f"C{V['tceb']}", f"={fr('tce_end', 'C')}", "num")
    vs.put(f"D{V['tceb']}", f"={fr('tce_end', 'D')}", "num")
    vput("chg", "Equity charge (cost of equity x beginning TCE x period length)", None, "num")
    vput("ri", "Residual income", None, "num")
    vput("df", "Discount factor", None, "0.0000")
    vput("pv", "Present value of residual income", None, "num")
    for c in "BCD":
        vs.put(f"{c}{V['chg']}", f"=$B${V['ke']}*{c}{V['tceb']}*{c}{V['len']}", "num")
        vs.put(f"{c}{V['ri']}", f"={c}{V['nic']}-{c}{V['chg']}", "num")
        vs.put(f"{c}{V['df']}", f"=1/(1+$B${V['ke']})^{c}{V['t']}", "0.0000")
        vs.put(f"{c}{V['pv']}", f"={c}{V['ri']}*{c}{V['df']}", "num")
    vrow += 1
    vput("tce0", "TCE at balance sheet date ($mm)", f"={fr('tce_end', 'G')}", "num")
    vput("sum_pv", "Sum of PV of residual income ($mm)", f"=SUM(B{V['pv']}:D{V['pv']})", "num")
    vput("ptbv_j", "Justified P/TBV at terminal: (ROTCE - g) / (ke - g)", f"=(B{V['rot']}-B{V['g']})/(B{V['ke']}-B{V['g']})", "x2")
    vput("ptbv_j27", "Memo: P/TBV consistent with FY2027E ROTCE if sustained: (ROTCE - g) / (ke - g)", f"=({fr('rotce', 'D')}-B{V['g']})/(B{V['ke']}-B{V['g']})", "x2")
    vput("tv", "Terminal value at end of FY2028E ($mm) = ending TCE x justified P/TBV", f"={fr('tce_end', 'E')}*B{V['ptbv_j']}", "num")
    vput("pv_tv", "PV of terminal value ($mm)", f"=B{V['tv']}*D{V['df']}", "num")
    vput("eqv", "Equity value at balance sheet date ($mm)", f"=B{V['tce0']}+B{V['sum_pv']}+B{V['pv_tv']}", "num")
    vput("v0", "Value per share at balance sheet date ($)", f"=B{V['eqv']}/{I['shares']}", "usd")
    vput("vpx", "Value per share carried to price date, less dividends paid ($)", f"=B{V['v0']}*(1+B{V['ke']})^B{V['tau']}-{I['dps_int']}", "usd")
    vput("v12", "Residual income value 12 months forward, less dividends ($)", f"=B{V['vpx']}*(1+B{V['ke']})-{I['dps_fwd']}", "usd")
    vrow += 1
    vs.put(f"A{vrow}", "Justified P/TBV method", bold=True)
    vrow += 1
    vput("w27", "Weight of FY2027E in horizon tangible book (price date + 12 months)", f"=(Inputs_Sources!$C${int(I['price'].split('$')[-1])}+365-DATE(2026,12,31))/365", "0.000")
    vput("tbv_h", "Tangible book value per share at horizon ($)", f"=(1-B{V['w27']})*{fr('tbvps', 'C')}+B{V['w27']}*{fr('tbvps', 'D')}", "usd")
    vput("v_pt", "Justified P/TBV x horizon tangible book ($)", f"=B{V['ptbv_j']}*B{V['tbv_h']}", "usd")
    vrow += 1
    vs.put(f"A{vrow}", "Price target and rating", bold=True)
    vrow += 1
    vput("pt_raw", "Blended value (weights on Inputs_Sources) ($)", f"={I['w_ri']}*B{V['v12']}+{I['w_pt']}*B{V['v_pt']}", "usd")
    vput("pt", "12-month price target ($, rounded to nearest dollar)", f"=ROUND(B{V['pt_raw']},0)", "usd", bold=True)
    vput("upside", "Upside to price target", f"=B{V['pt']}/B{V['price']}-1", "pct1")
    vput("dy", "Forward dividend yield", f"={I['dps_fwd']}/B{V['price']}", "pct1")
    vput("tr", "Expected total return", f"=B{V['upside']}+B{V['dy']}", "pct1")
    vput("rating", "Rating (Overweight / Neutral / Underweight per rule)", f'=IF(B{V["tr"]}>={I["ow"]},"Overweight",IF(B{V["tr"]}<={I["uw"]},"Underweight","Neutral"))', None, bold=True)
    vput("px_ow", "Share price at which expected total return equals the Overweight threshold ($)", f"=(B{V['pt']}+{I['dps_fwd']})/(1+{I['ow']})", "usd")
    vput("px_uw", "Share price at which expected total return equals the Underweight threshold ($)", f"=(B{V['pt']}+{I['dps_fwd']})/(1+{I['uw']})", "usd")
    vput("pe27", "Implied P/E at price target on FY2027E EPS", f"=B{V['pt']}/{fr('eps', 'D')}", "x2")
    vput("ptbv_pt", "Implied P/TBV at price target on Jun-26 tangible book", f"=B{V['pt']}/(B{V['tce0']}/{I['shares']})", "x2")
    vput("pe_now", "Current P/E on FY2027E EPS", f"=B{V['price']}/{fr('eps', 'D')}", "x2")
    vput("ptbv_now", "Current P/TBV on Jun-26 tangible book", f"=B{V['price']}/(B{V['tce0']}/{I['shares']})", "x2")
    return fs, vs, FR, V


# ---------------------------------------------------------------- Comps
def latest_instant(peer_json, tag):
    best = None
    for unit, rows in peer_json["us-gaap"].get(tag, {}).get("units", {}).items():
        for r in rows:
            if "start" in r:
                continue
            if best is None or r["end"] > best["end"] or (r["end"] == best["end"] and r["filed"] > best["filed"]):
                best = r
    return best


def build_comps(wb, hist, market, peers, HR, I):
    sh = Sheet(wb, "Comps", {"A": 8, "B": 20})
    ws = sh.ws
    K = COMPS_COL
    for c in range(3, 30):
        ws.column_dimensions[col(c)].width = 13
    sh.put("A1", "Comparable companies: consumer and card lenders", bold=True)
    sh.put("A2", f"Prices are last completed daily closes; TTM EPS and consensus EPS as of the Yahoo Finance retrieval on {market['retrieved_date']}. Tangible book: latest XBRL balance sheet per company (date shown). Discover is excluded (acquired by Capital One); Bread Financial is used.")
    r0 = 4
    for key, (letter, head) in COMPS_HEADS.items():
        sh.put(f"{letter}{r0}", head, bold=True, wrap=True)
    ws.row_dimensions[r0].height = 45
    CR = {}
    r = r0 + 1
    for t in [TICKER] + PEERS:
        d = market["tickers"][t]
        CR[t] = r
        trend = {x["period"]: x for x in d["earningsTrend"]["trend"]}
        pdate, px = close_before(t, dt.date.fromisoformat(market["retrieved_date"]))
        sh.put(f"{K['ticker']}{r}", t)
        sh.put(f"{K['name']}{r}", PEER_NAMES[t])
        sh.put(f"{K['price']}{r}", px, "usd")
        sh.put(f"{K['eps_ttm']}{r}", raw(d["defaultKeyStatistics"], "trailingEps"), "usd")
        sh.put(f"{K['eps26']}{r}", raw(trend["0y"]["earningsEstimate"], "avg"), "usd")
        sh.put(f"{K['eps27']}{r}", raw(trend["+1y"]["earningsEstimate"], "avg"), "usd")
        sh.put(f"{K['roe']}{r}", raw(d["financialData"], "returnOnEquity"), "pct1")
        sh.put(f"{K['pdate']}{r}", pdate, "date")
        if t == TICKER:
            sh.put(f"{K['shares']}{r}", f"={I['shares']}", "num1")
            k = HIST_COL["Q2_2026"]
            sh.put(f"{K['equity']}{r}", f"=Historical!{k}{HR['equity']}", "num")
            sh.put(f"{K['pref']}{r}", f"=Historical!{k}{HR['pref']}", "num")
            sh.put(f"{K['gwi']}{r}", f"=Historical!{k}{HR['gw']}+Historical!{k}{HR['intang']}", "num")
            sh.put(f"{K['bsdate']}{r}", BALANCE_DATE, "date")
            sh.put(f"{K['eps_ttm']}{r}", f"=Historical!{COL_LTM}{HR['eps']}", "usd")
        else:
            pj = peers[t]
            tags = PEER_TAGS[t]
            eq = latest_instant(pj, tags["equity"])
            pf = latest_instant(pj, tags["pref"]) if tags["pref"] else None
            gi = [latest_instant(pj, tg) for tg in tags["intang"]]
            sh_out = max(pj["dei"]["EntityCommonStockSharesOutstanding"]["units"]["shares"], key=lambda z: z["end"])
            sh.put(f"{K['shares']}{r}", sh_out["val"] / 1e6, "num1")
            sh.put(f"{K['equity']}{r}", eq["val"] / 1e6, "num")
            sh.put(f"{K['pref']}{r}", (pf["val"] / 1e6) if pf else 0, "num")
            sh.put(f"{K['gwi']}{r}", sum(g["val"] for g in gi) / 1e6, "num")
            sh.put(f"{K['bsdate']}{r}", dt.date.fromisoformat(eq["end"]), "date")
        c = lambda key: f"{K[key]}{r}"
        sh.put(c("mcap"), f"={c('price')}*{c('shares')}", "num")
        sh.put(c("pe_ttm"), f"={c('price')}/{c('eps_ttm')}", "x2")
        sh.put(c("pe26"), f"={c('price')}/{c('eps26')}", "x2")
        sh.put(c("pe27"), f"={c('price')}/{c('eps27')}", "x2")
        sh.put(c("tce"), f"={c('equity')}-{c('pref')}-{c('gwi')}", "num")
        sh.put(c("tbvps"), f"={c('tce')}/{c('shares')}", "usd")
        sh.put(c("ptbv"), f"={c('price')}/{c('tbvps')}", "x2")
        sh.put(c("rotce"), f"={c('eps_ttm')}/{c('tbvps')}", "pct1")
        r += 1
    first_peer, last_peer = CR[PEERS[0]], CR[PEERS[-1]]
    stat_keys = ["pe_ttm", "pe26", "pe27", "ptbv", "rotce"]
    r += 1
    CR["median"] = r
    sh.put(f"{K['name']}{r}", "Peer median (ex-SYF)", bold=True)
    for key in stat_keys:
        sh.put(f"{K[key]}{r}", f"=MEDIAN({K[key]}{first_peer}:{K[key]}{last_peer})", "pct1" if key == "rotce" else "x2")
    r += 1
    CR["mean"] = r
    sh.put(f"{K['name']}{r}", "Peer mean (ex-SYF)", bold=True)
    for key in stat_keys:
        sh.put(f"{K[key]}{r}", f"=AVERAGE({K[key]}{first_peer}:{K[key]}{last_peer})", "pct1" if key == "rotce" else "x2")
    r += 1
    CR["prem"] = r
    sh.put(f"{K['name']}{r}", "SYF premium / (discount) to peer median", bold=True)
    for key in ["pe_ttm", "pe26", "pe27", "ptbv"]:
        sh.put(f"{K[key]}{r}", f"={K[key]}{CR[TICKER]}/{K[key]}{CR['median']}-1", "pct0")
    r += 2
    sh.put(f"{K['ticker']}{r}", "Cross-checks on SYF (information only, not in the price target)", bold=True)
    r += 1
    CR["reg_s"], CR["reg_i"], CR["reg_c"], CR["reg_imp"], CR["pe_imp"] = r, r + 1, r + 2, r + 3, r + 4
    VC = COMPS_VALUE_COL
    y_rng = f"{K['ptbv']}{CR[TICKER]}:{K['ptbv']}{last_peer}"
    x_rng = f"{K['rotce']}{CR[TICKER]}:{K['rotce']}{last_peer}"
    sh.put(f"{K['name']}{r}", "Regression P/TBV on ROTCE, peers plus SYF: slope")
    sh.put(f"{VC}{r}", f"=SLOPE({y_rng},{x_rng})", "0.000")
    sh.put(f"{K['name']}{r + 1}", "Regression intercept")
    sh.put(f"{VC}{r + 1}", f"=INTERCEPT({y_rng},{x_rng})", "0.000")
    sh.put(f"{K['name']}{r + 2}", "Correlation")
    sh.put(f"{VC}{r + 2}", f"=CORREL({y_rng},{x_rng})", "0.000")
    sh.put(f"{K['name']}{r + 3}", "SYF P/TBV implied by regression at its ROTCE")
    sh.put(f"{VC}{r + 3}", f"={VC}{r + 1}+{VC}{r}*{K['rotce']}{CR[TICKER]}", "x2")
    sh.put(f"{K['name']}{r + 4}", "SYF value per share at peer median P/E on its own 2027E consensus EPS ($)")
    sh.put(f"{VC}{r + 4}", f"={K['pe27']}{CR['median']}*{K['eps27']}{CR[TICKER]}", "usd")
    r += 6
    sh.put(f"{K['ticker']}{r}", "Peer balance sheet sources (XBRL tags, latest instant fact)", bold=True)
    r += 1
    for t in PEERS:
        pj = peers[t]
        tags = PEER_TAGS[t]
        eq = latest_instant(pj, tags["equity"])
        gi = [latest_instant(pj, tg) for tg in tags["intang"]]
        pf = latest_instant(pj, tags["pref"]) if tags["pref"] else None
        parts = [f"equity {tags['equity']} ({eq['end']}, {eq['accn']})"] + [f"{tg} ({g['end']}, {g['accn']})" for tg, g in zip(tags["intang"], gi)]
        if pf:
            parts.append(f"preferred {tags['pref']} ({pf['end']}, {pf['accn']})")
        sh.put(f"{K['ticker']}{r}", t)
        sh.put(f"{K['name']}{r}", "; ".join(parts))
        r += 1
    sh.put(f"{K['ticker']}{r + 1}", "ROTCE here is TTM diluted EPS divided by period-end tangible book value per share, a simple consistent proxy across companies. SYF uses its own reported-basis LTM EPS and the recast Jun-26 tangible book.")
    return sh, CR


# ---------------------------------------------------------------- Sensitivity
def build_sensitivity(wb, FR, V, I, fname, vname, S):
    sh = Sheet(wb, "Sensitivity", {"A": 46})
    for c in "BCDEFGH":
        sh.ws.column_dimensions[c].width = 12
    vp = q(vname)
    fp = q(fname)
    sh.put("A1", "Sensitivity tables (live scenario, all cells are formulas)", bold=True)
    ke0, g0, rot0 = f"{vp}$B${V['ke']}", f"{vp}$B${V['g']}", f"{vp}$B${V['rot']}"
    nic, tceb, ln, tt = (f"{vp}$B${V[k]}:$D${V[k]}" for k in ("nic", "tceb", "len", "t"))
    tce0, tce_e, shares = f"{vp}$B${V['tce0']}", f"{fp}$E${FR['tce_end']}", I["shares"]
    tau, dpsi, dpsf = f"{vp}$B${V['tau']}", I["dps_int"], I["dps_fwd"]
    t3 = f"{vp}$D${V['t']}"
    tbv_h = f"{vp}$B${V['tbv_h']}"

    def ri12(ke, g, rot):
        v0 = f"(({tce0}+SUMPRODUCT(({nic}-{ke}*{tceb}*{ln})/(1+{ke})^{tt})+{tce_e}*(({rot}-{g})/({ke}-{g}))/(1+{ke})^{t3})/{shares})"
        return f"({v0}*(1+{ke})^{tau}-{dpsi})*(1+{ke})-{dpsf}"

    # Table 1: RI value 12m vs ke and g
    r = 3
    sh.put(f"A{r}", "Table 1: residual income value per share, 12 months forward ($), cost of equity (down) vs terminal growth (across)", bold=True)
    steps_ke = [-0.02, -0.01, 0, 0.01, 0.02]
    steps_g = [-0.01, -0.005, 0, 0.005, 0.01]
    r += 1
    T1 = r
    for j, sg in enumerate(steps_g):
        sh.put(f"{col(2 + j)}{r}", f"={g0}+({sg})", "pct1", bold=True)
    for i, sk in enumerate(steps_ke):
        rr = r + 1 + i
        sh.put(f"A{rr}", f"={ke0}+({sk})", "pct2", bold=True)
        for j in range(5):
            c = col(2 + j)
            sh.put(f"{c}{rr}", "=" + ri12(f"$A{rr}", f"{c}${r}", rot0), "usd")
    T1c = f"D{T1 + 3}"
    # Table 2: blended PT vs ke and terminal ROTCE
    r = T1 + 8
    sh.put(f"A{r}", "Table 2: blended price target before rounding ($), cost of equity (down) vs terminal ROTCE (across)", bold=True)
    steps_rot = [-0.04, -0.02, 0, 0.02, 0.04]
    r += 1
    T2 = r
    for j, sr in enumerate(steps_rot):
        sh.put(f"{col(2 + j)}{r}", f"={rot0}+({sr})", "pct1", bold=True)
    for i, sk in enumerate(steps_ke):
        rr = r + 1 + i
        sh.put(f"A{rr}", f"={ke0}+({sk})", "pct2", bold=True)
        for j in range(5):
            c = col(2 + j)
            rotc = f"{c}${r}"
            kk = f"$A{rr}"
            pt = f"{I['w_ri']}*({ri12(kk, g0, rotc)})+{I['w_pt']}*(({rotc}-{g0})/({kk}-{g0}))*{tbv_h}"
            sh.put(f"{c}{rr}", "=" + pt, "usd")
    T2c = f"D{T2 + 3}"
    # Table 3: FY2027E EPS vs NCO rate and loan yield
    r = T2 + 8
    sh.put(f"A{r}", "Table 3: FY2027E diluted EPS ($), net charge-off rate (down) vs loan yield (across). RSA re-sets with credit per the regression; reserve coverage and volumes held.", bold=True)
    r += 1
    T3 = r
    steps_n = [-0.015, -0.0075, 0, 0.0075, 0.015]
    steps_y = [-0.01, -0.005, 0, 0.005, 0.01]
    n0, y0 = f"{fp}$D${FR['d_nco']}", f"{fp}$D${FR['d_yield']}"
    for j, sy in enumerate(steps_y):
        sh.put(f"{col(2 + j)}{r}", f"={y0}+({sy})", "pct2", bold=True)
    for i, sn in enumerate(steps_n):
        rr = r + 1 + i
        sh.put(f"A{rr}", f"={n0}+({sn})", "pct2", bold=True)
        for j in range(5):
            c = col(2 + j)
            slope = f"Historical!$B${HRS['slope']}"
            f = (f"={fp}$D${FR['eps']}+(({c}${r}-{y0})*{fp}$D${FR['avg_loans']}-($A{rr}-{n0})*{fp}$D${FR['avg_loans']}*(1+{slope}))"
                 f"*(1-{fp}$D${FR['d_tax']})/{fp}$D${FR['sh_dil']}")
            sh.put(f"{c}{rr}", f, "usd")
    T3c = f"D{T3 + 3}"
    r = T3 + 8
    sh.put(f"A{r}", "Table 4: FY2027E diluted EPS rules of thumb (live scenario, one variable at a time, $ per share)", bold=True)
    slope = f"Historical!$B${HRS['slope']}"
    d = lambda k: f"{fp}$D${FR[k]}"
    sh.put(f"A{r + 1}", "Cost of funds +50bp")
    sh.put(f"B{r + 1}", f"=-0.005*{d('avg_ibl')}*(1-{d('d_tax')})/{d('sh_dil')}", "usd")
    sh.put(f"A{r + 2}", "Net charge-off rate +100bp (RSA offsets per regression)")
    sh.put(f"B{r + 2}", f"=-0.01*{d('avg_loans')}*(1+{slope})*(1-{d('d_tax')})/{d('sh_dil')}", "usd")
    sh.put(f"A{r + 3}", "Loan yield +100bp")
    sh.put(f"B{r + 3}", f"=0.01*{d('avg_loans')}*(1-{d('d_tax')})/{d('sh_dil')}", "usd")
    sh.put(f"A{r + 4}", "Other expense +1% of FY2027E base")
    sh.put(f"B{r + 4}", f"=-0.01*{d('oe')}*(1-{d('d_tax')})/{d('sh_dil')}", "usd")
    return sh, {"T1c": T1c, "T2c": T2c, "T3c": T3c, "T1": T1, "T2": T2, "T3": T3, "T4": r + 1}


# ---------------------------------------------------------------- Summary and Checks
def build_summary_checks(wb, HR, FR_live, V_live, scn, I, CR, SENS, S):
    sm = Sheet(wb, "Summary", {"A": 58, "B": 16, "C": 16, "D": 16, "E": 16})
    sm.put("A1", f"{COMPANY} ({TICKER}): initiation model summary (base case headline; all figures are formulas)", bold=True)
    sm.put("A2", "Illustrative author analysis for portfolio purposes. Not investment advice.")
    b = q("Scn_Base")
    bear, bull = q("Scn_Bear"), q("Scn_Bull")
    SR = {}
    r = 4
    items = [
        ("rating", "Rating (mechanical, from base-case price target)", f"={b}B{scn['Base']['rating']}", None),
        ("pt", "12-month price target ($)", f"={b}B{scn['Base']['pt']}", "usd"),
        ("price", "Share price ($)", f"={I['price']}", "usd"),
        ("pdate", "Price date", f"=Inputs_Sources!$C${int(I['price'].split('$')[-1])}", "date"),
        ("upside", "Upside to price target", f"={b}B{scn['Base']['upside']}", "pct1"),
        ("tr", "Expected total return including dividends", f"={b}B{scn['Base']['tr']}", "pct1"),
        ("bear", "Bear case value per share, 12-month blended ($)", f"={bear}B{scn['Bear']['pt_raw']}", "usd"),
        ("base", "Base case value per share, 12-month blended ($)", f"={b}B{scn['Base']['pt_raw']}", "usd"),
        ("bull", "Bull case value per share, 12-month blended ($)", f"={bull}B{scn['Bull']['pt_raw']}", "usd"),
        ("bear_vs", "Bear value versus share price", f"={bear}B{scn['Bear']['pt_raw']}/{I['price']}-1", "pct1"),
        ("base_vs", "Base value versus share price", f"={b}B{scn['Base']['pt_raw']}/{I['price']}-1", "pct1"),
        ("bull_vs", "Bull value versus share price", f"={bull}B{scn['Bull']['pt_raw']}/{I['price']}-1", "pct1"),
        ("px_ow", "Price above which the rating falls below Overweight ($)", f"={b}B{scn['Base']['px_ow']}", "usd"),
        ("px_uw", "Price above which the rating is Underweight ($)", f"={b}B{scn['Base']['px_uw']}", "usd"),
        ("ptbv27", "Memo: P/TBV consistent with FY2027E ROTCE if sustained", f"={b}B{scn['Base']['ptbv_j27']}", "x2"),
        ("ke", "Cost of equity", f"={b}B{scn['Base']['ke']}", "pct2"),
        ("g", "Terminal growth", f"={b}B{scn['Base']['g']}", "pct1"),
        ("rot", "Terminal ROTCE", f"={b}B{scn['Base']['rot']}", "pct1"),
        ("ptbv", "Justified P/TBV at terminal", f"={b}B{scn['Base']['ptbv_j']}", "x2"),
        ("v12", "Residual income value, 12 months forward ($)", f"={b}B{scn['Base']['v12']}", "usd"),
        ("vpt", "Justified P/TBV value at horizon ($)", f"={b}B{scn['Base']['v_pt']}", "usd"),
        ("pe_now", "Current P/E on base FY2027E EPS", f"={b}B{scn['Base']['pe_now']}", "x2"),
        ("ptbv_now", "Current P/TBV on Jun-26 tangible book", f"={b}B{scn['Base']['ptbv_now']}", "x2"),
    ]
    for key, label, f, fmt in items:
        SR[key] = r
        sm.put(f"A{r}", label)
        sm.put(f"B{r}", f, fmt)
        r += 1
    r += 1
    sm.put(f"A{r}", "Base-case forecast", bold=True)
    for c, t in zip("BCDE", ["FY2025A", "FY2026E", "FY2027E", "FY2028E"]):
        sm.put(f"{c}{r}", t, bold=True)
    r += 1
    fb = scn["Base_F"]
    for key, label, fmt in [("loans_end", "Loan receivables, year end ($mm)", "num"), ("pv", "Purchase volume ($mm)", "num"), ("nii", "Net interest income ($mm)", "num"), ("nim", "Net interest margin", "pct2"),
                            ("rsa", "Retailer share arrangements ($mm)", "num"), ("prov", "Provision for credit losses ($mm)", "num"), ("oe", "Other expense ($mm)", "num"), ("eff", "Efficiency ratio", "pct1"),
                            ("nic", "Net earnings to common ($mm)", "num"), ("eps", "Diluted EPS ($)", "usd"), ("d_nco", "Net charge-off rate", "pct2"), ("d_cov", "Allowance coverage ratio", "pct2"),
                            ("rotce", "ROTCE", "pct1"), ("tbvps", "Tangible book value per share ($)", "usd"), ("cet1r", "CET1 ratio", "pct1"), ("dps", "Dividends per share ($)", "usd")]:
        SR["f_" + key] = r
        sm.put(f"A{r}", label)
        for c in "BCDE":
            sm.put(f"{c}{r}", f"={q('Scn_Base')}{c}{fb[key]}", fmt)
        r += 1
    SR["cons"] = r
    sm.put(f"A{r}", "Consensus EPS (Yahoo Finance)")
    sm.put(f"C{r}", f"={I['cons26']}", "usd")
    sm.put(f"D{r}", f"={I['cons27']}", "usd")
    r += 1
    SR["vs_cons"] = r
    sm.put(f"A{r}", "Model EPS versus consensus")
    sm.put(f"C{r}", f"=C{SR['f_eps']}/C{SR['cons']}-1", "pct1")
    sm.put(f"D{r}", f"=D{SR['f_eps']}/D{SR['cons']}-1", "pct1")

    # ---- Checks
    ck = Sheet(wb, "Checks", {"A": 82, "B": 16, "C": 16, "D": 10})
    ck.put("A1", "Model integrity checks (all formula-driven)", bold=True)
    for c, t in zip("ABCD", ["Check", "Value / difference", "Tolerance", "Status"]):
        ck.put(f"{c}3", t, bold=True)
    rows = []
    H = "Historical!"
    fy_cols = [HIST_COL[f"FY{y}"] for y in FY_LIST]
    q_cols = [HIST_COL[qk] for qk in Q_LIST]
    all_cols = fy_cols + q_cols

    def rng(key, cols):
        return ",".join(f"{H}{c}{HR[key]}" for c in cols)

    def maxabs(key, cols):
        return "MAX(" + ",".join(f"ABS({H}{c}{HR[key]})" for c in cols) + ")"

    rows.append(("Historical NIM: recomputed vs reported, max abs difference (all periods)", "=" + maxabs("nim_d", all_cols), 0.0006))
    rows.append(("Historical NCO rate: recomputed vs reported, max abs difference", "=" + maxabs("nco_d", all_cols), 0.0006))
    rows.append(("Historical efficiency ratio: recomputed vs reported, max abs difference", "=" + maxabs("eff_d", all_cols), 0.0011))
    rows.append(("Historical pretax income: formula vs reported, max abs difference ($mm)", "=MAX(" + ",".join(f"ABS({H}{c}{HR['pretax_c']}-{H}{c}{HR['pretax']})" for c in all_cols) + ")", 1.5))
    rows.append(("Historical NII = total interest income - interest expense, max abs difference ($mm)", "=MAX(" + ",".join(f"ABS({H}{c}{HR['int_total']}-{H}{c}{HR['int_exp']}-{H}{c}{HR['nii']})" for c in all_cols) + ")", 1.5))
    rows.append(("Historical net earnings = pretax - tax, max abs difference ($mm)", "=MAX(" + ",".join(f"ABS({H}{c}{HR['pretax']}-{H}{c}{HR['tax']}-{H}{c}{HR['ni']})" for c in all_cols) + ")", 1.5))
    rows.append(("Historical loan yield: recomputed vs reported, max abs difference", "=MAX(" + ",".join(f"ABS({H}{c}{HR['yield']}-{H}{c}{HR['yld_rep']})" for c in all_cols) + ")", 0.0006))
    rows.append(("Historical cost of funds: recomputed vs reported, max abs difference", "=MAX(" + ",".join(f"ABS({H}{c}{HR['cof']}-{H}{c}{HR['cof_rep']})" for c in all_cols) + ")", 0.0006))
    rows.append(("LTM EPS (filings) vs Yahoo trailing EPS, abs difference ($)", f"=ABS({H}{COL_LTM}{HR['eps']}-{I['ttm_eps']})", 0.05))
    rows.append(("Historical ACL roll-forward plug, largest abs 'other movement' FY2021-FY2025 ($mm; portfolio sales, acquisitions and off-balance-sheet reserves)", "=MAX(" + ",".join(f"ABS({H}{c}{HR['acl_oth']})" for c in fy_cols[1:]) + ")", 400))
    for label_, FRx, Vx, sheetname in [("live", FR_live, V_live, "Forecast")] + [(k.lower(), scn[k + "_F"], scn[k], "Scn_" + k) for k in ("Base", "Bear", "Bull")]:
        p = q(sheetname)
        for c in "CDE":
            pass
        rows.append((f"[{label_}] Allowance roll-forward: beginning + provision - net charge-offs + other - ending, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['acl_beg']}+{p}{c}{FRx['acl_prov']}-{p}{c}{FRx['acl_nco']}+{p}{c}{FRx['acl_oth']}-{p}{c}{FRx['acl_end']})" for c in "CDE") + ")", 0.001))
        rows.append((f"[{label_}] Loan receivables tie: prior year x (1 + growth) - ending, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{pc}{FRx['loans_end']}*(1+{p}{c}{FRx['d_loan_g']})-{p}{c}{FRx['loans_end']})" for c, pc in zip("CDE", "BCD")) + ")", 0.001))
        rows.append((f"[{label_}] TCE roll-forward: beginning + flows - ending, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['tce_beg']}+{p}{c}{FRx['tce_nic']}+{p}{c}{FRx['tce_div']}+{p}{c}{FRx['tce_bb']}+{p}{c}{FRx['tce_oth']}+{p}{c}{FRx['tce_rec']}-{p}{c}{FRx['tce_end']})" for c in "CDE") + ")", 0.001))
        rows.append((f"[{label_}] Share count roll-forward, max abs (mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['sh_beg']}-{p}{c}{FRx['sh_bb']}+{p}{c}{FRx['sh_iss']}-{p}{c}{FRx['sh_end']})" for c in "CDE") + ")", 0.001))
        rows.append((f"[{label_}] Net income tie: NII - RSA - provision + other income - other expense - tax - net earnings, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['nii']}-{p}{c}{FRx['rsa']}-{p}{c}{FRx['prov']}+{p}{c}{FRx['oi']}-{p}{c}{FRx['oe']}-{p}{c}{FRx['tax']}-{p}{c}{FRx['ni']})" for c in "CDE") + ")", 0.001))
        rows.append((f"[{label_}] Forecast NIM = NII / average interest-earning assets vs displayed, max abs", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['nii']}/{p}{c}{FRx['avg_iea']}-{p}{c}{FRx['nim']})" for c in "CDE") + ")", 0.000001))
        rows.append((f"[{label_}] Provision = net charge-offs + reserve build, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['prov']}-{p}{c}{FRx['acl_nco']}-{p}{c}{FRx['res_b']})" for c in "CDE") + ")", 0.001))
        rows.append((f"[{label_}] Forecast NCO rate x average loans = charge-offs, and displayed NCO rate matches driver, max abs ($mm)", "=MAX(" + ",".join(f"ABS({p}{c}{FRx['acl_nco']}/{p}{c}{FRx['avg_loans']}-{p}{c}{FRx['d_nco']})" for c in "CDE") + ")", 0.000001))
        rows.append((f"[{label_}] CET1 ratio at least 10% in all forecast years (minimum CET1 ratio)", "=MIN(" + ",".join(f"{p}{c}{FRx['cet1r']}" for c in "CDE") + ")", None))
    # valuation ties
    Vb = q("Scn_Base")
    rows.append(("Live scenario price target equals the selected scenario sheet's price target ($)", f"=ABS({q('Valuation')}B{V_live['pt']}-CHOOSE({S['switch']},{q('Scn_Bear')}B{scn['Bear']['pt']},{q('Scn_Base')}B{scn['Base']['pt']},{q('Scn_Bull')}B{scn['Bull']['pt']}))", 0.001))
    rows.append(("Sensitivity table 1 center cell equals residual income value 12m forward, abs difference ($)", f"=ABS(Sensitivity!{SENS['T1c']}-{q('Valuation')}B{V_live['v12']})", 0.001))
    rows.append(("Sensitivity table 2 center cell equals blended value before rounding, abs difference ($)", f"=ABS(Sensitivity!{SENS['T2c']}-{q('Valuation')}B{V_live['pt_raw']})", 0.001))
    rows.append(("Sensitivity table 3 center cell equals FY2027E EPS, abs difference ($)", f"=ABS(Sensitivity!{SENS['T3c']}-{q('Forecast')}D{FR_live['eps']})", 0.0001))
    rows.append(("Scenario ordering: bear < base < bull blended value (1 = holds)", f"=IF(AND({q('Scn_Bear')}B{scn['Bear']['pt_raw']}<{q('Scn_Base')}B{scn['Base']['pt_raw']},{q('Scn_Base')}B{scn['Base']['pt_raw']}<{q('Scn_Bull')}B{scn['Bull']['pt_raw']}),1,0)", None))
    rows.append(("Terminal ROTCE exceeds terminal growth and cost of equity exceeds growth (1 = holds)", f"=IF(AND({q('Valuation')}B{V_live['ke']}>{q('Valuation')}B{V_live['g']},{q('Valuation')}B{V_live['rot']}>{q('Valuation')}B{V_live['g']}),1,0)", None))
    rows.append(("Residual income present values sum: PV(RI) row total equals sum used in equity value, abs diff ($mm)", f"=ABS(SUM({q('Valuation')}B{V_live['pv']}:D{V_live['pv']})-{q('Valuation')}B{V_live['sum_pv']})", 0.001))
    rows.append(("Comps: SYF P/TBV recomputed = price / (Jun-26 TCE / shares), abs difference (x)", f"=ABS(Comps!{COMPS_COL['ptbv']}{CR[TICKER]}-{I['price']}/({H}{HIST_COL['Q2_2026']}{HR['tce_r']}/{I['shares']}))", 0.0001))
    r = 4
    status_cells = []
    for label, f, tol in rows:
        ck.put(f"A{r}", label)
        ck.put(f"B{r}", f, "0.000000")
        if tol is not None:
            ck.put(f"C{r}", tol, "0.000000")
            ck.put(f"D{r}", f'=IF(B{r}<=C{r},"OK","CHECK")')
        elif "CET1 ratio" in label:
            ck.put(f"C{r}", 0.10, "0.0%")
            ck.put(f"D{r}", f'=IF(B{r}>=C{r},"OK","CHECK")')
        else:
            ck.put(f"C{r}", 1, "0")
            ck.put(f"D{r}", f'=IF(B{r}=C{r},"OK","CHECK")')
        status_cells.append(f"D{r}")
        r += 1
    r += 1
    ck.put(f"A{r}", "Overall", bold=True)
    ck.put(f"B{r}", f'=COUNTIF(D4:D{r - 2},"OK")', "0")
    ck.put(f"C{r}", f"=COUNTA(D4:D{r - 2})", "0")
    ck.put(f"D{r}", f'=IF(B{r}=C{r},"OK","CHECK")', bold=True)
    sm.put(f"A{SR['vs_cons'] + 2}", "Model integrity checks passing")
    sm.put(f"B{SR['vs_cons'] + 2}", f"=Checks!B{r}&\" of \"&Checks!C{r}")
    return sm, ck, SR, r


HRS = None


def crosscheck(scen="base"):
    """Recompute the base-case forecast and valuation in plain Python from the raw snapshots and compare with the recalculated workbook."""
    from openpyxl import load_workbook
    hist, market, _ = load_all()
    ann, qtr, misc = hist["annual"], hist["quarterly"], hist["misc"]
    ys = [str(y) for y in FY_LIST]
    nco_r = [ann[y][F_NCO] / ann[y][F_AVG_LOANS] for y in ys]
    rsa_r = [-ann[y][F_RSA] / ann[y][F_AVG_LOANS] for y in ys]
    mx, my = sum(nco_r) / len(ys), sum(rsa_r) / len(ys)
    slope = sum((x - mx) * (y - my) for x, y in zip(nco_r, rsa_r)) / sum((x - mx) ** 2 for x in nco_r)
    icpt = my - slope * mx
    drv = {k: vals[scen] for k, _, _, vals, _ in SCENARIO_DRIVERS}
    com = {k: vals for k, _, _, vals, _ in COMMON_ASSUMPTIONS}
    pdate, price = close_before(TICKER, dt.date.fromisoformat(market["retrieved_date"]))
    a25, qq = ann["2025"], qtr["Q2_2026"]
    rec = qq["_recast_prior_ye"]
    cover = misc["cover_shares_outstanding_m"]
    nic_h1 = qtr["Q1_2026"][F_NIC] + qq[F_NIC]
    div_h1, bb_h1 = misc["h1_2026_common_dividends_mm"], hist["buybacks_quarterly_text"]["Q1_2026"]["dollars_mm"] + hist["buybacks_quarterly_text"]["Q2_2026"]["dollars_mm"]
    bbsh_h1 = hist["buybacks_quarterly_text"]["Q1_2026"]["shares_m"] + hist["buybacks_quarterly_text"]["Q2_2026"]["shares_m"]
    dps_h1 = misc["dividends_quarterly_text"]["Q1_2026"]["dps"] + misc["dividends_quarterly_text"]["Q2_2026"]["dps"]
    dsh_h1 = (qtr["Q1_2026"][F_DSH] + qq[F_DSH]) / 2
    tce_jun = qq[F_EQ] - qq[F_PREF] - qq[F_GW] - qq[F_INT]
    loans, acl, oe = a25[F_LOANS_END], a25[F_ACL], a25[F_OE]
    tce = a25[F_EQ] - a25[F_PREF] - a25[F_GW] - rec["intangibles"]
    sh = a25[F_SH]
    px = price
    out = {"eps": [], "nic": [], "tce_end": [], "sh_end": []}
    nic_list = []
    for yi in range(3):
        loans_new = loans * (1 + drv["loan_g"][yi])
        avg_l = (loans + loans_new) / 2 * com["avgf"][yi]
        oth = avg_l * com["oth_iea"][yi]
        iea = avg_l + oth
        ibl = iea * com["ibl"][yi]
        nii = avg_l * drv["yield"][yi] + oth * com["oth_y"][yi] - ibl * drv["cof"][yi]
        rsa = (icpt + slope * drv["nco"][yi] + com["rsa_res"][yi]) * avg_l
        nco = drv["nco"][yi] * avg_l
        acl_new = drv["cov"][yi] * loans_new
        prov = nco + acl_new - acl
        oe = oe * (1 + drv["opex_g"][yi])
        pretax = nii - rsa - prov + avg_l * com["oi"][yi] - oe
        ni = pretax * (1 - com["tax"][yi])
        nic = ni - com["pref"][yi]
        if yi == 0:
            px_y = price * (1 + com["px_g"][0])
            bb_sh = bbsh_h1 + (drv["bb"][0] - bb_h1) / px_y
            iss = (cover - sh + bbsh_h1) + com["iss"][0]
            sh_new = sh - bb_sh + iss
            dil = (dsh_h1 + ((cover + sh_new) / 2 + com["dil"][0])) / 2
            div = div_h1 + (com["dps"][0] - dps_h1) * (cover + sh_new) / 2
            rec_adj = tce_jun - (tce + nic_h1 - div_h1 - bb_h1)
        else:
            px_y = px * (1 + com["px_g"][yi])
            bb_sh = drv["bb"][yi] / px_y
            sh_new = sh - bb_sh + com["iss"][yi]
            dil = (sh + sh_new) / 2 + com["dil"][yi]
            div = com["dps"][yi] * (sh + sh_new) / 2
            rec_adj = 0
        tce_beg = tce
        tce = tce + nic - div - drv["bb"][yi] + com["oth_tce"][yi] + rec_adj
        out["eps"].append(nic / dil)
        out["nic"].append(nic)
        out["tce_end"].append(tce)
        out["sh_end"].append(sh_new)
        nic_list.append((nic, tce_beg))
        loans, acl, sh, px = loans_new, acl_new, sh_new, px_y
    # valuation
    val = {a[0]: a[2] for a in VAL_ASSUMPTIONS}
    rf = market["risk_free"]["yield_pct"] / 100
    beta = market["tickers"][TICKER]["summaryDetail"]["beta"]["raw"]
    ke = rf + beta * val["erp"]
    g = val["g"]
    rot = SCENARIO_ROT[scen]
    tau = (pdate - BALANCE_DATE).days / 365
    lens, times = [0.5, 1.0, 1.0], [0.5, 1.5, 2.5]
    nics = [out["nic"][0] - nic_h1, out["nic"][1], out["nic"][2]]
    tceb = [tce_jun, out["tce_end"][0], out["tce_end"][1]]
    pv_ri = sum((n - ke * b * l) / (1 + ke) ** t for n, b, l, t in zip(nics, tceb, lens, times))
    pv_tv = out["tce_end"][2] * (rot - g) / (ke - g) / (1 + ke) ** 2.5
    v0 = (tce_jun + pv_ri + pv_tv) / cover
    v12 = (v0 * (1 + ke) ** tau - val["dps_int"]) * (1 + ke) - val["dps_fwd"]
    tbv26, tbv27 = out["tce_end"][0] / out["sh_end"][0], out["tce_end"][1] / out["sh_end"][1]
    w27 = ((pdate + dt.timedelta(days=365)) - dt.date(2026, 12, 31)).days / 365
    v_pt = (rot - g) / (ke - g) * ((1 - w27) * tbv26 + w27 * tbv27)
    raw_pt = val["w_ri"] * v12 + val["w_pt"] * v_pt
    wb = load_workbook(FINAL_XLSX, data_only=True)
    layout = json.loads((DATA_DIR / "model_layout.json").read_text(encoding="utf-8"))
    F, Vv = layout["SCN"][scen.capitalize() + "_F"], layout["SCN"][scen.capitalize()]
    ws = wb["Scn_" + scen.capitalize()]
    pairs = [("EPS FY2026E", out["eps"][0], ws[f"C{F['eps']}"].value), ("EPS FY2027E", out["eps"][1], ws[f"D{F['eps']}"].value), ("EPS FY2028E", out["eps"][2], ws[f"E{F['eps']}"].value),
             ("TCE FY2028E", out["tce_end"][2], ws[f"E{F['tce_end']}"].value), ("RI value 12m", v12, ws[f"B{Vv['v12']}"].value), ("Justified P/TBV value", v_pt, ws[f"B{Vv['v_pt']}"].value), ("Blended value", raw_pt, ws[f"B{Vv['pt_raw']}"].value)]
    ok = True
    for name, py_val, xl_val in pairs:
        good = abs(py_val - xl_val) < 1e-6 * max(1, abs(xl_val))
        ok &= good
        print(f"crosscheck {name}: python {py_val:.6f} vs workbook {xl_val:.6f} {'OK' if good else 'MISMATCH'}")
    return ok


def main():
    global HRS
    hist, market, peers = load_all()
    wb = Workbook()
    wb.remove(wb.active)
    Sheet_summary_placeholder = None
    # sheet order is fixed afterwards; build in dependency order
    hs, HR = build_historical(wb, hist)
    HRS = HR
    ins, I = build_inputs(wb, hist, market, peers)
    source_register(ins, I, hist, market, peers)
    scs, S = build_scenarios(wb)
    live_f, live_v, FR_live, V_live = build_engine(wb, "Forecast", "Valuation", S, I, HR, "live")
    scn = {}
    for name in ("Bear", "Base", "Bull"):
        fs_, vs_, FRs, Vs = build_engine(wb, "Scn_" + name, "Scn_" + name, S, I, HR, name.lower())
        scn[name] = Vs
        scn[name + "_F"] = FRs
    comps, CR = build_comps(wb, hist, market, peers, HR, I)
    sens, SENS = build_sensitivity(wb, FR_live, V_live, I, "Forecast", "Valuation", S)
    sm, ck, SR, ck_last = build_summary_checks(wb, HR, FR_live, V_live, scn, I, CR, SENS, S)
    order = ["Summary", "Inputs_Sources", "Historical", "Scenarios", "Forecast", "Valuation", "Comps", "Sensitivity", "Scn_Bear", "Scn_Base", "Scn_Bull", "Checks"]
    wb._sheets = [wb[n] for n in order]
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.properties.title = f"{COMPANY} initiation model"
    RAW_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(RAW_XLSX)
    layout = {"HR": HR, "FR": FR_live, "V": V_live, "SCN": {k: v for k, v in scn.items()}, "CR": CR, "SR": SR, "SENS": SENS, "I": {k: v for k, v in I.items() if k != "_next_row"}, "S": S, "checks_last": ck_last}
    (DATA_DIR / "model_layout.json").write_text(json.dumps(layout, indent=1), encoding="utf-8")
    recalc()
    print("independent python cross-check (base):", "PASS" if crosscheck("base") else "FAIL")


def recalc():
    """Recalculate with LibreOffice so cached values exist, then restore the author property."""
    out = RAW_XLSX.parent / "recalc"
    out.mkdir(exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(RAW_XLSX)], check=True, capture_output=True, timeout=300)
    shutil.copy(out / RAW_XLSX.name, FINAL_XLSX)
    fix_author(FINAL_XLSX)
    print("wrote", FINAL_XLSX)


def fix_author(path):
    import zipfile
    import re as _re
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                s = data.decode("utf-8")
                if "<dc:creator>" in s:
                    s = _re.sub(r"<dc:creator>.*?</dc:creator>", f"<dc:creator>{AUTHOR}</dc:creator>", s)
                else:
                    s = s.replace("</cp:coreProperties>", f"<dc:creator>{AUTHOR}</dc:creator></cp:coreProperties>")
                if "<cp:lastModifiedBy>" in s:
                    s = _re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", f"<cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy>", s)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


if __name__ == "__main__":
    main()
