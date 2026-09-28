"""Pull Synchrony (SYF) and peer data from SEC EDGAR and Yahoo Finance; write snapshots to data/."""
import csv
import datetime as dt
import json
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ============================ CONFIG (edit here) ============================
TICKER = "SYF"
CIK = "0001601712"
SEC_USER_AGENT = "Portfolio research contact@example.com"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
FILINGS_DIR = DATA_DIR / "filings"
COMPANYFACTS_FILE = DATA_DIR / "companyfacts.json"
HISTORY_FILE = DATA_DIR / "syf_history.json"
MARKET_FILE = DATA_DIR / "market_snapshot.json"
PRICES_FILE = DATA_DIR / "prices_1y.csv"
PEER_FACTS_FILE = DATA_DIR / "peer_xbrl.json"

# (key, accession, primary document, kind, period label)
FILINGS = [
    ("10K_FY2025", "0001601712-26-000006", "syf-20251231.htm", "10-K", "FY2025"),
    ("10K_FY2023", "0001601712-24-000047", "syf-20231231.htm", "10-K", "FY2023"),
    ("10K_FY2021", "0001601712-22-000053", "syf-20211231.htm", "10-K", "FY2021"),
    ("10Q_2026Q2", "0001601712-26-000033", "syf-20260630.htm", "10-Q", "Q2 2026"),
    ("10Q_2026Q1", "0001601712-26-000016", "syf-20260331.htm", "10-Q", "Q1 2026"),
]
# annual filing -> fiscal years in its 3 statement columns (newest first)
ANNUAL_COLS = {"10K_FY2025": [2025, 2024, 2023], "10K_FY2023": [2023, 2022, 2021], "10K_FY2021": [2021, 2020, 2019]}
# annual filing -> year-end balance sheet columns
BS_COLS = {"10K_FY2025": [2025, 2024], "10K_FY2023": [2023, 2022], "10K_FY2021": [2021, 2020]}
FISCAL_YEARS = [2020, 2021, 2022, 2023, 2024, 2025]
# quarterly filing -> quarter labels in three-month statement columns
Q_COLS = {"10Q_2026Q2": ["Q2_2026", "Q2_2025"], "10Q_2026Q1": ["Q1_2026", "Q1_2025"]}

PEERS = {  # ticker -> (CIK, label)
    "COF": ("0000927628", "Capital One"),
    "AXP": ("0000004962", "American Express"),
    "BFH": ("0001101215", "Bread Financial"),
    "OMF": ("0001584207", "OneMain Holdings"),
    "SOFI": ("0001818874", "SoFi Technologies"),
}
YAHOO_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
YAHOO_MODULES = "price,summaryDetail,defaultKeyStatistics,financialData,earningsTrend"
RF_TICKER = "^TNX"  # 10-year Treasury yield proxy for the risk-free rate
# ==========================================================================

SEC_HEADERS = {"User-Agent": SEC_USER_AGENT}
NUM_RE = re.compile(r"^\(?\s*-?[\d,]+(\.\d+)?\s*\)?%?$")


def sec_get(url, **kw):
    time.sleep(0.4)  # SEC fair-access rate limit
    r = requests.get(url, headers=SEC_HEADERS, timeout=120, **kw)
    r.raise_for_status()
    return r


def html_to_rows(content):
    soup = BeautifulSoup(content, "lxml")
    for t in soup(["script", "style"]):
        t.decompose()
    for tr in soup.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        cells = [c for c in cells if c not in ("", ")", "%")]
        tr.replace_with(soup.new_string("\n" + " | ".join(cells) + "\n"))
    txt = soup.get_text("\n")
    return re.sub(r"\n\s*\n+", "\n", txt)


def download_filings(refresh=False):
    FILINGS_DIR.mkdir(parents=True, exist_ok=True)
    for key, acc, doc, _, _ in FILINGS:
        txt_path = FILINGS_DIR / f"{key}.txt"
        if txt_path.exists() and not refresh:
            continue
        url = f"https://www.sec.gov/Archives/edgar/data/{int(CIK)}/{acc.replace('-', '')}/{doc}"
        content = sec_get(url).content
        (FILINGS_DIR / f"{key}.htm").write_bytes(content)
        txt_path.write_text(html_to_rows(content), encoding="utf-8")
        print("downloaded", key, url)


def download_companyfacts(refresh=False):
    if not COMPANYFACTS_FILE.exists() or refresh:
        COMPANYFACTS_FILE.write_bytes(sec_get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json").content)
        print("downloaded companyfacts")


# ---------------- text parsing helpers ----------------
def parse_nums(line):
    """Numbers in a pipe-delimited table row, skipping the label cell."""
    out = []
    for tok in line.split(" | ")[1:]:
        tok = tok.strip()
        if tok in ("$", ""):
            continue
        if tok in ("—", "-"):
            out.append(0.0)
            continue
        if NUM_RE.match(tok):
            neg = tok.startswith("(") or tok.startswith("-")
            val = float(re.sub(r"[^\d.]", "", tok))
            out.append(-val if neg else val)
    return out


def find_line(lines, pat, start=0, window=None):
    rx = re.compile(pat)
    end = len(lines) if window is None else min(len(lines), start + window)
    for i in range(start, end):
        if rx.search(lines[i]):
            return i
    raise LookupError(f"pattern not found: {pat} (start={start})")


def grab(lines, start, label_pat, nvals, window=600):
    rx = re.compile(label_pat)
    for i in range(start, min(len(lines), start + window)):
        if rx.search(lines[i]):
            nums = parse_nums(lines[i])
            if len(nums) >= nvals:
                return nums, lines[i]
    raise LookupError(f"row not found: {label_pat} from line {start}")


# (field, label regex, kind) per section; kind: "col" = one value per period column
IS_ROWS = [
    ("interest_and_fees_on_loans", r"^Interest and fees on loans \(Note"),
    ("total_interest_income", r"^Total interest income \|"),
    ("total_interest_expense", r"^Total interest expense \|"),
    ("net_interest_income", r"^Net interest income \|"),
    ("rsa", r"^Retailer share arrangements \|"),
    ("provision", r"^Provision for credit losses"),
    ("other_income", r"^Total other income \|"),
    ("other_expense", r"^Total other expense \|"),
    ("pretax_income", r"^Earnings before provision for income taxes"),
    ("income_tax", r"^Provision for income taxes"),
    ("net_earnings", r"^Net earnings \|"),
    ("net_earnings_common", r"^Net earnings available to common stockholders"),
    ("eps_diluted", r"^Diluted \|"),
]
STAT_ROWS = [
    ("avg_loans", r"^Loan receivables, including held for sale \|"),
    ("avg_assets", r"^Total assets \|"),
    ("avg_equity", r"^Total equity \|"),
    ("purchase_volume", r"^Purchase volume \(1\)"),
    ("avg_active_accounts", r"^Average active accounts"),
    ("nim_reported", r"^Net interest margin \(\d\)"),
    ("net_chargeoffs", r"^Net charge-offs \|"),
    ("nco_rate_reported", r"^Net charge-offs (\(annualized\) )?as a %"),
    ("acl_coverage_reported", r"^Allowance coverage ratio"),
    ("roa_reported", r"^Return on assets"),
    ("roe_reported", r"^Return on equity"),
    ("efficiency_reported", r"^Efficiency ratio"),
    ("loans_end", r"^Loan receivables \|"),
    ("acl_end", r"^Allowance for credit losses \|"),
    ("dq30_reported", r"^30\+ days past due"),
    ("dq90_reported", r"^90\+ days past due"),
    ("active_accounts_end", r"^Total active accounts"),
]
AVG_ROWS = [  # each row gives (avg balance, interest, yield) triples per period
    ("loans_avg_bal", "loans_int", "loan_yield_reported", r"^Total loan receivables, including held for sale \|"),
    ("iea_avg_bal", "iea_int", "iea_yield_reported", r"^Total interest-earning assets \|"),
    ("ibl_avg_bal", "ibl_int", "cost_of_funds_reported", r"^Total interest-bearing liabilities \|"),
    ("dep_avg_bal", "dep_int", "deposit_rate_reported", r"^Interest-bearing deposit accounts \|"),
]
BS_ROWS = [
    ("total_loans_bs", r"^Total loan receivables \|"),
    ("goodwill", r"^Goodwill( \(Note \d+\))? \|"),
    ("intangibles", r"^Intangible assets, net"),
    ("total_assets_bs", r"^Total assets \|"),
    ("total_deposits", r"^Total deposits \|"),
    ("total_borrowings", r"^Total borrowings \|"),
    ("preferred_equity", r"^Preferred stock, par"),
    ("total_equity", r"^Total equity \|"),
]


def extract_periods(lines, cols, ncols_is=None):
    """Extract every field for the given statement period columns from one filing's lines."""
    out = {c: {} for c in cols}
    src = {}
    n = len(cols)
    i_is = find_line(lines, r"^Interest and fees on loans \(Note")
    for field, pat in IS_ROWS:
        nums, raw = grab(lines, i_is, pat, n, window=60)
        for c, v in zip(cols, nums[:n]):
            out[c][field] = v
        src[field] = "Consolidated Statements of Earnings"
    i_stat = find_line(lines, r"^Financial Position Data \(Average\)")
    for field, pat in STAT_ROWS:
        nums, raw = grab(lines, i_stat, pat, n, window=80)
        for c, v in zip(cols, nums[:n]):
            out[c][field] = v
        src[field] = "MD&A Other Financial and Statistical Data"
    i_avg = find_line(lines, r"^Interest-earning assets:")
    for f_bal, f_int, f_yld, pat in AVG_ROWS:
        nums, raw = grab(lines, i_avg, pat, 3 * n, window=80)
        for k, c in enumerate(cols):
            out[c][f_bal], out[c][f_int], out[c][f_yld] = nums[3 * k: 3 * k + 3]
        src[f_bal] = "MD&A Average Balance Sheet"
    return out, src


def extract_bs(lines, cols):
    out = {c: {} for c in cols}
    i_bs = find_line(lines, r"^Loan receivables: \(Notes")
    for field, pat in BS_ROWS:
        nums, raw = grab(lines, i_bs, pat, len(cols), window=60)
        for c, v in zip(cols, nums[: len(cols)]):
            out[c][field] = v
    # shares outstanding from the common stock caption
    cs = lines[find_line(lines, r"^Common stock, par", i_bs, 60)]
    m = re.search(r"([\d,\.]+)( million)? and ([\d,\.]+)( million)? shares outstanding", cs)
    if not m:
        raise LookupError("shares outstanding caption not parsed")

    def to_m(s, mil):
        v = float(s.replace(",", ""))
        return v if mil else v / 1e6

    shares = [to_m(m.group(1), m.group(2)), to_m(m.group(3), m.group(4))]
    for c, v in zip(cols, shares):
        out[c]["shares_outstanding_m"] = v
    # CET1 and RWA (company-level table is the first one in the MD&A)
    i_cet = find_line(lines, r"^Common equity Tier 1 capital \|")
    nums, _ = grab(lines, i_cet, r"^Common equity Tier 1 capital \|", 4, window=5)
    out[cols[0]]["cet1_amount"], out[cols[0]]["cet1_ratio"] = nums[0], nums[1]
    if len(cols) > 1:
        out[cols[1]]["cet1_amount"], out[cols[1]]["cet1_ratio"] = nums[2], nums[3]
    nums, _ = grab(lines, i_cet, r"^Risk-weighted assets \|", 2, window=8)
    out[cols[0]]["rwa"] = nums[0]
    if len(cols) > 1:
        out[cols[1]]["rwa"] = nums[1]
    return out


def extract_weighted_shares(lines, ncols):
    i = find_line(lines, r"^Weighted average common shares outstanding, dilutive")
    return parse_nums(lines[i])[:ncols]


def read_lines(key):
    return (FILINGS_DIR / f"{key}.txt").read_text(encoding="utf-8").split("\n")


def build_history():
    hist = {"annual": {}, "quarterly": {}, "crosscheck": [], "source_notes": {}}
    annual_by_filing = {}
    for key, cols in ANNUAL_COLS.items():
        lines = read_lines(key)
        vals, src = extract_periods(lines, cols)
        bs = extract_bs(lines, BS_COLS[key])
        for y, d in bs.items():
            vals[y].update(d)
        ws = extract_weighted_shares(lines, len(cols))
        for y, v in zip(cols, ws):
            vals[y]["diluted_shares_m"] = v
        annual_by_filing[key] = vals
        hist["source_notes"].update(src)
    # newest filing wins per year; overlapping years are cross-checked between filings
    for y in FISCAL_YEARS:
        rec = None
        for key in ("10K_FY2025", "10K_FY2023", "10K_FY2021"):
            if y in annual_by_filing[key]:
                if rec is None:
                    rec = dict(annual_by_filing[key][y])
                    rec["_filing"] = key
                else:
                    for f, v in annual_by_filing[key][y].items():
                        if f not in rec:
                            rec[f] = v  # fields the newer filing does not carry (e.g. year-end balance sheet)
                        elif abs(rec[f] - v) > 1e-9:
                            hist["crosscheck"].append(f"FY{y} {f}: {rec['_filing']}={rec[f]} vs {key}={v}")
        hist["annual"][str(y)] = rec
    # quarterly (three-month columns)
    for key, cols in Q_COLS.items():
        lines = read_lines(key)
        vals, src = extract_periods(lines, cols)
        i_bs = find_line(lines, r"^Loan receivables: \(Notes")
        bs = extract_bs(lines, ["cur", "prior_ye"])
        vals[cols[0]].update(bs["cur"])
        vals[cols[0]]["_recast_prior_ye"] = bs["prior_ye"]
        ws = extract_weighted_shares(lines, 2)
        vals[cols[0]]["diluted_shares_m"], vals[cols[1]]["diluted_shares_m"] = ws[0], ws[1]
        for c in cols:
            vals[c]["_filing"] = key
        hist["quarterly"].update(vals)
    return hist


# ---------------- XBRL ----------------
def xbrl_annual(facts, tag, unit="USD"):
    """FY value by end-date year for 12-month duration facts filed on 10-K (latest filing wins)."""
    rows = facts["facts"]["us-gaap"].get(tag, {}).get("units", {}).get(unit, [])
    out = {}
    for r in rows:
        if r["form"] != "10-K" or "start" not in r:
            continue
        d0, d1 = dt.date.fromisoformat(r["start"]), dt.date.fromisoformat(r["end"])
        if 350 <= (d1 - d0).days <= 380:
            prev = out.get(d1.year)
            if prev is None or r["filed"] > prev["filed"]:
                out[d1.year] = {"val": r["val"], "filed": r["filed"], "accn": r["accn"]}
    return out


def xbrl_instant(facts, tag, unit="USD", forms=("10-K",)):
    rows = facts["facts"]["us-gaap"].get(tag, {}).get("units", {}).get(unit, [])
    out = {}
    for r in rows:
        if r["form"] in forms and "start" not in r:
            prev = out.get(r["end"])
            if prev is None or r["filed"] > prev["filed"]:
                out[r["end"]] = {"val": r["val"], "filed": r["filed"], "accn": r["accn"]}
    return out


def build_xbrl_crosscheck(hist):
    facts = json.loads(COMPANYFACTS_FILE.read_text(encoding="utf-8"))
    checks = [  # (field, tag, scale to $mm)
        ("total_interest_expense", "InterestExpense", 1e-6),
        ("net_interest_income", "InterestIncomeExpenseNet", 1e-6),
        ("provision", "ProvisionForLoanLossesExpensed", 1e-6),
        ("net_earnings", "NetIncomeLoss", 1e-6),
        ("eps_diluted", "EarningsPerShareDiluted", 1.0),
        ("interest_and_fees_on_loans", "InterestAndFeeIncomeLoansAndLeases", 1e-6),
    ]
    results = []
    for field, tag, scale in checks:
        unit = "USD/shares" if tag.startswith("EarningsPerShare") else "USD"
        ann = xbrl_annual(facts, tag, unit)
        for y in FISCAL_YEARS:
            if y in ann:
                x = ann[y]["val"] * scale
                h = hist["annual"][str(y)][field]
                results.append({"field": field, "year": y, "text": h, "xbrl": round(x, 4), "tag": tag, "accn": ann[y]["accn"], "match": abs(x - h) < 0.6})
    for field, tag, scale in [("total_equity", "StockholdersEquity", 1e-6), ("total_deposits", "Deposits", 1e-6), ("acl_end", "FinancingReceivableAllowanceForCreditLosses", 1e-6)]:
        ins = xbrl_instant(facts, tag)
        for y in FISCAL_YEARS:
            k = f"{y}-12-31"
            if k in ins:
                x = ins[k]["val"] * scale
                h = hist["annual"][str(y)][field]
                results.append({"field": field, "year": y, "text": h, "xbrl": round(x, 4), "tag": tag, "accn": ins[k]["accn"], "match": abs(x - h) < 0.6})
    # dividends per share and buybacks are XBRL-only inputs
    extra = {"dividends_per_share": {}, "buyback_cash_mm": {}}
    for y, v in xbrl_annual(facts, "CommonStockDividendsPerShareDeclared", "USD/shares").items():
        extra["dividends_per_share"][str(y)] = {"val": v["val"], "accn": v["accn"], "tag": "CommonStockDividendsPerShareDeclared"}
    for y, v in xbrl_annual(facts, "PaymentsForRepurchaseOfCommonStock").items():
        extra["buyback_cash_mm"][str(y)] = {"val": v["val"] / 1e6, "accn": v["accn"], "tag": "PaymentsForRepurchaseOfCommonStock"}
    extra["common_dividends_mm"] = {}
    for y, v in xbrl_annual(facts, "DividendsCommonStockCash").items():
        extra["common_dividends_mm"][str(y)] = {"val": v["val"] / 1e6, "accn": v["accn"], "tag": "DividendsCommonStockCash"}
    return results, extra


def parse_buybacks_from_text():
    """Quarterly buyback dollars and shares from the capital sections (2025 from 10-K, 2026 from 10-Qs)."""
    out = {}
    lines = read_lines("10K_FY2025")
    i = find_line(lines, r"^Common Shares Repurchased Under Publicly Announced Programs")
    for lab, q in [("March 31, 2025", "Q1_2025"), ("June 30, 2025", "Q2_2025"), ("September 30, 2025", "Q3_2025"), ("December 31, 2025", "Q4_2025")]:
        nums, _ = grab(lines, i, "^" + lab, 2, window=10)
        out[q] = {"shares_m": nums[0], "dollars_mm": nums[1]}
    lines = read_lines("10Q_2026Q2")
    i = find_line(lines, r"^Common Shares Repurchased Under Publicly Announced Programs")
    for lab, q in [("March 31, 2026", "Q1_2026"), ("June 30, 2026", "Q2_2026")]:
        nums, _ = grab(lines, i, "^" + lab, 2, window=10)
        out[q] = {"shares_m": nums[0], "dollars_mm": nums[1]}
    return out


def parse_misc_text():
    """Quarter dividends, Q2 2026 balance sheet, recast notes and the Q2 2026 Series C preferred."""
    misc = {}
    lines = read_lines("10Q_2026Q2")
    i = find_line(lines, r"^Common Stock Cash Dividends Declared")
    nums, _ = grab(lines, i, r"^March 31, 2026", 2, window=10)
    misc["dps_Q1_2026"] = nums[0]
    nums, _ = grab(lines, i, r"^June 30, 2026", 2, window=10)
    misc["dps_Q2_2026"] = nums[0]
    nums, _ = grab(lines, i, r"^Total dividends declared", 2, window=10)
    misc["h1_2026_common_dividends_mm"] = nums[1]
    quarterly_div = {}
    for lab, q in [("March 31, 2026", "Q1_2026"), ("June 30, 2026", "Q2_2026")]:
        n, _ = grab(lines, i, "^" + lab, 2, window=10)
        quarterly_div[q] = {"dps": n[0], "dollars_mm": n[1]}
    k_lines = read_lines("10K_FY2025")
    j = find_line(k_lines, r"^Common Stock Cash Dividends Declared")
    for lab, q in [("March 31, 2025", "Q1_2025"), ("June 30, 2025", "Q2_2025")]:
        n, _ = grab(k_lines, j, "^" + lab, 2, window=10)
        quarterly_div[q] = {"dps": n[0], "dollars_mm": n[1]}
    misc["dividends_quarterly_text"] = quarterly_div
    facts = json.loads(COMPANYFACTS_FILE.read_text(encoding="utf-8"))
    cover = [r for r in facts["facts"]["dei"]["EntityCommonStockSharesOutstanding"]["units"]["shares"] if r["form"] == "10-Q"][-1]
    misc["cover_shares_outstanding_m"] = cover["val"] / 1e6
    misc["cover_shares_date"] = cover["end"]
    misc["cover_shares_accn"] = cover["accn"]
    misc["q2_2026_dividend_raise_note"] = "Quarterly dividend raised to $0.34 per common share commencing Q3 2026 (10-Q Q2 2026, Dividend and Share Repurchases)"
    misc["q2_2026_buyback_authorization_remaining_bn"] = 5.7
    misc["q2_2026_series_c_preferred_note"] = "June 2026: $500mm 7.250% Series C non-cumulative perpetual preferred issued (10-Q Q2 2026)"
    txt = "\n".join(lines)
    m = re.search(r"reclassified to conform with the current period presentation", txt)
    misc["software_reclass_found"] = bool(m)
    return misc


# ---------------- market data ----------------
def yahoo_session():
    s = requests.Session()
    s.headers.update({"User-Agent": YAHOO_UA})
    s.get("https://fc.yahoo.com", timeout=20)
    crumb = s.get("https://query1.finance.yahoo.com/v1/test/getcrumb", timeout=20).text
    return s, crumb


def pull_market(tickers):
    s, crumb = yahoo_session()
    snap = {"retrieved_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"), "retrieved_date": dt.date.today().isoformat(), "source": "Yahoo Finance quoteSummary and chart API", "tickers": {}}
    price_rows = []
    for t in tickers:
        r = s.get(f"https://query1.finance.yahoo.com/v10/finance/quoteSummary/{t}", params={"modules": YAHOO_MODULES, "crumb": crumb}, timeout=30)
        r.raise_for_status()
        snap["tickers"][t] = r.json()["quoteSummary"]["result"][0]
        c = s.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{t}", params={"range": "1y", "interval": "1d"}, timeout=30)
        c.raise_for_status()
        res = c.json()["chart"]["result"][0]
        for ts, px in zip(res["timestamp"], res["indicators"]["quote"][0]["close"]):
            if px is not None:
                price_rows.append((t, dt.datetime.fromtimestamp(ts, dt.timezone.utc).date().isoformat(), round(px, 4)))
        snap["tickers"][t]["_last_chart_meta"] = {k: res["meta"].get(k) for k in ("regularMarketPrice", "regularMarketTime", "currency", "symbol")}
        time.sleep(0.5)
    rf = s.get(f"https://query1.finance.yahoo.com/v8/finance/chart/{RF_TICKER}", params={"range": "5d", "interval": "1d"}, timeout=30).json()["chart"]["result"][0]
    snap["risk_free"] = {"ticker": RF_TICKER, "yield_pct": rf["meta"]["regularMarketPrice"], "asof_epoch": rf["meta"]["regularMarketTime"], "note": "CBOE 10-year Treasury yield index via Yahoo, quoted in percent"}
    with PRICES_FILE.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ticker", "date", "close"])
        w.writerows(price_rows)
    return snap


def pull_peer_xbrl():
    out = {}
    for t, (cik, name) in PEERS.items():
        j = sec_get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json").json()
        keep = {"entityName": j["entityName"], "cik": cik, "us-gaap": {}, "dei": {}}
        for tag in ("NetIncomeLoss", "NetIncomeLossAvailableToCommonStockholdersBasic", "EarningsPerShareDiluted", "StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "PreferredStockValue", "PreferredStockIncludingAdditionalPaidInCapitalNetOfDiscount", "Goodwill", "IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet", "IntangibleAssetsNetIncludingGoodwill", "PreferredStockLiquidationPreferenceValue", "WeightedAverageNumberOfDilutedSharesOutstanding", "LoansAndLeasesReceivableNetReportedAmount", "ProvisionForLoanLossesExpensed", "ProvisionForCreditLosses"):
            if tag in j["facts"].get("us-gaap", {}):
                keep["us-gaap"][tag] = {"units": {u: [r for r in rows if r["end"] >= "2024-06-30"] for u, rows in j["facts"]["us-gaap"][tag]["units"].items()}}
        for tag in ("EntityCommonStockSharesOutstanding",):
            if tag in j["facts"].get("dei", {}):
                keep["dei"][tag] = j["facts"]["dei"][tag]
        out[t] = keep
        print("peer xbrl", t, j["entityName"])
    return out


def main(refresh=False):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    download_filings(refresh)
    download_companyfacts(refresh)
    hist = build_history()
    xb, extra = build_xbrl_crosscheck(hist)
    hist["xbrl_crosscheck"] = xb
    hist["xbrl_extra"] = extra
    hist["buybacks_quarterly_text"] = parse_buybacks_from_text()
    hist["misc"] = parse_misc_text()
    hist["filings"] = [{"key": k, "accession": a, "document": d, "form": f, "period": p, "url": f"https://www.sec.gov/Archives/edgar/data/{int(CIK)}/{a.replace('-', '')}/{d}"} for k, a, d, f, p in FILINGS]
    HISTORY_FILE.write_text(json.dumps(hist, indent=1), encoding="utf-8")
    bad = [x for x in xb if not x["match"]]
    print("xbrl cross-checks:", len(xb), "mismatches:", len(bad))
    for x in bad:
        print("  MISMATCH", x)
    print("filing cross-check diffs:", hist["crosscheck"])
    if refresh or not MARKET_FILE.exists():
        snap = pull_market([TICKER] + list(PEERS))
        MARKET_FILE.write_text(json.dumps(snap, indent=1), encoding="utf-8")
        print("market snapshot", snap["retrieved_date"])
    if refresh or not PEER_FACTS_FILE.exists():
        PEER_FACTS_FILE.write_text(json.dumps(pull_peer_xbrl()), encoding="utf-8")


if __name__ == "__main__":
    main(refresh="--refresh" in sys.argv)
