"""Build SYF_Initiation_Report.docx from the recalculated model, then verify every number against the xlsx."""
import datetime as dt
import json
import re
import subprocess
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook

import build_model as bm

# ============================ CONFIG (edit here) ============================
AUTHOR = "Yazad Madan"
COMPANY = "Synchrony Financial"
TICKER = "SYF"
BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR.parent
MODEL_XLSX = OUT_DIR / "SYF_Model.xlsx"
LAYOUT_FILE = BASE_DIR / "data" / "model_layout.json"
MARKET_FILE = BASE_DIR / "data" / "market_snapshot.json"
REPORT_DOCX = OUT_DIR / "SYF_Initiation_Report.docx"
CHART_DIR = BASE_DIR / "charts"
SOFFICE = bm.SOFFICE
BODY_FONT, BODY_PT = "Calibri", 10
HEADING_PT = {"Heading 1": 14, "Heading 2": 12, "Heading 3": 10.5}
CHART_DPI = 200
C_MAIN, C_GRAY, C_ACC, C_LIGHT = "#1F3A5F", "#6B7280", "#B45309", "#9FB3CC"
# ==========================================================================

LAY = json.loads(LAYOUT_FILE.read_text(encoding="utf-8"))
HR, FR, V, SCN, CR, SR, SENS, I = LAY["HR"], LAY["FR"], LAY["V"], LAY["SCN"], LAY["CR"], LAY["SR"], LAY["SENS"], LAY["I"]
WB = load_workbook(MODEL_XLSX, data_only=True)
MARKET = json.loads(MARKET_FILE.read_text(encoding="utf-8"))
REG = []  # (description, sheet, cell, kind, digits, rendered string)


def cell(sheet, addr):
    return WB[sheet][addr].value


def render(v, kind, d):
    if kind == "usd":
        return f"-${abs(v):,.{d}f}" if v < 0 else f"${v:,.{d}f}"
    if kind == "mm":
        return f"-${abs(v):,.{d}f}mm" if v < 0 else f"${v:,.{d}f}mm"
    if kind == "absusd":
        return f"${abs(v):,.{d}f}"
    if kind == "absnum":
        return f"{abs(v):,.{d}f}"
    if kind == "abspct":
        return f"{abs(v) * 100:.{d}f}%"
    if kind == "pct":
        return f"{v * 100:.{d}f}%"
    if kind == "num":
        return f"{v:,.{d}f}"
    if kind == "bn":
        return f"${v / 1000:,.{d}f}bn"
    if kind == "x":
        return f"{v:.{d}f}x"
    if kind == "txt":
        return str(v)
    if kind == "date":
        return v.strftime("%d %B %Y").lstrip("0")
    raise ValueError(kind)


def N(sheet, addr, kind, d=1, track=True):
    v = cell(sheet, addr)
    s = render(v, kind, d)
    if track:
        REG.append((f"{sheet}!{addr}", sheet, addr, kind, d, s))
    return s


def Hc(key, period, kind, d=1):
    return N("Historical", f"{bm.HIST_COL[period]}{HR[key]}", kind, d)


def Hl(key, kind, d=1, which="H1"):
    c = bm.COL_H1 if which == "H1" else bm.COL_LTM
    return N("Historical", f"{c}{HR[key]}", kind, d)


def Fb(key, col_letter, kind, d=1, scen="Base"):
    return N("Scn_" + scen, f"{col_letter}{SCN[scen + '_F'][key]}", kind, d)


def Vb(key, kind, d=1, scen="Base", col_letter="B"):
    return N("Scn_" + scen, f"{col_letter}{SCN[scen][key]}", kind, d)


def Cc(t, colkey, kind, d=1):
    return N("Comps", f"{bm.COMPS_COL[colkey]}{CR[t]}", kind, d)


def raw_hist(key, period):
    return cell("Historical", f"{bm.HIST_COL[period]}{HR[key]}")


def raw_f(key, col_letter, scen="Base"):
    return cell("Scn_" + scen, f"{col_letter}{SCN[scen + '_F'][key]}")


# ---------------------------------------------------------------- charts
def chart_style(ax):
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color="#E5E7EB", linewidth=0.7)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=8)


def chart_credit_margin():
    yrs = [2020, 2021, 2022, 2023, 2024, 2025]
    per = [f"FY{y}" for y in yrs]
    nim = [raw_hist("nim", p) * 100 for p in per]
    nco = [raw_hist("nco_rate", p) * 100 for p in per]
    rsa = [raw_hist("rsa_pct", p) * 100 for p in per]
    fx = ["C", "D", "E"]
    fyr = [2026, 2027, 2028]
    nim_f = [raw_f("nim", c) * 100 for c in fx]
    nco_f = [raw_f("d_nco", c) * 100 for c in fx]
    rsa_f = [raw_f("rsa_pct", c) * 100 for c in fx]
    fig, ax = plt.subplots(figsize=(6.6, 3.2))
    for series, fser, color, lab in [(nim, nim_f, C_MAIN, "Net interest margin"), (rsa, rsa_f, C_ACC, "RSA % of average loans"), (nco, nco_f, C_GRAY, "Net charge-off rate")]:
        ax.plot(yrs, series, marker="o", color=color, linewidth=1.8, label=lab)
        ax.plot([2025] + fyr, [series[-1]] + fser, marker="o", color=color, linewidth=1.8, linestyle="--", markerfacecolor="white")
        ax.annotate(f"{series[-1]:.1f}%", (2025, series[-1]), textcoords="offset points", xytext=(0, 6), ha="center", fontsize=7, color=color)
    ax.set_ylabel("% (annual)", fontsize=8)
    ax.set_xticks(yrs + fyr)
    ax.set_xticklabels([f"FY{y}" if y <= 2025 else f"FY{y}E" for y in yrs + fyr], fontsize=7.5)
    chart_style(ax)
    ax.legend(fontsize=7.5, frameon=False, loc="center left", bbox_to_anchor=(0.0, 0.62))
    fig.tight_layout()
    path = CHART_DIR / "chart1_margin_credit.png"
    fig.savefig(path, dpi=CHART_DPI)
    plt.close(fig)
    return path


def chart_growth():
    per = [f"FY{y}" for y in (2021, 2022, 2023, 2024, 2025)]
    lg = [raw_hist("loan_gr", p) * 100 for p in per] + [raw_f("d_loan_g", c) * 100 for c in "CDE"]
    pv = [raw_hist("pv_gr", p) * 100 for p in per] + [raw_f("d_pv_g", c) * 100 for c in "CDE"]
    labels = [p for p in per] + ["FY2026E", "FY2027E", "FY2028E"]
    fig, ax = plt.subplots(figsize=(6.6, 3.0))
    x = list(range(len(labels)))
    w = 0.38
    b1 = ax.bar([i - w / 2 for i in x], lg, w, color=[C_MAIN] * 5 + [C_LIGHT] * 3, label="Loan receivables growth (year end)")
    b2 = ax.bar([i + w / 2 for i in x], pv, w, color=[C_ACC] * 5 + ["#E3B98F"] * 3, label="Purchase volume growth")
    ax.axhline(0, color="#374151", linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=7.5)
    ax.set_ylabel("YoY growth, %", fontsize=8)
    chart_style(ax)
    ax.legend(fontsize=7.5, frameon=False, loc="upper right")
    fig.tight_layout()
    path = CHART_DIR / "chart2_growth.png"
    fig.savefig(path, dpi=CHART_DPI)
    plt.close(fig)
    return path


def chart_peers():
    tick = [bm.TICKER] + bm.PEERS
    xs = [cell("Comps", f"{bm.COMPS_COL['rotce']}{CR[t]}") * 100 for t in tick]
    ys = [cell("Comps", f"{bm.COMPS_COL['ptbv']}{CR[t]}") for t in tick]
    fig, ax = plt.subplots(figsize=(6.6, 3.2))
    for t, x, y in zip(tick, xs, ys):
        ax.scatter(x, y, s=55, color=C_ACC if t == bm.TICKER else C_MAIN, zorder=3)
        ax.annotate(t, (x, y), textcoords="offset points", xytext=(6, 4), fontsize=8)
    ax.set_xlabel("ROTCE proxy: TTM diluted EPS / tangible book value per share, %", fontsize=8)
    ax.set_ylabel("Price / tangible book (x)", fontsize=8)
    chart_style(ax)
    ax.grid(axis="x", color="#E5E7EB", linewidth=0.7)
    fig.tight_layout()
    path = CHART_DIR / "chart3_ptbv_rotce.png"
    fig.savefig(path, dpi=CHART_DPI)
    plt.close(fig)
    return path


def chart_scenarios():
    names = ["Bear", "Base", "Bull"]
    ri = [cell("Scn_" + n, f"B{SCN[n]['v12']}") for n in names]
    pt = [cell("Scn_" + n, f"B{SCN[n]['v_pt']}") for n in names]
    bl = [cell("Scn_" + n, f"B{SCN[n]['pt_raw']}") for n in names]
    price = cell("Summary", f"B{SR['price']}")
    fig, ax = plt.subplots(figsize=(6.6, 3.0))
    x = list(range(3))
    w = 0.26
    ax.bar([i - w for i in x], ri, w, color=C_LIGHT, label="Residual income value (12m)")
    ax.bar(x, pt, w, color=C_GRAY, label="Justified P/TBV value")
    ax.bar([i + w for i in x], bl, w, color=C_MAIN, label="Blended value")
    for i, v in enumerate(bl):
        ax.annotate(f"${v:,.0f}", (i + w, v), textcoords="offset points", xytext=(0, 3), ha="center", fontsize=8)
    ax.axhline(price, color=C_ACC, linewidth=1.4, linestyle="--", label=f"Share price ${price:,.2f}")
    ax.set_xticks(x)
    ax.set_xticklabels(names, fontsize=9)
    ax.set_ylabel("$ per share", fontsize=8)
    chart_style(ax)
    ax.legend(fontsize=7.5, frameon=False, loc="upper left")
    fig.tight_layout()
    path = CHART_DIR / "chart4_scenarios.png"
    fig.savefig(path, dpi=CHART_DPI)
    plt.close(fig)
    return path


# ---------------------------------------------------------------- docx helpers
def set_cell_text(c, text, bold=False, align=None, size=8.5):
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    if align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_table(doc, header, rows, widths=None, first_col_w=2.4, size=8.5, bold_rows=()):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        set_cell_text(t.rows[0].cells[i], h, bold=True, align="l" if i == 0 else "r", size=size)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            set_cell_text(cells[i], v, bold=(ri in bold_rows), align="l" if i == 0 else "r", size=size)
    if widths is None:
        rest = (6.6 - first_col_w) / max(1, len(header) - 1)
        widths = [first_col_w] + [rest] * (len(header) - 1)
    if len(rows) <= 16:
        for row in t.rows[:-1]:
            for c in row.cells:
                for pp in c.paragraphs:
                    pp.paragraph_format.keep_with_next = True
    t.autofit = False
    for i, w in enumerate(widths):
        t.columns[i].width = Inches(w)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def para(doc, text, bold_lead=None, size=None, italic=False, space_after=6, keep=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if bold_lead:
        r = p.add_run(bold_lead + " ")
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        for run in p.runs:
            run.font.size = Pt(size)
    if keep:
        p.paragraph_format.keep_with_next = True
    return p


def source_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(8)
    return p


def add_picture(doc, path, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Inches(6.3))
    c = doc.add_paragraph()
    c.paragraph_format.space_after = Pt(8)
    r = c.add_run(caption)
    r.italic = True
    r.font.size = Pt(8)


def add_page_number(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Illustrative equity research sample. Not investment advice. Page ")
    r.font.size = Pt(8)
    run = p.add_run()
    run.font.size = Pt(8)
    for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def style_document(doc):
    st = doc.styles["Normal"]
    st.font.name = BODY_FONT
    st.font.size = Pt(BODY_PT)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    for name, size in HEADING_PT.items():
        h = doc.styles[name]
        h.font.name = BODY_FONT
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = None  # no theme colour: size and bold only
        rf = h.element.rPr.find(qn("w:rFonts"))
        for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf is not None and rf.get(qn(att)) is not None:
                del rf.attrib[qn(att)]
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        rpr = h.element.rPr
        for el in rpr.findall(qn("w:color")):
            rpr.remove(el)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = Inches(0.95)
    sec.top_margin, sec.bottom_margin = Inches(0.8), Inches(0.8)
    add_page_number(sec)


# ---------------------------------------------------------------- report body
def build_doc():
    CHART_DIR.mkdir(exist_ok=True)
    c1, c2, c3, c4 = chart_credit_margin(), chart_growth(), chart_peers(), chart_scenarios()
    doc = Document()
    style_document(doc)
    price_date = N("Summary", f"B{SR['pdate']}", "date", track=False)
    price_date = cell("Summary", f"B{SR['pdate']}").strftime("%d %B %Y").lstrip("0")
    report_date = dt.date.fromisoformat(MARKET["retrieved_date"]).strftime("%d %B %Y").lstrip("0")
    S = lambda key, kind, d=1: N("Summary", f"B{SR[key]}", kind, d)
    SF = lambda key, c, kind, d=1: N("Summary", f"{c}{SR['f_' + key]}", kind, d)
    rating = S("rating", "txt")
    pt = S("pt", "usd", 0)
    price = S("price", "usd", 2)
    upside = S("upside", "pct", 1)
    tr = S("tr", "pct", 1)

    # ----- title block
    p = doc.add_paragraph()
    r = p.add_run(f"Equity Research  |  Initiation of Coverage  |  {report_date}")
    r.font.size = Pt(9)
    p.paragraph_format.space_after = Pt(2)
    p = doc.add_paragraph()
    r = p.add_run(f"{COMPANY} (NYSE: {TICKER})")
    r.bold = True
    r.font.size = Pt(20)
    p.paragraph_format.space_after = Pt(0)
    p = doc.add_paragraph()
    pe_now = S("pe_now", "x", 1)
    pbv_now = S("ptbv_now", "x", 1)
    r = p.add_run(f"A card lender earning {SF('rotce', 'C', 'pct', 0)} on tangible equity, priced at {pbv_now} tangible book")
    r.font.size = Pt(12)
    p.paragraph_format.space_after = Pt(6)

    add_table(doc, ["Rating", "12-month price target", "Share price", "Upside to target", "Expected total return"],
              [[rating, pt, price, upside, tr]], widths=[1.3, 1.4, 1.3, 1.3, 1.3], size=10, bold_rows=(0,))
    mcap = Cc(TICKER, "mcap", "bn", 1)
    add_table(doc, ["Key data", "", "Base-case estimates", "FY2025A", "FY2026E", "FY2027E", "FY2028E"],
              [["Market cap", mcap, "Diluted EPS ($)", SF("eps", "B", "usd", 2), SF("eps", "C", "usd", 2), SF("eps", "D", "usd", 2), SF("eps", "E", "usd", 2)],
               ["Shares out (mm)", N("Inputs_Sources", I["shares"].split("!")[1].replace("$", ""), "num", 1), "Net earnings to common ($mm)", SF("nic", "B", "num", 0), SF("nic", "C", "num", 0), SF("nic", "D", "num", 0), SF("nic", "E", "num", 0)],
               ["P/E on FY2027E EPS", pe_now, "ROTCE", SF("rotce", "B", "pct", 1), SF("rotce", "C", "pct", 1), SF("rotce", "D", "pct", 1), SF("rotce", "E", "pct", 1)],
               ["P/TBV (Jun-26)", pbv_now, "Net charge-off rate", SF("d_nco", "B", "pct", 2), SF("d_nco", "C", "pct", 2), SF("d_nco", "D", "pct", 2), SF("d_nco", "E", "pct", 2)],
               ["Cost of equity", S("ke", "pct", 1), "CET1 ratio", SF("cet1r", "B", "pct", 1), SF("cet1r", "C", "pct", 1), SF("cet1r", "D", "pct", 1), SF("cet1r", "E", "pct", 1)]],
              widths=[1.15, 0.7, 1.7, 0.75, 0.75, 0.75, 0.75], size=8)
    cons26 = N("Summary", f"C{SR['cons']}", "usd", 2)
    cons27 = N("Summary", f"D{SR['cons']}", "usd", 2)
    para(doc, f"We initiate coverage of Synchrony Financial with an {rating} rating and a 12-month price target of {pt}, {upside} above the {price} closing price on {price_date}. "
              f"Our base case has diluted EPS of {SF('eps', 'C', 'usd', 2)} in FY2026E and {SF('eps', 'D', 'usd', 2)} in FY2027E, against Yahoo Finance consensus of {cons26} and {cons27}. "
              f"The stock trades at {pe_now} our FY2027E EPS and {pbv_now} June 2026 tangible book value, while we forecast ROTCE of {SF('rotce', 'D', 'pct', 0)} in FY2027E. "
              f"The target blends a residual income model with a justified price-to-tangible-book value. Our bear and bull values are {S('bear', 'usd', 0)} and {S('bull', 'usd', 0)} per share. "
              f"The rating and target follow mechanically from the model: the rating rule is Overweight at or above {N('Inputs_Sources', I['ow'].split('!')[1].replace('$', ''), 'pct', 0)} expected total return.",
         bold_lead="Summary.")
    para(doc, "Forecasts, price target and rating are the author's illustrative analysis for portfolio purposes, prepared from public filings and market data retrieved on the date above (prices are the last completed close). "
              "They are not investment advice, and the author has no affiliation with any bank or research department. See the appendix for sources and the full disclaimer.",
         size=8, italic=True)

    # ----- 1 thesis
    doc.add_heading("1. Investment thesis", level=1)
    lgq2 = Hc("loan_gr", "Q2_2026", "pct", 1)
    para(doc, f"Synchrony's return on tangible equity has averaged {N('Historical', 'B' + str(HR['rotce_avg']), 'pct', 1)} over FY2021 to FY2025 and was {Hc('rotce', 'FY2025', 'pct', 1)} in FY2025, yet the stock trades at "
              f"{Cc(bm.TICKER, 'pe_ttm', 'x', 1)} trailing earnings and {Cc(bm.TICKER, 'ptbv', 'x', 1)} tangible book. The peer median is {Cc('median', 'pe27', 'x', 1)} FY2027E consensus earnings and {Cc('median', 'ptbv', 'x', 1)} tangible book, while "
              f"SYF's ROTCE proxy of {Cc(bm.TICKER, 'rotce', 'pct', 1)} sits above the peer median of {Cc('median', 'rotce', 'pct', 1)}. If our FY2027E ROTCE of {SF('rotce', 'D', 'pct', 1)} were sustained, the multiple consistent with it would be "
              f"{S('ptbv27', 'x', 2)} tangible book. Even a through-cycle ROTCE of {S('rot', 'pct', 0)} supports {S('ptbv', 'x', 2)}, so the market is capitalizing a return well below what the company is earning "
              f"(Comps sheet; market data from Yahoo Finance, retrieved {MARKET['retrieved_date']}).",
         bold_lead="1. Earnings power is not in the price.")
    para(doc, f"The net charge-off rate fell from {Hc('nco_rate', 'FY2024', 'pct', 2)} in FY2024 to {Hc('nco_rate', 'FY2025', 'pct', 2)} in FY2025 and {Hc('nco_rate', 'Q2_2026', 'pct', 2)} in Q2 2026, and 30+ day delinquencies were "
              f"{Hc('dq30', 'Q2_2026', 'pct', 2)} at June 2026 versus {Hc('dq30', 'Q2_2025', 'pct', 2)} a year earlier. Retailer share arrangements (RSAs) cushion the cycle: across FY2020 to FY2025 the RSA charge moved "
              f"{N('Historical', 'B' + str(HR['slope']), 'absnum', 2)} points for each point of charge-off rate, in the opposite direction (correlation {N('Historical', 'B' + str(HR['correl']), 'num', 2)}), so about {N('Historical', 'B' + str(HR['slope']), 'abspct', 0)} of a credit swing is shared with partners "
              f"(Historical sheet regression; 10-K FY2025 and 10-Q Q2 2026).",
         bold_lead="2. Credit has normalized, and partners share the pain.")
    para(doc, f"We forecast net earnings to common of {SF('nic', 'C', 'mm', 0)}, {SF('nic', 'D', 'mm', 0)} and {SF('nic', 'E', 'mm', 0)} in FY2026E to FY2028E, essentially flat, but EPS grows from {SF('eps', 'C', 'usd', 2)} to {SF('eps', 'E', 'usd', 2)} because "
              f"repurchases of {Fb('bb', 'C', 'mm', 0)}, {Fb('bb', 'D', 'mm', 0)} and {Fb('bb', 'E', 'mm', 0)} cut the diluted share count to {Fb('sh_dil', 'E', 'mm', 0)} by FY2028E from {Fb('sh_dil', 'B', 'mm', 0)} in FY2025. "
              f"The board approved a new repurchase program of up to $6.5 billion in April 2026 with no expiry, $5.7 billion of which remained at June 30, 2026, and raised the quarterly dividend to $0.34 (10-Q Q2 2026). "
              f"CET1 stays at {Fb('cet1r', 'E', 'pct', 1)} or higher in our base case through FY2028E.",
         bold_lead="3. Capital return, not net income growth, drives EPS.")
    para(doc, f"Purchase volume grew {Hc('pv_gr', 'Q2_2026', 'pct', 1)} year on year in Q2 2026 after {Hc('pv_gr', 'FY2025', 'pct', 1)} in FY2025, and loan receivables were up {lgq2} year on year at June 2026 after falling {Hc('loan_gr', 'FY2025', 'abspct', 1)} in FY2025. "
              f"New programs (OnePay at Walmart, launched September 2025), the Versatile Credit acquisition (October 2025) and the announced purchase of the Lowe's commercial co-branded card portfolio (about $0.8 billion of receivables) add to the base. We model year-end loan growth of "
              f"{Fb('d_loan_g', 'C', 'pct', 1)}, {Fb('d_loan_g', 'D', 'pct', 1)} and {Fb('d_loan_g', 'E', 'pct', 1)} (10-K FY2025 and 10-Q Q2 2026, MD&A).",
         bold_lead="4. Growth is re-accelerating from a low base.")

    # ----- 2 debates
    doc.add_heading("2. Key debates", level=1)
    para(doc, f"The bear case is that charge-offs re-accelerate in a downturn. Our bear scenario runs the net charge-off rate to {N('Scenarios', 'C' + str(bm_row('nco', 'bear')), 'pct', 1)} in FY2027E and builds the allowance to "
              f"{N('Scenarios', 'C' + str(bm_row('cov', 'bear')), 'pct', 1)} of loans, which cuts FY2027E EPS to {Fb('eps', 'D', 'usd', 2, 'Bear')}. We think much of that exposure is priced: even that scenario is worth "
              f"{S('bear', 'usd', 0)} per share on our blended method because the through-cycle ROTCE assumption falls to {N('Scenarios', 'B' + str(bm_row('rot', 'bear')), 'pct', 0)}.",
         bold_lead="Is the credit cycle really behind us?")
    para(doc, f"RSA expense rose {N('Historical', bm.HIST_COL['FY2025'] + str(HR['rsa']), 'mm', 0)} in FY2025, up 17.6%, as lower charge-offs raised partner payouts and product, pricing and policy changes shifted economics (10-K FY2025, MD&A). "
              f"Our RSA line is regression-fitted to the charge-off rate plus a {N('Scenarios', 'B' + str(LAY['S']['c_rsa_res']), 'pct', 2)} structural residual, so improving credit is only partly retained. In FY2028E the RSA charge is "
              f"{Fb('rsa_pct', 'E', 'pct', 2)} of average loans, up from {Hc('rsa_pct', 'FY2025', 'pct', 2)} in FY2025.",
         bold_lead="How much of the credit improvement does Synchrony keep?")
    para(doc, f"Late fees were {'$2.3'} billion in FY2025, down from $2.7 billion in FY2023, after the company pre-emptively changed pricing for the CFPB's 2024 late-fee rule, which was vacated on April 15, 2025 (10-K FY2025). "
              f"Our yield assumption of {Fb('d_yield', 'C', 'pct', 2)} to {Fb('d_yield', 'E', 'pct', 2)} assumes no further material fee compression; the bear case assumes {N('Scenarios', 'D' + str(bm_row('yield', 'bear')), 'pct', 1)} in FY2028E. "
              f"A renewed pricing rule is the main regulatory tail risk.",
         bold_lead="What is the regulatory overhang worth?")
    para(doc, f"Base-case payout of dividends plus buybacks runs above net earnings, so CET1 drifts from {Fb('cet1r', 'B', 'pct', 1)} at December 2025 (recast) to {Fb('cet1r', 'E', 'pct', 1)} in FY2028E. "
              f"The new $500 million Series C preferred (June 2026) adds Tier 1 capacity but not CET1 (10-Q Q2 2026). Repurchase capacity is the swing factor for EPS: our base case buys back "
              f"{Fb('bb', 'C', 'mm', 0)} in FY2026E versus {Fb('bb', 'B', 'mm', 0)} in FY2025.",
         bold_lead="Can the buyback pace be sustained?")

    # ----- 3 earnings drivers
    doc.add_heading("3. Earnings drivers", level=1)
    doc.add_heading("3.1 Net interest margin and retailer share arrangements", level=2)
    para(doc, f"Net interest margin was {Hc('nim', 'FY2025', 'pct', 2)} in FY2025 (reported {Hc('nim_rep', 'FY2025', 'pct', 2)}) and {Hc('nim', 'Q2_2026', 'pct', 2)} in Q2 2026. Loan yield was {Hc('yield', 'FY2025', 'pct', 2)} and cost of funds "
              f"{Hc('cof', 'FY2025', 'pct', 2)}; by Q2 2026 cost of funds had fallen to {Hc('cof', 'Q2_2026', 'pct', 2)} while loan yield was {Hc('yield', 'Q2_2026', 'pct', 2)}. "
              f"We hold NIM roughly flat at {Fb('nim', 'C', 'pct', 2)}, {Fb('nim', 'D', 'pct', 2)} and {Fb('nim', 'E', 'pct', 2)}: cost of funds of {Fb('d_cof', 'C', 'pct', 2)} to {Fb('d_cof', 'E', 'pct', 2)} offsets a slightly lower yield on liquidity assets "
              f"(cash and securities yield {Fb('d_oth_y', 'C', 'pct', 2)} in FY2026E). Each 50bp increase in cost of funds costs {N('Sensitivity', 'B' + str(SENS['T4']), 'absusd', 2)} of FY2027E EPS. "
              f"RSAs absorb a growing share of the spread: the RSA charge is {Fb('rsa', 'C', 'mm', 0)} in FY2026E ({Fb('rsa_pct', 'C', 'pct', 2)} of average loans) versus {Fb('rsa', 'B', 'mm', 0)} in FY2025 (10-K FY2025; model).",
         )
    add_picture(doc, c1, "Figure 1: Net interest margin, RSA and net charge-off rate, FY2020 to FY2025 actual and FY2026E to FY2028E base case (dashed). Source: SYF 10-K filings, model.")
    hist_rows = []
    for lab, key, kind, d in [("Loan yield", "yield", "pct", 2), ("Cost of funds", "cof", "pct", 2), ("Net interest margin (recomputed)", "nim", "pct", 2), ("Net interest margin (reported)", "nim_rep", "pct", 2),
                              ("RSA % of average loans", "rsa_pct", "pct", 2), ("Net charge-off rate", "nco_rate", "pct", 2), ("30+ day delinquency", "dq30", "pct", 2), ("Allowance coverage", "cov", "pct", 2), ("Efficiency ratio", "eff", "pct", 1)]:
        hist_rows.append([lab] + [Hc(key, p, kind, d) for p in ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "Q2_2026"]])
    add_table(doc, ["Historical ratios", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "Q2 2026"], hist_rows, first_col_w=2.1, size=8)
    source_note(doc, "Source: SYF 10-K FY2021, FY2023, FY2025 and 10-Q Q2 2026 (recomputed from reported line items; Q2 2026 ratios annualized 365/days). Model: Historical sheet.")

    doc.add_heading("3.2 Credit quality and the reserve build", level=2)
    para(doc, f"Charge-offs peaked at {Hc('nco_rate', 'FY2024', 'pct', 2)} in FY2024 after {Hc('nco_rate', 'FY2023', 'pct', 2)} in FY2023. Under CECL, provision equals charge-offs plus the change in the allowance. The allowance ended FY2025 at "
              f"{Hc('acl', 'FY2025', 'mm', 0)} ({Hc('cov', 'FY2025', 'pct', 2)} of loans), a reserve release of {Hc('res_b', 'FY2025', 'absnum', 0)}mm during the year, and was {Hc('cov', 'Q2_2026', 'pct', 2)} at June 2026. "
              f"We assume coverage of {Fb('d_cov', 'C', 'pct', 2)} at FY2026E, so the allowance grows only with loans: reserve builds of {Fb('res_b', 'C', 'mm', 0)}, {Fb('res_b', 'D', 'mm', 0)} and {Fb('res_b', 'E', 'mm', 0)}. "
              f"Provision is {Fb('prov', 'C', 'mm', 0)} in FY2026E, {Fb('prov', 'D', 'mm', 0)} in FY2027E and {Fb('prov', 'E', 'mm', 0)} in FY2028E. "
              f"A 100bp higher charge-off rate costs {N('Sensitivity', 'B' + str(SENS['T4'] + 1), 'absusd', 2)} of FY2027E EPS after the RSA offset (10-K FY2025; model).")

    doc.add_heading("3.3 Volume growth", level=2)
    para(doc, f"Purchase volume was {Hc('pv', 'FY2025', 'mm', 0)} in FY2025 (essentially flat) and {Hl('pv', 'mm', 0)} in H1 2026, up 6.9% year on year, led by Diversified & Value (+10.3%) and Digital (+8.7%) (10-Q Q2 2026). "
              f"We forecast purchase volume growth of {Fb('d_pv_g', 'C', 'pct', 1)}, {Fb('d_pv_g', 'D', 'pct', 1)} and {Fb('d_pv_g', 'E', 'pct', 1)}, and loan growth that lags spend because higher payment rates absorb part of it "
              f"(purchase volume is {Fb('pv_int', 'C', 'x', 2)} average loans in FY2026E versus {Fb('pv_int', 'B', 'x', 2)} in FY2025).")
    add_picture(doc, c2, "Figure 2: Loan receivables growth versus purchase volume growth. Lighter bars are base-case forecasts. Source: SYF 10-K filings, model.")

    doc.add_heading("3.4 Capital return", level=2)
    para(doc, f"SYF repurchased {Hc('bb', 'FY2025', 'mm', 0)} of stock in FY2025 and {Hl('bb', 'mm', 0)} in H1 2026, reducing common shares from {Hc('sh', 'FY2025', 'num', 1)}mm at December 2025 to "
              f"{N('Inputs_Sources', I['shares'].split('!')[1].replace('$', ''), 'num', 1)}mm at the July 2026 10-Q cover date. CET1 was {Hc('cet1r', 'Q2_2026', 'pct', 1)} at June 30, 2026. "
              f"Our base case pays dividends per share of {Fb('dps', 'C', 'usd', 2)}, {Fb('dps', 'D', 'usd', 2)} and {Fb('dps', 'E', 'usd', 2)} and total payout of {Fb('payout', 'C', 'pct', 0)}, {Fb('payout', 'D', 'pct', 0)} and {Fb('payout', 'E', 'pct', 0)} of net earnings to common, "
              f"drawing the CET1 ratio down to {Fb('cet1r', 'E', 'pct', 1)} (10-Q Q2 2026; model).")

    # ----- 4 forecasts
    doc.add_heading("4. Forecasts", level=1)
    para(doc, "The model is annual, driver-based and anchored to the June 2026 balance sheet and H1 2026 results. Loans, the allowance, tangible common equity, CET1 and share count tie to roll-forwards; "
              "we do not forecast a full balance sheet. Historical columns are as reported. Base case shown.")
    cols_h = ["FY2023", "FY2024", "FY2025"]
    rows = []
    spec = [("Loan receivables, year end ($mm)", "loans_end", "loans_end", "num", 0), ("Purchase volume ($mm)", "pv", "pv", "num", 0), ("Net interest income ($mm)", "nii", "nii", "num", 0), ("Net interest margin", "nim", "nim", "pct", 2),
            ("Retailer share arrangements ($mm)", "rsa", "rsa", "num", 0), ("Provision for credit losses ($mm)", "prov", "prov", "num", 0), ("Net charge-off rate", "nco_rate", "d_nco", "pct", 2),
            ("Other expense ($mm)", "oe", "oe", "num", 0), ("Efficiency ratio", "eff", "eff", "pct", 1), ("Net earnings to common ($mm)", "nic", "nic", "num", 0), ("Diluted EPS ($)", "eps", "eps", "usd", 2),
            ("Dividends per share ($)", "dps", "dps", "usd", 2), ("ROTCE", "rotce", "rotce", "pct", 1), ("CET1 ratio", "cet1r", "cet1r", "pct", 1)]
    for lab, hk, fk, kind, d in spec:
        row = [lab] + [Hc(hk, p, kind, d) for p in cols_h] + [Fb(fk, c, kind, d) for c in "CDE"]
        rows.append(row)
    add_table(doc, ["", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E"], rows, first_col_w=2.3, size=8)
    source_note(doc, "Source: SYF filings (history); model, base case (forecast). Historical ROTCE and CET1 use the as-originally-reported basis; forecasts use the recast basis after the Q2 2026 software reclassification.")
    srows = []
    for n in ("Bear", "Base", "Bull"):
        srows.append([n] + [Fb("eps", c, "usd", 2, n) for c in "CDE"] + [Fb("d_nco", "D", "pct", 2, n), Fb("cet1r", "E", "pct", 1, n)])
    add_table(doc, ["Scenario", "EPS FY2026E", "EPS FY2027E", "EPS FY2028E", "NCO rate FY2027E", "CET1 FY2028E"], srows, first_col_w=1.4, size=8.5)
    source_note(doc, "Source: model, Scenarios sheet. Bear assumes a 2027 credit downturn, weaker growth, lower yield, reserve build and reduced buybacks; bull assumes faster growth, lower charge-offs and higher buybacks.")

    vc26 = N("Summary", f"C{SR['vs_cons']}", "abspct", 1)
    vc27 = N("Summary", f"D{SR['vs_cons']}", "abspct", 1)
    para(doc, f"Our FY2026E EPS of {SF('eps', 'C', 'usd', 2)} is {vc26} below consensus of {cons26}, and FY2027E EPS of {SF('eps', 'D', 'usd', 2)} is {vc27} below {cons27}. The gap reflects our reserve build and expense growth assumptions for H2 2026 and FY2027E. "
              f"H1 2026 was reported; H2 2026E below is the FY2026E forecast less H1 (consensus source: Yahoo Finance).",
         bold_lead="Versus consensus.")
    bridge = []
    for lab, key, kind in [("Net interest income", "nii", "mm"), ("Retailer share arrangements", "rsa", "mm"), ("Provision for credit losses", "prov", "mm"), ("  of which net charge-offs", "acl_nco", "mm"),
                           ("Other income", "oi", "mm"), ("Other expense", "oe", "mm"), ("Earnings before income taxes", "pretax", "mm"), ("Net earnings to common", "nic", "mm"), ("Diluted EPS", "eps", "usd")]:
        d_ = 2 if kind == "usd" else 0
        bridge.append([lab, N("Scn_Base", f"G{SCN['Base_F'][key]}", kind, d_), N("Scn_Base", f"H{SCN['Base_F'][key]}", kind, d_), Fb(key, "C", kind, d_), Fb(key, "B", kind, d_)])
    add_table(doc, ["FY2026E bridge, base case", "H1 2026A", "H2 2026E (implied)", "FY2026E", "FY2025A"], bridge, first_col_w=2.4, size=8.5)
    source_note(doc, "Source: 10-Q Q1 and Q2 2026 (H1); model, Forecast columns. H1 EPS is the sum of reported quarterly diluted EPS; the reported six-month figure differs by $0.01 because of share count rounding.")
    para(doc, f"Our rating changes mechanically: with the target unchanged, the expected total return falls below the Overweight threshold if the shares rise above {S('px_ow', 'usd', 2)}, and turns Underweight above {S('px_uw', 'usd', 2)}. "
              "Items that would move our estimates are quarterly charge-off and delinquency trends, the pace of buybacks against the capital plan, RSA payouts on new programs, and any renewed pricing regulation.",
         bold_lead="What would change our view.")

    # ----- 5 valuation
    doc.add_heading("5. Valuation", level=1)
    para(doc, "We do not use a free-cash-flow DCF. For a lender, debt is raw material, not financing: interest expense, deposit growth and loan originations are operating flows, so unlevered free cash flow is not well defined, "
              "and growth in loans consumes regulatory capital rather than generating cash. We value the equity directly with a residual income model, where value equals tangible common equity plus the present value of earnings "
              "above the cost of equity, and cross-check with the justified price-to-tangible-book multiple implied by ROTCE versus cost of equity.",
         bold_lead="Approach.")
    ri_rows = []
    ri_rows.append(["Net earnings to common ($mm)"] + [N("Scn_Base", f"{c}{SCN['Base']['nic']}", "num", 0) for c in "BCD"])
    ri_rows.append(["Beginning TCE ($mm)"] + [N("Scn_Base", f"{c}{SCN['Base']['tceb']}", "num", 0) for c in "BCD"])
    ri_rows.append(["Equity charge at cost of equity ($mm)"] + [N("Scn_Base", f"{c}{SCN['Base']['chg']}", "num", 0) for c in "BCD"])
    ri_rows.append(["Residual income ($mm)"] + [N("Scn_Base", f"{c}{SCN['Base']['ri']}", "num", 0) for c in "BCD"])
    ri_rows.append(["Present value ($mm)"] + [N("Scn_Base", f"{c}{SCN['Base']['pv']}", "num", 0) for c in "BCD"])
    add_table(doc, ["Residual income, base case", "H2 2026E", "FY2027E", "FY2028E"], ri_rows, first_col_w=2.8, size=8.5)
    para(doc, f"Cost of equity is {S('ke', 'pct', 1)}: a {N('Inputs_Sources', I['rf'].split('!')[1].replace('$', ''), 'pct', 2)} 10-year Treasury yield plus a Yahoo Finance beta of {N('Inputs_Sources', I['beta'].split('!')[1].replace('$', ''), 'num', 2)} "
              f"times a {N('Inputs_Sources', I['erp'].split('!')[1].replace('$', ''), 'pct', 1)} equity risk premium (the premium and terminal assumptions are the author's). The valuation base is tangible common equity of {Vb('tce0', 'mm', 0)} at June 30, 2026. "
              f"Terminal growth is {S('g', 'pct', 1)} and terminal ROTCE {S('rot', 'pct', 0)}, which gives a justified price-to-tangible-book of {S('ptbv', 'x', 2)}. Residual income of {Vb('sum_pv', 'mm', 0)} (present value) over the explicit period and a terminal value with a present value of "
              f"{Vb('pv_tv', 'mm', 0)} give an equity value of {Vb('eqv', 'mm', 0)}, or {Vb('v0', 'usd', 2)} per share at June 30, 2026. Carried to the price date and rolled forward 12 months net of dividends, the residual income value is "
              f"{S('v12', 'usd', 2)} per share (model: Valuation sheet).",
         bold_lead="Residual income.")
    para(doc, f"Applying the justified {S('ptbv', 'x', 2)} to our horizon tangible book value of {Vb('tbv_h', 'usd', 2)} per share gives {S('vpt', 'usd', 2)}. This is deliberately below the residual income value because it capitalizes only the "
              f"through-cycle ROTCE and ignores the excess returns we forecast for FY2026E to FY2028E. We weight the two methods equally, which gives a blended value of {Vb('pt_raw', 'usd', 2)} and a price target of {pt}. "
              f"The target implies {Vb('pe27', 'x', 1)} FY2027E EPS and {Vb('ptbv_pt', 'x', 1)} June 2026 tangible book, against {pe_now} and {pbv_now} today.",
         bold_lead="Justified P/TBV and price target.")
    add_picture(doc, c4, "Figure 3: Value per share by scenario and method versus the current share price. Source: model, Scenarios and Scn_ sheets.")
    srow = []
    for n in ("Bear", "Base", "Bull"):
        vs_txt = S({"Bear": "bear_vs", "Bull": "bull_vs", "Base": "base_vs"}[n], "pct", 0)
        srow.append([n, N("Scn_" + n, f"B{SCN[n]['rot']}", "pct", 0), N("Scn_" + n, f"B{SCN[n]['v12']}", "usd", 0), N("Scn_" + n, f"B{SCN[n]['v_pt']}", "usd", 0), N("Scn_" + n, f"B{SCN[n]['pt_raw']}", "usd", 0), vs_txt])
    add_table(doc, ["Scenario", "Terminal ROTCE", "Residual income value", "Justified P/TBV value", "Blended value", "Blended value vs share price"], srow, first_col_w=1.2, size=8.5)
    source_note(doc, "Source: model. Values are 12-month forward per-share values; the price target is the base-case blended value rounded to the nearest dollar.")

    doc.add_heading("Sensitivity", level=2)
    t2 = SENS["T2"]
    hdr = [""] + [N("Sensitivity", f"{c}{t2}", "pct", 0) for c in "BCDEF"]
    rows = []
    for i in range(5):
        rr = t2 + 1 + i
        rows.append([N("Sensitivity", f"A{rr}", "pct", 1)] + [N("Sensitivity", f"{c}{rr}", "usd", 0) for c in "BCDEF"])
    para(doc, "Blended value per share (before rounding) by cost of equity (rows) and terminal ROTCE (columns), base case:", keep=True, space_after=2)
    add_table(doc, ["Cost of equity \\ terminal ROTCE"] + hdr[1:], rows, first_col_w=2.0, size=8.5)
    t1 = SENS["T1"]
    hdr = [N("Sensitivity", f"{c}{t1}", "pct", 1) for c in "BCDEF"]
    rows = []
    for i in range(5):
        rr = t1 + 1 + i
        rows.append([N("Sensitivity", f"A{rr}", "pct", 1)] + [N("Sensitivity", f"{c}{rr}", "usd", 0) for c in "BCDEF"])
    para(doc, "Residual income value per share, 12 months forward, by cost of equity (rows) and terminal growth (columns):", keep=True, space_after=2)
    add_table(doc, ["Cost of equity \\ terminal growth"] + hdr, rows, first_col_w=2.0, size=8.5)
    t3 = SENS["T3"]
    hdr = [N("Sensitivity", f"{c}{t3}", "pct", 1) for c in "BCDEF"]
    rows = []
    for i in range(5):
        rr = t3 + 1 + i
        rows.append([N("Sensitivity", f"A{rr}", "pct", 2)] + [N("Sensitivity", f"{c}{rr}", "usd", 2) for c in "BCDEF"])
    para(doc, "FY2027E diluted EPS by net charge-off rate (rows) and loan yield (columns); RSA re-sets with credit, reserve coverage and volumes held:", keep=True, space_after=2)
    add_table(doc, ["NCO rate \\ loan yield"] + hdr, rows, first_col_w=2.0, size=8.5)
    source_note(doc, "Source: model, Sensitivity sheet. All cells are live formulas; the center cell of each table equals the base case.")

    doc.add_heading("Comparable companies", level=2)
    rows = []
    for t in [bm.TICKER] + bm.PEERS:
        rows.append([f"{bm.PEER_NAMES[t]} ({t})", Cc(t, "price", "usd", 2), Cc(t, "mcap", "bn", 1), Cc(t, "pe_ttm", "x", 1), Cc(t, "pe27", "x", 1), Cc(t, "ptbv", "x", 1), Cc(t, "rotce", "pct", 1)])
    rows.append(["Peer median (ex-SYF)", "", "", Cc("median", "pe_ttm", "x", 1), Cc("median", "pe27", "x", 1), Cc("median", "ptbv", "x", 1), Cc("median", "rotce", "pct", 1)])
    add_table(doc, ["Company", "Price", "Market cap", "P/E TTM", "P/E FY2027E", "P/TBV", "ROTCE proxy"], rows, first_col_w=2.3, size=8.5, bold_rows=(0, len(rows) - 1))
    source_note(doc, f"Source: Yahoo Finance (prices, trailing and consensus EPS; retrieved {MARKET['retrieved_date']}); SEC XBRL balance sheets for tangible book (latest filed period per company); model, Comps sheet. Discover is excluded because Capital One acquired it, so Bread Financial is used.")
    para(doc, f"SYF trades at a {Cc('prem', 'pe27', 'abspct', 0)} discount to the peer median on FY2027E consensus earnings and a {Cc('prem', 'ptbv', 'abspct', 0)} discount on tangible book, "
              f"despite an ROTCE proxy above the median. The peer set is heterogeneous (American Express and OneMain distort the P/TBV to ROTCE regression), so we use it as a cross-check only: the fitted line implies "
              f"{N('Comps', bm.COMPS_VALUE_COL + str(CR['reg_imp']), 'x', 1)} for SYF, and the peer median forward P/E applied to SYF's consensus FY2027E EPS gives {N('Comps', bm.COMPS_VALUE_COL + str(CR['pe_imp']), 'usd', 0)} per share.")
    add_picture(doc, c3, "Figure 4: Price to tangible book versus ROTCE proxy across SYF and peers. Source: Yahoo Finance, SEC XBRL, model.")

    # ----- 6 risks
    doc.add_heading("6. Risks", level=1)
    para(doc, f"A recession would raise charge-offs and force reserve builds. Our bear case has FY2027E EPS of {Fb('eps', 'D', 'usd', 2, 'Bear')} and a blended value of {S('bear', 'usd', 0)}. Every 100bp on the net charge-off rate costs "
              f"{N('Sensitivity', 'B' + str(SENS['T4'] + 1), 'absusd', 2)} of FY2027E EPS after the RSA offset.",
         bold_lead="Credit cycle.")
    para(doc, "The CFPB's 2024 late-fee rule was vacated in April 2025, but the 10-K notes that legislative and regulatory proposals continue to seek limits on pricing of consumer credit products, and that competitive or regulatory changes could limit the rates and fees it charges. "
              f"Late fees were $2.3 billion of the {Hc('int_loans', 'FY2025', 'mm', 0)} of FY2025 interest and fees on loans. Reductions in interchange fees are also listed among the risk factors (10-K FY2025).",
         bold_lead="Regulation.")
    para(doc, "The five largest programs have expirations between 2030 and 2035, and 22 of the 25 largest program agreements expire in 2028 or later, representing 97% of interest and fees on loans from the 25 largest programs (10-K FY2025). "
              "Loss or repricing of a large partner, or a partner's own sales weakness, would hit volume and RSA economics. RSAs also transfer part of the upside, so a partner can capture improving credit.",
         bold_lead="Partner concentration and renewals.")
    para(doc, f"Deposits were 84% of funding at December 2025 (10-K FY2025), and part of the base comes from a savings partnership with PayPal. Each 50bp increase in cost of funds costs "
              f"{N('Sensitivity', 'B' + str(SENS['T4']), 'absusd', 2)} of FY2027E EPS. Deposit competition, a reversal in benchmark rates or securitization market dislocation would compress NIM.",
         bold_lead="Funding and deposit costs.")
    para(doc, f"Our base case runs total payout above earnings and CET1 to {Fb('cet1r', 'E', 'pct', 1)}. Supervisory stress test participation begins in 2026 (10-K FY2025), and a capital plan constraint would reduce the buyback contribution to EPS growth.",
         bold_lead="Capital return and regulatory capital.")

    # ----- appendix
    doc.add_heading("Appendix A: Sources", level=1)
    src = [
        "SEC EDGAR, Synchrony Financial (CIK 0001601712): Form 10-K for FY2025 (accession 0001601712-26-000006), FY2023 (0001601712-24-000047) and FY2021 (0001601712-22-000053); Form 10-Q for Q1 2026 (0001601712-26-000016) and Q2 2026 (0001601712-26-000033). Financial statements, MD&A statistical tables, capital tables and dividend and repurchase tables.",
        "SEC XBRL companyfacts API (data.sec.gov) for Synchrony: used to cross-check reported line items (NII, interest expense, provision, net earnings, EPS, equity, deposits, allowance) and to source dividends per share, common dividends and cash share repurchases. 52 of 52 cross-checks matched.",
        f"Yahoo Finance quoteSummary and chart APIs, retrieved {MARKET['retrieved_date']}: share prices, beta, trailing EPS, consensus EPS estimates for SYF, COF, AXP, BFH, OMF and SOFI, and the 10-year Treasury yield (^TNX).",
        "SEC XBRL companyfacts for Capital One, American Express, Bread Financial, OneMain Holdings and SoFi: equity, preferred stock, goodwill, intangibles and cover-page share counts for tangible book value.",
        "Company definitions: ROTCE in the model is net earnings available to common divided by average tangible common equity (equity less preferred stock, goodwill and intangibles). The Q2 2026 10-Q reclassified internal-use software out of intangibles, which raises TCE and CET1 relative to earlier periods; history is shown as originally reported.",
    ]
    for s_ in src:
        para(doc, s_, space_after=3, size=9)
    doc.add_heading("Appendix B: Model notes and simplifications", level=1)
    notes = [
        "The model is SYF_Model.xlsx with live formulas: Inputs_Sources, Historical, Scenarios (switch cell B3), Forecast, Valuation, Comps, Sensitivity, three scenario sheets and a Checks sheet. All Checks read OK.",
        "Annual forecast granularity. FY2026E share count, tangible equity and buybacks are anchored to reported H1 2026; income statement lines are full-year, driver-based. No full balance sheet is forecast: funding is implied by cost of funds on interest-bearing liabilities set as a share of interest-earning assets.",
        "RSA is regression-fitted (FY2020 to FY2025) to the net charge-off rate plus a structural residual. Provision equals charge-offs plus the change in the allowance, with other allowance movements set to zero in the forecast.",
        "Market inputs are as of the retrieval date and will move. The equity risk premium, terminal growth, terminal ROTCE, blend weights and rating thresholds are the author's assumptions, shown on Inputs_Sources and Scenarios.",
        "Items not sourced from a primary filing are marked as assumptions in the model. The 10-year yield and consensus EPS are from Yahoo Finance and were not cross-checked against a second source.",
    ]
    for s_ in notes:
        para(doc, s_, space_after=3, size=9)
    doc.add_heading("Appendix C: Disclaimer", level=1)
    para(doc, f"This document is an illustrative equity research sample prepared by {AUTHOR} for portfolio and educational purposes, dated {report_date}, with prices as of the {price_date} close. Forecasts, the price target and the rating are the author's own analysis, derived mechanically from the accompanying model, and are not investment advice or a recommendation to buy or sell any security. "
              "The author has no affiliation with any bank, broker-dealer or research department, and no position in the securities discussed is disclosed. Data comes from public filings and third-party market data that may contain errors or change without notice. Past performance is not indicative of future results.", size=9)

    doc.core_properties.author = AUTHOR
    doc.core_properties.last_modified_by = AUTHOR
    doc.core_properties.title = f"{COMPANY} ({TICKER}): Initiation of coverage"
    doc.core_properties.comments = ""
    doc.save(REPORT_DOCX)


def bm_row(key, scen):
    return LAY["S"][key][scen]


# ---------------------------------------------------------------- verification
def docx_text(path):
    from docx import Document as D
    d = D(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                parts.append(c.text)
    for sec in d.sections:
        parts += [p.text for p in sec.footer.paragraphs]
    return "\n".join(parts)


def verify():
    text = docx_text(REPORT_DOCX)
    wb2 = load_workbook(MODEL_XLSX, data_only=True)
    bad = []
    for desc, sheet, addr, kind, d, s in REG:
        v = wb2[sheet][addr].value
        again = render(v, kind, d)
        if again != s:
            bad.append(("xlsx mismatch", desc, s, again))
        if s not in text:
            bad.append(("not in docx", desc, s))
    dashes = [ch for ch in text if ch in "—–"]
    double = " -- " in text or "--" in text
    non_ascii = sorted({ch for ch in text if ord(ch) > 127})
    print(f"registered numbers checked: {len(REG)}; problems: {len(bad)}")
    for b in bad[:20]:
        print("  ", b)
    print("em/en dashes:", len(dashes), "double hyphen:", double, "non-ascii chars:", non_ascii)
    # style check: no colour in heading styles
    with zipfile.ZipFile(REPORT_DOCX) as z:
        styles = z.read("word/styles.xml").decode("utf8")
    for hid in ("Heading1", "Heading2", "Heading3"):
        m = re.search(rf'<w:style [^>]*w:styleId="{hid}".*?</w:style>', styles, re.S)
        print(hid, "has w:color:", bool(m and "<w:color" in m.group(0)))
    return not bad and not dashes and not double


def page_count():
    out = BASE_DIR / "_build" / "pdf"
    out.mkdir(parents=True, exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(REPORT_DOCX)], check=True, capture_output=True, timeout=300)
    import pymupdf
    n = pymupdf.open(out / (REPORT_DOCX.stem + ".pdf")).page_count
    print("pages:", n)
    return n


if __name__ == "__main__":
    build_doc()
    ok = verify()
    page_count()
    print("VERIFIED" if ok else "PROBLEMS FOUND")
