# Synchrony Financial (SYF) Initiation of Coverage


> **Disclaimer.** Illustrative analysis for educational and portfolio purposes only. Not investment advice, not a recommendation to buy or sell any security, and not affiliated with any bank, broker or research department. Forecasts, rating and price target rest on stated assumptions and public data as of 2026-09-28.

## What it is

A sell-side style equity research initiation sample on Synchrony Financial (NYSE: SYF), built end to end from primary sources. It has two deliverables that tie to each other exactly: a driver-based financial model with valuation (`SYF_Model.xlsx`, live Excel formulas) and an 8-page initiation report (`SYF_Initiation_Report.docx`) that is generated programmatically from the recalculated model. Rating and price target follow mechanically from the valuation. Every historical figure comes from SEC EDGAR filings (10-K FY2021, FY2023, FY2025; 10-Q Q1 and Q2 2026), cross-checked against SEC XBRL companyfacts. Market data (price, beta, consensus EPS, 10-year yield) is from Yahoo Finance, retrieved 2026-09-28, with prices as the last completed close (2026-09-25).

This is an illustrative author analysis for portfolio purposes, not investment advice, and has no affiliation with any bank or research department.

## Tech and methods

- Financial modelling: driver-based annual forecast (loan growth, purchase volume, yield, cost of funds, NIM, retailer share arrangements (RSA) as % of average receivables, net charge-off rate, CECL reserve build, opex growth, buyback pace), allowance roll-forward, TCE and CET1 roll-forwards, bear/base/bull scenario switch.
- Valuation for a lender: residual income model on tangible common equity, justified P/TBV from ROTCE versus cost of equity, peer comps (Capital One, American Express, Bread Financial, OneMain, SoFi), sensitivity tables built with formulas. A plain FCF DCF is deliberately not used, with the reasoning in the report.
- Consumer finance domain: NIM and RSA dynamics (regression of RSA % on charge-off rate, correlation -0.95), CECL reserve mechanics, credit cycle, capital return, regulatory and partner-concentration risks read from the 10-K.
- Python: `requests`, `BeautifulSoup` (filing table parsing), `openpyxl` (formula workbook), LibreOffice headless recalculation, `matplotlib` charts, `python-docx` report, automated number verification (docx versus xlsx), independent Python re-computation of the model.
- Data hygiene: 52 of 52 XBRL cross-checks match the parsed filing text; 54 of 54 formula-driven model checks read OK.

## Scope and layout

Scope: data pull, model, valuation, charts, report, verification. Kept deliberately tight: annual granularity (FY2026E to FY2028E) anchored to reported H1 2026, no full balance sheet forecast, and one rating framework. Discover is no longer listed (acquired by Capital One), so Bread Financial is used in the comps.

**Folder layout:** `SYF_Model.xlsx`, `SYF_Initiation_Report.docx`, `python/` (`pull_data.py`, `build_model.py`, `build_report.py`, `data/` snapshots, `charts/`). Run order: `python python/pull_data.py` (optional, needs network; `--refresh` to re-pull), then `python python/build_model.py`, then `python python/build_report.py`. The build runs offline from `python/data/`.

## Headline outputs (base case)

- Rating: Overweight (mechanical rule: Overweight at 15% or more expected total return, Underweight at -10% or less).
- 12-month price target: $89, versus $72.94 close on 2026-09-25 (upside 22.0%, expected total return 23.9% including dividends).
- Value per share, 12-month blended (residual income and justified P/TBV, equal weights): bear $62.72, base $89.49, bull $111.12.
- Diluted EPS forecast: FY2026E $9.31, FY2027E $10.02, FY2028E $11.49 (Yahoo consensus $9.42 and $10.44 for FY2026E and FY2027E).
- Valuation inputs: cost of equity 11.8% (10-year yield 5.26% plus beta 1.31 times 5.0% ERP), terminal growth 3.0%, terminal ROTCE 16% (base), justified P/TBV 1.48x at terminal versus 1.74x at the current price.
- Forecast highlights: NIM about 15.2%, net charge-off rate 5.25% falling to 5.00%, ROTCE about 22%, net earnings to common roughly flat while EPS grows on buybacks, CET1 12.6% falling to 11.6% by FY2028E.
- Checks: NIM, charge-off rate, efficiency ratio, loan yield and cost of funds recompute to reported figures within rounding; allowance, loan, TCE and share roll-forwards tie exactly; report contains 542 model-linked numbers, all re-verified against the workbook.

## Status and limitations

Complete as of 2026-09-28. Known limitations: consensus EPS and the 10-year Treasury yield come from Yahoo Finance and were not cross-checked against a second source; the equity risk premium, terminal assumptions, blend weights and rating thresholds are author assumptions; OneMain's tangible book uses its 2026-03-31 balance sheet because its Q2 2026 filing was not in the SEC XBRL feed; the Q2 2026 10-Q reclassified internal-use software out of intangibles, so TCE and CET1 after that date are not on the same basis as earlier history (history is as originally reported, the forecast uses the recast basis).
