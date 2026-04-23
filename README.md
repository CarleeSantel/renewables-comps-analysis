# Renewables Comparables Analysis

A Python-based comparable company analysis of 10 publicly traded renewable energy firms. The energy transition is inevitable, and AI datacenter power demand are accelerating capital flows into the sector. The script pulls live market and financial data via the yfinance API, calculates key valuation metrics, and exports a formatted, print-ready Excel workbook with segment groupings, alternating row shading, frozen panes, and color-coded N/A flags for meaningless multiples.

---

## Sector Coverage

| Segment | Tickers |
|---|---|
| Solar | ENPH, FSLR, RUN |
| Diversified Renewables | NEE, BEP, CWEN, AES |
| Utilities | ED |
| Clean Energy Tech | PLUG, BE |

---

## Methodology

**Enterprise Value** is calculated as:

```
EV = Market Cap + Total Debt − Cash & Cash Equivalents
```

Three valuation multiples are computed for each company:

- **EV/Revenue** — included because several companies in the set are pre-profitability or have compressed margins; revenue is the most consistent denominator across the universe.
- **EV/EBITDA** — the primary operating multiple for capital-intensive businesses like utilities and diversified renewables, where depreciation is significant and distorts net income comparisons.
- **P/E** — included as a secondary check for companies with positive, stable earnings (primarily the utilities and large diversified players).

PLUG and BE display **N/A** on EV/EBITDA and P/E because both companies carry negative EBITDA and negative net income; dividing by a negative denominator would produce a mathematically valid but analytically meaningless result.

---

## Key Observations

- **FSLR trades at a discount to solar peers** on EV/Revenue, likely reflecting the capital intensity of its domestic manufacturing model relative to asset-light installers like ENPH and RUN.
- **BE commands a notable premium** on a revenue basis, consistent with a market thesis around its solid oxide fuel cells as a high-reliability power source for data center and critical infrastructure buildouts.
- **NEE commands a premium to utility peers** (e.g., ED) across EV/EBITDA and P/E, reflecting its scale, regulated Florida utility base, and one of the largest renewables development pipelines in North America.
- **PLUG's negative EBITDA and net income** underscore that the green hydrogen segment remains in heavy investment mode — the stock is priced on optionality and long-term addressable market rather than current earnings.

---

## Data Source

All financial and market data is sourced from **Yahoo Finance via the yfinance library**. Select figures (market cap, enterprise value, EV multiples for NEE and FSLR) were cross-verified against publicly available filings and financial data providers to confirm directional accuracy.

---

## Tech Stack

- Python 3
- [yfinance](https://github.com/ranaroussi/yfinance) — market & financial data
- [pandas](https://pandas.pydata.org/) — data handling
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel workbook generation

---

## How to Run

```bash
pip install -r requirements.txt
python3 comps.py
```

The script will fetch live data for all 10 tickers, print a console summary with any flagged values (missing or negative line items), and save `renewables_comps.xlsx` to the working directory.
