# Renewables Comparables Analysis

I used python to create a comparable company analysis of 10 publicly traded renewable energy firms. The energy transition is inevitable, and AI datacenter power demand is accelerating capital flows into the sector. The script I constructed pulls data with yfinance, calculates valuations, and exports a formatted Excel workbook. For gaps in yfinance data, color-coded N/A flags appear for meaningless multiples.

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

Three valuation measures are computed for each company:

- **EV/Revenue** — because several companies in the set are pre-profitability and revenue is the most consistent denominator.
- **EV/EBITDA** — the primary operating multiple for capital-intensive businesses like utilities and diversified renewables, where depreciation is significant and distorts net income comparisons.
- **P/E** — as a check for companies with positiveand  stable earnings (mostly the utilities and large diversified players

PLUG and BE display **N/A** on EV/EBITDA and P/E because both companies carry negative EBITDA and negative net income; dividing by a negative denominator would produce a mathematically valid but analytically meaningless result.

---

## Key Observations

- **FSLR trades lower than solar competitors** on EV/Revenue, reflecting the intensity of its domestic manufacturing model relative to asset-light installers like ENPH and RUN.
- **BE commands a notable premium** on a revenue basis, consistent with a market thesis around its solid oxide fuel cells as a high-reliability power source for data center and critical infrastructure buildouts.
- **NEE commands a premium to utility peers** (e.g., ED) across EV/EBITDA and P/E, reflecting its scale, regulated Florida utility base, and one of the largest renewables development pipelines in North America.
- **PLUG's negative EBITDA and net income** underscore that the green hydrogen segment remains in heavy investment mode — the stock is priced on optionality and long-term addressable market rather than current earnings.

---

## Data Source

data is sourced from yfinance. Select figures (market cap, enterprise value, EV multiples for NEE and FSLR) were cross-verified against publicly available filings and financial data providers to confirm directional accuracy.

---

## Tech Stack

- Python 3
- [yfinance](https://github.com/ranaroussi/yfinance) — market data
- [pandas](https://pandas.pydata.org/) — data handling
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel workbook 

---

## How to Run

```bash
pip3 install -r requirements.txt
python3 comps.py
```

script will get live data for all 10 tickers, print a summary with any flagged values, and save `renewables_comps.xlsx` to the working directory
