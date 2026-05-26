# Professional DCF Valuation Model

Currently freelancing. I studied undergrad in FinTech and Accoutning. Most people don't have access to Bloomberg Terminal, so I built one with real time access.

Two commands. Any public company. Professional output.

```bash
python build_dcf.py          # builds DCF_Model.xlsx
python update_dcf.py         # pulls live data — enter your ticker when prompted
```

---

## What You Get

A 7-sheet Excel workbook. Color-coded, fully linked, no manual data entry required.

| Sheet | What it does |
|---|---|
| **Guide** | Color key and how to use it |
| **Stock Search** | Live snapshot — price, market cap, revenue, debt, multiples |
| **Assumptions** | Every input in one place. Gold cells = yours to edit |
| **Income Statement** | 10-year revenue → EBITDA → FCFF projection |
| **WACC** | CAPM cost of equity, after-tax cost of debt, blended WACC |
| **DCF Valuation** | Discounted FCFs + terminal value → Enterprise Value → Implied Price |
| **Sensitivity** | Implied price across WACC × TGR and EBITDA Margin × Revenue Growth |

Gold cells are inputs. White cells are formulas. Don't touch the white cells.

---

## Install

```bash
pip install openpyxl yfinance
```

Python 3.9+. That's it.

---

## Usage

### Step 1 — Build the model

```bash
python build_dcf.py
```

Saves `DCF_Model.xlsx` to your Desktop. Run this once, or any time you want a clean slate.

### Step 2 — Load a stock

```bash
python update_dcf.py AAPL
python update_dcf.py MSFT
python update_dcf.py NVDA
```

Pulls from Yahoo Finance. Populates Stock Search and pre-fills Assumptions with price, shares, debt, cash, beta, and revenue. Works with any ticker Yahoo supports.

No argument — it'll ask:

```bash
python update_dcf.py
# Enter ticker symbol (e.g. AAPL, MSFT, NVDA): TSLA
```

### Step 3 — Enter your thesis

Open `DCF_Model.xlsx`. Go to **Assumptions**. Adjust the gold cells:

- Revenue growth Y1–Y10
- EBITDA and gross margins
- CapEx, D&A, NWC, SBC, tax rate
- WACC inputs (Rf, beta, ERP, cost of debt, capital structure)
- Terminal growth rate

This is where your view of the business goes. The model is just the math.

### Step 4 — Read the output

- **DCF Valuation** → Implied share price vs current price, upside/downside
- **Sensitivity** → How the price moves across scenarios. Green = upside, red = downside.

---

## Workflow

```
update_dcf.py TICKER
      ↓
Live data populates Stock Search + Assumptions
      ↓
You adjust growth rates and margins to reflect your thesis
      ↓
DCF Valuation shows implied price vs market
      ↓
Sensitivity shows the range across assumptions
```

---

## Microsoft 365

If you have M365, you can skip the Python script for price data:

1. Type your ticker in cell B5 of Stock Search
2. `Data` tab → `Stocks` in the ribbon
3. Use `=B5.Price`, `=B5.MarketCap`, etc. to pull any field live

The Python script works with any Excel version.

---

## Notes

- Everything in USD millions. Percentages as decimals — 10% = `0.10`
- International stocks (TSM, ASML, etc.): Yahoo returns financials in local currency. Convert manually.
- The WACC cell in Assumptions pulls from the WACC sheet automatically. Don't overwrite it.
- `yfinance` installs itself if it's missing.

---

## Why I Built This

This is made for a finance student, an analyst, or just someone who wants to run a real valuation without paying for Bloomberg. Enjoy it. 

---

*Python 3.13 · openpyxl · yfinance*
