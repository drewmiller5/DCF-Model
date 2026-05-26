# Professional DCF Valuation Model

Currently freelancing. I earned my undergraduate degree in Accounting and Financial Technology from Elon University, and built a freelance real-time market platform to give users access to financial information.

One command. Any public company. Professional output.

```bash
python build_dcf.py AAPL     # builds DCF_Model.xlsx and populates live data
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

**Step 1 — Make sure Python is installed**

Open a terminal (Mac: search "Terminal", Windows: search "Command Prompt") and run:

```bash
python --version
```

You should see `Python 3.9` or higher. If you get an error, download Python at [python.org/downloads](https://python.org/downloads) — just click Install and keep all defaults.

**Step 2 — Install the two required packages**

```bash
pip install openpyxl yfinance
```

That's it. You only need to do this once.

---

## Usage

> **First time?** Open your terminal, navigate to the folder where you downloaded these files, then run the commands below. On Windows you can also right-click the folder and choose "Open in Terminal".

### Step 1 — Build the model and load a stock

```bash
python build_dcf.py AAPL
python build_dcf.py MSFT
python build_dcf.py NVDA
```

Saves the file to the current directory, then pulls live data from Yahoo Finance and populates Stock Search and Assumptions with price, shares, debt, cash, beta, and revenue. Works with any ticker Yahoo supports.

No argument — it'll ask:

```bash
python build_dcf.py
# Enter ticker to auto-populate (or press Enter to skip): TSLA
```

To refresh an existing model with new data without rebuilding:

```bash
python update_dcf.py AAPL
```

### Step 2 — Enter your thesis

Open `DCF_Model.xlsx`. Go to **Assumptions**. Adjust the gold cells:

- Revenue growth Y1–Y10
- EBITDA and gross margins
- CapEx, D&A, NWC, SBC, tax rate
- WACC inputs (Rf, beta, ERP, cost of debt, capital structure)
- Terminal growth rate

This is where your view of the business goes. The model is just the math.

### Step 3 — Read the output

- **DCF Valuation** → Implied share price vs current price, upside/downside
- **Sensitivity** → How the price moves across scenarios. Green = upside, red = downside.

---

## Workflow

```
build_dcf.py TICKER
      ↓
Model built + live data populates Stock Search + Assumptions
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
- The sample you see in Excel was built on 5/25/2026 and the stock was AAPL at close. This is a sample output just to show the accuracy of it. This sheet will not update often.
---

## Why I Built This

This is made for a finance student, an analyst, or just someone who wants to run a real valuation without paying for Bloomberg. Enjoy it. 

---

*Python 3.13 · openpyxl · yfinance*
