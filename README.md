# Professional DCF Valuation Model

Currently freelancing. I studied undergrad in FinTech and Accoutning. Most people don't have access to Bloomberg Terminal, so I built one with real time access.

Two commands. Any public company. Professional output.

```bash
python build_dcf.py AAPL     # build + populate in one shot
python update_dcf.py         # refresh data on an existing model
```

---

## What You Get

A 10-sheet Excel workbook. Color-coded, fully linked, no manual data entry required.

| Sheet | What it does |
|---|---|
| **Guide** | Color key and how to use it |
| **Stock Search** | Live snapshot — price, market cap, revenue, debt, multiples |
| **Assumptions** | Every input in one place. Gold cells = yours to edit |
| **Income Statement** | 10-year revenue → EBITDA → FCFF projection with FCF chart |
| **WACC** | CAPM cost of equity, after-tax cost of debt, blended WACC |
| **DCF Valuation** | Discounted FCFs + terminal value → Enterprise Value → Implied Price |
| **Sensitivity** | Implied price across WACC × TGR and EBITDA Margin × Revenue Growth |
| **Scenarios** | Bear / Base / Bull side-by-side with implied prices |
| **Reverse DCF** | What growth rate is the market currently pricing in |
| **Price** | Intraday price tracker — open, high, low, alpha vs S&P 500 |

Gold cells are inputs. White cells are formulas. Don't touch the white cells.

---

## Install

```bash
pip install -r requirements.txt
```

Python 3.10+. That's it.

---

## Usage

### Build + populate in one shot

```bash
python build_dcf.py AAPL
python build_dcf.py MSFT --date 2020-03-23   # COVID crash low
python build_dcf.py NVDA --date 2008         # financial crisis
python build_dcf.py AAPL --date 6/1/2025     # slash date format works too
```

You'll be prompted to name the output file. The model is saved to the same folder as the script.

### Refresh data on an existing model

```bash
python update_dcf.py
```

Interactive flow — picks up your existing models automatically:

```
Found 2 DCF models:
    [1] DCF_Model_AAPL_2026-06-02_11-09AM.xlsx  (AAPL)
    [2] DCF_Model_Meta.xlsx  (META)

  Pick a file [1]: 1
  Ticker in file: AAPL. Use this? [Y/n]: y

  Update AAPL with:
    [1] Live data
    [2] Historical date
  Choice [1]:

  Save as:
    [1] Overwrite  DCF_Model_AAPL_2026-06-02_11-09AM.xlsx
    [2] New file   DCF_Model_AAPL_2026-06-02_02-30PM.xlsx
  Choice [1]:
```

Or skip the prompts entirely:

```bash
python update_dcf.py AAPL                    # live data, overwrites newest file
python update_dcf.py AAPL --date 2020-03-23  # historical
python update_dcf.py AAPL --price-only       # quick price refresh only
```

### Enter your thesis

Open the model. Go to **Assumptions**. Adjust the gold cells:

- Revenue growth Y1–Y10
- EBITDA and gross margins
- CapEx, D&A, NWC, SBC, tax rate
- WACC inputs (Rf, beta, ERP, cost of debt, capital structure)
- Terminal growth rate

This is where your view of the business goes. The model is just the math.

### Read the output

- **DCF Valuation** → Implied share price vs current price, upside/downside
- **Sensitivity** → How the price moves across scenarios. Green = upside, red = downside.
- **Scenarios** → Bear/base/bull implied prices side by side
- **Reverse DCF** → Find what growth rate the market is pricing in at current price

---

## Workflow

```
build_dcf.py TICKER
      ↓
Fresh 10-sheet model populated with live data
      ↓
You adjust growth rates and margins in Assumptions
      ↓
DCF Valuation shows implied price vs market
      ↓
update_dcf.py to refresh data anytime (keeps your thesis intact)
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

I did my undergraduate studies in accounting and fintech, and often in classes models like these are models I wish I had. The DCF model started as something I learned in Finance and now it's something I use frequently for my personal projects.

If you're a finance student, an analyst, or just someone who wants to run a real valuation without paying for Bloomberg, this is for you. Go enjoy it. 

---

*Python 3.13 · openpyxl · yfinance*
