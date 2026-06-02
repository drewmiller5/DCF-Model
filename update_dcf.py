"""
Stock Data Updater for DCF_Model.xlsx

Usage:
  python update_dcf.py AAPL                     # live data
  python update_dcf.py AAPL --date 2020-03-23   # COVID crash low
  python update_dcf.py AAPL --date 2008         # crisis (price/beta OK; financials may need manual entry)
  python update_dcf.py META --output path/to/DCF_Model.xlsx

Data sources:
  Price / beta / risk-free rate  ->  Yahoo Finance  (full history)
  Financial statements           ->  yfinance (~4 yrs), then SEC EDGAR XBRL (~2009+) as fallback
  Pre-2009 financials            ->  not available via free APIs; enter manually from 10-K

Requires: pip install yfinance openpyxl pandas requests
"""

import sys, os, argparse, datetime
import requests

try:
    import yfinance as yf
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "yfinance"])
    import yfinance as yf

try:
    import pandas as pd
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
    import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

INPUT_GOLD = "FFF9C4"


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def parse_date(date_str: str) -> datetime.date | None:
    """Parse a user date string. Returns None for 'present'."""
    if date_str.lower() in ("present", "now", "today", "current"):
        return None
    if len(date_str) == 4 and date_str.isdigit():
        return datetime.date(int(date_str), 12, 31)  # year-only → Dec 31
    # ISO format: YYYY-MM-DD
    try:
        return datetime.date.fromisoformat(date_str)
    except ValueError:
        pass
    # Slash formats: M/D/YYYY or M/D/YY
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    raise ValueError(
        f"Unrecognized date format: '{date_str}'. "
        "Use YYYY, YYYY-MM-DD, M/D/YYYY, or 'present'."
    )


# ═══════════════════════════════════════════════════════════════════════════════
# SEC EDGAR XBRL  (free, no API key, ~2009+)
# ═══════════════════════════════════════════════════════════════════════════════

# SEC EDGAR requires a real contact in the User-Agent — update this with your email
_HDRS = {"User-Agent": "DCF-Model-Script contact@example.com"}


def _edgar_cik(ticker: str) -> str | None:
    try:
        r = requests.get(
            "https://www.sec.gov/files/company_tickers.json",
            headers=_HDRS, timeout=10,
        )
        return next(
            (str(v["cik_str"]).zfill(10)
             for v in r.json().values()
             if v["ticker"].upper() == ticker.upper()),
            None,
        )
    except Exception:
        return None


def fetch_edgar_financials(ticker: str, as_of_date: datetime.date) -> dict:
    """
    Pull annual financials from SEC EDGAR XBRL company facts.

    Coverage: US public companies filing XBRL — large accelerated filers from FY2009,
    all domestic filers from FY2011. Returns zeros for any unavailable concept.
    """
    empty = dict(revenue_m=0, ebitda_m=0, ni_m=0, debt_m=0, cash_m=0, shares_m=0)

    cik = _edgar_cik(ticker)
    if not cik:
        print(f"  EDGAR: no CIK found for '{ticker}'.")
        return empty

    try:
        r = requests.get(
            f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json",
            headers=_HDRS, timeout=30,
        )
        us_gaap = r.json().get("facts", {}).get("us-gaap", {})
    except Exception as exc:
        print(f"  EDGAR fetch failed: {exc}")
        return empty

    def pick_usd(*concepts) -> float:
        """Return the most recent 10-K USD value <= as_of_date for the first matching concept."""
        for concept in concepts:
            rows = us_gaap.get(concept, {}).get("units", {}).get("USD", [])
            valid = [
                row for row in rows
                if row.get("form") in ("10-K", "10-K/A")
                and datetime.date.fromisoformat(row["end"]) <= as_of_date
            ]
            if valid:
                # Most recent period, then most recently filed (handles amendments)
                valid.sort(key=lambda r: (r["end"], r.get("filed", "")), reverse=True)
                return valid[0]["val"] / 1e6
        return 0

    def pick_shares(*concepts) -> float:
        for concept in concepts:
            for unit in ("shares", "Shares"):
                rows = us_gaap.get(concept, {}).get("units", {}).get(unit, [])
                valid = [
                    r for r in rows
                    if r.get("form") in ("10-K", "10-K/A")
                    and datetime.date.fromisoformat(r["end"]) <= as_of_date
                ]
                if valid:
                    valid.sort(key=lambda r: (r["end"], r.get("filed", "")), reverse=True)
                    return valid[0]["val"] / 1e6
        return 0

    revenue_m = pick_usd(
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueGoodsNet",
    )
    ni_m    = pick_usd("NetIncomeLoss", "ProfitLoss")
    ebit_m  = pick_usd(
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    )
    da_m    = pick_usd(
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
        "Depreciation",
    )
    lt_debt = pick_usd(
        "LongTermDebt", "LongTermDebtNoncurrent",
        "LongTermDebtAndCapitalLeaseObligations",
    )
    st_debt = pick_usd("ShortTermBorrowings", "DebtCurrent", "NotesPayableCurrent")
    cash_m  = pick_usd(
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "CashAndCashEquivalents",
    )
    shares_m = pick_shares(
        "CommonStockSharesOutstanding",
        "WeightedAverageNumberOfSharesOutstandingBasic",
    )

    return {
        "revenue_m": revenue_m,
        "ebitda_m":  (ebit_m + da_m) if ebit_m else 0,
        "ni_m":      ni_m,
        "debt_m":    (lt_debt or 0) + (st_debt or 0),
        "cash_m":    cash_m,
        "shares_m":  shares_m,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# YAHOO FINANCE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _calc_beta(stock_closes: pd.Series, as_of: datetime.date) -> float:
    """1-year trailing beta vs S&P 500 ending at as_of."""
    start = (as_of - datetime.timedelta(days=365)).strftime("%Y-%m-%d")
    end   = (as_of + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    sp = yf.download("^GSPC", start=start, end=end, progress=False, auto_adjust=True)
    if sp.empty:
        return 1.0
    sp_close   = sp["Close"].squeeze()

    def _drop_tz(s: pd.Series) -> pd.Series:
        if hasattr(s.index, "tz") and s.index.tz is not None:
            s = s.copy()
            s.index = s.index.tz_localize(None)
        return s

    stock_ret  = _drop_tz(stock_closes.pct_change().dropna())
    market_ret = _drop_tz(sp_close.pct_change().dropna())
    both = pd.concat([stock_ret, market_ret], axis=1).dropna()
    both.columns = ["stock", "market"]
    if len(both) < 20:
        return 1.0
    cov = both.cov()
    return round(cov.loc["stock", "market"] / cov.loc["market", "market"], 2)


def _hist_rf_rate(as_of: datetime.date) -> float:
    """10-year treasury yield (^TNX) as of a given date."""
    start = (as_of - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
    end   = (as_of + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    tnx = yf.download("^TNX", start=start, end=end, progress=False, auto_adjust=True)
    if tnx.empty:
        return 0.043
    return round(float(tnx["Close"].iloc[-1].item()), 4) / 100


def _get_financials_as_of(
    tk: yf.Ticker, as_of_date: datetime.date, ticker: str = ""
) -> dict:
    """
    yfinance financial statements with automatic SEC EDGAR fallback.

    yfinance typically covers the last 4 annual periods.
    EDGAR covers from ~FY2009 for large filers, ~FY2011 for all others.
    Pre-2009 data requires manual entry from 10-K filings.
    """
    result = dict(revenue_m=0, ebitda_m=0, ni_m=0, debt_m=0, cash_m=0, shares_m=0)

    def nearest_col(df: pd.DataFrame):
        valid = [c for c in df.columns if hasattr(c, "date") and c.date() <= as_of_date]
        return max(valid) if valid else None

    def safe_row(df, *candidates):
        for name in candidates:
            if name in df.index:
                return df.loc[name]
        return None

    # Try yfinance first
    try:
        inc = tk.income_stmt
        col = nearest_col(inc)
        if col is not None:
            rev  = safe_row(inc, "Total Revenue", "TotalRevenue")
            ebit = safe_row(inc, "EBITDA", "Ebitda")
            ni   = safe_row(inc, "Net Income", "NetIncome")
            result["revenue_m"] = float(rev[col])  / 1e6 if rev  is not None else 0
            result["ebitda_m"]  = float(ebit[col]) / 1e6 if ebit is not None else 0
            result["ni_m"]      = float(ni[col])   / 1e6 if ni   is not None else 0
    except Exception:
        pass

    try:
        bs  = tk.balance_sheet
        col = nearest_col(bs)
        if col is not None:
            debt = safe_row(bs, "Total Debt", "TotalDebt", "LongTermDebt")
            cash = safe_row(bs, "Cash And Cash Equivalents", "CashAndCashEquivalents",
                            "Cash", "CashEquivalentsAndShortTermInvestments")
            result["debt_m"] = float(debt[col]) / 1e6 if debt is not None else 0
            result["cash_m"] = float(cash[col]) / 1e6 if cash is not None else 0
    except Exception:
        pass

    # EDGAR fallback when yfinance comes up empty
    if result["revenue_m"] == 0 and ticker:
        print("  yfinance: no data for this period — trying SEC EDGAR XBRL ...")
        edgar = fetch_edgar_financials(ticker, as_of_date)
        if edgar["revenue_m"]:
            print(f"  EDGAR: revenue ${edgar['revenue_m']:,.0f}M "
                  f"(nearest 10-K ending ≤ {as_of_date}).")
            result.update(edgar)
        else:
            print(
                f"  EDGAR: no XBRL data found for {as_of_date.year}.\n"
                f"  XBRL coverage starts FY2009 (large filers) / FY2011 (all).\n"
                f"  Enter income / balance sheet figures manually from the company's 10-K."
            )

    return result


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN DATA FETCH
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_stock_data(ticker: str, as_of_date: datetime.date | None = None) -> dict:
    """Pull key metrics — live from Yahoo Finance, or historical via Yahoo + EDGAR."""
    tk  = yf.Ticker(ticker)
    inf = tk.info

    def safe(key, default=None):
        val = inf.get(key, default)
        return val if val is not None else default

    # ── HISTORICAL MODE ───────────────────────────────────────────────────────
    if as_of_date is not None and as_of_date >= datetime.date.today():
        print("Date is today or future — switching to live data.")
        as_of_date = None

    if as_of_date is not None:
        print(f"  Fetching {ticker.upper()} as of {as_of_date} ...")
        start_1y = as_of_date - datetime.timedelta(days=365)

        hist = tk.history(
            start=start_1y.strftime("%Y-%m-%d"),
            end=(as_of_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d"),
            auto_adjust=True,
        )
        if hist.empty:
            print(f"WARNING: No price data for {ticker} around {as_of_date}.")
            price = hi52 = lo52 = 0.0
            open_price = day_high = day_low = 0.0
            beta = 1.0
            sp500_move = 0.0
        else:
            last        = hist.iloc[-1]
            price       = round(float(last["Close"].item()), 2)
            open_price  = round(float(last["Open"].item()),  2)
            day_high    = round(float(last["High"].item()),  2)
            day_low     = round(float(last["Low"].item()),   2)
            hi52        = round(float(hist["High"].max()),   2)
            lo52        = round(float(hist["Low"].min()),    2)
            beta        = _calc_beta(hist["Close"], as_of_date)

            # S&P 500 daily move on that specific date
            try:
                sp_day = yf.download(
                    "^GSPC",
                    start=as_of_date.strftime("%Y-%m-%d"),
                    end=(as_of_date + datetime.timedelta(days=2)).strftime("%Y-%m-%d"),
                    progress=False, auto_adjust=True,
                )
                if not sp_day.empty:
                    sp_o = float(sp_day["Open"].iloc[0].item())
                    sp_c = float(sp_day["Close"].iloc[0].item())
                    sp500_move = round((sp_c / sp_o) - 1, 4) if sp_o else 0.0
                else:
                    sp500_move = 0.0
            except Exception:
                sp500_move = 0.0

        rf_rate  = _hist_rf_rate(as_of_date)
        fins     = _get_financials_as_of(tk, as_of_date, ticker=ticker)

        shares_m  = fins.get("shares_m") or (safe("sharesOutstanding", 0) / 1e6)
        mktcap_m  = price * shares_m
        net_debt  = fins["debt_m"] - fins["cash_m"]
        ev        = mktcap_m + net_debt
        pe        = price / (fins["ni_m"] / shares_m) if shares_m and fins["ni_m"] else 0
        ev_ebitda = round(ev / fins["ebitda_m"], 1) if fins["ebitda_m"] else 0

        return {
            "ticker":       ticker.upper(),
            "name":         safe("longName", safe("shortName", ticker)),
            "price":        price,
            "hi52":         hi52,
            "lo52":         lo52,
            "mktcap_m":     mktcap_m,
            "shares_m":     shares_m,
            "float_pct":    0.99,
            "pe":           round(pe, 1) if pe else 0,
            "ev_ebitda":    ev_ebitda,
            "ps":           0,
            "revenue_m":    fins["revenue_m"],
            "ebitda_m":     fins["ebitda_m"],
            "ni_m":         fins["ni_m"],
            "debt_m":       fins["debt_m"],
            "cash_m":       fins["cash_m"],
            "net_debt_m":   fins["debt_m"] - fins["cash_m"],
            "div_yield":    0,
            "rev_growth":   0,
            "eps_growth":   0,
            "beta":         beta,
            "rf_rate":      rf_rate,
            "updated":      f"{as_of_date} (historical)",
            "open_price":   open_price,
            "session_high": day_high,
            "session_low":  day_low,
            "market_state": "HISTORICAL",
            "sp500_move":   sp500_move,
        }

    # ── LIVE MODE ─────────────────────────────────────────────────────────────
    if not inf or inf.get("quoteType") is None:
        print(f"\n  WARNING: '{ticker}' not found on Yahoo Finance — building blank model.")
        print("  Check the ticker symbol (e.g. AAPL, MSFT, NVDA) or proceed with a blank DCF.\n")

    hist = tk.history(period="1y")
    hi52 = float(hist["High"].max()) if not hist.empty else safe("fiftyTwoWeekHigh", 0)
    lo52 = float(hist["Low"].min())  if not hist.empty else safe("fiftyTwoWeekLow",  0)

    shares    = safe("sharesOutstanding", 0)
    price     = safe("previousClose") or safe("regularMarketPrice") or safe("currentPrice", 0)
    mktcap    = safe("marketCap", 0)
    ebitda    = safe("ebitda", 0)
    rev       = safe("totalRevenue", 0)
    ni        = safe("netIncomeToCommon", 0)
    debt      = safe("totalDebt", 0)
    cash      = safe("totalCash", 0)
    beta      = safe("beta", 1.0)
    float_pct = safe("floatShares", shares) / shares if shares else 0.99

    # Intraday data for Price Tracker sheet
    open_price    = safe("regularMarketOpen", price)
    session_high  = safe("regularMarketDayHigh", price)
    session_low   = safe("regularMarketDayLow", price)
    market_state  = safe("marketState", "UNKNOWN")

    # S&P 500 today's move
    sp500_move = 0.0
    try:
        sp = yf.download("^GSPC", period="1d", interval="1d",
                         progress=False, auto_adjust=True)
        if not sp.empty:
            sp_open  = float(sp["Open"].iloc[-1].item() if hasattr(sp["Open"].iloc[-1], "item") else sp["Open"].iloc[-1])
            sp_close = float(sp["Close"].iloc[-1].item() if hasattr(sp["Close"].iloc[-1], "item") else sp["Close"].iloc[-1])
            if sp_open != 0:
                sp500_move = round(sp_close / sp_open - 1, 4)
    except Exception:
        pass

    return {
        "ticker":       ticker.upper(),
        "name":         safe("longName", safe("shortName", ticker)),
        "price":        price,
        "hi52":         hi52,
        "lo52":         lo52,
        "mktcap_m":     mktcap / 1e6 if mktcap else 0,
        "shares_m":     shares / 1e6 if shares else 0,
        "float_pct":    float_pct,
        "pe":           safe("trailingPE", 0),
        "ev_ebitda":    safe("enterpriseToEbitda", 0),
        "ps":           safe("priceToSalesTrailing12Months", 0),
        "revenue_m":    rev    / 1e6 if rev    else 0,
        "ebitda_m":     ebitda / 1e6 if ebitda else 0,
        "ni_m":         ni     / 1e6 if ni     else 0,
        "debt_m":       debt   / 1e6 if debt   else 0,
        "cash_m":       cash   / 1e6 if cash   else 0,
        "net_debt_m":   (debt - cash) / 1e6 if debt and cash else 0,
        "div_yield":    safe("dividendYield", 0) or 0,
        "rev_growth":   safe("revenueGrowth", 0) or 0,
        "eps_growth":   safe("earningsGrowth", 0) or 0,
        "beta":         beta,
        "rf_rate":      None,  # live: leave Assumptions risk-free rate cell untouched
        "updated":      datetime.date.today().strftime("%Y-%m-%d"),
        "open_price":   open_price,
        "session_high": session_high,
        "session_low":  session_low,
        "market_state": market_state,
        "sp500_move":   sp500_move,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════════════════════════════════════════════

def update_excel(path: str, data: dict):
    """Write fetched data into DCF_Model.xlsx."""
    wb = load_workbook(path)

    if "Stock Search" not in wb.sheetnames:
        print(f"ERROR: 'Stock Search' sheet not found in {path}")
        return

    ws   = wb["Stock Search"]
    gold = fill(INPUT_GOLD)

    def write(ref, value, sheet=ws):
        c = sheet[ref]
        c.value = value
        c.fill  = gold

    # Stock Search sheet
    write("B5", data["ticker"])
    write("C5", round(data["price"], 2))
    write("D5", data["name"])
    write("E5", data["updated"])

    write("C8",  round(data["price"],    2))
    write("D8",  round(data["hi52"],     2))
    write("E8",  round(data["lo52"],     2))

    write("C10", round(data["mktcap_m"],  1))
    write("D10", round(data["shares_m"],  1))
    write("E10", round(data["float_pct"], 4))

    write("C12", round(data["revenue_m"], 0))
    write("D12", round(data["ebitda_m"],  0))
    write("E12", round(data["ni_m"],      0))

    write("C14", round(data["debt_m"],    0))
    write("D14", round(data["cash_m"],    0))
    # E14 = formula =C14-D14, leave it

    write("C16", round(data["pe"],        1) if data["pe"] else "N/A")
    write("D16", round(data["ev_ebitda"], 1) if data["ev_ebitda"] else "N/A")
    write("E16", round(data["beta"],      2) if data["beta"] else 1.0)

    # Assumptions sheet
    if "Assumptions" in wb.sheetnames:
        wa = wb["Assumptions"]

        def wa_w(ref, val):
            write(ref, val, sheet=wa)

        wa_w("C5",  data["name"])
        wa_w("E5",  data["ticker"])
        wa_w("C7",  round(data["price"],    2))
        wa_w("E7",  round(data["shares_m"], 1))
        wa_w("C8",  round(data["debt_m"],   0))
        wa_w("E8",  round(data["cash_m"],   0))
        wa_w("C22", round(data["beta"],      2) if data["beta"] else 1.0)
        wa_w("C28", round(data["revenue_m"], 0))

        # Historical-only: also write risk-free rate
        if data.get("rf_rate") is not None:
            wa_w("C21", data["rf_rate"])
            print(f"  Risk-Free Rate → {data['rf_rate']:.2%}  (10-yr treasury on that date)")

    # Price Tracker sheet — cell layout mirrors build_price_tracker() in build_dcf.py
    # Section headers consume rows 4, 7, 12, 16; data rows:
    #   C8  = open price     E8  = market status
    #   C10 = session high   E10 = session low
    #   C17 = S&P 500 move today
    if "Price" in wb.sheetnames:
        wp = wb["Price"]
        def wp_w(ref, val, fmt=None):
            c = wp[ref]
            c.value = val
            c.fill  = gold
            if fmt:
                c.number_format = fmt
        wp_w("C8",  round(data.get("open_price",   data["price"]), 2))
        wp_w("E8",  data.get("market_state", "UNKNOWN"))
        wp_w("C10", round(data.get("session_high",  data["price"]), 2))
        wp_w("E10", round(data.get("session_low",   data["price"]), 2))
        wp_w("C17", round(data.get("sp500_move", 0.0), 4), "0.00%")

    try:
        wb.save(path)
    except PermissionError:
        print(f"\n  ERROR: Could not save — {os.path.basename(path)} is open in Excel.")
        print("  Close the file and run the command again.")
        sys.exit(1)
    print(f"\nSaved: {path}")
    print(f"  {data['ticker']} -- {data['name']}  |  {data['updated']}")
    print(f"  Price: ${data['price']:.2f}  |  Open: ${data.get('open_price', data['price']):.2f}"
          f"  |  Beta: {data['beta']:.2f}"
          + (f"  |  RF: {data['rf_rate']:.2%}" if data.get("rf_rate") else ""))
    print(f"  S&P 500 today: {data.get('sp500_move', 0):.2%}"
          f"  |  Market: {data.get('market_state', 'UNKNOWN')}")
    print(f"  Revenue: ${data['revenue_m']:,.0f}M  |  EBITDA: ${data['ebitda_m']:,.0f}M")
    print(f"  Net Debt: ${data['debt_m'] - data['cash_m']:,.0f}M")
    if data["revenue_m"] == 0:
        print("\n  Income/balance sheet cells are 0 -- enter manually from the 10-K.")
    print("\n  Adjust growth rates and margins in the Assumptions tab.")


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

def _read_ticker_from_file(path: str) -> str | None:
    """Read the ticker stored in B5 of the Stock Search sheet."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if "Stock Search" in wb.sheetnames:
            val = wb["Stock Search"]["B5"].value
            wb.close()
            return str(val).upper().strip() if val else None
    except Exception:
        pass
    return None


def _update_price_only(path: str, ticker: str):
    """Quick refresh — update current price and Last Updated only."""
    import yfinance as yf
    tk  = yf.Ticker(ticker)
    inf = tk.info
    price = inf.get("currentPrice") or inf.get("regularMarketPrice") or inf.get("previousClose", 0)
    today = datetime.date.today().strftime("%Y-%m-%d")

    wb = load_workbook(path)
    gold = fill(INPUT_GOLD)

    def write(sheet, ref, val):
        c = sheet[ref]
        c.value = val
        c.fill  = gold

    if "Stock Search" in wb.sheetnames:
        ws = wb["Stock Search"]
        write(ws, "C5", round(price, 2))
        write(ws, "E5", today)
        write(ws, "C8", round(price, 2))
    if "Assumptions" in wb.sheetnames:
        write(wb["Assumptions"], "C7", round(price, 2))

    wb.save(path)
    print(f"  Price updated: {ticker} → ${price:.2f}  ({today})")


def _find_dcf_files() -> list[str]:
    """Return DCF_Model*.xlsx files in the script directory, newest first."""
    import glob
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return sorted(
        glob.glob(os.path.join(script_dir, "DCF_Model*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )


def _pick_file(candidates: list[str]) -> str:
    """Show a numbered list of DCF files and return the one the user picks."""
    print(f"\n  Found {len(candidates)} DCF model{'s' if len(candidates) > 1 else ''}:\n")
    for i, path in enumerate(candidates, 1):
        ticker = _read_ticker_from_file(path) or "?"
        print(f"    [{i}] {os.path.basename(path)}  ({ticker})")
    print()
    raw = input(f"  Pick a file [1]: ").strip()
    idx = int(raw) if raw.isdigit() and 1 <= int(raw) <= len(candidates) else 1
    return candidates[idx - 1]


def main():
    parser = argparse.ArgumentParser(
        description="Update DCF_Model.xlsx with Yahoo Finance + SEC EDGAR data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python update_dcf.py                         interactive — pick file, date, overwrite/new
  python update_dcf.py AAPL                    full refresh, live data (auto-picks newest file)
  python update_dcf.py AAPL --date 2020-03-23  historical data, overwrites selected file
  python update_dcf.py AAPL --date 6/1/2026    slash date format also accepted
  python update_dcf.py AAPL --date 2008        year-only → Dec 31 of that year
  python update_dcf.py AAPL --price-only       quick price refresh only
  python update_dcf.py AAPL -o path/file.xlsx  explicit output file

Date formats:  2008  |  2020-03-23  |  6/1/2026  |  present
        """,
    )
    parser.add_argument("ticker", nargs="?", default=None,
                        help="Stock ticker — omit to read from file or be prompted")
    parser.add_argument("--date", "-d", default=None,
                        help="Reference date (YYYY, YYYY-MM-DD, M/D/YYYY, or 'present')")
    parser.add_argument("--price-only", "-p", action="store_true",
                        help="Quick refresh: update current price only")
    parser.add_argument("--output", "-o", default=None,
                        help="Path to target .xlsx (auto-detected if omitted)")
    args = parser.parse_args()

    # ── Resolve output file ───────────────────────────────────────────────────
    if args.output:
        output_path = args.output
        if not os.path.exists(output_path):
            print(f"ERROR: File not found: {output_path}")
            sys.exit(1)
    else:
        candidates = _find_dcf_files()
        if not candidates:
            print("  No DCF_Model*.xlsx files found in this directory.")
            print("  Run  python build_dcf.py  first to create a model.")
            sys.exit(1)
        # Non-interactive (ticker provided via CLI): silently pick newest
        if args.ticker:
            output_path = candidates[0]
        else:
            output_path = _pick_file(candidates)

    # ── Resolve ticker ────────────────────────────────────────────────────────
    interactive = (args.ticker is None)   # capture before prompts mutate args.ticker
    if not args.ticker:
        existing = _read_ticker_from_file(output_path)
        if existing:
            ans = input(f"  Ticker in file: {existing}. Use this? [Y/n]: ").strip().lower()
            args.ticker = existing if ans in ("", "y", "yes") else None
        if not args.ticker:
            args.ticker = input("  Ticker (e.g. AAPL, MSFT, NVDA): ").strip().upper()
        if not args.ticker:
            sys.exit(1)

    ticker = args.ticker.upper()

    # ── Price-only shortcut ───────────────────────────────────────────────────
    if args.price_only:
        print(f"\n  Quick price refresh for {ticker} ...")
        _update_price_only(output_path, ticker)
        return

    # ── Resolve date ──────────────────────────────────────────────────────────
    as_of = None
    if args.date:
        try:
            as_of = parse_date(args.date)
        except ValueError as exc:
            print(f"ERROR: {exc}")
            sys.exit(1)
    elif interactive and not args.output:
        print(f"\n  Update {ticker} with:")
        print("    [1] Live data")
        print("    [2] Historical date")
        mode = input("  Choice [1]: ").strip()
        if mode == "2":
            date_str = input("  Date (YYYY, YYYY-MM-DD, or M/D/YYYY): ").strip()
            if date_str:
                try:
                    as_of = parse_date(date_str)
                except ValueError as exc:
                    print(f"ERROR: {exc}")
                    sys.exit(1)

    # ── Overwrite or new file (always ask in interactive mode) ────────────────
    final_path = output_path
    if interactive and not args.output:
        if as_of is not None:
            new_name = f"DCF_Model_{ticker}_{as_of.strftime('%Y-%m-%d')}.xlsx"
        else:
            stamp    = datetime.datetime.now().strftime("%Y-%m-%d_%I-%M%p")
            new_name = f"DCF_Model_{ticker}_{stamp}.xlsx"
        new_path = os.path.join(os.path.dirname(os.path.abspath(output_path)), new_name)
        print(f"\n  Save as:")
        print(f"    [1] Overwrite  {os.path.basename(output_path)}")
        print(f"    [2] New file   {new_name}")
        raw = input("  Choice [1]: ").strip()
        if raw == "2":
            import shutil
            shutil.copy2(output_path, new_path)
            final_path = new_path
            print(f"  Copied → {new_name}")

    # ── Fetch and write ───────────────────────────────────────────────────────
    label = f"{ticker} as of {as_of}" if as_of else f"{ticker} (live)"
    print(f"\n  Fetching {label} ...")
    data = fetch_stock_data(ticker, as_of_date=as_of)
    update_excel(final_path, data)


if __name__ == "__main__":
    main()
