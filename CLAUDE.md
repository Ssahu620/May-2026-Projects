# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Amazon vendor analysis tool for BHF. Identifies items where conversion rates (CR) increased month-over-month **without** any active Deal or Coupon Rebate, then further classifies those organic increases by whether retail price fell in the same month — distinguishing truly organic growth from price-cut-driven lift.

## How to Run

```
python cr_rebate_analysis.py
```

Dependencies: `pip install pandas openpyxl`

Close `CR_Increase_No_Rebate_Analysis.xlsx` in Excel before running or a clear `PermissionError` is raised.

## Input Files

| File | Key Columns |
|---|---|
| `v2.Conversion Rate.xlsx` | `Item`, `ASIN`, `Brand`, `Program-Category`, `Description \| Color \| Size`, then one column per month (`Jan 2026`–`Dec 2026`; add columns as months complete) |
| `v2.2026 VC Master Data.xlsx` | `Item`, `MonthYear`, `ASIN`, `Brand`, `Description \| Color \| Size`, `Deals Rebate`, `Coupons Rebate` |
| `v2.Retail Prices.xlsx` | `Item`, `ASIN`, `Brand`, `Program-Category`, `Description \| Color \| Size`, then one column per month (`Dec 2025`–present; add columns as months complete) |

## Code Architecture (`cr_rebate_analysis.py`)

The script is a single-file pipeline with no classes. Execution flows through `main()` in seven steps:

1. **`load_data()`** — reads all three Excel files into DataFrames (`cr_df`, `vc_df`, `rp_df`).
2. **`find_cr_increases(cr_df)`** — iterates over consecutive month pairs defined in the `MONTHS` constant, recording one row per item per month where `curr > prev`. Each row includes a `CR Change` column (absolute pp difference = `curr - prev`).
3. **`find_rebate_months(vc_df)`** — extracts `(Item, MonthYear)` pairs where `Deals Rebate > 0` OR `Coupons Rebate > 0`, deduplicated to one row per item-month.
4. **`split_by_rebate(cr_increases, rebate_months)`** — left-joins on `(Item, Month)`; unmatched rows → `no_rebate`, matched → `with_rebate`. Both Item columns are cast to `float` before the merge to avoid type-mismatch misses.
5. **`classify_by_price(no_rebate, rp_df)`** — for each no-rebate row, looks up retail price in the current and previous month; tags each record with a `Price Direction` field (`"Down"`, `"Up"`, or `"Flat / No Data"`). Records where `curr < prev` → `price_down`; all others → `price_stable`. Adds `Prev Price`, `Curr Price`, `Price Change`, `Price Direction` columns. Rows where May 2026 price is unavailable default to `"Flat / No Data"` and land in `price_stable`.
6. **`build_rebate_detail(vc_df)`** / **`build_brand_summary(price_stable)`** — aggregate supporting tables. The brand summary uses `price_stable` (no rebate + price did NOT decrease), which is the same dataset as the D5 "Unique Items" KPI box — both always show consistent numbers.
7. **`save_output()`** → **`write_executive_summary()`** + five **`write_data_sheet()`** calls → saves the workbook.

### Key Constant — `MONTHS`

```python
MONTHS = [
    "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026",
    "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026", "Oct 2026",
    "Nov 2026", "Dec 2026",
]
```

All 12 months of 2026 are pre-listed. `find_cr_increases` skips any pair where either column is absent from the CR file, so **no code change is needed as you add months throughout the year** — just add the new column to the Excel input and re-run. `main()` prints which months were detected at runtime. The Executive Summary methodology note is built dynamically from the months actually present.

### Excel Formatting Layer

All openpyxl styling is done via four helpers at the top of the file:

- `_fill(hex_)`, `_font(...)`, `_align(...)`, `_border(...)` — return openpyxl style objects.
- `_write_title_rows`, `_write_header_row`, `_write_data_row` — apply those styles to worksheet rows.
- `_auto_width(ws)` — sizes columns by content length.

Column formatting (percent, currency, text alignment) is controlled by the `pct_col_names`, `cur_col_names`, and `text_col_names` arguments passed to `write_data_sheet()`.

### Color-Matching Design Rule

Each KPI box in the Executive Summary uses the **exact same color** as the tab of its corresponding detail sheet — white text on all boxes. The palette constants used throughout:

| Constant | Hex | Used for |
|---|---|---|
| `NAVY` | `1F3864` | Executive Summary tab, title banner |
| `BLUE` | `2E75B6` | Rebate Detail tab; Total CR Increases KPI box |
| `GREEN_TAB` | `538135` | CR Up - No Rebate tab; "CR Increased No Rebate" KPI box; Brand Summary header |
| `LIME_TAB` | `70AD47` | No Rebate - Price Stable tab; "Unique Items (No Price Decrease)" and "Price Stable" KPI boxes |
| `RED_TAB` | `C00000` | No Rebate - Price Down tab; "Price Decreased" KPI box |
| `ORANGE_TAB` | `C55A11` | CR Up - Had Rebate tab; "CR Increased With Rebate" KPI box |

Light pastel variants (`GREEN_BG`, `GREEN_FG`, `ORANGE_BG`, etc.) were removed; the tab colors are used directly with white text.

## Output Sheets

| Sheet | Tab Color | Contents |
|---|---|---|
| `Executive Summary` | Navy | 6 KPI boxes (total, with rebate, no rebate, **unique items = `price_stable`** — no rebate AND no price decrease, price-down count, price-stable count) + brand summary + methodology note |
| `CR Up - No Rebate` | Green (`538135`) | All no-rebate CR increases (parent list) |
| `No Rebate - Price Down` | Red (`C00000`) | No-rebate items where retail price fell that month; includes `Prev Price`, `Curr Price`, `Price Change`, `Price Direction` columns |
| `No Rebate - Price Stable` | Lime (`70AD47`) | No-rebate items where price held or rose; same extra price columns |
| `CR Up - Had Rebate` | Orange (`C55A11`) | CR increases where a rebate was active |
| `Rebate Detail` | Blue (`2E75B6`) | All item-months with non-zero rebate, with dollar totals |

## Multi-Year Support

### What works automatically (no code changes needed)

| Function | Why it is year-agnostic |
|---|---|
| `find_rebate_months` | Reads the `MonthYear` column from VC Master Data — accepts any year |
| `split_by_rebate` | Joins on `(Item, Month)` string — works for any month label |
| `classify_by_price` | Looks up retail price column by name — works for any year present in the file |
| All formatting / output functions | Driven entirely by the data passed in |

### What requires changes for 2026 (within-year)

Nothing. `MONTHS` already contains all 12 months of 2026. Just add the new month column to each Excel input file and re-run — the script detects which columns are present and skips the rest.

### What requires changes for prior/future years

**1. `MONTHS` constant** — add the new year's months in strict chronological order. No calendar gaps between consecutive entries:

```python
MONTHS = [
    "Jan 2026", ..., "Dec 2026",
    "Jan 2027", ..., "Dec 2027",   # extend when needed
]
```

**2. `INPUT_VC` file name** — hardcoded as `"v2.2026 VC Master Data.xlsx"`. If prior/future year rebate data lives in a separate file, those CR increases would be silently treated as "no rebate." Options:
- Combine all years into one file and update the constant.
- Load multiple files and `pd.concat()` them in `load_data()`.

**3. Retail Prices file** — `classify_by_price` looks up month columns by name. Price data for months absent from `v2.Retail Prices.xlsx` defaults to `"Flat / No Data"` (safe, no evidence of a price drop).

### Critical risk — the gap problem

The code compares **every adjacent pair** in `MONTHS` as a true month-over-month change. If bridging years with only partial data, a gap is created:

```
Dec 2026  →  Jan 2027   ← OK if both columns exist in CR file
May 2026  →  Jan 2027   ← 8-month gap if Jun–Dec 2026 columns are absent ✗
```

**Safe rule:** Only extend `MONTHS` across a year boundary once you have complete 12-month data for the prior year in the CR file.

## Known Open Items

- **Retail Prices lag**: `v2.Retail Prices.xlsx` typically trails by one month; the most recent CR-increase month has no current price and defaults to `price_stable` (Flat / No Data).
- **Multi-year VC Master Data**: `INPUT_VC` is named `v2.2026 VC Master Data.xlsx`; if prior-year rebate files exist they must be merged before running.
