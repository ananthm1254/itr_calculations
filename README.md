# ITR Calculations - Foreign Assets Tax Calculator

Python tool for calculating ITR values for foreign assets (dividends, ESPP, RSU, Schedule FA). Automatically converts USD to INR using SBI TT Buy rates.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Input Format](#input-format)
- [Output](#output)
- [Calculations](#calculations)
- [Troubleshooting](#troubleshooting)

---

## Features

- **Dividend Income**: Converts foreign dividends and tax to INR using SBI TT Buy rate from last day of preceding month
- **ESPP/RSU**: Processes buy/vest and sale transactions with FIFO matching
- **Capital Gains**: Calculates LTCG (≥24 months) and STCG (<24 months)
- **Schedule FA**: 
  - Opening/Closing portfolio values in INR
  - Peak value calculation (max daily portfolio value during year)
  - Positive cash summary
- **Smart Exchange Rates**: Auto-fallback to nearest date for weekends/holidays

---

## Installation

**Prerequisites**: Python 3.7+

```bash
# Install dependencies
python -m pip install pandas openpyxl python-dateutil yfinance

# Verify installation
python calculate_itr_values.py --help
```

**Note**: The tool automatically fetches the latest SBI exchange rates from GitHub. No additional setup required.

---

## Usage

```bash
# Basic usage (default files)
python calculate_itr_values.py

# Custom files and ticker
python calculate_itr_values.py -i FY2023-24.xlsx -o Results.xlsx -t AAPL
```

**Options**:
- `-i FILE` - Input Excel file (default: `ITR_Foreign_Assets.xlsx`)
- `-o FILE` - Output Excel file (default: `ITR_Calculated_Values.xlsx`)
- `-t TICKER` - Stock ticker for peak value (default: `MU`)

---

## Input Format

Input Excel file should contain these sheets:

### 1. Dividend_FY
```
Date of Dividend | Value (in $) | Tax (in $)
23/10/2024      | 1.79         | 0.45
```

### 2. ESPP-Buy / ESPP-Sale
```
Transaction date | Purchase/Sale FMV (in $) | No. of Shares
31/01/2022      | 79.27                    | 14.000
```

### 3. RSU-Vest / RSU-Sale
Same format as ESPP sheets.

### 4. ESPP-Assets (Schedule FA)
```
Date       | Cash/Share | No. of Shares | Cash (in $) | Market Value (in $)
01/01/2023 | Opening    | 31.674        |             | 1583.07
19/01/2023 | Cash       |               | 3.64        |
31/01/2023 | Share      | 17.145        |             |
31/12/2023 | Closing    | 66.225        |             | 5678.90
```

**Row Types**: Opening (year start), Closing (year end), Share (additions), Cash (deposits/fees)

### 5. RSU-Assets (Schedule FA)
Same format as ESPP-Assets sheet.

### 6. Cash (Combined Portfolio)
```
Date       | Type    | ESPP value | RSU value
01/01/2023 | Opening | 4.99       | 0.0
19/01/2023 | Cash    | 2.73       |
31/12/2023 | Closing | 21.85      | 0.0
```

**Row Types**: Opening (year start), Closing (year end), Cash (cash transactions for ESPP/RSU)

**Date Format**: DD/MM/YYYY

---

## Output

Generated file `ITR_Calculated_Values.xlsx` contains:

1. **Summary** - Comprehensive ITR summary (first sheet)
   - Dividend income with India tax (31.2%)
   - Dividend quarterly breakdown (Q1-Q4, with 16-31 Mar separately)
   - LTCG summary (combined ESPP+RSU): Sale proceeds, purchase cost, total gain
   - LTCG quarterly breakdown
   - STCG summary (combined ESPP+RSU): Sale proceeds, purchase cost, total gain
   - STCG quarterly breakdown
   - Schedule FA - ESPP: Opening, Peak, Closing values, Positive cash
   - Schedule FA - RSU: Opening, Peak, Closing values, Positive cash
   - Schedule FA - Cash: Combined portfolio opening, peak, closing values
   
2. **Dividend_Calculated** - Dividends with INR conversion
3. **ESPP_Buy_Calculated** / **ESPP_Sale_Calculated** - Transactions with exchange rates
4. **ESPP_Matched_Transactions** - FIFO-matched sales with capital gains (LTCG/STCG)
5. **RSU_Vest_Calculated** / **RSU_Sale_Calculated** / **RSU_Matched_Transactions** - Same as ESPP
6. **Schedule_FA_ESPP_Details** / **Schedule_FA_ESPP_Summary** - ESPP asset transactions and summary
7. **Schedule_FA_RSU_Details** / **Schedule_FA_RSU_Summary** - RSU asset transactions and summary
8. **Cash_Details** / **Cash_Summary** - Combined ESPP+RSU portfolio values and peak calculation

---

## Calculations

### Exchange Rate
- **Reference Date**: Last day of preceding month
- **Example**: Transaction on 15/01/2025 → uses rate for 31/12/2024
- **Fallback**: Nearest preceding date if exact date unavailable

### Capital Gains
- **LTCG**: Holding ≥24 months (taxed at 10% if >₹1 lakh)
- **STCG**: Holding <24 months (taxed at slab rates)
- **Holding Period**: `(Sale Date - Purchase Date) / 30.44` months

### FIFO Matching
Sales matched to earliest purchases/vests first.

**Example**:
```
Purchases: 14 shares (2022-01-31), 18 shares (2022-07-29)
Sale: 20 shares (2024-03-25)
→ Matches: 14 from first purchase, 6 from second
```

### Peak Value
1. Track daily share balances
2. Fetch daily stock prices for the year
3. Calculate: `max(Daily Price × Shares Held)`
4. Convert to INR using peak date's exchange rate

### Cash Sheet Peak Value
For the combined portfolio (Cash sheet):
1. Track cash balance timeline (Opening + Cash transactions + Closing)
2. Convert each entry to INR using that date's exchange rate
3. Find maximum INR value across all timeline entries
4. Report the date and value with highest INR amount

**Note**: Peak is based on maximum INR value, not USD value, since exchange rates vary.

**Example**:
```
2023-10-25: $21.85 @ 82.68 = ₹1,806.56
2023-12-31: $21.85 @ 82.78 = ₹1,808.74 ← Peak (higher INR)
```

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Excel file not found | Ensure file exists or use `-i` option |
| Network error fetching rates | Check internet connection; rates are fetched from GitHub |
| yfinance not installed | Run `python -m pip install yfinance` (optional, for peak value) |
| No exchange rate found | Script uses nearest preceding rate; check your transaction dates |
| Incorrect peak value | Verify ESPP-Assets sheet format and ticker symbol (`-t` option) |
| Console encoding errors | Cosmetic only; Excel file has correct values |

---

## Tax Filing Reference (ITR-2)

**Schedule FA**:
- Opening/Closing/Peak Balance: From Schedule_FA_Summary
- Total Investment: Positive Cash Total (INR)

**Schedule CG**:
- Long-term gains: Sum LTCG (INR) from matched transactions
- Short-term gains: Sum STCG (INR) from matched transactions

**Schedule OS (Dividend)**:
- Gross Dividend: Sum Value (INR) from Dividend_Calculated
- Foreign Tax Paid: Sum Tax (INR) from Dividend_Calculated

---

## Examples

### Dividend Calculation
```
Input: 15/01/2025, $1.79 dividend, $0.45 tax
Process: Reference 31/12/2024, Rate 85.20
Output: ₹152.51 dividend, ₹38.34 tax
```

### ESPP Capital Gains
```
Buy: 31/01/2022, 14 shares @ $79.27
Sale: 25/03/2024, 14 shares @ $119.50
Result: 25.8 months → LTCG of ₹55,827
```

### Peak Value
```
Timeline: Jan 1 (30 shares @ $100), Feb 1 (50 shares @ $90), 
          Mar 1 (50 shares @ $95 ← PEAK), Dec 31 (50 shares @ $85)
Peak: $4,750 on Mar 1
```

---

## Notes

- **Date Format**: DD/MM/YYYY
- **Currency**: USD only
- **FIFO**: Required by Indian tax law
- **Precision**: 2 decimal places
- **Backup**: Keep copy of input file

---

**Happy Tax Filing! 🎉**
