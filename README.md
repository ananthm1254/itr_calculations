# ITR Calculations - Foreign Assets Tax Calculator

A comprehensive Python tool for calculating Income Tax Return (ITR) values for foreign assets, including dividends, ESPP, RSU, and Schedule FA reporting. Automatically converts USD to INR using SBI TT Buy rates with proper date-based exchange rate lookup.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Detailed Usage](#detailed-usage)
- [Input File Format](#input-file-format)
- [Output Files](#output-files)
- [Command Line Options](#command-line-options)
- [Calculations Explained](#calculations-explained)
- [Troubleshooting](#troubleshooting)
- [Examples](#examples)

---

## Features

### 💰 Dividend Income Calculation
- Converts foreign dividend income and tax to INR
- Uses SBI TT Buy rate from the last day of preceding month
- Automatically handles date-based exchange rate lookup

### 📊 ESPP (Employee Stock Purchase Plan)
- **Buy Transactions**: Converts purchase prices to INR
- **Sale Transactions**: Converts sale proceeds to INR
- **FIFO Matching**: Automatically matches sales to purchases
- **Capital Gains**: Calculates LTCG (>24 months) and STCG (<24 months)

### 🎁 RSU (Restricted Stock Units)
- **Vest Transactions**: Converts vesting values to INR
- **Sale Transactions**: Converts sale proceeds to INR
- **FIFO Matching**: Automatically matches sales to vests
- **Capital Gains**: Calculates LTCG and STCG with holding period

### 📋 Schedule FA (Foreign Assets)
- **Opening & Closing Values**: Converts portfolio values to INR
- **Peak Value Calculation**: Finds maximum portfolio value during the year
  - Tracks daily share balances
  - Fetches historical stock prices
  - Calculates: `max(Daily Stock Price × Shares Held)`
- **Positive Cash Summary**: Totals all deposits/credits in INR

### 🔄 Smart Exchange Rate Lookup
- Uses SBI TT Buy rates from `sbi-fx-ratekeeper` repository
- Reference date: Last day of preceding month
- Falls back to nearest available rate if exact date missing
- Handles weekends and holidays automatically

---

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)

### Step 1: Clone or Download
```bash
cd your-workspace-directory
git clone <your-repo-url>
cd itr_calculations
```

### Step 2: Install Dependencies
```bash
python -m pip install pandas openpyxl python-dateutil yfinance
```

**Required packages:**
- `pandas` - Data processing
- `openpyxl` - Excel file handling
- `python-dateutil` - Date calculations
- `yfinance` - Stock price data (optional, for peak value calculation)

### Step 3: Verify Installation
```bash
python calculate_itr_values.py --help
```

You should see the help message with available options.

---

## Quick Start

### Basic Usage (Default Files)
```bash
python calculate_itr_values.py
```

This will:
1. Read `ITR_Foreign_Assets.xlsx`
2. Use exchange rates from `sbi-fx-ratekeeper/csv_files/SBI_REFERENCE_RATES_USD.csv`
3. Generate `ITR_Calculated_Values.xlsx`

### Custom Files
```bash
python calculate_itr_values.py -i my_assets.xlsx -o my_results.xlsx
```

### Custom Stock Ticker (for Schedule FA)
```bash
python calculate_itr_values.py -t AAPL
```

---

## Detailed Usage

### Input File Structure

Your input Excel file (`ITR_Foreign_Assets.xlsx`) should contain the following sheets:

#### 1. **Dividend_FY** Sheet
Columns:
- `Date of Dividend` - Transaction date (DD/MM/YYYY)
- `Value (in $)` - Dividend amount in USD
- `Tax (in $)` - Tax withheld in USD

Example:
```
Date of Dividend | Value (in $) | Tax (in $)
23/10/2024      | 1.79         | 0.45
15/01/2025      | 1.79         | 0.45
```

#### 2. **ESPP-Buy** Sheet
Columns:
- `Transaction date` - Purchase date (DD/MM/YYYY)
- `Purchase/Sale FMV (in $)` - Fair Market Value per share
- `No. of Shares` - Number of shares purchased

Example:
```
Transaction date | Purchase/Sale FMV (in $) | No. of Shares
31/01/2022      | 79.27                    | 14.000
29/07/2022      | 62.07                    | 17.674
```

#### 3. **ESPP-Sale** Sheet
Same columns as ESPP-Buy, but for sale transactions.

#### 4. **RSU-Vest** Sheet
Columns:
- `Transaction date` - Vesting date (DD/MM/YYYY)
- `Purchase/Sale FMV (in $)` - Fair Market Value per share
- `No. of Shares` - Number of shares vested

#### 5. **RSU-Sale** Sheet
Same columns as RSU-Vest, but for sale transactions.

#### 6. **ESPP-Assets** Sheet (for Schedule FA)
Columns:
- `Date` - Transaction date
- `Cash/Share` - Type: "Opening", "Closing", "Share", or "Cash"
- `No. of Shares` - Number of shares (for Opening/Share rows)
- `Cash (in $)` - Cash amount (for Cash rows)
- `Market Value (in $)` - Portfolio value (for Opening/Closing rows)

Example:
```
Date       | Cash/Share | No. of Shares | Cash (in $) | Market Value (in $)
01/01/2023 | Opening    | 31.674        |             | 1583.07
19/01/2023 | Cash       |               | 3.64        |
31/01/2023 | Share      | 17.145        |             |
31/12/2023 | Closing    | 66.225        |             | 5678.90
```

**Row Types:**
- **Opening**: Portfolio value at start of year
- **Closing**: Portfolio value at end of year
- **Share**: Share additions (vesting, purchases)
- **Cash**: Cash transactions (positive = deposits, negative = fees)

---

## Output Files

### Generated Excel File: `ITR_Calculated_Values.xlsx`

Contains multiple sheets:

#### 1. **Dividend_Calculated**
- Transaction Date
- Reference Date (Month End)
- SBI Rate Date
- SBI TT Buy Rate
- Value (Foreign) & Value (INR)
- Tax (Foreign) & Tax (INR)

#### 2. **ESPP_Buy_Calculated**
- Transaction Date
- Reference Date
- SBI TT Buy Rate
- FMV per Share (USD & INR)
- Total Purchase Price (USD & INR)

#### 3. **ESPP_Sale_Calculated**
- Similar to Buy, but for sales

#### 4. **ESPP_Matched_Transactions**
- Sale Date & Purchase Date
- Holding Period (Days & Months)
- Gain Type (LTCG or STCG)
- Shares Sold
- Purchase & Sale Prices
- Capital Gain/Loss (INR)
- LTCG (INR) & STCG (INR)

#### 5. **RSU_Vest_Calculated** & **RSU_Sale_Calculated**
- Similar to ESPP sheets

#### 6. **RSU_Matched_Transactions**
- Similar to ESPP matched transactions

#### 7. **Schedule_FA_Details**
- All ESPP-Assets transactions with INR conversions

#### 8. **Schedule_FA_Summary**
- Opening Date & Value (INR)
- Closing Date & Value (INR)
- Total Shares
- Positive Cash Total (USD & INR)
- Peak Date, Price, & Value (INR)

---

## Command Line Options

```bash
python calculate_itr_values.py [OPTIONS]
```

### Options:

| Option | Long Form | Description | Default |
|--------|-----------|-------------|---------|
| `-h` | `--help` | Show help message | - |
| `-i FILE` | `--input FILE` | Input Excel file path | `ITR_Foreign_Assets.xlsx` |
| `-o FILE` | `--output FILE` | Output Excel file path | `ITR_Calculated_Values.xlsx` |
| `-t TICKER` | `--ticker TICKER` | Stock ticker for peak value | `MU` |

### Examples:

```bash
# Use default files
python calculate_itr_values.py

# Custom input file
python calculate_itr_values.py -i FY2023-24.xlsx

# Custom input and output
python calculate_itr_values.py -i FY2023-24.xlsx -o Results_2023-24.xlsx

# Different stock ticker
python calculate_itr_values.py -t AAPL

# All options combined
python calculate_itr_values.py -i assets.xlsx -o results.xlsx -t MSFT
```

---

## Calculations Explained

### 1. Exchange Rate Lookup

**Reference Date**: Last day of preceding month

Example:
- Transaction Date: 15/01/2025
- Reference Date: 31/12/2024
- Uses SBI TT Buy rate for 31/12/2024

**Fallback Logic**:
- If exact date not found (weekend/holiday), uses nearest preceding date
- Skips rates with value 0.00

### 2. Capital Gains Classification

**Long-Term Capital Gains (LTCG)**:
- Holding period ≥ 24 months
- Taxed at 10% (without indexation) if gains > ₹1 lakh

**Short-Term Capital Gains (STCG)**:
- Holding period < 24 months
- Taxed at applicable slab rates

**Holding Period Calculation**:
```
Holding Period (months) = (Sale Date - Purchase Date) in days / 30.44
```

### 3. FIFO Matching

Sales are matched to purchases/vests using First-In-First-Out:

Example:
```
Purchases:
- 2022-01-31: 14 shares @ $79.27
- 2022-07-29: 18 shares @ $62.07

Sale:
- 2024-03-25: 20 shares @ $119.50

Matching:
- 14 shares from 2022-01-31 purchase
- 6 shares from 2022-07-29 purchase
```

### 4. Peak Value Calculation

**Algorithm**:
1. Build share timeline (track when shares were added)
2. For each trading day in the year:
   - Get stock closing price
   - Determine shares held on that date
   - Calculate: `Portfolio Value = Price × Shares`
3. Find the maximum value across all days
4. Convert to INR using that date's exchange rate

**Example**:
```
Timeline:
- Jan 1: 30 shares @ $100 = $3,000
- Feb 1: 50 shares @ $90 = $4,500 (added 20 shares)
- Mar 1: 50 shares @ $95 = $4,750 ← PEAK
- Dec 31: 50 shares @ $85 = $4,250

Peak Value: $4,750 on Mar 1
```

---

## Troubleshooting

### Issue: "Excel file not found"
**Solution**: Ensure `ITR_Foreign_Assets.xlsx` is in the same directory, or use `-i` to specify path.

### Issue: "CSV file not found"
**Solution**: Ensure `sbi-fx-ratekeeper` folder exists with the CSV file at:
```
sbi-fx-ratekeeper/csv_files/SBI_REFERENCE_RATES_USD.csv
```

### Issue: "yfinance not installed" warning
**Solution**: Install yfinance for peak value calculation:
```bash
python -m pip install yfinance
```
Peak value will be skipped if yfinance is not available.

### Issue: "No exchange rate found"
**Solution**: 
- Check if SBI CSV has data for your transaction dates
- Ensure dates are in DD/MM/YYYY format
- The script will use the nearest preceding available rate

### Issue: Incorrect peak value
**Solution**:
- Verify ESPP-Assets sheet has correct Opening/Share/Closing rows
- Check stock ticker symbol is correct (use `-t` option)
- Ensure dates are in correct format

### Issue: Character encoding errors in console
**Solution**: This is cosmetic only. The Excel file will have correct values. You can ignore Rupee symbol (₹) display issues in console.

---

## Examples

### Example 1: Basic Dividend Calculation

**Input** (Dividend_FY sheet):
```
Date of Dividend | Value (in $) | Tax (in $)
15/01/2025      | 1.79         | 0.45
```

**Process**:
1. Reference Date: 31/12/2024
2. SBI TT Buy Rate: 85.20
3. Value (INR) = 1.79 × 85.20 = ₹152.51
4. Tax (INR) = 0.45 × 85.20 = ₹38.34

### Example 2: ESPP Capital Gains

**Input**:
```
Buy:  31/01/2022, 14 shares @ $79.27
Sale: 25/03/2024, 14 shares @ $119.50
```

**Process**:
1. Holding Period = 25.8 months → LTCG
2. Purchase (Jan 2022): Rate = 74.15, Cost = ₹82,179
3. Sale (Mar 2024): Rate = 82.75, Proceeds = ₹138,006
4. Capital Gain = ₹55,827 (LTCG)

### Example 3: Schedule FA Peak Value

**Input** (ESPP-Assets):
```
01/01/2023 | Opening | 31.67 shares | Market Value: $1,583
31/01/2023 | Share   | 17.15 shares added
31/12/2023 | Closing | 66.22 shares | Market Value: $5,679
```

**Process**:
1. Fetch MU stock prices for 2023
2. For each day, calculate: Price × Shares held
3. Peak found: 26/12/2023, Price $86.33, Shares 66.22
4. Peak Value = $5,716 = ₹4,73,190

---

## Tax Filing Reference

### For ITR-2 Form

**Schedule FA (Foreign Assets)**:
- Opening Balance: Use "Opening Value (INR)"
- Closing Balance: Use "Closing Value (INR)"
- Peak Balance: Use "Peak Value (INR)"
- Total Investment: Use "Positive Cash Total (INR)"

**Schedule CG (Capital Gains)**:
- Long-term gains: Sum of all LTCG (INR) from matched transactions
- Short-term gains: Sum of all STCG (INR) from matched transactions

**Schedule OS (Other Sources - Dividend)**:
- Gross Dividend: Sum of all "Value (INR)" from Dividend_Calculated
- Foreign Tax Paid: Sum of all "Tax (INR)" from Dividend_Calculated

---

## Notes

1. **Date Format**: Always use DD/MM/YYYY in Excel
2. **Decimal Precision**: Values rounded to 2 decimal places
3. **Currency**: All foreign amounts assumed to be in USD
4. **Tax Year**: Based on transaction dates in your input file
5. **FIFO**: Cannot be changed; required by Indian tax law
6. **Backup**: Always keep a copy of your input Excel file

---

## Support

For issues or questions:
1. Check the [Troubleshooting](#troubleshooting) section
2. Verify your input file format matches the examples
3. Review the console output for specific error messages
4. Check the generated Excel file for detailed calculations

---

## License

Personal use only. Not for commercial distribution.

---

## Changelog

### Latest Version
- ✅ Dividend income calculation with INR conversion
- ✅ ESPP buy/sell processing with FIFO matching
- ✅ RSU vest/sale processing with FIFO matching
- ✅ LTCG/STCG classification (24-month threshold)
- ✅ Schedule FA with peak value calculation
- ✅ Command-line arguments support
- ✅ Automatic exchange rate fallback
- ✅ Comprehensive Excel output with multiple sheets

---

**Happy Tax Filing! 🎉**
