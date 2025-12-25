import pandas as pd
import datetime
from pathlib import Path
import os
import argparse
import sys
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False
    print("Warning: yfinance not installed. Peak value calculation will be skipped.")

# Paths
BASE_DIR = Path(r"c:\Users\anant\OneDrive\Documents\itr_calculations")
EXCEL_PATH = BASE_DIR / "ITR_Foreign_Assets.xlsx"
CSV_PATH = BASE_DIR / "sbi-fx-ratekeeper" / "csv_files" / "SBI_REFERENCE_RATES_USD.csv"
OUTPUT_PATH = BASE_DIR / "ITR_Calculated_Values.xlsx"

def get_last_day_of_preceding_month(date_val):
    """
    Returns the last day of the month preceding the date_val.
    E.g. 2023-05-15 -> 2023-04-30, 2025-01-15 -> 2024-12-31
    """
    # Ensure it's a pandas Timestamp
    ts = pd.Timestamp(date_val)
    # Move to the first day of the current month, then subtract 1 day
    first_of_current_month = ts.replace(day=1)
    last_of_previous_month = first_of_current_month - pd.Timedelta(days=1)
    return last_of_previous_month

def get_exchange_rate(target_date, rates_df):
    """
    Finds the TT BUY rate for the target_date.
    If rate is 0 or missing, looks for the previous available non-zero rate.
    Returns (rate, date_of_rate)
    """
    # Normalize target_date to just the date (no time component)
    target_ts = pd.Timestamp(target_date).normalize()
    
    # Create a normalized date column for comparison (if not already exists)
    if 'DATE_NORMALIZED' not in rates_df.columns:
        rates_df['DATE_NORMALIZED'] = rates_df['DATE'].dt.normalize()
    
    # Filter rates on or before target_date (comparing date-only)
    available = rates_df[rates_df['DATE_NORMALIZED'] <= target_ts]
    
    if available.empty:
        return None, None
        
    # Sort descending by date to find the closest one
    available = available.sort_values('DATE', ascending=False)
    
    # Iterate to find first non-zero rate
    for _, row in available.iterrows():
        rate = row['TT BUY']
        # Check if rate is valid and > 0
        if pd.notna(rate) and rate > 0:
            return rate, row['DATE']
            
    return None, None

def process_dividend_sheet(df, rates_df):
    """Process the Dividend_FY sheet"""
    cols = df.columns
    
    # Heuristics for column names
    date_col = next((c for c in cols if 'date' in str(c).lower() and 'txn' not in str(c).lower()), 
                    next((c for c in cols if 'date' in str(c).lower()), None))
    val_col = next((c for c in cols if 'value' in str(c).lower()), 
                   next((c for c in cols if 'amount' in str(c).lower()), None))
    tax_col = next((c for c in cols if 'tax' in str(c).lower()), None)
    
    if not date_col:
        print("WARNING: Could not identify 'Date' column. Trying first column.")
        date_col = cols[0]
    
    print(f"Dividend Mapping -> Date: '{date_col}', Value: '{val_col}', Tax: '{tax_col}'")
    
    results = []
    
    for idx, row in df.iterrows():
        try:
            date_val = pd.to_datetime(row[date_col], dayfirst=True)
        except Exception as e:
            print(f"Row {idx}: Invalid date '{row[date_col]}'. Skipping.")
            continue
            
        ref_date = get_last_day_of_preceding_month(date_val)
        rate, found_date = get_exchange_rate(ref_date, rates_df)
        
        val_foreign = 0.0
        if val_col and pd.notna(row[val_col]):
             val_foreign = pd.to_numeric(row[val_col], errors='coerce') or 0.0
             
        tax_foreign = 0.0
        if tax_col and pd.notna(row[tax_col]):
            tax_foreign = pd.to_numeric(row[tax_col], errors='coerce') or 0.0
            
        val_inr = val_foreign * rate if rate else 0.0
        tax_inr = tax_foreign * rate if rate else 0.0
        
        res_row = {
            'Transaction Date': date_val,
            'Reference Date (Month End)': ref_date.strftime('%Y-%m-%d'),
            'SBI Rate Date': found_date.strftime('%Y-%m-%d') if found_date else 'N/A',
            'SBI TT Buy Rate': rate if rate else 'N/A',
            'Value (Foreign)': val_foreign,
            'Tax (Foreign)': tax_foreign,
            'Value (INR)': round(val_inr, 2),
            'Tax (INR)': round(tax_inr, 2)
        }
        results.append(res_row)

    return pd.DataFrame(results)

def process_espp_buy_sheet(df, rates_df):
    """Process the ESPP-Buy sheet"""
    results = []
    
    for idx, row in df.iterrows():
        try:
            date_val = pd.to_datetime(row['Transaction date'], dayfirst=True)
        except Exception as e:
            print(f"ESPP-Buy Row {idx}: Invalid date '{row['Transaction date']}'. Skipping.")
            continue
            
        ref_date = get_last_day_of_preceding_month(date_val)
        rate, found_date = get_exchange_rate(ref_date, rates_df)
        
        fmv_usd = pd.to_numeric(row['Purchase/Sale FMV (in $)'], errors='coerce') or 0.0
        shares = pd.to_numeric(row['No. of Shares'], errors='coerce') or 0.0
        
        total_usd = fmv_usd * shares
        total_inr = total_usd * rate if rate else 0.0
        fmv_inr = fmv_usd * rate if rate else 0.0
        
        res_row = {
            'Transaction Date': date_val,
            'Reference Date (Month End)': ref_date.strftime('%Y-%m-%d'),
            'SBI Rate Date': found_date.strftime('%Y-%m-%d') if found_date else 'N/A',
            'SBI TT Buy Rate': rate if rate else 'N/A',
            'FMV per Share (USD)': fmv_usd,
            'No. of Shares': shares,
            'Total Purchase Price (USD)': round(total_usd, 2),
            'FMV per Share (INR)': round(fmv_inr, 2),
            'Total Purchase Price (INR)': round(total_inr, 2)
        }
        results.append(res_row)

    return pd.DataFrame(results)

def process_espp_sale_sheet(df, rates_df):
    """Process the ESPP-Sale sheet"""
    results = []
    
    for idx, row in df.iterrows():
        try:
            date_val = pd.to_datetime(row['Transaction date'], dayfirst=True)
        except Exception as e:
            print(f"ESPP-Sale Row {idx}: Invalid date '{row['Transaction date']}'. Skipping.")
            continue
            
        ref_date = get_last_day_of_preceding_month(date_val)
        rate, found_date = get_exchange_rate(ref_date, rates_df)
        
        fmv_usd = pd.to_numeric(row['Purchase/Sale FMV (in $)'], errors='coerce') or 0.0
        shares = pd.to_numeric(row['No. of Shares'], errors='coerce') or 0.0
        
        total_usd = fmv_usd * shares
        total_inr = total_usd * rate if rate else 0.0
        fmv_inr = fmv_usd * rate if rate else 0.0
        
        res_row = {
            'Transaction Date': date_val,
            'Reference Date (Month End)': ref_date.strftime('%Y-%m-%d'),
            'SBI Rate Date': found_date.strftime('%Y-%m-%d') if found_date else 'N/A',
            'SBI TT Buy Rate': rate if rate else 'N/A',
            'FMV per Share (USD)': fmv_usd,
            'No. of Shares': shares,
            'Total Sale Price (USD)': round(total_usd, 2),
            'FMV per Share (INR)': round(fmv_inr, 2),
            'Total Sale Price (INR)': round(total_inr, 2)
        }
        results.append(res_row)

    return pd.DataFrame(results)

def match_sales_to_purchases(buy_df, sale_df):
    """
    Match sales to purchases using FIFO (First In First Out) method
    Returns a DataFrame with matched transactions
    """
    # Create copies to track remaining shares
    buy_remaining = buy_df.copy()
    buy_remaining['Shares Remaining'] = buy_remaining['No. of Shares']
    
    matched_results = []
    
    for sale_idx, sale_row in sale_df.iterrows():
        shares_to_sell = sale_row['No. of Shares']
        sale_date = sale_row['Transaction Date']
        sale_price_per_share_inr = sale_row['FMV per Share (INR)']
        
        # FIFO: Match with earliest purchases first
        for buy_idx, buy_row in buy_remaining.iterrows():
            if buy_row['Shares Remaining'] <= 0:
                continue
                
            # Determine how many shares to match
            shares_matched = min(shares_to_sell, buy_row['Shares Remaining'])
            
            if shares_matched > 0:
                # Calculate purchase cost for these shares
                purchase_price_per_share_inr = buy_row['FMV per Share (INR)']
                purchase_cost_inr = shares_matched * purchase_price_per_share_inr
                sale_proceeds_inr = shares_matched * sale_price_per_share_inr
                capital_gain_inr = sale_proceeds_inr - purchase_cost_inr
                
                # Calculate holding period
                purchase_date = buy_row['Transaction Date']
                holding_period_days = (sale_date - purchase_date).days
                holding_period_months = holding_period_days / 30.44  # Average days per month
                
                # Classify as LTCG or STCG (24 months threshold)
                is_long_term = holding_period_months >= 24
                ltcg = capital_gain_inr if is_long_term else 0
                stcg = capital_gain_inr if not is_long_term else 0
                
                matched_results.append({
                    'Sale Date': sale_date,
                    'Purchase Date': purchase_date,
                    'Holding Period (Days)': holding_period_days,
                    'Holding Period (Months)': round(holding_period_months, 1),
                    'Gain Type': 'LTCG' if is_long_term else 'STCG',
                    'Shares Sold': shares_matched,
                    'Purchase Price per Share (INR)': round(purchase_price_per_share_inr, 2),
                    'Sale Price per Share (INR)': round(sale_price_per_share_inr, 2),
                    'Total Purchase Cost (INR)': round(purchase_cost_inr, 2),
                    'Total Sale Proceeds (INR)': round(sale_proceeds_inr, 2),
                    'Capital Gain/Loss (INR)': round(capital_gain_inr, 2),
                    'LTCG (INR)': round(ltcg, 2),
                    'STCG (INR)': round(stcg, 2)
                })
                
                # Update remaining shares
                buy_remaining.at[buy_idx, 'Shares Remaining'] -= shares_matched
                shares_to_sell -= shares_matched
                
                if shares_to_sell <= 0:
                    break
        
        if shares_to_sell > 0:
            print(f"Warning: Sale on {sale_date} has {shares_to_sell} unmatched shares!")
    
    return pd.DataFrame(matched_results)

def process_schedule_fa(df, rates_df, ticker_symbol='MU'):
    """
    Process ESPP-Assets sheet for Schedule FA calculations
    1. Convert Opening/Closing market values to INR
    2. Calculate peak value using year's stock prices
    3. Sum positive cash values in INR
    """
    results = []
    opening_value_inr = 0
    closing_value_inr = 0
    closing_date = None
    opening_date = None
    total_shares = 0
    positive_cash_total_usd = 0
    
    # Process each row
    for idx, row in df.iterrows():
        try:
            date_val = pd.to_datetime(row['Date'], dayfirst=True)
        except:
            continue
            
        cash_share = str(row['Cash/Share']).strip() if pd.notna(row['Cash/Share']) else ''
        shares = pd.to_numeric(row['No. of Shares'], errors='coerce') or 0
        cash_usd = pd.to_numeric(row['Cash (in $)'], errors='coerce') or 0
        market_value_usd = pd.to_numeric(row['Market Value (in $)'], errors='coerce') or 0
        
        # Get exchange rate for this date
        rate, found_date = get_exchange_rate(date_val, rates_df)
        
        # Handle Opening
        if cash_share.lower() == 'opening':
            opening_value_inr = market_value_usd * rate if rate else 0
            opening_date = date_val
            total_shares = shares
            
        # Handle Closing
        elif cash_share.lower() == 'closing':
            closing_value_inr = market_value_usd * rate if rate else 0
            closing_date = date_val
            
        # Handle Share additions
        elif cash_share.lower() == 'share':
            total_shares += shares
            
        # Handle Cash
        elif cash_share.lower() == 'cash':
            if cash_usd > 0:
                positive_cash_total_usd += cash_usd
        
        # Record transaction
        results.append({
            'Date': date_val,
            'Type': cash_share,
            'Shares': shares,
            'Cash (USD)': cash_usd,
            'Market Value (USD)': market_value_usd,
            'Exchange Rate': rate if rate else 'N/A',
            'Market Value (INR)': round(market_value_usd * rate, 2) if rate else 0
        })
    
    # Calculate positive cash total in INR using closing date rate
    positive_cash_inr = 0
    if closing_date:
        rate, _ = get_exchange_rate(closing_date, rates_df)
        positive_cash_inr = positive_cash_total_usd * rate if rate else 0
    
    # Calculate peak value for the year
    peak_value_inr = 0
    peak_date = None
    peak_price_usd = 0
    peak_shares = 0
    
    if YFINANCE_AVAILABLE and closing_date and opening_date:
        try:
            year = closing_date.year
            start_date = f"{year}-01-01"
            end_date = f"{year}-12-31"
            
            print(f"\nFetching {ticker_symbol} stock prices for {year}...")
            stock = yf.Ticker(ticker_symbol)
            hist = stock.history(start=start_date, end=end_date)
            
            if not hist.empty:
                # Build a timeline of share balances
                # Start with opening shares
                share_timeline = []
                current_shares = 0
                
                # Process transactions chronologically to track share balance
                for idx, row in df.iterrows():
                    try:
                        date_val = pd.to_datetime(row['Date'], dayfirst=True)
                    except:
                        continue
                    
                    cash_share = str(row['Cash/Share']).strip() if pd.notna(row['Cash/Share']) else ''
                    shares = pd.to_numeric(row['No. of Shares'], errors='coerce') or 0
                    
                    if cash_share.lower() == 'opening':
                        current_shares = shares
                        share_timeline.append((date_val.normalize(), current_shares))
                        print(f"  Opening: {date_val.strftime('%Y-%m-%d')} - {current_shares:.2f} shares")
                    elif cash_share.lower() == 'share':
                        current_shares += shares
                        share_timeline.append((date_val.normalize(), current_shares))
                        print(f"  Added shares: {date_val.strftime('%Y-%m-%d')} - Total: {current_shares:.2f} shares")
                
                print(f"Share timeline has {len(share_timeline)} entries")
                
                # For each day in stock history, find the share balance and calculate value
                max_value_usd = 0
                max_value_date = None
                max_value_price = 0
                max_value_shares = 0
                
                for stock_date, stock_row in hist.iterrows():
                    stock_price = stock_row['Close']
                    # Convert to timezone-naive datetime for comparison
                    if hasattr(stock_date, 'tz_localize'):
                        stock_date_clean = stock_date.tz_localize(None) if stock_date.tz else stock_date
                    else:
                        stock_date_clean = pd.Timestamp(stock_date)
                    
                    # Find shares held on this date
                    shares_on_date = 0
                    for timeline_date, timeline_shares in share_timeline:
                        # Ensure both are comparable
                        timeline_ts = pd.Timestamp(timeline_date)
                        stock_ts = pd.Timestamp(stock_date_clean)
                        
                        if timeline_ts.date() <= stock_ts.date():
                            shares_on_date = timeline_shares
                        else:
                            break
                    
                    # Calculate portfolio value on this date
                    value_on_date = stock_price * shares_on_date
                    
                    if value_on_date > max_value_usd:
                        max_value_usd = value_on_date
                        max_value_date = stock_date_clean
                        max_value_price = stock_price
                        max_value_shares = shares_on_date
                
                if max_value_date:
                    peak_date = max_value_date
                    peak_price_usd = max_value_price
                    peak_shares = max_value_shares
                    
                    # Convert to INR using peak date's exchange rate
                    peak_rate, _ = get_exchange_rate(peak_date, rates_df)
                    peak_value_inr = max_value_usd * peak_rate if peak_rate else 0
                    
                    print(f"Peak date: {peak_date.strftime('%Y-%m-%d')}")
                    print(f"Peak price: ${peak_price_usd:.2f}")
                    print(f"Shares held on peak date: {peak_shares:.2f}")
                    print(f"Peak value: ${max_value_usd:,.2f} = ₹{peak_value_inr:,.2f}")
                else:
                    print("Warning: No peak value found - share timeline may be empty")

        except Exception as e:
            print(f"Warning: Could not fetch stock data for {ticker_symbol}: {e}")

    
    # Create summary
    summary = {
        'Opening Date': opening_date.strftime('%Y-%m-%d') if opening_date else 'N/A',
        'Opening Value (INR)': round(opening_value_inr, 2),
        'Closing Date': closing_date.strftime('%Y-%m-%d') if closing_date else 'N/A',
        'Closing Value (INR)': round(closing_value_inr, 2),
        'Total Shares': total_shares,
        'Positive Cash Total (USD)': round(positive_cash_total_usd, 2),
        'Positive Cash Total (INR)': round(positive_cash_inr, 2),
        'Peak Date': peak_date.strftime('%Y-%m-%d') if peak_date else 'N/A',
        'Peak Price (USD)': round(peak_price_usd, 2) if peak_price_usd else 0,
        'Peak Value (INR)': round(peak_value_inr, 2)
    }
    
    return pd.DataFrame(results), summary



def main(excel_input=None, excel_output=None, ticker_symbol='MU'):
    # Use provided paths or defaults
    input_path = Path(excel_input) if excel_input else EXCEL_PATH
    output_path = Path(excel_output) if excel_output else OUTPUT_PATH
    
    if not input_path.exists():
        print(f"Error: Excel file not found at {input_path}")
        return
    if not CSV_PATH.exists():
        print(f"Error: CSV file not found at {CSV_PATH}")
        return

    print(f"Input file: {input_path}")
    print(f"Output file: {output_path}")
    
    # 1. Load Rates
    print(f"\nLoading rates from {CSV_PATH.name}...")
    try:
        rates_df = pd.read_csv(CSV_PATH)
        rates_df['DATE'] = pd.to_datetime(rates_df['DATE'], format='%Y-%m-%d %H:%M')
        rates_df['TT BUY'] = pd.to_numeric(rates_df['TT BUY'], errors='coerce').fillna(0.0)
    except Exception as e:
        print(f"Error loading rates CSV: {e}")
        return

    # 2. Load Excel sheets
    print(f"Loading excel from {input_path.name}...")
    try:
        xl_file = pd.ExcelFile(input_path)
        print(f"Available sheets: {xl_file.sheet_names}")
        
        # Process Dividend sheet
        df_dividend = pd.read_excel(input_path, sheet_name='Dividend_FY')
        print("\n=== Processing Dividend_FY ===")
        dividend_results = process_dividend_sheet(df_dividend, rates_df)
        
        # Process ESPP-Buy sheet
        df_espp_buy = pd.read_excel(input_path, sheet_name='ESPP-Buy')
        print("\n=== Processing ESPP-Buy ===")
        espp_buy_results = process_espp_buy_sheet(df_espp_buy, rates_df)
        
        # Process ESPP-Sale sheet
        df_espp_sale = pd.read_excel(input_path, sheet_name='ESPP-Sale')
        print("\n=== Processing ESPP-Sale ===")
        espp_sale_results = process_espp_sale_sheet(df_espp_sale, rates_df)
        
        # Process RSU-Vest sheet
        df_rsu_vest = pd.read_excel(input_path, sheet_name='RSU-Vest')
        print("\n=== Processing RSU-Vest ===")
        rsu_vest_results = process_espp_buy_sheet(df_rsu_vest, rates_df)  # Same logic as ESPP-Buy
        
        # Process RSU-Sale sheet
        df_rsu_sale = pd.read_excel(input_path, sheet_name='RSU-Sale')
        print("\n=== Processing RSU-Sale ===")
        rsu_sale_results = process_espp_sale_sheet(df_rsu_sale, rates_df)  # Same logic as ESPP-Sale
        
        # Match sales to purchases (ESPP)
        espp_matched = pd.DataFrame()
        if not espp_sale_results.empty:
            print("\n=== Matching ESPP Sales to Purchases (FIFO) ===")
            espp_matched = match_sales_to_purchases(espp_buy_results, espp_sale_results)
        else:
            print("\n=== No ESPP sales to match ===")
        
        # Match sales to purchases (RSU)
        rsu_matched = pd.DataFrame()
        if not rsu_sale_results.empty:
            print("\n=== Matching RSU Sales to Vests (FIFO) ===")
            rsu_matched = match_sales_to_purchases(rsu_vest_results, rsu_sale_results)
        else:
            print("\n=== No RSU sales to match ===")
        
        # Process Schedule FA (ESPP-Assets) if sheet exists
        schedule_fa_results = pd.DataFrame()
        schedule_fa_summary = None
        try:
            df_schedule_fa = pd.read_excel(input_path, sheet_name='ESPP-Assets')
            print("\n=== Processing Schedule FA (ESPP-Assets) ===")
            schedule_fa_results, schedule_fa_summary = process_schedule_fa(df_schedule_fa, rates_df, ticker_symbol)
        except ValueError:
            print("\n=== ESPP-Assets sheet not found, skipping Schedule FA ===")
        except Exception as e:
            print(f"\n=== Error processing Schedule FA: {e} ===")

        
    except Exception as e:
        print(f"Error processing Excel: {e}")
        import traceback
        traceback.print_exc()
        return

    # 3. Display summaries
    print("\n--- Dividend Summary (First 5 Rows) ---")
    if not dividend_results.empty:
        print(dividend_results[['Transaction Date', 'SBI TT Buy Rate', 'Value (INR)', 'Tax (INR)']].head().to_string())
    
    print("\n--- ESPP Buy Summary (First 5 Rows) ---")
    if not espp_buy_results.empty:
        print(espp_buy_results[['Transaction Date', 'No. of Shares', 'Total Purchase Price (INR)']].head().to_string())
    
    print("\n--- ESPP Sale Summary (First 5 Rows) ---")
    if not espp_sale_results.empty:
        print(espp_sale_results[['Transaction Date', 'No. of Shares', 'Total Sale Price (INR)']].head().to_string())
    else:
        print("No ESPP sales data")
    
    print("\n--- RSU Vest Summary (First 5 Rows) ---")
    if not rsu_vest_results.empty:
        print(rsu_vest_results[['Transaction Date', 'No. of Shares', 'Total Purchase Price (INR)']].head().to_string())
    
    print("\n--- RSU Sale Summary (First 5 Rows) ---")
    if not rsu_sale_results.empty:
        print(rsu_sale_results[['Transaction Date', 'No. of Shares', 'Total Sale Price (INR)']].head().to_string())
    else:
        print("No RSU sales data")
    
    # Display ESPP matched transactions
    if not espp_matched.empty:
        print("\n--- ESPP Matched Transactions (First 5 Rows) ---")
        print(espp_matched[['Sale Date', 'Purchase Date', 'Holding Period (Months)', 'Gain Type', 'Shares Sold', 'Capital Gain/Loss (INR)']].head().to_string())
    
    # Display RSU matched transactions
    if not rsu_matched.empty:
        print("\n--- RSU Matched Transactions (First 5 Rows) ---")
        print(rsu_matched[['Sale Date', 'Purchase Date', 'Holding Period (Months)', 'Gain Type', 'Shares Sold', 'Capital Gain/Loss (INR)']].head().to_string())
    
    # Combined capital gains summary
    if not espp_matched.empty or not rsu_matched.empty:
        espp_total_gain = espp_matched['Capital Gain/Loss (INR)'].sum() if not espp_matched.empty else 0
        espp_total_ltcg = espp_matched['LTCG (INR)'].sum() if not espp_matched.empty else 0
        espp_total_stcg = espp_matched['STCG (INR)'].sum() if not espp_matched.empty else 0
        
        rsu_total_gain = rsu_matched['Capital Gain/Loss (INR)'].sum() if not rsu_matched.empty else 0
        rsu_total_ltcg = rsu_matched['LTCG (INR)'].sum() if not rsu_matched.empty else 0
        rsu_total_stcg = rsu_matched['STCG (INR)'].sum() if not rsu_matched.empty else 0
        
        total_gain = espp_total_gain + rsu_total_gain
        total_ltcg = espp_total_ltcg + rsu_total_ltcg
        total_stcg = espp_total_stcg + rsu_total_stcg
        
        print(f"\n{'='*60}")
        print(f"CAPITAL GAINS SUMMARY")
        print(f"{'='*60}")
        if not espp_matched.empty:
            print(f"ESPP:")
            print(f"  Total Gain/Loss:        ₹{espp_total_gain:>15,.2f}")
            print(f"    - LTCG:               ₹{espp_total_ltcg:>15,.2f}")
            print(f"    - STCG:               ₹{espp_total_stcg:>15,.2f}")
        if not rsu_matched.empty:
            print(f"RSU:")
            print(f"  Total Gain/Loss:        ₹{rsu_total_gain:>15,.2f}")
            print(f"    - LTCG:               ₹{rsu_total_ltcg:>15,.2f}")
            print(f"    - STCG:               ₹{rsu_total_stcg:>15,.2f}")
        print(f"{'-'*60}")
        print(f"COMBINED TOTAL:")
        print(f"  Total Capital Gain/Loss:₹{total_gain:>15,.2f}")
        print(f"    - Long Term (LTCG):   ₹{total_ltcg:>15,.2f}")
        print(f"    - Short Term (STCG):  ₹{total_stcg:>15,.2f}")
        print(f"{'='*60}")
    
    # Schedule FA Summary
    if schedule_fa_summary:
        print(f"\n{'='*60}")
        print(f"SCHEDULE FA SUMMARY (Foreign Assets)")
        print(f"{'='*60}")
        print(f"Opening Date:             {schedule_fa_summary['Opening Date']}")
        print(f"Opening Value (INR):      ₹{schedule_fa_summary['Opening Value (INR)']:>15,.2f}")
        print(f"Closing Date:             {schedule_fa_summary['Closing Date']}")
        print(f"Closing Value (INR):      ₹{schedule_fa_summary['Closing Value (INR)']:>15,.2f}")
        print(f"Total Shares:             {schedule_fa_summary['Total Shares']:>18,.2f}")
        print(f"{'-'*60}")
        print(f"Positive Cash (USD):      ${schedule_fa_summary['Positive Cash Total (USD)']:>15,.2f}")
        print(f"Positive Cash (INR):      ₹{schedule_fa_summary['Positive Cash Total (INR)']:>15,.2f}")
        print(f"{'-'*60}")
        if schedule_fa_summary['Peak Date'] != 'N/A':
            print(f"Peak Date:                {schedule_fa_summary['Peak Date']}")
            print(f"Peak Price (USD):         ${schedule_fa_summary['Peak Price (USD)']:>15,.2f}")
            print(f"Peak Value (INR):         ₹{schedule_fa_summary['Peak Value (INR)']:>15,.2f}")
        else:
            print(f"Peak Value:               Not calculated (yfinance not available)")
        print(f"{'='*60}")

    
    # 4. Save to Excel with multiple sheets
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            dividend_results.to_excel(writer, sheet_name='Dividend_Calculated', index=False)
            espp_buy_results.to_excel(writer, sheet_name='ESPP_Buy_Calculated', index=False)
            espp_sale_results.to_excel(writer, sheet_name='ESPP_Sale_Calculated', index=False)
            rsu_vest_results.to_excel(writer, sheet_name='RSU_Vest_Calculated', index=False)
            rsu_sale_results.to_excel(writer, sheet_name='RSU_Sale_Calculated', index=False)
            
            if not espp_matched.empty:
                espp_matched.to_excel(writer, sheet_name='ESPP_Matched_Transactions', index=False)
            if not rsu_matched.empty:
                rsu_matched.to_excel(writer, sheet_name='RSU_Matched_Transactions', index=False)
            
            # Add Schedule FA if available
            if not schedule_fa_results.empty:
                schedule_fa_results.to_excel(writer, sheet_name='Schedule_FA_Details', index=False)
            if schedule_fa_summary:
                # Create a summary DataFrame
                summary_df = pd.DataFrame([schedule_fa_summary])
                summary_df.to_excel(writer, sheet_name='Schedule_FA_Summary', index=False)
        
        print(f"\n✓ Successfully saved calculated data to: {output_path}")
        print(f"  - Dividend_Calculated: {len(dividend_results)} rows")
        print(f"  - ESPP_Buy_Calculated: {len(espp_buy_results)} rows")
        print(f"  - ESPP_Sale_Calculated: {len(espp_sale_results)} rows")
        print(f"  - RSU_Vest_Calculated: {len(rsu_vest_results)} rows")
        print(f"  - RSU_Sale_Calculated: {len(rsu_sale_results)} rows")
        if not espp_matched.empty:
            print(f"  - ESPP_Matched_Transactions: {len(espp_matched)} rows")
        if not rsu_matched.empty:
            print(f"  - RSU_Matched_Transactions: {len(rsu_matched)} rows")
        if not schedule_fa_results.empty:
            print(f"  - Schedule_FA_Details: {len(schedule_fa_results)} rows")
        if schedule_fa_summary:
            print(f"  - Schedule_FA_Summary: 1 row")
    except Exception as e:
        print(f"Error saving output file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Calculate ITR values for foreign assets with USD to INR conversion',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Use default files
  python calculate_itr_values.py
  
  # Specify custom input file
  python calculate_itr_values.py -i my_assets.xlsx
  
  # Specify both input and output files
  python calculate_itr_values.py -i my_assets.xlsx -o my_results.xlsx
        """
    )
    
    parser.add_argument(
        '-i', '--input',
        type=str,
        help='Input Excel file path (default: ITR_Foreign_Assets.xlsx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Output Excel file path (default: ITR_Calculated_Values.xlsx)'
    )
    
    parser.add_argument(
        '-t', '--ticker',
        type=str,
        default='MU',
        help='Stock ticker symbol for peak value calculation (default: MU)'
    )
    
    args = parser.parse_args()
    
    main(excel_input=args.input, excel_output=args.output, ticker_symbol=args.ticker)
