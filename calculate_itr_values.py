import pandas as pd
import datetime
from pathlib import Path
import os
import argparse
import sys
import urllib.request
import io
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False
    print("Warning: yfinance not installed. Peak value calculation will be skipped.")

# Paths and URLs
BASE_DIR = Path(r"c:\Users\anant\OneDrive\Documents\itr_calculations")
EXCEL_PATH = BASE_DIR / "ITR_Foreign_Assets.xlsx"
OUTPUT_PATH = BASE_DIR / "ITR_Calculated_Values.xlsx"
SBI_RATES_URL = "https://raw.githubusercontent.com/sahilgupta/sbi-fx-ratekeeper/main/csv_files/SBI_REFERENCE_RATES_USD.csv"

def load_sbi_rates():
    """
    Load SBI exchange rates from GitHub URL.
    Returns a DataFrame with DATE and TT BUY columns.
    """
    try:
        print(f"Fetching SBI rates from GitHub...")
        with urllib.request.urlopen(SBI_RATES_URL) as response:
            csv_data = response.read().decode('utf-8')
        
        rates_df = pd.read_csv(io.StringIO(csv_data))
        rates_df['DATE'] = pd.to_datetime(rates_df['DATE'], format='%Y-%m-%d %H:%M')
        rates_df['TT BUY'] = pd.to_numeric(rates_df['TT BUY'], errors='coerce').fillna(0.0)
        print(f"Successfully loaded {len(rates_df)} exchange rates")
        return rates_df
    except Exception as e:
        print(f"Error loading SBI rates from GitHub: {e}")
        raise

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

def process_cash_sheet(df, rates_df, ticker_symbol='MU'):
    """
    Process Cash sheet to calculate combined ESPP and RSU portfolio values.
    Calculates closing value and peak value for the combined portfolio.
    """
    results = []
    opening_date = None
    closing_date = None
    opening_espp = 0
    opening_rsu = 0
    closing_espp = 0
    closing_rsu = 0
    
    # Track cash transactions for timeline
    cash_transactions = []
    
    # Get column names (handle potential unnamed columns)
    cols = df.columns.tolist()
    date_col = cols[0]  # First column is Date
    type_col = cols[1]  # Second column is type (Opening/Closing/Cash)
    espp_col = cols[2]  # Third column is ESPP value
    rsu_col = cols[3]   # Fourth column is RSU value
    
    # Process each row
    for idx, row in df.iterrows():
        try:
            date_val = pd.to_datetime(row[date_col], dayfirst=True)
        except:
            continue
        
        row_type = str(row[type_col]).strip() if pd.notna(row[type_col]) else ''
        espp_val = pd.to_numeric(row[espp_col], errors='coerce')
        espp_val = 0 if pd.isna(espp_val) else espp_val
        rsu_val = pd.to_numeric(row[rsu_col], errors='coerce')
        rsu_val = 0 if pd.isna(rsu_val) else rsu_val
        
        # Get exchange rate for this date
        rate, found_date = get_exchange_rate(date_val, rates_df)
        
        # Handle Opening
        if row_type.lower() == 'opening':
            opening_date = date_val
            opening_espp = espp_val
            opening_rsu = rsu_val
            
        # Handle Closing
        elif row_type.lower() == 'closing':
            closing_date = date_val
            closing_espp = espp_val
            closing_rsu = rsu_val
            
        # Handle Cash transactions
        elif row_type.lower() == 'cash':
            cash_transactions.append({
                'date': date_val,
                'espp': espp_val,
                'rsu': rsu_val
            })
        
        # Record transaction
        combined_val = espp_val + rsu_val
        results.append({
            'Date': date_val,
            'Type': row_type,
            'ESPP Value (USD)': espp_val,
            'RSU Value (USD)': rsu_val,
            'Combined Value (USD)': combined_val,
            'Exchange Rate': rate if rate else 'N/A',
            'Combined Value (INR)': round(combined_val * rate, 2) if rate else 0
        })
    
    # Calculate closing value in INR
    closing_combined_usd = closing_espp + closing_rsu
    closing_rate, _ = get_exchange_rate(closing_date, rates_df) if closing_date else (None, None)
    closing_combined_inr = closing_combined_usd * closing_rate if closing_rate else 0
    
    # Calculate opening value in INR
    opening_combined_usd = opening_espp + opening_rsu
    opening_rate, _ = get_exchange_rate(opening_date, rates_df) if opening_date else (None, None)
    opening_combined_inr = opening_combined_usd * opening_rate if opening_rate else 0
    
    
    # Calculate peak value - find maximum combined cash value in INR
    peak_value_inr = 0
    peak_date = None
    peak_combined_usd = 0
    
    if closing_date and opening_date:
        # Build cash balance timeline with INR conversions
        cash_timeline = []
        current_espp_cash = opening_espp
        current_rsu_cash = opening_rsu
        
        print(f"\nCalculating peak value for combined portfolio...")
        
        # Add opening
        opening_combined = current_espp_cash + current_rsu_cash
        opening_rate, _ = get_exchange_rate(opening_date, rates_df)
        opening_inr = opening_combined * opening_rate if opening_rate else 0
        cash_timeline.append({
            'date': opening_date,
            'espp': current_espp_cash,
            'rsu': current_rsu_cash,
            'combined_usd': opening_combined,
            'rate': opening_rate,
            'combined_inr': opening_inr
        })
        print(f"  Opening: {opening_date.strftime('%Y-%m-%d')} - Combined: ${opening_combined:.2f} @ {opening_rate:.2f} = ₹{opening_inr:,.2f}")
        
        # Add cash transactions chronologically
        for tx in sorted(cash_transactions, key=lambda x: x['date']):
            current_espp_cash += tx['espp']
            current_rsu_cash += tx['rsu']
            combined_usd = current_espp_cash + current_rsu_cash
            rate, _ = get_exchange_rate(tx['date'], rates_df)
            combined_inr = combined_usd * rate if rate else 0
            cash_timeline.append({
                'date': tx['date'],
                'espp': current_espp_cash,
                'rsu': current_rsu_cash,
                'combined_usd': combined_usd,
                'rate': rate,
                'combined_inr': combined_inr
            })
            print(f"  {tx['date'].strftime('%Y-%m-%d')}: Combined: ${combined_usd:.2f} @ {rate:.2f} = ₹{combined_inr:,.2f}")
        
        # Add closing
        closing_combined = closing_espp + closing_rsu
        closing_rate_val, _ = get_exchange_rate(closing_date, rates_df)
        closing_inr = closing_combined * closing_rate_val if closing_rate_val else 0
        cash_timeline.append({
            'date': closing_date,
            'espp': closing_espp,
            'rsu': closing_rsu,
            'combined_usd': closing_combined,
            'rate': closing_rate_val,
            'combined_inr': closing_inr
        })
        print(f"  Closing: {closing_date.strftime('%Y-%m-%d')} - Combined: ${closing_combined:.2f} @ {closing_rate_val:.2f} = ₹{closing_inr:,.2f}")
        
        print(f"Cash timeline has {len(cash_timeline)} entries")
        
        # Find maximum INR value
        max_entry = max(cash_timeline, key=lambda x: x['combined_inr'])
        peak_date = max_entry['date']
        peak_combined_usd = max_entry['combined_usd']
        peak_value_inr = max_entry['combined_inr']
        
        print(f"Peak date: {peak_date.strftime('%Y-%m-%d')}")
        print(f"Peak combined value: ${peak_combined_usd:.2f} = ₹{peak_value_inr:,.2f}")


    
    # Create summary
    summary = {
        'Opening Date': opening_date.strftime('%Y-%m-%d') if opening_date else 'N/A',
        'Opening ESPP (USD)': round(opening_espp, 2),
        'Opening RSU (USD)': round(opening_rsu, 2),
        'Opening Combined (USD)': round(opening_combined_usd, 2),
        'Opening Combined (INR)': round(opening_combined_inr, 2),
        'Closing Date': closing_date.strftime('%Y-%m-%d') if closing_date else 'N/A',
        'Closing ESPP (USD)': round(closing_espp, 2),
        'Closing RSU (USD)': round(closing_rsu, 2),
        'Closing Combined (USD)': round(closing_combined_usd, 2),
        'Closing Combined (INR)': round(closing_combined_inr, 2),
        'Peak Date': peak_date.strftime('%Y-%m-%d') if peak_date else 'N/A',
        'Peak Combined Value (USD)': round(peak_combined_usd, 2),
        'Peak Combined Value (INR)': round(peak_value_inr, 2)
    }
    
    return pd.DataFrame(results), summary



def main(excel_input=None, excel_output=None, ticker_symbol='MU'):
    # Use provided paths or defaults
    input_path = Path(excel_input) if excel_input else EXCEL_PATH
    output_path = Path(excel_output) if excel_output else OUTPUT_PATH
    
    if not input_path.exists():
        print(f"Error: Excel file not found at {input_path}")
        return

    print(f"Input file: {input_path}")
    print(f"Output file: {output_path}")
    
    # 1. Load Rates from GitHub
    print(f"\nLoading SBI exchange rates...")
    try:
        rates_df = load_sbi_rates()
    except Exception as e:
        print(f"Error loading rates: {e}")
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
        schedule_fa_espp_results = pd.DataFrame()
        schedule_fa_espp_summary = None
        try:
            df_schedule_fa_espp = pd.read_excel(input_path, sheet_name='ESPP-Assets')
            print("\n=== Processing Schedule FA (ESPP-Assets) ===")
            schedule_fa_espp_results, schedule_fa_espp_summary = process_schedule_fa(df_schedule_fa_espp, rates_df, ticker_symbol)
        except ValueError:
            print("\n=== ESPP-Assets sheet not found, skipping ESPP Schedule FA ===")
        except Exception as e:
            print(f"\n=== Error processing ESPP Schedule FA: {e} ===")

        # Process Schedule FA (RSU-Assets) if sheet exists
        schedule_fa_rsu_results = pd.DataFrame()
        schedule_fa_rsu_summary = None
        try:
            df_schedule_fa_rsu = pd.read_excel(input_path, sheet_name='RSU-Assets')
            print("\n=== Processing Schedule FA (RSU-Assets) ===")
            schedule_fa_rsu_results, schedule_fa_rsu_summary = process_schedule_fa(df_schedule_fa_rsu, rates_df, ticker_symbol)
        except ValueError:
            print("\n=== RSU-Assets sheet not found, skipping RSU Schedule FA ===")
        except Exception as e:
            print(f"\n=== Error processing RSU Schedule FA: {e} ===")

        # Process Cash sheet if it exists
        cash_results = pd.DataFrame()
        cash_summary = None
        try:
            df_cash = pd.read_excel(input_path, sheet_name='Cash')
            print("\n=== Processing Cash (Combined Portfolio) ===")
            cash_results, cash_summary = process_cash_sheet(df_cash, rates_df, ticker_symbol)
        except ValueError:
            print("\n=== Cash sheet not found, skipping combined portfolio calculation ===")
        except Exception as e:
            print(f"\n=== Error processing Cash sheet: {e} ===")

        
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
    
    # Schedule FA Summary - ESPP
    if schedule_fa_espp_summary:
        print(f"\n{'='*60}")
        print(f"SCHEDULE FA SUMMARY - ESPP (Foreign Assets)")
        print(f"{'='*60}")
        print(f"Opening Date:             {schedule_fa_espp_summary['Opening Date']}")
        print(f"Opening Value (INR):      ₹{schedule_fa_espp_summary['Opening Value (INR)']:>15,.2f}")
        print(f"Closing Date:             {schedule_fa_espp_summary['Closing Date']}")
        print(f"Closing Value (INR):      ₹{schedule_fa_espp_summary['Closing Value (INR)']:>15,.2f}")
        print(f"Total Shares:             {schedule_fa_espp_summary['Total Shares']:>18,.2f}")
        print(f"{'-'*60}")
        print(f"Positive Cash (USD):      ${schedule_fa_espp_summary['Positive Cash Total (USD)']:>15,.2f}")
        print(f"Positive Cash (INR):      ₹{schedule_fa_espp_summary['Positive Cash Total (INR)']:>15,.2f}")
        print(f"{'-'*60}")
        if schedule_fa_espp_summary['Peak Date'] != 'N/A':
            print(f"Peak Date:                {schedule_fa_espp_summary['Peak Date']}")
            print(f"Peak Price (USD):         ${schedule_fa_espp_summary['Peak Price (USD)']:>15,.2f}")
            print(f"Peak Value (INR):         ₹{schedule_fa_espp_summary['Peak Value (INR)']:>15,.2f}")
        else:
            print(f"Peak Value:               Not calculated (yfinance not available)")
        print(f"{'='*60}")
    
    # Schedule FA Summary - RSU
    if schedule_fa_rsu_summary:
        print(f"\n{'='*60}")
        print(f"SCHEDULE FA SUMMARY - RSU (Foreign Assets)")
        print(f"{'='*60}")
        print(f"Opening Date:             {schedule_fa_rsu_summary['Opening Date']}")
        print(f"Opening Value (INR):      ₹{schedule_fa_rsu_summary['Opening Value (INR)']:>15,.2f}")
        print(f"Closing Date:             {schedule_fa_rsu_summary['Closing Date']}")
        print(f"Closing Value (INR):      ₹{schedule_fa_rsu_summary['Closing Value (INR)']:>15,.2f}")
        print(f"Total Shares:             {schedule_fa_rsu_summary['Total Shares']:>18,.2f}")
        print(f"{'-'*60}")
        print(f"Positive Cash (USD):      ${schedule_fa_rsu_summary['Positive Cash Total (USD)']:>15,.2f}")
        print(f"Positive Cash (INR):      ₹{schedule_fa_rsu_summary['Positive Cash Total (INR)']:>15,.2f}")
        print(f"{'-'*60}")
        if schedule_fa_rsu_summary['Peak Date'] != 'N/A':
            print(f"Peak Date:                {schedule_fa_rsu_summary['Peak Date']}")
            print(f"Peak Price (USD):         ${schedule_fa_rsu_summary['Peak Price (USD)']:>15,.2f}")
            print(f"Peak Value (INR):         ₹{schedule_fa_rsu_summary['Peak Value (INR)']:>15,.2f}")
        else:
            print(f"Peak Value:               Not calculated (yfinance not available)")
        print(f"{'='*60}")
    
    # Cash Summary - Combined Portfolio
    if cash_summary:
        print(f"\n{'='*60}")
        print(f"COMBINED PORTFOLIO SUMMARY (ESPP + RSU)")
        print(f"{'='*60}")
        print(f"Opening Date:             {cash_summary['Opening Date']}")
        print(f"  ESPP Value (USD):       ${cash_summary['Opening ESPP (USD)']:>15,.2f}")
        print(f"  RSU Value (USD):        ${cash_summary['Opening RSU (USD)']:>15,.2f}")
        print(f"  Combined (USD):         ${cash_summary['Opening Combined (USD)']:>15,.2f}")
        print(f"  Combined (INR):         ₹{cash_summary['Opening Combined (INR)']:>15,.2f}")
        print(f"{'-'*60}")
        print(f"Closing Date:             {cash_summary['Closing Date']}")
        print(f"  ESPP Value (USD):       ${cash_summary['Closing ESPP (USD)']:>15,.2f}")
        print(f"  RSU Value (USD):        ${cash_summary['Closing RSU (USD)']:>15,.2f}")
        print(f"  Combined (USD):         ${cash_summary['Closing Combined (USD)']:>15,.2f}")
        print(f"  Combined (INR):         ₹{cash_summary['Closing Combined (INR)']:>15,.2f}")
        print(f"{'-'*60}")
        if cash_summary['Peak Date'] != 'N/A':
            print(f"Peak Date:                {cash_summary['Peak Date']}")
            print(f"  Peak Combined (USD):    ${cash_summary['Peak Combined Value (USD)']:>15,.2f}")
            print(f"  Peak Combined (INR):    ₹{cash_summary['Peak Combined Value (INR)']:>15,.2f}")
        else:
            print(f"Peak Value:               Not calculated")
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
            
            # Add Schedule FA - ESPP if available
            if not schedule_fa_espp_results.empty:
                schedule_fa_espp_results.to_excel(writer, sheet_name='Schedule_FA_ESPP_Details', index=False)
            if schedule_fa_espp_summary:
                summary_df = pd.DataFrame([schedule_fa_espp_summary])
                summary_df.to_excel(writer, sheet_name='Schedule_FA_ESPP_Summary', index=False)
            
            # Add Schedule FA - RSU if available
            if not schedule_fa_rsu_results.empty:
                schedule_fa_rsu_results.to_excel(writer, sheet_name='Schedule_FA_RSU_Details', index=False)
            if schedule_fa_rsu_summary:
                summary_df = pd.DataFrame([schedule_fa_rsu_summary])
                summary_df.to_excel(writer, sheet_name='Schedule_FA_RSU_Summary', index=False)
            
            # Add Cash (Combined Portfolio) if available
            if not cash_results.empty:
                cash_results.to_excel(writer, sheet_name='Cash_Details', index=False)
            if cash_summary:
                summary_df = pd.DataFrame([cash_summary])
                summary_df.to_excel(writer, sheet_name='Cash_Summary', index=False)
        
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
        if not schedule_fa_espp_results.empty:
            print(f"  - Schedule_FA_ESPP_Details: {len(schedule_fa_espp_results)} rows")
        if schedule_fa_espp_summary:
            print(f"  - Schedule_FA_ESPP_Summary: 1 row")
        if not schedule_fa_rsu_results.empty:
            print(f"  - Schedule_FA_RSU_Details: {len(schedule_fa_rsu_results)} rows")
        if schedule_fa_rsu_summary:
            print(f"  - Schedule_FA_RSU_Summary: 1 row")
        if not cash_results.empty:
            print(f"  - Cash_Details: {len(cash_results)} rows")
        if cash_summary:
            print(f"  - Cash_Summary: 1 row")
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
