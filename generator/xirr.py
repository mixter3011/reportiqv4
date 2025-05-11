import os
import sys
import glob
import pandas as pd
from datetime import datetime
import numpy as np
import warnings
from scipy import optimize
warnings.filterwarnings('ignore')

def mk_dir(path):
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except Exception as e:
            print(f"{path}:{e}")
    return path

def conv(file, op = None):
    try:
        if op is None:
            op = os.path.splitext(file)[0] + '.csv'
        
        df = pd.read_excel(file)
        df.to_csv(op, index=False)
        return op
    except Exception as e:
        print(f"Excel conversion error: {e}")
        return None
    
def get_files(code):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    
    ldg_dir = os.path.join(desktop, 'Ledger')
    mf_dir = os.path.join(desktop, 'MF Transactions')
    
    ldg_pat_lower = os.path.join(ldg_dir, f"*{code}_ledger*.csv")
    ldg_pat_upper = os.path.join(ldg_dir, f"*{code}_Ledger*.csv")
    mf_pat = os.path.join(mf_dir, f"*{code}_MFTrans*.xlsx")
    
    ldg_fs = glob.glob(ldg_pat_lower) + glob.glob(ldg_pat_upper)
    mf_fs = glob.glob(mf_pat)
    
    ldg_f = sorted(ldg_fs)[-1] if ldg_fs else None
    mf_f = sorted(mf_fs)[-1] if mf_fs else None
    
    print(f"Files for client {code}:")
    print(f"  Ledger: {ldg_f}")
    print(f"  MF Transaction: {mf_f}")
    
    if not ldg_fs:
        print(f"No ledger files found in: {ldg_dir}")
        print(f"Searched patterns: {ldg_pat_lower} and {ldg_pat_upper}")
    
    if not mf_fs:
        print(f"No MF transaction files found in: {mf_dir}")
        print(f"Searched pattern: {mf_pat}")
    
    return ldg_f, mf_f

def get_all():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    ldg_dir = os.path.join(desktop, 'Ledger')
    
    if not os.path.exists(ldg_dir):
        print(f"Ledger directory not found: {ldg_dir}")
        return []
    
    ldg_files_lower = glob.glob(os.path.join(ldg_dir, "*_ledger*.csv"))
    ldg_files_upper = glob.glob(os.path.join(ldg_dir, "*_Ledger*.csv"))
    ldg_files = ldg_files_lower + ldg_files_upper
    
    codes = set()
    for f in ldg_files:
        fname = os.path.basename(f)
        parts = fname.split('_')
        if len(parts) > 0:
            codes.add(parts[0])
    
    print(f"Found {len(codes)} client codes: {codes}")
    return list(codes)

def parse_float(value):
    if isinstance(value, (int, float)):
        return float(value)
    elif isinstance(value, str):
        return float(value.replace(',', ''))
    else:
        return 0.0

def calculate_xirr(values, dates, guess=0.1):
    try:
        dates = [pd.to_datetime(d) for d in dates]
        days = [(d - dates[0]).days for d in dates]
        
        def xnpv(rate):
            return sum(values[i] / (1 + rate) ** (days[i] / 365.0) for i in range(len(values)))
        
        result = optimize.newton(lambda r: xnpv(r), guess)
        return result
    except Exception as e:
        print(f"Python XIRR calculation error: {e}")
        return None

def process_ledger_data(ldg, start_date, today_date):
    print("Processing Ledger data...")
    ledger_transactions = []
    credit_col = None
    debit_col = None
    
    for col in ldg.columns:
        col_low = str(col).lower()
        if 'credit' in col_low:
            credit_col = col
        elif 'debit' in col_low:
            debit_col = col
    
    if credit_col is None or debit_col is None:
        print(f"Warning: Credit or Debit columns not found. Credit: {credit_col}, Debit: {debit_col}")
        print("Will use balance column as fallback.")
        
        bal_col = None
        for col in ldg.columns:
            if 'balance' in str(col).lower():
                bal_col = col
                break
        
        if bal_col is None:
            print("Balance column not found. Using first numeric column.")
            for col in ldg.columns:
                if pd.api.types.is_numeric_dtype(ldg[col]):
                    bal_col = col
                    print(f"Using {col} as balance column.")
                    break
        
        if bal_col is None:
            print("No suitable balance column found. Using 0.")
            first_bal = 0
        else:
            try:
                first_bal = ldg[bal_col].iloc[0]
                if pd.isna(first_bal):
                    print("First balance is NaN, using 0")
                    first_bal = 0
            except Exception as e:
                print(f"Error getting first balance: {e}")
                first_bal = 0
    else:
        print(f"Found Credit column: {credit_col} and Debit column: {debit_col}")
        
        bal_col = None
        for col in ldg.columns:
            if 'balance' in str(col).lower():
                bal_col = col
                break
        
        if bal_col is None:
            first_bal = 0
        else:
            try:
                first_bal = ldg[bal_col].iloc[0]
                if pd.isna(first_bal):
                    first_bal = 0
            except:
                first_bal = 0
    
    print(f"First balance from ledger: {first_bal}")
    
    vch_type_col = None
    eff_date_col = None
    
    for col in ldg.columns:
        col_low = str(col).lower()
        if 'voucher' in col_low and 'type' in col_low:
            vch_type_col = col
        elif 'effective' in col_low and 'date' in col_low:
            eff_date_col = col
    
    if vch_type_col is None:
        for col in ldg.columns:
            if 'type' in str(col).lower():
                vch_type_col = col
                print(f"Using {col} as voucher type column")
                break
    
    if eff_date_col is None:
        for col in ldg.columns:
            if 'date' in str(col).lower():
                eff_date_col = col
                print(f"Using {col} as effective date column")
                break
    
    if vch_type_col is not None and eff_date_col is not None:
        try:
            vch_vals = ldg[vch_type_col].unique()
            print(f"Unique voucher types: {vch_vals}")
            
            pay_vals = [v for v in vch_vals if 'pay' in str(v).lower()]
            
            if pay_vals:
                pay_df = ldg[ldg[vch_type_col].isin(pay_vals)]
            else:
                pay_df = ldg.copy()
            
            if eff_date_col in pay_df.columns:
                try:
                    pay_df[eff_date_col] = pd.to_datetime(pay_df[eff_date_col], errors='coerce')
                    pay_df = pay_df.sort_values(by=eff_date_col, ascending=True)
                    print(f"Sorted transactions by date (ascending order - oldest first)")
                except Exception as e:
                    print(f"Error sorting by date: {e}")
            
            print(f"Found {len(pay_df)} transactions to process from ledger")
            
            for idx, row in pay_df.iterrows():
                try:
                    vch_type = str(row[vch_type_col])
                    
                    eff_date = row[eff_date_col]
                    if isinstance(eff_date, str):
                        try:
                            eff_date = pd.to_datetime(eff_date).date()
                        except:
                            eff_date = today_date
                    
                    is_payin = 'payin' in vch_type.lower() or 'pay in' in vch_type.lower()
                    is_payout = 'payout' in vch_type.lower() or 'pay out' in vch_type.lower()
                    
                    if is_payin:
                        vch_type = "Ledger Buy"
                        
                        if credit_col and not pd.isna(row[credit_col]):
                            value = parse_float(row[credit_col])
                            value = -abs(value)
                        else:
                            value = parse_float(row[bal_col]) if bal_col else 0
                            value = -abs(value)  
                    elif is_payout:
                        vch_type = "Ledger Sell"
                        
                        if debit_col and not pd.isna(row[debit_col]):
                            value = parse_float(row[debit_col])
                            value = abs(value)
                        else:
                            value = parse_float(row[bal_col]) if bal_col else 0
                            value = abs(value)  
                    else:
                        if bal_col:
                            value = parse_float(row[bal_col])
                        else:
                            value = 0
                    
                    ledger_transactions.append([eff_date, value, vch_type])
                except Exception as e:
                    print(f"Error processing ledger row {idx}: {e}")
        except Exception as e:
            print(f"Error processing ledger data: {e}")
    else:
        print("Required columns not found in ledger. Skipping ledger processing.")
    
    return ledger_transactions

def process_mf_data(mft, today_date):
    print("Processing MF Transaction data...")
    mf_transactions = []
    tr_type_col = None
    tr_date_col = None
    tr_val_col = None
    
    for col in mft.columns:
        if 'unnamed' in str(col).lower():
            try:
                for idx, val in enumerate(mft[col]):
                    if isinstance(val, str) and 'transaction type' in val.lower():
                        tr_type_col = col
                        
                        try:
                            tr_date_col = mft.columns[0]  
                            tr_val_col = mft.columns[6]   
                            print(f"Found MF columns: Type={tr_type_col}, Date={tr_date_col}, Value={tr_val_col}")
                            break
                        except Exception as e:
                            print(f"Error setting MF columns: {e}")
                if tr_type_col is not None:
                    break
            except Exception as e:
                print(f"Error checking column {col}: {e}")
    
    if tr_type_col is None:
        for col in mft.columns:
            col_low = str(col).lower()
            if 'type' in col_low and ('trans' in col_low or 'action' in col_low):
                tr_type_col = col
            elif 'date' in col_low:
                tr_date_col = col
            elif 'value' in col_low or 'amount' in col_low:
                tr_val_col = col
    
    if tr_type_col is not None and tr_date_col is not None and tr_val_col is not None:
        try:
            found_header = False
            header_idx = -1
            
            if 'unnamed' in str(tr_type_col).lower():
                for idx, val in enumerate(mft[tr_type_col]):
                    if isinstance(val, str) and 'transaction type' in val.lower():
                        found_header = True
                        header_idx = idx
                        continue
                    
                    if found_header and idx > header_idx and not pd.isna(val):
                        if isinstance(val, str) and ('buy' in val.lower() or 'sell' in val.lower()):
                            try:
                                tr_date = mft.iloc[idx][tr_date_col]
                                tr_value = mft.iloc[idx][tr_val_col]
                                
                                if not pd.isna(tr_date) and not pd.isna(tr_value):
                                    try:
                                        tr_date = pd.to_datetime(tr_date).date()
                                    except:
                                        tr_date = today_date
                                        
                                    try:
                                        tr_value = parse_float(tr_value)
                                    except:
                                        tr_value = 0
                                        
                                    is_buy = 'buy' in val.lower()
                                    rem = 'MF BUY' if is_buy else 'MF SELL'
                                    
                                    if is_buy:
                                        tr_value = -abs(tr_value)
                                    else:
                                        tr_value = abs(tr_value)
                                    
                                    mf_transactions.append([tr_date, tr_value, rem])
                                    print(f"Added MF transaction: {tr_date} - {tr_value} - {rem}")
                            except Exception as e:
                                print(f"Error processing MF row {idx}: {e}")
            else:
                tr_types = mft[tr_type_col].unique()
                print(f"Unique transaction types: {tr_types}")
                
                bs_vals = [v for v in tr_types if 'buy' in str(v).lower() or 'sell' in str(v).lower()]
                
                if bs_vals:
                    mf_trans = mft[mft[tr_type_col].isin(bs_vals)]
                else:
                    mf_trans = mft.copy()
                
                if tr_date_col in mf_trans.columns:
                    try:
                        mf_trans[tr_date_col] = pd.to_datetime(mf_trans[tr_date_col], errors='coerce')
                        mf_trans = mf_trans.sort_values(by=tr_date_col, ascending=True)
                        print(f"Sorted MF transactions by date (ascending order - oldest first)")
                    except Exception as e:
                        print(f"Error sorting MF transactions by date: {e}")
                
                print(f"Found {len(mf_trans)} transactions to process from MF")
                
                for idx, row in mf_trans.iterrows():
                    try:
                        tr_type = str(row[tr_type_col])
                        
                        tr_date = row[tr_date_col]
                        if isinstance(tr_date, str):
                            try:
                                tr_date = pd.to_datetime(tr_date).date()
                            except:
                                tr_date = today_date
                        
                        tr_val = parse_float(row[tr_val_col])
                        
                        is_buy = 'buy' in tr_type.lower()
                        rem = 'MF BUY' if is_buy else 'MF SELL'
                        
                        if is_buy:
                            tr_val = -abs(tr_val)
                        else:
                            tr_val = abs(tr_val)
                        
                        mf_transactions.append([tr_date, tr_val, rem])
                    except Exception as e:
                        print(f"Error processing MF row {idx}: {e}")
        except Exception as e:
            print(f"Error processing MF transactions: {e}")
    else:
        print("Required columns not found in MF data. Skipping MF processing.")
    
    return mf_transactions
        
def main(ldg, mft, init_val, curr_val, out_dir=None, cl_code=None, start_date=None):
    print("Ledger columns:", ldg.columns.tolist())
    print("MF Transaction columns:", mft.columns.tolist())
    
    if start_date is None:
        today = datetime.now()
        start_date = datetime(today.year - 1, today.month, today.day).date()
        print(f"Using default start date (one year ago): {start_date.strftime('%d/%m/%Y')}")
    
    initial_date = start_date
    today_date = datetime.now().date()
    
    print(f"Initial portfolio date: {initial_date.strftime('%d/%m/%Y')}")
    print(f"Current date: {today_date.strftime('%d/%m/%Y')}")
    
    print(f"Initial portfolio value: {init_val}")
    print(f"Current portfolio value: {curr_val}")
    
    
    res_df = pd.DataFrame(columns=['Date', 'Fund', 'Remarks'])
    
    
    res_df.loc[0] = [initial_date, -abs(init_val), 'Initial Value']
    
    
    ledger_trans = process_ledger_data(ldg, start_date, today_date)
    
    
    mf_trans = process_mf_data(mft, today_date)
    
    
    for idx, (date, value, remark) in enumerate(ledger_trans):
        new_idx = len(res_df)
        res_df.loc[new_idx] = [date, value, remark]
    
    
    for idx, (date, value, remark) in enumerate(mf_trans):
        new_idx = len(res_df)
        res_df.loc[new_idx] = [date, value, remark]
    
    
    new_idx = len(res_df)
    res_df.loc[new_idx] = [today_date, abs(curr_val), 'Current Value']
    
    print("\nGenerated transactions table:")
    print(res_df)
    
    
    res_df_xl = res_df.copy()
    res_df_xl['Date'] = pd.to_datetime(res_df_xl['Date'], errors='coerce')
    res_df_xl['Fund'] = pd.to_numeric(res_df_xl['Fund'], errors='coerce')
    
    
    res_df_xl_sorted = res_df_xl.sort_values(by='Date', ascending=True)
    print("\nSorted transaction table for XIRR calculation (oldest first):")
    print(res_df_xl_sorted)
    
    
    res_df_xl_sorted = res_df_xl_sorted.dropna(subset=['Date', 'Fund'])
    
    
    has_positive = (res_df_xl_sorted['Fund'] > 0).any()
    has_negative = (res_df_xl_sorted['Fund'] < 0).any()
    
    if not (has_positive and has_negative):
        print("\nWARNING: XIRR calculation requires both positive and negative cash flows")
        print(f"Has positive cash flows: {has_positive}, Has negative cash flows: {has_negative}")
    else:
        print("\nData has both positive and negative cash flows, XIRR calculation should work")
    
    
    python_xirr = None
    try:
        values = res_df_xl_sorted['Fund'].tolist()
        dates = res_df_xl_sorted['Date'].tolist()
        python_xirr = calculate_xirr(values, dates)
        print(f"\nPython calculated XIRR: {python_xirr:.2%}" if python_xirr is not None else "\nCould not calculate Python XIRR")
    except Exception as e:
        print(f"\nError calculating Python XIRR: {e}")
    
    try:
        try:
            import xlsxwriter
        except ImportError:
            print("\nWarning: xlsxwriter module not found. You need to install it:")
            print("pip install xlsxwriter")
            print("Saving to CSV instead...")
            
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            xirr_reports_dir = os.path.join(desktop, "xirr_reports")
            mk_dir(xirr_reports_dir)
            
            if cl_code:
                out_file = os.path.join(xirr_reports_dir, f"{cl_code}_xirr_report.csv")
            else:
                out_file = os.path.join(xirr_reports_dir, "xirr_report.csv")
                
            res_df_xl['Date'] = res_df_xl['Date'].dt.strftime('%d/%m/%Y')
            if python_xirr is not None:
                xirr_row = pd.DataFrame([['XIRR Calculation', 'XIRR Value', f'{python_xirr:.2%}']], 
                                        columns=res_df_xl.columns)
                result_df = pd.concat([res_df_xl, xirr_row])
                result_df.to_csv(out_file, index=False)
            else:
                res_df_xl.to_csv(out_file, index=False)
                
            print(f"Analysis saved to {out_file}")
            return out_file
            
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        xirr_reports_dir = os.path.join(desktop, "xirr_reports")
        mk_dir(xirr_reports_dir)  
        
        if cl_code:
            out_file = os.path.join(xirr_reports_dir, f"{cl_code}_xirr_report.xlsx")
        else:
            out_file = os.path.join(xirr_reports_dir, "xirr_report.xlsx")
        
        with pd.ExcelWriter(out_file, engine='xlsxwriter') as writer:
            
            res_df_xl['Date'] = res_df_xl['Date'].dt.strftime('%d/%m/%Y')
            res_df_xl.to_excel(writer, sheet_name='Portfolio Analysis', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Portfolio Analysis']
            
            row_count = len(res_df_xl)
            
            
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            worksheet.set_column('A:A', 20, date_format)  
            worksheet.set_column('B:B', 25)  
            worksheet.set_column('C:C', 30)  
            
            percent_format = workbook.add_format({'num_format': '0.00%'})
            
            second_last_row = row_count + 1  
            if second_last_row >= 2:  
                
                xirr_formula = f'=IFERROR(XIRR(B2:B{second_last_row},A2:A{second_last_row}), "Could not calculate")'
                
                worksheet.write(row_count + 1, 0, 'XIRR Calculation')
                worksheet.write(row_count + 1, 1, 'XIRR Value')
                
                worksheet.write_formula(row_count + 1, 2, xirr_formula, percent_format)
            
        print(f"Analysis saved to {out_file}")
        return out_file
    except Exception as e:
        print(f"Error saving file: {e}")
        return "Error saving results"
    
def proc(cl_code=None, init_val=100000, curr_val=None, start_date=None):
    if curr_val is None:
        curr_val = init_val 
        
    if cl_code:
        ldg_f, mf_f = get_files(cl_code)
        
        if not all([ldg_f, mf_f]):
            print(f"Error: Missing files for client {cl_code}")
            return None
        
        mf_csv = conv(mf_f)
        
        ldg_df = pd.read_csv(ldg_f)
        mf_df = pd.read_csv(mf_csv)
        
        out_file = main(ldg_df, mf_df, init_val, curr_val, cl_code=cl_code, start_date=start_date)
        
        return out_file
    else:
        cl_codes = get_all()
        out_files = []
        
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        xirr_reports_dir = os.path.join(desktop, "xirr reports")
        mk_dir(xirr_reports_dir)  
        
        for code in cl_codes:
            print(f"\nProcessing client: {code}")
            ldg_f, mf_f = get_files(code)
            
            if not all([ldg_f, mf_f]):
                print(f"Skipping client {code} due to missing files")
                continue
            
            mf_csv = conv(mf_f)
            
            try:
                ldg_df = pd.read_csv(ldg_f)
                mf_df = pd.read_csv(mf_csv)
                
                out_file = main(ldg_df, mf_df, init_val, curr_val, cl_code=code, start_date=start_date)
                out_files.append(out_file)
            except Exception as e:
                print(f"Error processing client {code}: {e}")
        
        return out_files
    
if __name__ == "__main__":
    print("XIRR Calculation Tool")
    print("--------------------")
    print("Checking required modules...")
    
    modules_to_check = ['pandas', 'numpy', 'scipy', 'xlsxwriter']
    missing_modules = []
    
    for module in modules_to_check:
        try:
            __import__(module)
            print(f"✓ {module} is installed")
        except ImportError:
            print(f"✗ {module} is NOT installed")
            missing_modules.append(module)
    
    if missing_modules:
        print("\nSome required modules are missing. Please install them using pip:")
        for module in missing_modules:
            print(f"pip install {module}")
        print("\nContinuing anyway, but the script might fail...")
    
    cl_code = sys.argv[1] if len(sys.argv) > 1 else None
    init_val = float(sys.argv[2]) if len(sys.argv) > 2 else 100000
    curr_val = float(sys.argv[3]) if len(sys.argv) > 3 else init_val
    
    today = datetime.now()
    default_start_date = datetime(today.year - 1, today.month, today.day).date()
    start_date = default_start_date
    
    if len(sys.argv) > 4:
        try:
            start_date = datetime.strptime(sys.argv[4], '%d/%m/%Y').date()
        except:
            print(f"Invalid date format. Use DD/MM/YYYY format. Using one year ago date: {default_start_date.strftime('%d/%m/%Y')}")
    
    if cl_code:
        print(f"Processing data for client: {cl_code}")
        out_file = proc(cl_code, init_val, curr_val, start_date)
        print(f"Output saved to: {out_file}")
    else:
        print("Processing data for all clients")
        out_files = proc(init_val=init_val, curr_val=curr_val, start_date=start_date)
        print(f"Outputs saved to: {out_files}")