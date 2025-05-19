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
            pass
    return path

def conv(file, op=None):
    try:
        if op is None:
            op = os.path.splitext(file)[0] + '.csv'
        
        df = pd.read_excel(file)
        df.to_csv(op, index=False)
        return op
    except Exception:
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
    
    return ldg_f, mf_f

def get_all_codes():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    ldg_dir = os.path.join(desktop, 'Ledger')
    
    if not os.path.exists(ldg_dir):
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
    
    return list(codes)

def parse_float(value):
    try:
        if pd.isna(value) or value == '':
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = ''.join(c for c in str(value) 
                        if c.isdigit() or c in {'.', '-'})
        return float(cleaned) if cleaned else 0.0
    except Exception:
        return 0.0

def calc_xirr(values, dates, guess=0.1):
    if len(values) < 2:
        return None
        
    if not (any(v > 0 for v in values) and any(v < 0 for v in values)):
        return None
    
    try:
        dates = [pd.to_datetime(d) for d in dates]
        days = [(d - dates[0]).days for d in dates]
        
        if len(set(days)) < len(days):
            day_value_map = {}
            
            for i, day in enumerate(days):
                if day in day_value_map:
                    day_value_map[day] += values[i]
                else:
                    day_value_map[day] = values[i]
            
            days = sorted(day_value_map.keys())
            values = [day_value_map[day] for day in days]
        
        def xnpv(rate):
            result = 0
            for i in range(len(values)):
                result += values[i] / (1 + rate) ** (days[i] / 365.0)
            return result
            
        def xnpv_derivative(rate):
            result = 0
            for i in range(len(values)):
                d = days[i] / 365.0
                result -= d * values[i] / ((1 + rate) ** (d + 1))
            return result
        
        try:
            left, right = -0.999, 9.0  
            
            if xnpv(left) * xnpv(right) > 0:
                result = optimize.newton(xnpv, guess, xnpv_derivative, 
                                         maxiter=1000, tol=1.48e-08, disp=False)
            else:
                result = optimize.brentq(xnpv, left, right, 
                                         xtol=1.48e-08, rtol=1.48e-08, maxiter=1000)
            
            return result
        except Exception:
            guesses = [0.1, 0.05, 0.01, 0.2, 0.3, -0.1, -0.2, 0.5, -0.5]
            
            for guess in guesses:
                try:
                    result = optimize.newton(xnpv, guess, xnpv_derivative,
                                             maxiter=1000, tol=1.48e-08, disp=False)
                    if abs(xnpv(result)) < 1e-6:  
                        return result
                except Exception:
                    continue
                    
        def secant_method(f, x0, x1, tol=1.48e-8, max_iter=1000):
            f_x0 = f(x0)
            f_x1 = f(x1)
            
            for i in range(max_iter):
                if abs(f_x1) < tol:
                    return x1
                if abs(f_x1 - f_x0) < 1e-10:  
                    x0, x1 = x1, x1 * 1.1 + 0.01
                    f_x0, f_x1 = f_x1, f(x1)
                    continue
                    
                x_new = x1 - f_x1 * (x1 - x0) / (f_x1 - f_x0)
                x0, x1 = x1, x_new
                f_x0, f_x1 = f_x1, f(x1)
                
            return None  
            
        result = secant_method(xnpv, 0.1, 0.2)
        if result is not None:
            return result
                    
        return None
    except Exception:
        return None

def process_ldg(ldg, start_date, today_date):
    ldg_trans = []
    credit_col = None
    debit_col = None
    
    for col in ldg.columns:
        col_low = str(col).lower()
        if 'credit' in col_low:
            credit_col = col
        elif 'debit' in col_low:
            debit_col = col
    
    if credit_col is None or debit_col is None:
        bal_col = None
        for col in ldg.columns:
            if 'balance' in str(col).lower():
                bal_col = col
                break
        
        if bal_col is None:
            for col in ldg.columns:
                if pd.api.types.is_numeric_dtype(ldg[col]):
                    bal_col = col
                    break
        
        if bal_col is None:
            first_bal = 0
        else:
            try:
                first_bal = ldg[bal_col].iloc[0]
                if pd.isna(first_bal):
                    first_bal = 0
            except Exception:
                first_bal = 0
    else:
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
                break
    
    if eff_date_col is None:
        for col in ldg.columns:
            if 'date' in str(col).lower():
                eff_date_col = col
                break
    
    if vch_type_col is not None and eff_date_col is not None:
        try:
            vch_vals = ldg[vch_type_col].unique()
            pay_vals = [v for v in vch_vals if 'pay' in str(v).lower()]
            
            if pay_vals:
                pay_df = ldg[ldg[vch_type_col].isin(pay_vals)]
            else:
                pay_df = ldg.copy()
            
            if eff_date_col in pay_df.columns:
                try:
                    pay_df[eff_date_col] = pd.to_datetime(pay_df[eff_date_col], errors='coerce')
                    pay_df = pay_df.sort_values(by=eff_date_col, ascending=True)
                except Exception:
                    pass
            
            for _, row in pay_df.iterrows():
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
                    
                    ldg_trans.append([eff_date, value, vch_type])
                except Exception:
                    pass
        except Exception:
            pass
    
    return ldg_trans, bal_col

def process_mf(mft, today_date):
    mf_trans = []
    
    if 'Transaction Date' in mft.columns:
        try:
            for _, row in mft.iterrows():
                try:
                    if pd.isna(row['Transaction Date']) or 'Total' in str(row['Transaction Date']):
                        continue
                    
                    tr_date = pd.to_datetime(row['Transaction Date'], errors='coerce').date()
                    tr_value = parse_float(row['Transaction Value'])
                    tr_type = str(row['Transaction Type']).lower() if 'Transaction Type' in row else ''
                    
                    if 'buy' in tr_type:
                        mf_trans.append([tr_date, -abs(tr_value), 'MF BUY'])
                    elif 'sell' in tr_type:
                        mf_trans.append([tr_date, abs(tr_value), 'MF SELL'])
                        
                except Exception:
                    pass
        except Exception:
            pass
    elif 'Unnamed: 3' in mft.columns and 'Unnamed: 6' in mft.columns:
        try:
            date_col = 'Unnamed: 0'
                
            for _, row in mft.iterrows():
                try:
                    if pd.isna(row['Unnamed: 3']):
                        continue
                    
                    tr_type = str(row['Unnamed: 3']).lower()
                    
                    if 'buy' not in tr_type and 'sell' not in tr_type:
                        continue
                    
                    tr_value = parse_float(row['Unnamed: 6'])
                    
                    if date_col:
                        try:
                            tr_date = pd.to_datetime(row[date_col], errors='coerce').date()
                            if pd.isna(tr_date):
                                tr_date = today_date
                        except:
                            tr_date = today_date
                    else:
                        tr_date = today_date
                    
                    if 'buy' in tr_type:
                        mf_trans.append([tr_date, -abs(tr_value), 'MF BUY'])
                    elif 'sell' in tr_type:
                        mf_trans.append([tr_date, abs(tr_value), 'MF SELL'])
                        
                except Exception:
                    pass
                    
        except Exception:
            pass
    else:
        col_map = {
            'type': ['transaction type', 'tr type', 'type'],
            'date': ['date', 'tr date', 'effective date'],
            'value': ['amount', 'value', 'nav']
        }
        
        cols = {}
        for col_type, keywords in col_map.items():
            for col in mft.columns:
                if any(kw in str(col).lower() for kw in keywords):
                    cols[col_type] = col
                    break

        if not all(cols.get(k) for k in ['type', 'date', 'value']):
            return mf_trans

        try:
            mft = mft.rename(columns={
                cols['type']: 'TransactionType',
                cols['date']: 'Date',
                cols['value']: 'Value'
            })
            
            for _, row in mft.iterrows():
                try:
                    tr_date = pd.to_datetime(row['Date'], errors='coerce').date()
                    tr_value = parse_float(row['Value'])
                    tr_type = str(row['TransactionType']).lower()
                    
                    if 'buy' in tr_type:
                        mf_trans.append([tr_date, -abs(tr_value), 'MF BUY'])
                    elif 'sell' in tr_type:
                        mf_trans.append([tr_date, abs(tr_value), 'MF SELL'])
                        
                except Exception:
                    pass

        except Exception:
            pass
    
    return mf_trans

def get_curr_val(code, ldg):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    cons_path = os.path.join(desktop, "Holding", "Consolidated_Holdings.xlsx")
    
    holdings_val = None
    
    if os.path.exists(cons_path):
        try:
            cons_df = pd.read_excel(cons_path)
            
            code_col = None
            for col in cons_df.columns:
                if 'client' in str(col).lower() or 'code' in str(col).lower():
                    code_col = col
                    break
            
            val_col = None
            for col in cons_df.columns:
                if 'portfolio value' in str(col).lower() or 'port val' in str(col).lower():
                    val_col = col
                    break
            
            if code_col and val_col:
                code_rows = cons_df[cons_df[code_col].astype(str).str.strip() == str(code).strip()]
                if not code_rows.empty:
                    holdings_val = parse_float(code_rows.iloc[0][val_col])
        except Exception:
            pass
    
    bal_col = None
    ledger_bal = 0
    
    for col in ldg.columns:
        if 'balance' in str(col).lower():
            bal_col = col
            break
    
    if bal_col:
        try:
            first_idx = ldg.index[0]
            ledger_bal = parse_float(ldg.loc[first_idx, bal_col])
        except Exception:
            pass
    
    final_val = None
    if holdings_val is not None:
        final_val = holdings_val + ledger_bal
    
    return final_val

def run_xirr(ldg, mft, init_val, curr_val, out_dir=None, code=None, start_date=None):
    if start_date is None:
        today = datetime.now()
        start_date = datetime(today.year - 1, today.month, today.day).date()
    
    initial_date = start_date
    today_date = datetime.now().date()
    
    res_df = pd.DataFrame(columns=['Date', 'Fund', 'Remarks'])
    
    res_df.loc[0] = [initial_date, -abs(init_val), 'Initial Value']
    
    ldg_trans, _ = process_ldg(ldg, start_date, today_date)
    
    mf_trans = process_mf(mft, today_date)
    
    for date, value, remark in ldg_trans:
        new_idx = len(res_df)
        res_df.loc[new_idx] = [date, value, remark]
    
    for date, value, remark in mf_trans:
        new_idx = len(res_df)
        res_df.loc[new_idx] = [date, value, remark]
    
    new_idx = len(res_df)
    res_df.loc[new_idx] = [today_date, abs(curr_val), 'Current Value']
    
    res_df_xl = res_df.copy()
    res_df_xl['Date'] = pd.to_datetime(res_df_xl['Date'], errors='coerce')
    res_df_xl['Fund'] = pd.to_numeric(res_df_xl['Fund'], errors='coerce')
    
    res_df_xl = res_df_xl.dropna(subset=['Date', 'Fund'])
    
    res_df_xl_sorted = res_df_xl.sort_values(by='Date', ascending=True)
    
    has_positive = (res_df_xl_sorted['Fund'] > 0).any()
    has_negative = (res_df_xl_sorted['Fund'] < 0).any()
    
    python_xirr = None
    try:
        values = res_df_xl_sorted['Fund'].tolist()
        dates = res_df_xl_sorted['Date'].tolist()
        python_xirr = calc_xirr(values, dates)
    except Exception:
        pass
    
    try:
        try:
            import xlsxwriter
        except ImportError:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            reports_dir = os.path.join(desktop, "xirr_reports")
            mk_dir(reports_dir)
            
            if code:
                out_file = os.path.join(reports_dir, f"{code}_xirr_report.csv")
            else:
                out_file = os.path.join(reports_dir, "xirr_report.csv")
                
            res_df_xl['Date'] = res_df_xl['Date'].dt.strftime('%d/%m/%Y')
            if python_xirr is not None:
                xirr_row = pd.DataFrame([['XIRR Calculation', 'XIRR Value', f'{python_xirr:.2%}']], 
                                        columns=res_df_xl.columns)
                result_df = pd.concat([res_df_xl, xirr_row])
                result_df.to_csv(out_file, index=False)
            else:
                res_df_xl.to_csv(out_file, index=False)
                
            return out_file
            
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        reports_dir = os.path.join(desktop, "xirr_reports")
        mk_dir(reports_dir)  
        
        if code:
            out_file = os.path.join(reports_dir, f"{code}_xirr_report.xlsx")
        else:
            out_file = os.path.join(reports_dir, "xirr_report.xlsx")
        
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
                xirr_formula = f'=XIRR(B2:B{second_last_row},A2:A{second_last_row})'
                
                worksheet.write(row_count + 1, 0, 'XIRR Calculation')
                worksheet.write(row_count + 1, 1, 'XIRR Value')
                
                worksheet.write_formula(row_count + 1, 2, xirr_formula, percent_format)
            
        return out_file
    except Exception:
        return "Error saving results"

def proc_dir(mf_dir, init_val, start_date=None):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    ledger_dir = os.path.join(desktop, 'Ledger')
    results = []
    
    for mf_file in glob.glob(os.path.join(mf_dir, "*.xlsx")) + glob.glob(os.path.join(mf_dir, "*.xls")):
        csv_file = conv(mf_file)
        if csv_file:
            pass
    
    for mf_file in glob.glob(os.path.join(mf_dir, "*.csv")):
        try:
            code = os.path.basename(mf_file).split('_')[0]
            ledger_files = glob.glob(os.path.join(ledger_dir, f"{code}_Ledger*"))
            
            if not ledger_files:
                continue
                
            ldg_df = pd.read_csv(ledger_files[0])
            mf_df = pd.read_csv(mf_file)
            
            curr_val = get_curr_val(code, ldg_df)
            
            if curr_val is None:
                curr_val = init_val
            
            out_file = run_xirr(ldg_df, mf_df, init_val, curr_val, 
                           code=code, start_date=start_date)
            results.append(out_file)
        except Exception:
            pass
    
    return results

def proc(code=None, init_val=100000, start_date=None, input_dir=None):
    if input_dir: 
        return proc_dir(input_dir, init_val, start_date)
    
    if code:
        ldg_f, mf_f = get_files(code)
        
        if not all([ldg_f, mf_f]):
            return None
        
        mf_csv = conv(mf_f)
        
        ldg_df = pd.read_csv(ldg_f)
        mf_df = pd.read_csv(mf_csv)
        
        curr_val = get_curr_val(code, ldg_df)
        
        if curr_val is None:
            curr_val = init_val
        
        out_file = run_xirr(ldg_df, mf_df, init_val, curr_val, code=code, start_date=start_date)
        
        return out_file
    else:
        codes = get_all_codes()
        out_files = []
        
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        reports_dir = os.path.join(desktop, "xirr_reports")
        mk_dir(reports_dir)  
        
        for code in codes:
            ldg_f, mf_f = get_files(code)
            
            if not all([ldg_f, mf_f]):
                continue
            
            mf_csv = conv(mf_f)
            
            try:
                ldg_df = pd.read_csv(ldg_f)
                mf_df = pd.read_csv(mf_csv)
                
                curr_val = get_curr_val(code, ldg_df)
                
                if curr_val is None:
                    curr_val = init_val
                
                out_file = run_xirr(ldg_df, mf_df, init_val, curr_val, code=code, start_date=start_date)
                out_files.append(out_file)
            except Exception:
                pass
        
        return out_files
    
if __name__ == "__main__":    
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
    
    code = sys.argv[1] if len(sys.argv) > 1 else None
    init_val = float(sys.argv[2]) if len(sys.argv) > 2 else 100000
    
    today = datetime.now()
    default_start_date = datetime(today.year - 1, today.month, today.day).date()
    start_date = default_start_date
    
    if len(sys.argv) > 3:
        try:
            start_date = datetime.strptime(sys.argv[3], '%d/%m/%Y').date()
        except:
            print(f"Invalid date format. Use DD/MM/YYYY format. Using one year ago date: {default_start_date.strftime('%d/%m/%Y')}")
    
    if code:
        print(f"Processing data for client: {code}")
        out_file = proc(code, init_val, start_date)
        print(f"Output saved to: {out_file}")
    else:
        print("Processing data for all clients")
        out_files = proc(init_val=init_val, start_date=start_date)
        print(f"Outputs saved to: {out_files}")