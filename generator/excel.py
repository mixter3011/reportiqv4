import pandas as pd
from openpyxl import Workbook
from utils.format import format_num
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side

def excel_generator(df):
    client_info_row = df[df['Unnamed: 0'] == 'Client Equity Code/UCID/Name'].index[0]
    client_info = str(df.iloc[client_info_row, 1]).strip()
    parts = client_info.split('/')
    client_code = parts[0].strip()  
    client_name = parts[-1].strip()  

    equity_row = df[df['Unnamed: 0'] == 'Equity:-'].index[0]
    mf_row = df[df['Unnamed: 0'] == 'Mutual Fund:-'].index[0]
    fno_row = df[df['Unnamed: 0'] == 'FnO:-'].index[0]
    bond_row = df[df['Unnamed: 0'] == 'Bond:-'].index[0]

    equity_header = df.iloc[equity_row + 1].tolist()
    equity_end = mf_row - 4  # Skip empty rows
    equity_data = df.iloc[equity_row + 2:equity_end].copy()

    mf_header = df.iloc[mf_row + 1].tolist()
    mf_end = fno_row - 4
    mf_data = df.iloc[mf_row + 2:mf_end].copy()

    bond_header = df.iloc[bond_row + 1].tolist()
    bond_end = len(df)
    for i in range(bond_row + 2, len(df)):
        if i >= len(df) or pd.isna(df.iloc[i, 0]) or df.iloc[i, 0] == '':
            bond_end = i
            break
    bond_data = df.iloc[bond_row + 2:bond_end].copy()

    equity_data = equity_data[equity_data['Unnamed: 0'] != 'Total:']
    mf_data = mf_data[mf_data['Unnamed: 0'] != 'Total:']
    bond_data = bond_data[bond_data['Unnamed: 0'] != 'Total:']

    wb = Workbook()
    ws = wb.active

    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    ws.merge_cells('A1:J1')
    ws.merge_cells('A2:J2')

    market_value_cells = {}

    ws.cell(row=1, column=1, value=client_name)
    ws.cell(row=1, column=1).fill = light_blue_fill
    ws.cell(row=1, column=1).font = Font(bold=True)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    ws.cell(row=2, column=1, value=client_code)
    ws.cell(row=2, column=1).fill = light_blue_fill
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

    direct_equity = equity_data[~equity_data['Unnamed: 0'].str.contains('ETF', na=False)]
    direct_equity = direct_equity[~direct_equity['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]

    equity_cols_to_keep = [0, 1, 2, 4, 5, 10]  
    equity_rename = {2: "Buy Price", 10: "P&L"}

    direct_equity_total = ['Total:']
    for col_idx, col in enumerate(equity_cols_to_keep[1:], 1):  
        try:
            values = direct_equity[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                direct_equity_total.append(total)
            else:
                direct_equity_total.append('')
        except (ValueError, TypeError):
            direct_equity_total.append('')

    direct_equity_market_value = direct_equity_total[4] if len(direct_equity_total) > 4 else 0

    etf_equity = equity_data[equity_data['Unnamed: 0'].str.contains('ETF', na=False)]
    etf_equity = etf_equity[~etf_equity['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]
    etf_equity = etf_equity[~etf_equity['Unnamed: 0'].str.contains('Nippon India ETF Nifty 8-13 yr G-Sec LongTerm Gilt', na=False)]

    etf_equity_total = ['Total:']
    for col_idx, col in enumerate(equity_cols_to_keep[1:], 1):  
        try:
            values = etf_equity[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                etf_equity_total.append(total)
            else:
                etf_equity_total.append('')
        except (ValueError, TypeError):
            etf_equity_total.append('')

    etf_equity_market_value = etf_equity_total[4] if len(etf_equity_total) > 4 else 0

    debt_etf = equity_data[equity_data['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]

    debt_etf_total = ['Total:']
    for col_idx, col in enumerate(equity_cols_to_keep[1:], 1):  
        try:
            values = debt_etf[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                debt_etf_total.append(total)
            else:
                debt_etf_total.append('')
        except (ValueError, TypeError):
            debt_etf_total.append('')

    debt_etf_market_value = debt_etf_total[4] if len(debt_etf_total) > 4 else 0

    gilt_etf = equity_data[equity_data['Unnamed: 0'].str.contains('Nippon India ETF Nifty 8-13 yr G-Sec LongTerm Gilt', na=False)]

    mf_cols_to_keep = [1, 2, 3, 5, 6, 12]  
    mf_rename = {3: "Buy Price", 12: "P&L"}

    equity_mf = mf_data[mf_data['Unnamed: 0'] == 'Equity']

    equity_mf_total = ['Total:']
    for col_idx, col in enumerate(mf_cols_to_keep[1:], 1):  
        try:
            values = equity_mf[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                equity_mf_total.append(total)
            else:
                equity_mf_total.append('')
        except (ValueError, TypeError, KeyError):
            equity_mf_total.append('')

    equity_mf_market_value = equity_mf_total[4] if len(equity_mf_total) > 4 else 0

    debt_mf = mf_data[mf_data['Unnamed: 0'] == 'Debt']

    if not gilt_etf.empty:
        gilt_for_debt = pd.DataFrame(columns=debt_mf.columns)

        column_mapping = {
            'Unnamed: 0': 'Unnamed: 1',  
            'Unnamed: 1': 'Unnamed: 2',  
            'Unnamed: 2': 'Unnamed: 3',  
            'Unnamed: 4': 'Unnamed: 5',  
            'Unnamed: 5': 'Unnamed: 6',  
            'Unnamed: 10': 'Unnamed: 12' 
        }
    
        for old_col, new_col in column_mapping.items():
            if old_col in gilt_etf.columns and new_col in debt_mf.columns:
                gilt_for_debt[new_col] = gilt_etf[old_col].values
    
        debt_mf = pd.concat([debt_mf, gilt_for_debt], ignore_index=True)

    debt_mf_total = ['Total:']
    for col_idx, col in enumerate(mf_cols_to_keep[1:], 1):  
        try:
            values = debt_mf[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                debt_mf_total.append(total)
            else:
                debt_mf_total.append('')
        except (ValueError, TypeError, KeyError):
            debt_mf_total.append('')

    debt_mf_market_value = debt_mf_total[4] if len(debt_mf_total) > 4 else 0

    bond_cols_to_keep = [0, 1, 2, 4, 5, 10]  
    bond_rename = {2: "Buy Price", 10: "P&L"}

    bond_total = ['Total:']
    for col_idx, col in enumerate(bond_cols_to_keep[1:], 1):  
        try:
            values = bond_data[f'Unnamed: {col}']
            values = pd.to_numeric(values, errors='coerce')
        
            if not values.isna().all():
                total = values.sum()
                bond_total.append(total)
            else:
                bond_total.append('')
        except (ValueError, TypeError):
            bond_total.append('')

    bond_market_value = bond_total[4] if len(bond_total) > 4 else 0

    total_portfolio_value = (
        direct_equity_market_value + 
        etf_equity_market_value + 
        debt_etf_market_value + 
        equity_mf_market_value + 
        debt_mf_market_value + 
        bond_market_value
    )

    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    total_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
    portfolio_value_fill = PatternFill(start_color="B0E0E6", end_color="B0E0E6", fill_type="solid") 
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

    row = 5
    ws.cell(row=row, column=1, value="EQUITY").font = Font(bold=True)

    row_portfolio = row - 1  
    ws.merge_cells(f'A{row_portfolio}')
    ws.cell(row=row_portfolio, column=1, value="Portfolio Value").font = Font(bold=True)
    ws.cell(row=row_portfolio, column=1).border = thin_border
    ws.cell(row=row_portfolio, column=1).alignment = Alignment(horizontal="left")

    ws.merge_cells(f'B{row_portfolio}')
    ws.cell(row=row_portfolio, column=2, value=format_num(total_portfolio_value))
    ws.cell(row=row_portfolio, column=2).border = thin_border
    ws.cell(row=row_portfolio, column=2).alignment = Alignment(horizontal="center")
    ws.cell(row=row_portfolio, column=2).font = Font(bold=True)

    row += 1
    ws.cell(row=row, column=1, value="Direct Equity").font = Font(bold=True)
    row += 1

    equity_col_map = {}
    for new_idx, old_idx in enumerate(equity_cols_to_keep, 1):
        equity_col_map[old_idx] = new_idx

    for old_idx in equity_cols_to_keep:
        new_idx = equity_col_map[old_idx]
        header = equity_header[old_idx]
    
        if old_idx in equity_rename:
            header = equity_rename[old_idx]
        
        ws.cell(row=row, column=new_idx, value=header)
        ws.cell(row=row, column=new_idx).fill = header_fill
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).border = thin_border
        ws.column_dimensions[get_column_letter(new_idx)].width = 15

    alloc_col = len(equity_cols_to_keep) + 1
    ws.cell(row=row, column=alloc_col, value="% Allocation")
    ws.cell(row=row, column=alloc_col).fill = header_fill
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border
    ws.column_dimensions[get_column_letter(alloc_col)].width = 12

    row += 1

    for _, data_row in direct_equity.iterrows():
        for old_idx in equity_cols_to_keep:
            new_idx = equity_col_map[old_idx]
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx, value=cell_value)
                ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx).border = thin_border
        row += 1

    for idx, value in enumerate(direct_equity_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1, value=cell_value)
            ws.cell(row=row, column=idx + 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1).border = thin_border
            ws.cell(row=row, column=idx + 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1).font = Font(bold=True)

    if total_portfolio_value > 0:
        direct_equity_percent = (direct_equity_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=alloc_col, value=f"{format_num(direct_equity_percent)}%")
        ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=alloc_col).border = thin_border
        ws.cell(row=row, column=alloc_col).fill = total_fill

    row += 1

    row += 1
    ws.cell(row=row, column=1, value="Equity ETF").font = Font(bold=True)
    row += 1

    for old_idx in equity_cols_to_keep:
        new_idx = equity_col_map[old_idx]
        header = equity_header[old_idx]
    
        if old_idx in equity_rename:
            header = equity_rename[old_idx]
        
        ws.cell(row=row, column=new_idx, value=header)
        ws.cell(row=row, column=new_idx).fill = header_fill
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).border = thin_border

    ws.cell(row=row, column=alloc_col, value="% Allocation")
    ws.cell(row=row, column=alloc_col).fill = header_fill
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border

    row += 1

    for _, data_row in etf_equity.iterrows():
        for old_idx in equity_cols_to_keep:
            new_idx = equity_col_map[old_idx]
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx, value=cell_value)
                ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx).border = thin_border
        row += 1

    for idx, value in enumerate(etf_equity_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1, value=cell_value)
            ws.cell(row=row, column=idx + 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1).border = thin_border
            ws.cell(row=row, column=idx + 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1).font = Font(bold=True)
            
    if total_portfolio_value > 0:
        etf_equity_percent = (etf_equity_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=alloc_col, value=f"{format_num(etf_equity_percent)}%")
        ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=alloc_col).border = thin_border
        ws.cell(row=row, column=alloc_col).fill = total_fill
        
    row += 1

    row += 1
    ws.cell(row=row, column=1, value="Equity Mutual Fund").font = Font(bold=True)
    row += 1

    mf_col_map = {}
    for new_idx, old_idx in enumerate(mf_cols_to_keep, 1):
        mf_col_map[old_idx] = new_idx

    for old_idx in mf_cols_to_keep:
        new_idx = mf_col_map[old_idx]
        header = mf_header[old_idx]
    
        if old_idx in mf_rename:
            header = mf_rename[old_idx]
        
        ws.cell(row=row, column=new_idx, value=header)
        ws.cell(row=row, column=new_idx).fill = header_fill
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).border = thin_border

    ws.cell(row=row, column=alloc_col, value="% Allocation")
    ws.cell(row=row, column=alloc_col).fill = header_fill
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border

    row += 1

    for _, data_row in equity_mf.iterrows():
        for old_idx in mf_cols_to_keep:
            new_idx = mf_col_map[old_idx]
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx, value=cell_value)
                ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx).border = thin_border
        row += 1

    for idx, value in enumerate(equity_mf_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1, value=cell_value)
            ws.cell(row=row, column=idx + 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1).border = thin_border
            ws.cell(row=row, column=idx + 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1).font = Font(bold=True)
            
    if total_portfolio_value > 0:
        equity_mf_percent = (equity_mf_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=alloc_col, value=f"{format_num(equity_mf_percent)}%")
        ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=alloc_col).border = thin_border
        ws.cell(row=row, column=alloc_col).fill = total_fill
        
    row += 1

    col_offset = 9  
    row = 5
    ws.cell(row=row, column=col_offset, value="DEBT").font = Font(bold=True)
    row += 1
    ws.cell(row=row, column=col_offset, value="Debt ETF").font = Font(bold=True)
    row += 1

    for old_idx in equity_cols_to_keep:
        new_idx = equity_col_map[old_idx]
        header = equity_header[old_idx]
    
        if old_idx in equity_rename:
            header = equity_rename[old_idx]
        
        ws.cell(row=row, column=new_idx + col_offset - 1, value=header)
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = header_fill
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        ws.column_dimensions[get_column_letter(new_idx + col_offset - 1)].width = 15

    right_alloc_col = len(equity_cols_to_keep) + col_offset
    ws.cell(row=row, column=right_alloc_col, value="% Allocation") 
    ws.cell(row=row, column=right_alloc_col).fill = header_fill
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border

    row += 1

    for _, data_row in debt_etf.iterrows():
        for old_idx in equity_cols_to_keep:
            new_idx = equity_col_map[old_idx]
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx + col_offset - 1, value=cell_value)
                ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        row += 1

    for idx, value in enumerate(debt_etf_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1, value=cell_value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1 + col_offset - 1).border = thin_border
            ws.cell(row=row, column=idx + 1 + col_offset - 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1 + col_offset - 1).font = Font(bold=True)
            
    if total_portfolio_value > 0:
        debt_etf_percent = (debt_etf_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=right_alloc_col, value=f"{format_num(debt_etf_percent)}%")
        ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=right_alloc_col).border = thin_border
        ws.cell(row=row, column=right_alloc_col).fill = total_fill
        
    row += 1

    row += 1
    ws.cell(row=row, column=col_offset, value="Debt Mutual Fund").font = Font(bold=True)
    row += 1

    for old_idx in mf_cols_to_keep:
        new_idx = mf_col_map[old_idx]
        header = mf_header[old_idx]
    
        if old_idx in mf_rename:
            header = mf_rename[old_idx]
        
        ws.cell(row=row, column=new_idx + col_offset - 1, value=header)
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = header_fill
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border

    ws.cell(row=row, column=right_alloc_col, value="% Allocation")
    ws.cell(row=row, column=right_alloc_col).fill = header_fill
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border

    row += 1

    for _, data_row in debt_mf.iterrows():
        for old_idx in mf_cols_to_keep:
            new_idx = mf_col_map[old_idx]
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx + col_offset - 1, value=cell_value)
                ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        row += 1

    for idx, value in enumerate(debt_mf_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1, value=cell_value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1 + col_offset - 1).border = thin_border
            ws.cell(row=row, column=idx + 1 + col_offset - 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1 + col_offset - 1).font = Font(bold=True)
            
    if total_portfolio_value > 0:
        debt_mf_percent = (debt_mf_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=len(mf_cols_to_keep) + col_offset, value=f"{format_num(debt_mf_percent)}%")
        ws.cell(row=row, column=len(mf_cols_to_keep) + col_offset).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=len(mf_cols_to_keep) + col_offset).border = thin_border
        ws.cell(row=row, column=len(mf_cols_to_keep) + col_offset).fill = total_fill
        
    row += 1

    row += 2
    ws.cell(row=row, column=col_offset, value="BONDS").font = Font(bold=True)
    row += 1

    for old_idx in bond_cols_to_keep:
        new_idx = equity_col_map[old_idx]  
        header = bond_header[old_idx]
    
        if old_idx in bond_rename:
            header = bond_rename[old_idx]
        
        ws.cell(row=row, column=new_idx + col_offset - 1, value=header)
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = header_fill
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border

    ws.cell(row=row, column=right_alloc_col, value="% Allocation")
    ws.cell(row=row, column=right_alloc_col).fill = header_fill
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border

    row += 1

    for _, data_row in bond_data.iterrows():
        for old_idx in bond_cols_to_keep:
            new_idx = equity_col_map[old_idx]  
            value = data_row[old_idx]
            if not pd.isna(value):
                cell_value = format_num(value)
                ws.cell(row=row, column=new_idx + col_offset - 1, value=cell_value)
                ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        row += 1

    for idx, value in enumerate(bond_total):
        if value is not None and value != '':
            cell_value = format_num(value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1, value=cell_value)
            ws.cell(row=row, column=idx + 1 + col_offset - 1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=idx + 1 + col_offset - 1).border = thin_border
            ws.cell(row=row, column=idx + 1 + col_offset - 1).fill = total_fill
            if idx == 0:
                ws.cell(row=row, column=idx + 1 + col_offset - 1).font = Font(bold=True)
            
    if total_portfolio_value > 0:
        bond_percent = (bond_market_value / total_portfolio_value) * 100
        ws.cell(row=row, column=len(bond_cols_to_keep) + col_offset, value=f"{format_num(bond_percent)}%")
        ws.cell(row=row, column=len(bond_cols_to_keep) + col_offset).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=len(bond_cols_to_keep) + col_offset).border = thin_border
        ws.cell(row=row, column=len(bond_cols_to_keep) + col_offset).fill = total_fill

    wb.save('portfolio_summary.xlsx')
    print("Excel file created: portfolio_summary.xlsx")