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
    equity_end = mf_row - 4  
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

    etf_equity = equity_data[equity_data['Unnamed: 0'].str.contains('ETF', na=False)]
    etf_equity = etf_equity[~etf_equity['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]
    etf_equity = etf_equity[~etf_equity['Unnamed: 0'].str.contains('Nippon India ETF Nifty 8-13 yr G-Sec LongTerm Gilt', na=False)]

    debt_etf = equity_data[equity_data['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]

    gilt_etf = equity_data[equity_data['Unnamed: 0'].str.contains('Nippon India ETF Nifty 8-13 yr G-Sec LongTerm Gilt', na=False)]

    mf_cols_to_keep = [1, 2, 3, 5, 6, 12]  
    mf_rename = {3: "Buy Price", 12: "P&L"}

    equity_mf = mf_data[mf_data['Unnamed: 0'] == 'Equity']

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

    bond_cols_to_keep = [0, 1, 2, 4, 5, 10]  
    bond_rename = {2: "Buy Price", 10: "P&L"}

    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    total_fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
    portfolio_value_fill = PatternFill(start_color="B0E0E6", end_color="B0E0E6", fill_type="solid") 
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

    row = 5
    ws.cell(row=row, column=1, value="EQUITY").font = Font(bold=True)

    row_portfolio = row - 1
    ws.cell(row=row_portfolio, column=1, value="Portfolio Value").font = Font(bold=True)
    ws.cell(row=row_portfolio, column=1).border = thin_border
    ws.cell(row=row_portfolio, column=1).alignment = Alignment(horizontal="left")
    
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
    direct_equity_start_row = row

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

    direct_equity_total_row = row
    ws.cell(row=row, column=1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1).fill = total_fill
    
    for col_idx, old_idx in enumerate(equity_cols_to_keep[1:], 1):
        new_idx = equity_col_map[old_idx]
        col_letter = get_column_letter(new_idx)
        ws.cell(row=row, column=new_idx, value=f"=SUM({col_letter}{direct_equity_start_row}:{col_letter}{row-1})")
        ws.cell(row=row, column=new_idx).border = thin_border
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).fill = total_fill

    col_letter = get_column_letter(5)  
    ws.cell(row=row, column=alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border
    ws.cell(row=row, column=alloc_col).fill = total_fill
    ws.cell(row=row, column=alloc_col).number_format = '0.00"%"'

    row += 2

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
    etf_equity_start_row = row

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

    etf_equity_total_row = row
    ws.cell(row=row, column=1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1).fill = total_fill
    
    for col_idx, old_idx in enumerate(equity_cols_to_keep[1:], 1):
        new_idx = equity_col_map[old_idx]
        col_letter = get_column_letter(new_idx)
        ws.cell(row=row, column=new_idx, value=f"=SUM({col_letter}{etf_equity_start_row}:{col_letter}{row-1})")
        ws.cell(row=row, column=new_idx).border = thin_border
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).fill = total_fill

    col_letter = get_column_letter(5)  
    ws.cell(row=row, column=alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border
    ws.cell(row=row, column=alloc_col).fill = total_fill
    ws.cell(row=row, column=alloc_col).number_format = '0.00"%"'
        
    row += 2

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
    equity_mf_start_row = row

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

    equity_mf_total_row = row
    ws.cell(row=row, column=1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1).fill = total_fill
    
    for col_idx, old_idx in enumerate(mf_cols_to_keep[1:], 1):
        new_idx = mf_col_map[old_idx]
        col_letter = get_column_letter(new_idx)
        ws.cell(row=row, column=new_idx, value=f"=SUM({col_letter}{equity_mf_start_row}:{col_letter}{row-1})")
        ws.cell(row=row, column=new_idx).border = thin_border
        ws.cell(row=row, column=new_idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx).fill = total_fill

    col_letter = get_column_letter(6)  
    ws.cell(row=row, column=alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border
    ws.cell(row=row, column=alloc_col).fill = total_fill
    ws.cell(row=row, column=alloc_col).number_format = '0.00"%"'
        
    row += 2

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
    debt_etf_start_row = row

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

    debt_etf_total_row = row
    ws.cell(row=row, column=1 + col_offset - 1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=1 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1 + col_offset - 1).fill = total_fill
    
    for col_idx, old_idx in enumerate(equity_cols_to_keep[1:], 1):
        new_idx = equity_col_map[old_idx]
        col_letter = get_column_letter(new_idx + col_offset - 1)
        start_row = debt_etf_start_row
        end_row = row - 1
        if end_row >= start_row:  
            ws.cell(row=row, column=new_idx + col_offset - 1, value=f"=SUM({col_letter}{start_row}:{col_letter}{end_row})")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = total_fill

    col_letter = get_column_letter(5 + col_offset - 1)  
    ws.cell(row=row, column=right_alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border
    ws.cell(row=row, column=right_alloc_col).fill = total_fill
    ws.cell(row=row, column=right_alloc_col).number_format = '0.00"%"'
        
    row += 2

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
    debt_mf_start_row = row

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

    debt_mf_total_row = row
    ws.cell(row=row, column=1 + col_offset - 1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=1 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1 + col_offset - 1).fill = total_fill
    
    for col_idx, old_idx in enumerate(mf_cols_to_keep[1:], 1):
        new_idx = mf_col_map[old_idx]
        col_letter = get_column_letter(new_idx + col_offset - 1)
        start_row = debt_mf_start_row
        end_row = row - 1
        if end_row >= start_row:  
            ws.cell(row=row, column=new_idx + col_offset - 1, value=f"=SUM({col_letter}{start_row}:{col_letter}{end_row})")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = total_fill

    col_letter = get_column_letter(6 + col_offset - 1)  
    ws.cell(row=row, column=right_alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border
    ws.cell(row=row, column=right_alloc_col).fill = total_fill
    ws.cell(row=row, column=right_alloc_col).number_format = '0.00"%"'
        
    row += 2

    row += 1
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
    bond_start_row = row

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

    bond_total_row = row
    ws.cell(row=row, column=1 + col_offset - 1, value="Total:").font = Font(bold=True)
    ws.cell(row=row, column=1 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=1 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1 + col_offset - 1).fill = total_fill
    
    for col_idx, old_idx in enumerate(bond_cols_to_keep[1:], 1):
        new_idx = equity_col_map[old_idx]
        col_letter = get_column_letter(new_idx + col_offset - 1)
        start_row = bond_start_row
        end_row = row - 1
        if end_row >= start_row:  
            ws.cell(row=row, column=new_idx + col_offset - 1, value=f"=SUM({col_letter}{start_row}:{col_letter}{end_row})")
        ws.cell(row=row, column=new_idx + col_offset - 1).border = thin_border
        ws.cell(row=row, column=new_idx + col_offset - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=new_idx + col_offset - 1).fill = total_fill

    col_letter = get_column_letter(5 + col_offset - 1)  
    ws.cell(row=row, column=right_alloc_col, value=f"={col_letter}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border
    ws.cell(row=row, column=right_alloc_col).fill = total_fill
    ws.cell(row=row, column=right_alloc_col).number_format = '0.00"%"'

    row += 2
    equity_total_row = row
    ws.cell(row=row, column=1, value="TOTAL EQUITY").font = Font(bold=True)
    ws.cell(row=row, column=1).border = thin_border
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1).fill = portfolio_value_fill

    market_val_col = get_column_letter(5)  
    mf_market_val_col = get_column_letter(6)  
    ws.cell(row=row, column=5, value=f"={market_val_col}{direct_equity_total_row}+{market_val_col}{etf_equity_total_row}+{mf_market_val_col}{equity_mf_total_row}")
    ws.cell(row=row, column=5).border = thin_border
    ws.cell(row=row, column=5).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=5).fill = portfolio_value_fill

    pnl_col = get_column_letter(6)  
    mf_pnl_col = get_column_letter(6)
    ws.cell(row=row, column=6, value=f"={pnl_col}{direct_equity_total_row}+{pnl_col}{etf_equity_total_row}+{mf_pnl_col}{equity_mf_total_row}")
    ws.cell(row=row, column=6).border = thin_border
    ws.cell(row=row, column=6).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=6).fill = portfolio_value_fill

    ws.cell(row=row, column=alloc_col, value=f"={market_val_col}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=alloc_col).border = thin_border
    ws.cell(row=row, column=alloc_col).fill = portfolio_value_fill
    ws.cell(row=row, column=alloc_col).number_format = '0.00"%"'

    debt_total_row = row
    ws.cell(row=row, column=1 + col_offset - 1, value="TOTAL DEBT").font = Font(bold=True)
    ws.cell(row=row, column=1 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=1 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=1 + col_offset - 1).fill = portfolio_value_fill

    debt_market_val_col = get_column_letter(5 + col_offset - 1)  
    debt_mf_market_val_col = get_column_letter(6 + col_offset - 1)  
    bond_market_val_col = get_column_letter(5 + col_offset - 1) 
    ws.cell(row=row, column=5 + col_offset - 1, value=f"={debt_market_val_col}{debt_etf_total_row}+{debt_mf_market_val_col}{debt_mf_total_row}+{bond_market_val_col}{bond_total_row}")
    ws.cell(row=row, column=5 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=5 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=5 + col_offset - 1).fill = portfolio_value_fill

    debt_pnl_col = get_column_letter(6 + col_offset - 1)  
    debt_mf_pnl_col = get_column_letter(6 + col_offset - 1)  
    bond_pnl_col = get_column_letter(6 + col_offset - 1)  
    ws.cell(row=row, column=6 + col_offset - 1, value=f"={debt_pnl_col}{debt_etf_total_row}+{debt_mf_pnl_col}{debt_mf_total_row}+{bond_pnl_col}{bond_total_row}")
    ws.cell(row=row, column=6 + col_offset - 1).border = thin_border
    ws.cell(row=row, column=6 + col_offset - 1).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=6 + col_offset - 1).fill = portfolio_value_fill

    ws.cell(row=row, column=right_alloc_col, value=f"={debt_market_val_col}{row}/$B${row_portfolio}*100")
    ws.cell(row=row, column=right_alloc_col).alignment = Alignment(horizontal="center")
    ws.cell(row=row, column=right_alloc_col).border = thin_border
    ws.cell(row=row, column=right_alloc_col).fill = portfolio_value_fill
    ws.cell(row=row, column=right_alloc_col).number_format = '0.00"%"'

    market_val_col = get_column_letter(5)  
    debt_market_val_col = get_column_letter(5 + col_offset - 1)  
    ws.cell(row=row_portfolio, column=2, value=f"={market_val_col}{equity_total_row}+{debt_market_val_col}{debt_total_row}")
    ws.cell(row=row_portfolio, column=2).fill = portfolio_value_fill
    ws.cell(row=row_portfolio, column=2).number_format = '#,##0.00'

    for col_idx in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15

    output_filename = f"{client_code}_{client_name}_Portfolio.xlsx"
    wb.save(output_filename)
    
    return output_filename