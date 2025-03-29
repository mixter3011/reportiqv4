from typing import Optional, List
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def plot_table_and_pie(df, Holding, xirr, cash_equivalent_value, cash_equivalent_percentage, equity_allocation_percentage):
    df = df.replace([np.inf, -np.inf], 0)
    df = df.fillna(0)
    df["Portfolio Value"] = pd.to_numeric(df["Portfolio Value"], errors='coerce').fillna(0)
    df["Portfolio Value"] = df["Portfolio Value"].round().astype(int)
    
    order = ["Available Cash", "Debt", "Equity", "Gold"]
    ordered_df = df[df["Portfolio Component"].isin(order)].copy()
    ordered_df["Order"] = ordered_df["Portfolio Component"].map(lambda x: order.index(x))
    ordered_df = ordered_df.sort_values(by="Order").drop(columns=["Order"])
    
    pie_values = ordered_df["Portfolio Value"].values
    pie_labels = ordered_df["Portfolio Component"].values
    grand_total_value = ordered_df["Portfolio Value"].sum()
    
    ordered_df["Portfolio Value"] = ordered_df["Portfolio Value"].apply(lambda x: f"{x:,}")
    
    grand_total_row = pd.DataFrame({
        "Portfolio Component": ["Grand Total"],
        "Portfolio Value": [f"{grand_total_value:,}"]
    })
    
    result_df = pd.concat([ordered_df, grand_total_row], ignore_index=True)
    
    fig = plt.figure(figsize=(15.8, 8))
    fig.patch.set_facecolor('white')
    
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_facecolor('white')
    ax.axis('off')

    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)

    try:
        logo = plt.imread('logo.png')
        header_img = plt.imread('header.png')
        logo_ax = fig.add_axes([0.79, 0.85, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
        header_ax = fig.add_axes([0.02, 0.72, 0.9, 0.3])
        header_ax.imshow(header_img)
        header_ax.axis('off')
    except:
        print("Warning: One or more image files not found")

    ax.text(0.11, 0.865, "Holding Summary and Performance",
            fontsize=28, color='#CD0000', weight='light')
    ax.text(0.56, 0.87, f"For Period (As On {pd.Timestamp.now().strftime('%d-%b-%Y')})",
            fontsize=12, color='#000000', weight='normal')

    ax.text(0.05, 0.80, "Investment Summary",
            fontsize=16, fontweight='bold', color='#000000')
    ax.add_patch(plt.Rectangle((0.035, 0.80), 0.004, 0.015,
                              facecolor='#CD0000'))

    metrics_data = [
        ["Metric", "Value"],
        ["Cash Equivalent Value", f"₹{cash_equivalent_value:,.2f}"],
        ["Cash Equivalent %", f"{cash_equivalent_percentage:.2f}%"],
        ["Equity Allocation %", f"{equity_allocation_percentage:.2f}%"]
    ]

    metrics_table_ax = fig.add_axes([0.05, 0.1, 0.4, 0.35])
    metrics_table_ax.axis('off')

    metrics_table = metrics_table_ax.table(
        cellText=metrics_data[1:],
        colLabels=metrics_data[0],
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1],
        colWidths=[0.6, 0.4]
    )

    metrics_table.auto_set_font_size(False)
    metrics_table.set_fontsize(10)
    for (row, col), cell in metrics_table._cells.items():
        cell.set_edgecolor('black')
        cell.set_height(0.2)
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        else:
            cell.set_facecolor('white')
            if col == 0:
                cell.get_text().set_weight('bold')
 
    table_ax = fig.add_axes([0.05, 0.55, 0.4, 0.2])
    table_ax.axis('off')

    table = table_ax.table(
        cellText=result_df.values,
        colLabels=["Portfolio Component", "Portfolio Value"],
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1],
        colWidths=[0.5, 0.5]
    )

    table.auto_set_font_size(False)
    table.set_fontsize(10)
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        if row == len(result_df):
            cell.get_text().set_weight('bold')

    pie_ax = fig.add_axes([0.5, 0.2, 0.45, 0.45])
    
    percentages = (pie_values / grand_total_value) * 100
    labels = [f'{label}\n({percent:.1f}%)' for label, percent in zip(pie_labels, percentages)]
    
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']
    
    wedges, texts, autotexts = pie_ax.pie(
        pie_values,
        labels=labels,
        colors=colors,
        autopct='',
        startangle=90
    )
    
    for text in texts:
        text.set_fontsize(9)

    pie_ax.text(0, 1.6, "Portfolio Allocation",
                fontsize=12, fontweight='bold', ha='center')

    footer_text = (
        "This document is not valid without disclosure. Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"In case of any query/feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {pd.Timestamp.now().strftime('%d-%b-%Y')}"
    )
    ax.text(
        0.5, 0.02, footer_text,
        horizontalalignment='center',
        fontsize=10,
        color='#2F4F4F',
        wrap=True
    )

    return fig

# def create_holdings_summary(equity_file, debt_file, holding_df):
#     fig = plt.figure(figsize=(15.8, 10))
#     ax = fig.add_subplot(111)
#     ax.set_position([0, 0, 1, 1])
#     ax.set_xlim(0, 1)
#     ax.set_ylim(0, 1)
#     ax.axis('off')
    
#     gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
#     ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
#     current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
#     current_time = pd.Timestamp.now().strftime('%I:%M %p')
    
#     ax.text(0.11, 0.915, 'Holding Summary and Performance', 
#             color='#CD0000', fontsize=28, fontweight='light')
#     ax.text(0.56, 0.92, f"For Period (As On {current_date})",
#             fontsize=12, color='#000000', weight='normal')
    
#     ax.text(0.05, 0.77, "Productwise Performance", 
#             fontsize=16, fontweight='bold', color='#000000')
#     ax.text(0.83, 0.77, "(Amount in Lacs)", 
#             fontsize=12, color='#666666')
#     ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
#                                 facecolor='#CD0000'))
    
#     try:
#         logo = plt.imread('logo.png')
#         header = plt.imread('header.png')
        
#         logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
#         logo_ax.imshow(logo)
#         logo_ax.axis('off')
        
#         header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
#         header_ax.imshow(header)
#         header_ax.axis('off')
#     except:
#         print("Warning: Image files not found")
    
#     summary_tables = []
#     for df, category_label in [(equity_file, 'Equity'), (debt_file, 'Debt')]:
#         temp_summary = pd.DataFrame()
#         for category in df['Category'].unique():
#             category_data = df[df['Category'] == category]
#             numeric_cols = category_data.select_dtypes(include=['number']).columns
#             summary_row = category_data[numeric_cols].sum().to_frame().T
#             summary_row.insert(0, 'Category', category)
#             temp_summary = pd.concat([temp_summary, summary_row], ignore_index=True)
#         summary_tables.append(temp_summary)
    
#     combined_summary = pd.concat(summary_tables, ignore_index=True)
#     combined_summary = combined_summary.drop(columns=['PandL', 'CMP', 'Quantity', 'Type'], errors='ignore')
#     numeric_cols = combined_summary.select_dtypes(include=['number']).columns
#     combined_summary[numeric_cols] = combined_summary[numeric_cols].apply(lambda x: (x / 1000).round(2))
#     combined_summary = combined_summary.rename(columns={
#         'Category': 'Asset Class',
#         'Buy Price': 'Investment at Cost'
#     })
    
#     combined_summary['Dividend Since Inception'] = '-'  
#     combined_summary['Unrealised G/L %'] = '-'

#     combined_summary['Asset Class'] = combined_summary['Asset Class'].str.upper()
    
#     if 'Market Value' in numeric_cols:
#         combined_summary['Market Value'] = combined_summary['Market Value'].apply(lambda x: round(x, 3))
#     combined_summary['Unrealised G/L %'] = combined_summary['Unrealised G/L %'].round(2)
    
#     total_row = pd.DataFrame(combined_summary.select_dtypes(include=['number']).sum()).T
#     total_row.insert(0, 'Asset Class', 'TOTAL')
#     total_row['Dividend Since Inception'] = '-'
#     if 'Market Value' in numeric_cols:
#         total_row['Market Value'] = round(total_row['Market Value'].iloc[0], 2)  
#     total_row['Unrealised G/L %'] = '-'  
#     combined_summary = pd.concat([combined_summary, total_row], ignore_index=True)
    
#     table_ax = fig.add_axes([0.03, 0.25, 0.9, 0.5])
#     table_ax.axis('off')
    
#     table = table_ax.table(
#         cellText=combined_summary.values,
#         colLabels=combined_summary.columns,
#         cellLoc='center',
#         loc='center',
#         bbox=[0, 0, 1, 1]
#     )
    
#     table.auto_set_font_size(False)
#     table.set_fontsize(10)
#     table.auto_set_column_width(col=list(range(len(combined_summary.columns))))
    
#     for (row, col), cell in table._cells.items():
#         cell.set_edgecolor('black')
#         if row == 0:  
#             cell.set_facecolor('#E6E6E6')
#             cell.set_text_props(weight='bold')
#         elif row == len(combined_summary):  
#             cell.set_facecolor('#ADD8E6')  
#             cell.set_text_props(weight='bold')
#         elif 'Market Value' in combined_summary.columns and col == combined_summary.columns.get_loc('Market Value'):
#             if row != len(combined_summary) - 1:  
#                 cell.get_text().set_text(f"{float(cell.get_text().get_text()):.3f}")
    
#     current_time = pd.Timestamp.now()
#     footer_text = (
#         "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
#         "Strictly Private & Confidential.\n"
#         f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
#         f"Generated Date & Time : {current_time.strftime('%d-%b-%Y')} & {current_time.strftime('%I:%M %p')}"
#     )
    
#     ax.text(0.5, 0.02, footer_text,
#             horizontalalignment='center',
#             fontsize=10,
#             color='#2F4F4F',
#             wrap=True)
    
#     return fig

def create_combined_portfolio_table(
    df: pd.DataFrame,
    portfolio_type: str = 'equity',
    target_stocks: Optional[List[str]] = None,
    label: str = None  
) -> plt.Figure:

    
    if portfolio_type.lower() == 'equity':
        equity_data = []
        equity_section = False
        
        for idx, row in df.iterrows():
            if isinstance(row['Unnamed: 0'], str) and 'Equity' in row['Unnamed: 0']:
                equity_section = True
                continue
            if equity_section and isinstance(row['Unnamed: 0'], str) and (not target_stocks or any(stock in row['Unnamed: 0'] for stock in target_stocks)):
                equity_data.append(row.tolist())
            if equity_section and isinstance(row['Unnamed: 0'], str) and 'Total' in row['Unnamed: 0']:
                equity_section = False
        
        cols = [
            "Instrument Name", "Quantity", "Purchase Price", "Purchase Value",
            "Market Price", "Market Value", "ST Qty", "ST G/L", "LT Qty", "LT G/L",
            "Unrealised Gain/Loss", "Unrealised Gain/Loss %", "ISIN", "Unused_1", "Unused_2"
        ]
        portfolio_data = pd.DataFrame(equity_data, columns=df.columns)
        portfolio_data.columns = cols
        
        if target_stocks:
            portfolio_data = portfolio_data[portfolio_data["Instrument Name"].isin(target_stocks)]
        
        display_cols = [
            "Instrument Name", "Quantity", "Purchase Price", "Market Price", 
            "Unrealised Gain/Loss", "Market Value"
        ]
        
        column_rename_map = {
            "Purchase Price": "Buy Price(Sum)",
            "Market Price": "CMP(Sum)",
            "Unrealised Gain/Loss": "PandL(Sum)"
        }
        
    else:  
        eq_strt = df[df.iloc[:, 0] == 'Mutual Fund:-'].index[0]
        eq_end = df[df.iloc[:, 0] == 'FnO:-'].index[0]
        
        portfolio_data = df.iloc[eq_strt+1:eq_end-4].copy()
        portfolio_data.columns = portfolio_data.iloc[0]
        portfolio_data = portfolio_data.iloc[1:]
        portfolio_data = portfolio_data.reset_index(drop=True)
        portfolio_data = portfolio_data[~portfolio_data['Asset Type'].isin(['Debt']) & ~portfolio_data['Asset Type'].isna()]
        
        display_cols = [
            'Scheme Name', 'Units', 'Purchase NAV', 'Current NAV',
            'Unrealised GainLoss', 'Market Value'
        ]
        
        column_rename_map = {
            "Purchase NAV": "Buy Price(Sum)",
            "Current NAV": "CMP(Sum)",
            "Unrealised GainLoss": "PandL(Sum)"
        }

    fig = plt.figure(figsize=(15.8, 10))
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    current_time = pd.Timestamp.now().strftime('%I:%M %p')
    
    ax.text(0.11, 0.915, 'Detailed Holdings and Performance', 
            color='#CD0000', fontsize=28, fontweight='light')
    ax.text(0.56, 0.92, f"For Period (As On {current_date})",
            fontsize=12, color='#000000', weight='normal')
    

    if label is None:
        label_text = "Direct Equity" if portfolio_type.lower() == 'equity' else "Mutual Fund"
    else:
        label_text = label  
    
    ax.text(0.05, 0.77, "Equity - ", 
            fontsize=16, fontweight='bold', color='#CD0000')
    ax.text(0.13, 0.77, label_text, 
            fontsize=16, fontweight='light', color='#666666')
    ax.text(0.83, 0.77, "(Amount in Lacs)", 
            fontsize=12, color='#666666')
    ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
                              facecolor='#CD0000'))
    
    try:
        logo = plt.imread('logo.png')
        logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
            
        header = plt.imread('header.png')
        header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
        header_ax.imshow(header)
        header_ax.axis('off')
    except Exception as e:
        print(f"Warning: Image loading error: {e}")
    
    table_ax = fig.add_axes([0.03, 0.25, 0.94, 0.5])
    table_ax.axis('off')
    
    table_data = portfolio_data[display_cols].values.tolist()
    
    display_cols_renamed = [column_rename_map.get(col, col) for col in display_cols]
    
    table = table_ax.table(
        cellText=table_data,
        colLabels=display_cols_renamed,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    col_widths = [0.25] + [0.15] * (len(display_cols) - 1)
    for i, width in enumerate(col_widths):
        table.auto_set_column_width([i])
    
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        
        if col == 0 and row != 0:
            cell._loc = 'left'
            
        if row > 0 and col == 4:
            text = cell.get_text().get_text()
            if text.startswith('(') or (text.replace('.','').replace('-','').isdigit() and float(text) < 0):
                cell.get_text().set_color('red')
    
    footer_text = (
        "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} & {current_time}"
    )
    
    ax.text(0.5, 0.02, footer_text,
            horizontalalignment='center',
            fontsize=10,
            color='#2F4F4F',
            wrap=True)
    
    return fig

def create_portfolio_table(df):
    target_stocks = [
        "Bajaj Finserv Ltd",
        "Central Depository Services (India) Ltd",
        "HDFC Bank Ltd",
        "IDFC First Bank Ltd",
        "Kotak Mahindra Bank Ltd",
        "Shriram Finance Ltd",
        "Bharat Forge Ltd",
        "Tata Consultancy",
    ]
    
    equity_data = []
    equity_section = False
    for idx, row in df.iterrows():
        if isinstance(row['Unnamed: 0'], str) and 'Equity' in row['Unnamed: 0']:
            equity_section = True
            continue
        if equity_section and isinstance(row['Unnamed: 0'], str) and any(stock in row['Unnamed: 0'] for stock in target_stocks):
            equity_data.append(row.tolist())
        if equity_section and isinstance(row['Unnamed: 0'], str) and 'Total' in row['Unnamed: 0']:
            equity_section = False
    
    cols = [
        "Instrument Name", "Quantity", "Purchase Price", "Purchase Value",
        "Market Price", "Market Value", "ST Qty", "ST G/L", "LT Qty", "LT G/L",
        "Unrealised Gain/Loss", "Unrealised Gain/Loss %", "ISIN", "Unused_1", "Unused_2"
    ]
    equity = pd.DataFrame(equity_data, columns=df.columns)
    equity.columns = cols
    
    if target_stocks:
        equity = equity[equity["Instrument Name"].isin(target_stocks)]
    
    display_cols = [
        "Instrument Name", "Quantity", "Purchase Price", "Purchase Value",
        "Market Price", "Market Value", "Unrealised Gain/Loss",
        "Unrealised Gain/Loss %", "ISIN"
    ]
    
    fig = plt.figure(figsize=(15.8, 10))
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    current_time = pd.Timestamp.now().strftime('%I:%M %p')
    
    ax.text(0.11, 0.915, 'Detailed Holdings and Performance', 
            color='#CD0000', fontsize=28, fontweight='light')
    ax.text(0.56, 0.92, f"For Period (As On {current_date})",
            fontsize=12, color='#000000', weight='normal')
    
    ax.text(0.05, 0.77, "Equity - ", 
            fontsize=16, fontweight='bold', color='#CD0000')
    ax.text(0.12, 0.77, "Direct Equity", 
            fontsize=16, fontweight='light', color='#666666')
    ax.text(0.83, 0.77, "(Amount in Lacs)", 
            fontsize=12, color='#666666')
    ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
                              facecolor='#CD0000'))
    
    try:
        logo = plt.imread('logo.png')
        logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
            
        header = plt.imread('header.png')
        header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
        header_ax.imshow(header)
        header_ax.axis('off')
    except Exception as e:
        print(f"Warning: Image loading error: {e}")
    
    table_data = equity[display_cols].values.tolist()
    
    table_ax = fig.add_axes([0.03, 0.25, 0.94, 0.5])
    table_ax.axis('off')
    
    table = table_ax.table(
        cellText=table_data,
        colLabels=display_cols,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    col_widths = [0.25] + [0.09375] * (len(display_cols) - 1)
    for i, width in enumerate(col_widths):
        table.auto_set_column_width([i])
    
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        
        if col == 0 and row != 0:
            cell._loc = 'left'
            
        if row > 0 and col in [6, 7]:  
            text = cell.get_text().get_text()
            if text.startswith('(') or (text.replace('.','').replace('-','').isdigit() and float(text) < 0):
                cell.get_text().set_color('red')
    
    footer_text = (
        "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} & {current_time}"
    )
    
    ax.text(0.5, 0.02, footer_text,
            horizontalalignment='center',
            fontsize=10,
            color='#2F4F4F',
            wrap=True)
    
    return fig

def analyze_fno_holdings(df):
    fno_start = df[df.iloc[:, 0] == 'FnO:-'].index[0]
    fno_end = df[df.iloc[:, 0] == 'Currency:-'].index[0]
    
    fno_data = df.iloc[fno_start+1:fno_end-4].copy()
    fno_data.columns = fno_data.iloc[0]
    fno_data = fno_data.iloc[1:]
    fno_data = fno_data.reset_index(drop=True)
    fno_data = fno_data.dropna(subset=['Instrument Name'])
    
    fig = plt.figure(figsize=(15.8, 10))
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    current_time = pd.Timestamp.now().strftime('%I:%M %p')
    
    ax.text(0.11, 0.915, 'Detailed Holdings and Performance', 
            color='#CD0000', fontsize=28, fontweight='light')
    ax.text(0.56, 0.92, f"For Period (As On {current_date})",
            fontsize=12, color='#000000', weight='normal')
    
    ax.text(0.05, 0.77, "Derivatives", 
            fontsize=16, fontweight='bold', color='#000000')
    ax.text(0.83, 0.77, "(Amount in Lacs)", 
            fontsize=12, color='#666666')
    ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
                              facecolor='#CD0000'))
    
    try:
        logo = plt.imread('logo.png')
        logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
            
        header = plt.imread('header.png')
        header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
        header_ax.imshow(header)
        header_ax.axis('off')
    except Exception as e:
        print(f"Warning: Image loading error: {e}")
    
    columns = [
        'Instrument Name', 'B/S', 'Quantity', 'Rate', 'Value',
        'Market Price', 'Market Value', 'UnrealisedGain/Loss',
        'Unrealised Gain/Loss%'
    ]
    
    table_ax = fig.add_axes([0.03, 0.25, 0.94, 0.5])
    table_ax.axis('off')
    
    table = table_ax.table(
        cellText=fno_data[columns].values,
        colLabels=columns,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    col_widths = [0.25] + [0.09375] * (len(columns) - 1)
    for i, width in enumerate(col_widths):
        table.auto_set_column_width([i])
    
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        
        if col == 0 and row != 0:
            cell._loc = 'left'
            
        if row > 0 and col in [7, 8]:  
            text = cell.get_text().get_text()
            if text.startswith('(') or (text.replace('.','').replace('-','').isdigit() and float(text) < 0):
                cell.get_text().set_color('red')
    
    footer_text = (
        "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} & {current_time}"
    )
    
    ax.text(0.5, 0.02, footer_text,
            horizontalalignment='center',
            fontsize=10,
            color='#2F4F4F',
            wrap=True)
    
    return fig

def eqmf(df):
    eq_strt = df[df.iloc[:, 0] == 'Mutual Fund:-'].index[0]
    eq_end = df[df.iloc[:, 0] == 'FnO:-'].index[0]
    
    eq_data = df.iloc[eq_strt+1:eq_end-4].copy()
    eq_data.columns = eq_data.iloc[0]
    eq_data = eq_data.iloc[1:]
    eq_data = eq_data.reset_index(drop=True)
    
    eq_data = eq_data.iloc[:-1]
    
    eq_data = eq_data[~eq_data['Asset Type'].isin(['Debt']) & ~eq_data['Asset Type'].isna()]
    
    column_rename_map = {
        "Purchase NAV": "Buy Price(Sum)",
        "Current NAV": "CMP(Sum)",
        "Unrealised GainLoss": "PandL(Sum)"
    }

    numeric_cols = ['Units', 'Purchase Value', 'Market Value', 
                   'ST Qty', 'ST G/L', 'LT Qty', 'LT G/L', 
                   'Dividend', 'Unrealised GainLoss']
    for col in numeric_cols:
        eq_data[col] = pd.to_numeric(eq_data[col], errors='coerce').fillna(0)
    
    original_columns = [
        'Scheme Name', 'Units', 'Purchase NAV', 'Purchase Value', 'Current NAV',
        'Market Value', 'ST Qty', 'ST G/L', 'LT Qty', 'LT G/L', 'Dividend',
        'Unrealised GainLoss', 'Unrealised GainLoss Per'
    ]
    
    display_cols_renamed = [column_rename_map.get(col, col) for col in original_columns]
    
    fig = plt.figure(figsize=(15.8, 10))
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    current_time = pd.Timestamp.now().strftime('%I:%M %p')

    ax.text(0.11, 0.915, 'Detailed Holdings and Performance', 
            color='#CD0000', fontsize=28, fontweight='light')
    ax.text(0.56, 0.92, f"For Period (As On {current_date})",
            fontsize=12, color='#000000', weight='normal')
    
    ax.text(0.05, 0.77, "Equity - ", 
            fontsize=16, fontweight='bold', color='#CD0000')
    ax.text(0.12, 0.77, "Mutual Fund", 
            fontsize=16, fontweight='light', color='#666666')
    ax.text(0.83, 0.77, "(Amount in Lacs)", 
            fontsize=12, color='#666666')
    ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
                              facecolor='#CD0000'))
    
    try:
        logo = plt.imread('logo.png')
        logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
            
        header = plt.imread('header.png')
        header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
        header_ax.imshow(header)
        header_ax.axis('off')
    except Exception as e:
        print(f"Warning: Image loading error: {e}")
    
    columns = [
        'Scheme Name', 'Units', 'Purchase NAV', 'Purchase Value', 'Current NAV',
        'Market Value', 'ST Qty', 'ST G/L', 'LT Qty', 'LT G/L', 'Dividend',
        'Unrealised GainLoss', 'Unrealised GainLoss Per'
    ]        

    table_ax = fig.add_axes([0.03, 0.25, 0.94, 0.5])
    table_ax.axis('off')
    
    table_data = eq_data[original_columns].values.tolist()
    
    table_data = eq_data[columns].values.tolist()
    table = table_ax.table(
        cellText=table_data,
        colLabels=display_cols_renamed,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    col_widths = [0.25] + [0.09375] * (len(columns) - 1)
    for i, width in enumerate(col_widths):
        table.auto_set_column_width([i])
    
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
            
        if col == 0 and row != 0:
            cell._loc = 'left'
            
        if row > 0 and col in [7, 8]:  
            text = cell.get_text().get_text()
            if text.startswith('(') or (text.replace('.','').replace('-','').isdigit() and float(text) < 0):
                cell.get_text().set_color('red')
    
    footer_text = (
        "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} & {current_time}"
    )
    
    ax.text(0.5, 0.02, footer_text,
            horizontalalignment='center',
            fontsize=10,
            color='#2F4F4F',
            wrap=True)
    
    return fig

def dmf(df: pd.DataFrame, label: str = None) -> plt.Figure:
    d_strt = df[df.iloc[:, 0] == 'Mutual Fund:-'].index[0]
    d_end = df[df.iloc[:, 0] == 'FnO:-'].index[0]
    
    d_data = df.iloc[d_strt+1:d_end-4].copy()
    d_data.columns = d_data.iloc[0]
    d_data = d_data.iloc[1:]
    d_data = d_data.reset_index(drop=True)
    d_data = d_data[~d_data['Asset Type'].isin(['Equity']) & ~d_data['Asset Type'].isna()]
    
    d_data = d_data.iloc[:-1]
    
    numeric_cols = ['Units', 'Purchase NAV', 'Market Value', 'Unrealised GainLoss']
    
    for col in numeric_cols:
        d_data[col] = pd.to_numeric(d_data[col], errors='coerce').fillna(0)
    
    total_values = d_data[numeric_cols].sum()
    total_row = {
        'Scheme Name': 'TOTAL',
        'Units': total_values['Units'],
        'Purchase NAV': total_values['Purchase NAV'],
        'Current NAV': '-',
        'Unrealised GainLoss': total_values['Unrealised GainLoss'],
        'Market Value': total_values['Market Value']
    }
    d_data = pd.concat([d_data, pd.DataFrame([total_row])], ignore_index=True)
    
    display_cols = [
        'Scheme Name', 'Units', 'Purchase NAV', 'Current NAV',
        'Unrealised GainLoss', 'Market Value'
    ]
    
    column_rename_map = {
        "Purchase NAV": "Buy Price(Sum)",
        "Current NAV": "CMP(Sum)",
        "Unrealised GainLoss": "PandL(Sum)"
    }
    
    fig = plt.figure(figsize=(15.8, 10))
    ax = fig.add_subplot(111)
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    
    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(gradient, extent=(0, 1, 0, 1), cmap='Blues', aspect='auto', alpha=0.3)
    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    current_time = pd.Timestamp.now().strftime('%I:%M %p')

    ax.text(0.11, 0.915, 'Detailed Holdings and Performance', 
            color='#CD0000', fontsize=28, fontweight='light')
    ax.text(0.56, 0.92, f"For Period (As On {current_date})",
            fontsize=12, color='#000000', weight='normal')
    
    label_text = label if label is not None else ""
    ax.text(0.05, 0.77, "Debt - ", 
            fontsize=16, fontweight='bold', color='#CD0000')
    ax.text(0.12, 0.77, label_text, 
            fontsize=16, fontweight='light', color='#666666')
    ax.text(0.83, 0.77, "(Amount in Lacs)", 
            fontsize=12, color='#666666')
    ax.add_patch(plt.Rectangle((0.035, 0.77), 0.004, 0.015,
                              facecolor='#CD0000'))
    
    try:
        logo = plt.imread('logo.png')
        logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])
        logo_ax.imshow(logo)
        logo_ax.axis('off')
            
        header = plt.imread('header.png')
        header_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
        header_ax.imshow(header)
        header_ax.axis('off')
    except Exception as e:
        print(f"Warning: Image loading error: {e}")
    
    table_data = d_data[display_cols].values.tolist()
    display_cols_renamed = [column_rename_map.get(col, col) for col in display_cols]
    
    table_ax = fig.add_axes([0.03, 0.25, 0.94, 0.5])
    table_ax.axis('off')
    
    table = table_ax.table(
        cellText=table_data,
        colLabels=display_cols_renamed,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1]
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    
    col_widths = [0.25] + [0.15] * (len(display_cols) - 1)
    for i, width in enumerate(col_widths):
        table.auto_set_column_width([i])
    
    for (row, col), cell in table._cells.items():
        cell.set_edgecolor('black')
        
        if row == 0:
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        elif row == len(table_data):
            cell.set_facecolor('#E6E6E6')
            cell.set_text_props(weight='bold')
        
        if col == 0 and row != 0:
            cell._loc = 'left'
            
        if row > 0 and col == 4:  
            text = cell.get_text().get_text()
            if text.startswith('(') or (text.replace('.','').replace('-','').isdigit() and float(text) < 0):
                cell.get_text().set_color('red')
    
    footer_text = (
        "This document is not valid without disclosure, Please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} & {current_time}"
    )
    
    ax.text(0.5, 0.02, footer_text,
            horizontalalignment='center',
            fontsize=10,
            color='#2F4F4F',
            wrap=True)
    
    return fig
    