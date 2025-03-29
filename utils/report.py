import numpy as np
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from utils.plotting import (
    plot_table_and_pie,
    create_combined_portfolio_table,
    eqmf,
    dmf
)

def get_customer_details(holding_df):
    client_row = holding_df[holding_df['Unnamed: 0'] == 'Client Equity Code/UCID/Name']
    
    if client_row.empty:
        raise ValueError("Could not find client information row in the Holding dataframe")
    
    client_info = client_row['Unnamed: 1'].iloc[0]
    
    parts = client_info.split('/')
    if len(parts) != 3:
        raise ValueError(f"Unexpected format in client information: {client_info}")
    
    ucid = parts[1]
    customer_name = parts[2]
    
    return customer_name, ucid

def create_cover_page(pdf, customer_name, ucid):
    fig = plt.figure(figsize=(16, 10))
    ax = fig.add_subplot(111)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(
        gradient,
        extent=(0, 1, 0, 1),
        cmap='Blues',
        aspect='auto',
        alpha=0.3
    )
    
    logo_img = plt.imread('logo.png')
    logo_ax = fig.add_axes([0.78, 0.88, 0.2, 0.08])  
    logo_ax.imshow(logo_img)
    logo_ax.axis('off')

    ax.text(
        0.04, 0.92, "CUSTOMER STATEMENT",
        fontsize=32, color='#8B0000',
        weight='light'
    )

    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    ax.text(
        0.045, 0.88,  
        f"Report Level : Member | Report Period : Since Inception to {current_date}",
        fontsize=16, color='#2F4F4F'
    )

    ax.text(0.05, 0.6, customer_name, fontsize=28, color='#2F4F4F', weight='bold')  
    ax.text(0.05, 0.55, f"UCID : {ucid}", fontsize=20, color='#2F4F4F') 

    footer_text = (
        "This document is not valid without disclosure, please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} | {pd.Timestamp.now().strftime('%I:%M %p')}"
    )
    ax.text(
        0.5, 0.1, footer_text,
        horizontalalignment='center',
        fontsize=12,
        color='#2F4F4F',
        wrap=True
    )

    ax.text(
        0.056, 0.23, "WINNING PORTFOLIOS",
        fontsize=14,
        color='#2F4F4F',
        weight='bold'
    )

    ax.text(
        0.05, 0.2, "POWERED BY KNOWLEDGE",  
        fontsize=14,
        color='white',
        bbox=dict(
            facecolor='#FF0000',
            edgecolor='none',
            boxstyle='round,pad=0.5'
        )
    )

    footer_img = plt.imread('footer.png')
    footer_logo_ax = fig.add_axes([0.23, 0.12, 0.8, 0.2])  
    footer_logo_ax.imshow(footer_img)
    footer_logo_ax.axis('off')

    plt.subplots_adjust(top=1, bottom=0, left=0, right=1, hspace=0, wspace=0)
    pdf.savefig(fig, bbox_inches=None, pad_inches=0)
    plt.close(fig)

def create_footer_page(pdf):
    fig = plt.figure(figsize=(16, 10))
    ax = fig.add_subplot(111)
    
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(
        gradient,
        extent=(0, 1, 0, 1),
        cmap='Blues',
        aspect='auto',
        alpha=0.3
    )

    logo_img = plt.imread('logo.png')
    logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])  
    logo_ax.imshow(logo_img)
    logo_ax.axis('off')
    
    ax.text(
        0.11, 0.92, "Notes & Assumptions",
        fontsize=32, color='#8B0000',
        weight='light'
    )

    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    
    main_txt = (
        "1. All valuation as on last available Price / NAV.\n\n"
        "2. Unrealised Gain / Loss = Market Value - Investment at cost. It does not account for Dividend / Interest Paid out. XIRR will be frozen for that day till next valuation is available.\n\n"
        "3. XIRR and Benchmark Values are calculated on balance units for open scripts / folios. In case there are reinvestments in a prior closed script / folio, then all historical cashflow of that script / foliow will be considered\n\n     for XIRR & Benchmark Calculations.\n\n"
        "4. All Private Equity Funds(RE) and Real Estate Funds(RE) are classified as Alternates (Unquoted).Since Quoted Valuations of PE & RE Funds are not available its Market value is computed as Total Drawdown(Cost) -\n\n    Capital Returned.XIRR and benchmarks are not computed for PE & RE Funds.\n\n"
        "5. If bifurcation between Capital Return & Profit is not provided by Manufacturers for AIF / PE / RE distribution, they will appear as Profit / Unallocated distribution in report.\n\n"
        "6. Net investments for PMS might not match with AMC statement due to difference in calculation methodology.In case if cost in folio is less than equal to zero and Market Value is less than 50000, we will consider this\n\n    as a cloud investment and will not appear in report.\n\n"
        "7. For Mutual Funds, less than 1 units or fraction unit is considered as closed.\n\n"
        "8. Capital Gain / (Loss) – for Equity Stocks / Mutual Fund: Short Term < 1 Year, Long Term > 1 Year; & for Debt Mutual Fund Only, Short Term < 3 Years, Long Term > 3 Years. Long Term Capital Gains (LTCG) and Short\n\n     Term Capital Gains (STCG) does not account for Exit Loads.\n\n"
        "9. Interest accrued - all unpaid interest is accrued on FV from the last Interest Payment date.\n\n"
        "10. Bonds & Other listed instruments: Valuation of instruments that are actively traded on the exchange will be shown at the market price. Due to market dependent bid - ask spreads,the availability of the price\n\n      displayed cannot be guaranteed. For debt instruments which are not actively traded on the exchange, valuation will be shown at Face Value.\n\n"
        "11. For Direct equity transactions, balance quantity is adjusted on trade date, while the Demat Service provider may account for it within T + 3 days.\n\n"
        "12. Any case where DP is not with Motilal Oswal Financial Services Limited, all stock trades will be squared off at the end of day.\n\n"
        "13. In Case of a corporate action on a Direct Equity Stock, Holdings in Report might not Tally with DP due to delay in receiving corporate action transactions in DP.\n\n"
        "14. In case of any security held as collateral, this will appear in the client portfolio and might not tally with actual DP holdings.\n\n"
        "15. Unit rate of Stocks acquired in ESOP / bought via other brokers / Off Market Transfers might not tally with actual buying price. Please contact your Advisor to update the same.\n\n"
        "16. Since values are displayed in Lakhs, total of individual rows might not match with Grand Total Row on account of rounding off differences.\n\n"
        "17. In case of IPO issue price declared will be taken as cost.\n\n"
        "18. Arbitrage Funds are classified in the report as Debt based on underlying risk. This differs from SEBI classification.\n\n"
        "19. Benchmark XIRR calculations are done by imitating the cash flows of each security on underlying benchmark (as per benchmark list).\n\n"
        "20. The ledger balances depicted above is held with Motilal Oswal Financial Services Limited: No changes.\n\n"
        "21. The Bank balances depicted above are held with HDFC Bank having Power of Attorney with Motilal Oswal Wealth Limited/HDFC Custody.\n\n"
        "22. We are showing the Bank and Ledger balances August 1st, 2023 onwards. We are not displaying back-dated bank and ledger balances.\n\n"
    )
    
    ax.text(
        0.04, 0.04,
        main_txt,
        horizontalalignment='left',
        fontsize=10,
        color='#2F4F4F'
    )

    footer_text = (
        "This document is not valid without disclosure, please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} | {pd.Timestamp.now().strftime('%I:%M %p')}"
    )
    ax.text(
        0.5, 0.02, footer_text,
        horizontalalignment='center',
        fontsize=10,
        color='#2F4F4F',
        wrap=True
    )
    
    footer = plt.imread('header.png')
    footer_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
    footer_ax.imshow(footer)
    footer_ax.axis('off')

    pdf.savefig(fig, bbox_inches=None, pad_inches=0)
    plt.close(fig)

def create_benchmark_tables_page(pdf):
    fig = plt.figure(figsize=(16, 10))
    ax = fig.add_subplot(111)
    
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(
        gradient,
        extent=(0, 1, 0, 1),
        cmap='Blues',
        aspect='auto',
        alpha=0.3
    )

    logo_img = plt.imread('logo.png')
    logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])  
    logo_ax.imshow(logo_img)
    logo_ax.axis('off')
    
    ax.text(
    0.11, 0.92, "List of Benchmarks used for comparison",
    fontsize=32, color='#8B0000',
    weight='light'
    )
    
    footer = plt.imread('header.png')
    footer_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
    footer_ax.imshow(footer)
    footer_ax.axis('off')

    indices_data = [
        ['List of Indices', ''],
        ['All Direct Equity / Stocks', 'Nifty 50'],
        ['All Bonds', 'CRISIL Composite Bond Fund Index'],
        ['All Structure Products', 'Crisil Short Term Bond Fund Index']
    ]

    mutual_fund_data = [
        ['List of Mutual Fund Indices', ''],
        ['Mutual Fund Categories', 'Benchmark'],
        ['Equity Savings Fund', 'Nifty Equity Saving TRI'],
        ['Thematic Fund', 'Nifty 500 TRI'],
        ['Arbitrage Fund', 'Nifty 50 Arbitrage'],
        ['Banking and PSU', 'CRISIL Short Term Bond Fund Index'],
        ['Corporate Bond', 'Crisil Composite Bond Fund Index'],
        ['Credit Risk', 'Crisil Short Term Bond Fund Index'],
        ['Debt Hybrid Fund', 'Crisil Short Term Bond Fund Index'],
        ['Dynamic Bond Fund', 'Crisil Composite Bond Fund Index'],
        ['ELSS', 'Nifty 500 TRI'],
        ['Equity Hybrid Fund', 'CRISIL Hybrid 35+65 - Aggressive Index'],
        ['Fixed Maturity Fund', 'Crisil Short Term Bond Fund Index'],
        ['Flexi Cap Fund', 'Nifty 500 TRI'],
        ['Floating Rate Fund', 'Crisil Short Term Bond Fund Index'],
        ['GILT Fund', 'CRISIL GILT Index'],
        ['GOLD FUND', 'MCX GOLD SPOT'],
        ['Income Fund', 'Crisil Composite Bond Fund Index'],
        ['International Fund', 'S&P 500 INR'],
        ['Large Cap Fund', 'Nifty 50 TRI'],
        ['Liquid Fund', 'Crisil Liquid Fund Index'],
        ['Long Duration', 'Crisil Composite Bond Fund Index'],
        ['Low Duration', 'Crisil Liquid Fund Index'],
        ['Medium Duration', 'Crisil Composite Bond Fund Index'],
        ['Medium to Long Duration', 'Crisil Composite Bond Fund Index'],
        ['Mid Cap Fund', 'Nifty Midcap 150 TRI'],
        ['Multi Cap Fund', 'Nifty 500 TRI'],
        ['Overnight', 'Crisil Liquid Fund Index'],
        ['Short Duration', 'Crisil Short Term Bond Fund Index'],
        ['Silver Fund', 'MCX Silver SPOT'],
        ['Small Cap Fund', 'Nifty Smallcap 250 TRI'],
        ['Target Maturity Fund', 'Crisil Composite Bond Fund Index'],
        ['Thematic Fund', 'Nifty 500 TRI'],
        ['Ultra Short Duration Fund', 'Crisil Liquid Fund Index'],
    ]

    pms_data = [
        ['List of PMS / AIF Indices', ''],
        ['Schemes', 'Benchmark'],
        ['Ashmore India Opportunities Fund Class B', 'BSE Small Cap'],
        ['ASK Growth Portfolio', 'S&P BSE 500'],
        ['Ask India Vision Portfolio', 'S&P BSE 500'],
        ['ASK India Select Portfolio', 'S&P BSE 500'],
        ['ASK Indian Entrepreneur Portfolio', 'S&P BSE 500'],
        ['Avendus Absolute Return Fund', 'Crisil Short Term Bond Fund Index'],
        ['Avendus Enhanced Return Fund', 'Nifty 50'],
        ['DHFL Pramerica Deep Value Strategy', 'Nifty 500'],
        ['Edelweiss Stressed Troubled Assets Revival Fund Estar', 'Nifty 50'],
        ['India Invest Opportunity -Citi Bank', 'Nifty 50'],
        ['India Opportunities Portfolio Strategy', 'S&P BSE 500'],
        ['India Opportunity Portfolio Strategy V2', 'Nifty Free Float Smallcap 100'],
        ['Invesco India Dawn Portfolio', 'S&P BSE 500'],
        ['Invesco India Rise Portfolio', 'S&P BSE 500'],
        ['Liquid Strategy', 'CRISIL Liquid Fund Index'],
        ['Motilal Oswal Focused Emergence Fund', 'BSE Small Cap'],
        ['Motilal Oswal Focused Multicap Opportunities Fund', 'Nifty 500'],
        ['Next Trillion Dollar Opportunity Strategy', 'S&P BSE 500'],
        ['OBCMPL All Cap Strategy', 'S&P BSE 500'],
        ['OBCMPL Thematic Portfolio', 'Nifty 50'],
        ['Old Bridge Nri Vantage Equity Plan', 'S&P BSE 500'],
        ['Old Bridge Vantage Equity Fund', 'S&P BSE 500'],
        ['Reliance Yield Maximiser All Schemes', 'Crisil Liquid Fund Index + 2%'],
        ['Renaissance India Next Portfolio', 'Nifty 50'],
        ['UTI Structured Debt Opportunities Fund I', 'Crisil Short Term Bond Fund Index'],
        ['Value Strategy', 'Nifty 50'],
        ['Unifi Blend Fund', 'S&P BSE Midcap'],
        ['WO Pioneers PMS', 'S&P BSE 200']
    ]

    indices_table = ax.table(
        cellText=indices_data,
        loc='upper left',
        bbox=[0.05, 0.78, 0.3, 0.1],
        cellLoc='left'
    )
    indices_table.auto_set_font_size(False)
    indices_table.set_fontsize(8)
    
    mutual_fund_table = ax.table(
        cellText=mutual_fund_data,
        loc='upper left',
        bbox=[0.05, 0.06, 0.4, 0.7],
        cellLoc='left'
    )
    mutual_fund_table.auto_set_font_size(False)
    mutual_fund_table.set_fontsize(8)

    pms_table = ax.table(
        cellText=pms_data,
        loc='upper right',
        bbox=[0.55, 0.06, 0.4, 0.8],
        cellLoc='left'
    )
    pms_table.auto_set_font_size(False)
    pms_table.set_fontsize(8)

    for table in [indices_table, mutual_fund_table, pms_table]:
        for cell in table._cells:
            table._cells[cell].set_linewidth(0.5)
            if cell[0] == 0 or (cell[0] == 1 and table != indices_table):
                table._cells[cell].set_facecolor('#D3D3D3')
                table._cells[cell].set_text_props(weight='bold')

    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    footer_text = (
        "This document is not valid without disclosure, please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} | {pd.Timestamp.now().strftime('%I:%M %p')}"
    )
    ax.text(
        0.5, 0.02, footer_text,
        horizontalalignment='center',
        fontsize=10,
        color='#2F4F4F',
        wrap=True
    )

    pdf.savefig(fig, bbox_inches=None, pad_inches=0)
    plt.close(fig)
    
def create_benchmark_tables_page2(pdf):
    fig = plt.figure(figsize=(16, 10))
    ax = fig.add_subplot(111)
    
    ax.set_position([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')

    gradient = np.linspace(1, 0.9, 500).reshape(1, -1)
    ax.imshow(
        gradient,
        extent=(0, 1, 0, 1),
        cmap='Blues',
        aspect='auto',
        alpha=0.3
    )

    logo_img = plt.imread('logo.png')
    logo_ax = fig.add_axes([0.79, 0.9, 0.2, 0.08])  
    logo_ax.imshow(logo_img)
    logo_ax.axis('off')
    
    ax.text(
    0.11, 0.92, "List of Benchmarks used for comparison",
    fontsize=32, color='#8B0000',
    weight='light'
    )
    
    footer = plt.imread('header.png')
    footer_ax = fig.add_axes([0.02, 0.77, 0.9, 0.3])
    footer_ax.imshow(footer)
    footer_ax.axis('off')

    pms_data = [
        ['List of PMS / AIF Indices', ''],
        ['Schemes', 'Benchmark'],
        ['Unifi Blend PMS', 'S&P BSE 500'],
        ['Motilal Oswal Value PMS', 'Nifty 50 TRI'],
        ['Motilal Oswal NTDOC PMS', 'Nifty 500'],
        ['ASK India Select PMS', 'S&P BSE 500'],
        ['ASk IEP PMS', 'S&P BSE 500'],
        ['ASK India Vision', 'S&P BSE 500'],
        ['ASK Indian Entrepreneur Portfolio', 'S&P BSE 500'],
        ['Abakkus All Cap', 'S&P  BSE 200'],
        ['ENAM IDEA', 'Nifty 500'],
        ['Motilal Oswal BOP PMS', 'Nifty 50 TRI'],
        ['Marcellus CC PMS', 'Nifty 50'],
        ['Renaissance Oppurtunities PMS', 'Nifty 50'],
        ['Renaissance India Next Portfolio', 'Nifty 50'],
        ['Invesco India DAWN', 'S&P BSE 500'],
        ['Invesco India RISE', 'S&P BSE 500'],
        ['ASK India 2025', 'S&P BSE 500'],
        ['Renaissance Midcap', 'Nifty Free Float Midcap 100'],
        ['Unifi Blended PMS', 'S&P BSE Midcap'],
        ['Unifi BCAD PMS', 'S&P BSE Midcap'],
        ['Unifi Blend AIF', 'S&P BSE Midcap'],
        ['Motilal Oswal IOP PMS', 'Nifty Small Cap 50 Tri'],
        ['MO EOP 2', 'Nifty 500'],
        ['Marcellus Little Champs', 'S&P BSE 500'],
        ['Old Bridge Long Term Equility', 'Nifty 50'],
        ['Alchemy High Growth Select Stock', 'S&P BSE 500'],
        ['Alchemy High Growth', 'S&P BSE 500'],
        ['MO BAF 2(Anti Fragile)', 'Nifty 500'],
    ]
    
    pms_table = ax.table(
        cellText=pms_data,
        loc='upper right',
        bbox=[0.25, 0.06, 0.4, 0.8],
        cellLoc='left'
    )
    pms_table.auto_set_font_size(False)
    pms_table.set_fontsize(8)

    for table in [pms_table]:
        for cell in table._cells:
            table._cells[cell].set_linewidth(0.5)
            if cell[0] == 0 or (cell[0] == 1 and table != pms_table):
                table._cells[cell].set_facecolor('#D3D3D3')
                table._cells[cell].set_text_props(weight='bold')

    
    current_date = pd.Timestamp.now().strftime('%d-%b-%Y')
    footer_text = (
        "This document is not valid without disclosure, please refer to the last page for the disclaimer. | "
        "Strictly Private & Confidential.\n"
        f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | "
        f"Generated Date & Time : {current_date} | {pd.Timestamp.now().strftime('%I:%M %p')}"
    )
    ax.text(
        0.5, 0.02, footer_text,
        horizontalalignment='center',
        fontsize=10,
        color='#2F4F4F',
        wrap=True
    )

    pdf.savefig(fig, bbox_inches=None, pad_inches=0)
    plt.close(fig)

def rearrange_and_add_total(portfolio_value_df, debt_df, bond_df):
    portfolio_value_df["Portfolio Value"] = pd.to_numeric(
        portfolio_value_df["Portfolio Value"], errors="coerce"
    )
    portfolio_value_df["Portfolio Value"] = portfolio_value_df["Portfolio Value"].round().astype(int)
    
    order = ["Available Cash", "Debt", "Equity", "Gold"]
    ordered_df = portfolio_value_df[portfolio_value_df["Portfolio Component"].isin(order)].copy()
    ordered_df["Order"] = ordered_df["Portfolio Component"].map(lambda x: order.index(x))
    ordered_df = ordered_df.sort_values(by="Order").drop(columns=["Order"])
    
    grand_total_value = ordered_df["Portfolio Value"].sum()
    
    cash_components = ["Available Cash", "Debt", "Gold"]
    cash_equivalent_value = ordered_df[
        ordered_df["Portfolio Component"].isin(cash_components)
    ]["Portfolio Value"].sum()
    
    
    cash_equivalent_percentage = round(
        (cash_equivalent_value / grand_total_value) * 100
    ) if grand_total_value != 0 else 0
        
    equity_value = ordered_df.loc[
        ordered_df["Portfolio Component"] == "Equity", "Portfolio Value"
    ].sum()
    equity_allocation_percentage = round(
        (equity_value / grand_total_value) * 100
    ) if grand_total_value != 0 else 0
    
    grand_total_row = {
        "Portfolio Component": "Grand Total",
        "Portfolio Value": f"{grand_total_value:,}",
    }
    result_df = pd.concat([ordered_df, pd.DataFrame([grand_total_row])], ignore_index=True)
    
    return (
        result_df,
        cash_equivalent_value,  
        cash_equivalent_percentage,
        equity_allocation_percentage,
    )

def preprocess_table(dataframe, category_name):
    dataframe = dataframe.sort_values(by='Market Value', ascending=False)
    numeric_columns = ['Quantity', 'Buy Price', 'CMP', 'PandL', 'Market Value']
    sums = dataframe[numeric_columns].sum()
    dataframe['Category'] = category_name

    total_row = pd.DataFrame(
        {col: [sums[col] if col in sums else ''] for col in dataframe.columns}
    )
    total_row['Category'] = f"{category_name} Total"
    result_df = pd.concat([dataframe, total_row], ignore_index=True)

    return result_df, sums

def create_portfolio_reports(data, portfolio_dir):
    try:
        Portfolio_Value = data['Portfolio Value']
        Holding = data['Holding']
        xirr = data['XIRR']
        equity = data['Equity']
        debt = data['Debt']
        bond = data.get('bond', pd.DataFrame())
        fno = data['FNO']
        profit = data['Profits']
        
        customer_name, ucid = get_customer_details(Holding)
        
        Portfolio_Value_Modified, cash_equivalent_value, cash_equivalent_percentage, equity_allocation_percentage = rearrange_and_add_total(
            Portfolio_Value, debt, bond
        )
        
        portfolio_dir = Path(portfolio_dir)
        
        with PdfPages(portfolio_dir / 'portfolio_report.pdf') as pdf:
            create_cover_page(pdf, customer_name, ucid)

            fig1 = plot_table_and_pie(
                Portfolio_Value_Modified,
                Holding,
                xirr,
                cash_equivalent_value,
                cash_equivalent_percentage,
                equity_allocation_percentage,
            )
            pdf.savefig(fig1, bbox_inches='tight')
            plt.close(fig1)
            
            target_stocks = [ "Bajaj Finserv Ltd","Central Depository Services (India) Ltd","HDFC Bank Ltd","IDFC First Bank Ltd","Kotak Mahindra Bank Ltd","Shriram Finance Ltd","Bharat Forge Ltd","Tata Consultancy Services",]
            
            fig2 = create_combined_portfolio_table(Holding, portfolio_type='equity', target_stocks=target_stocks, label="Direct Equity")
            pdf.savefig(fig2, bbox_inches ='tight')
            plt.close(fig2)
            
            target_etf = [ "Nippon India ETF Nifty IT", "Nippon India ETF Hang Seng Bees"]
            
            fig8 = create_combined_portfolio_table(Holding, portfolio_type='equity', target_stocks=target_etf, label="ETF")
            pdf.savefig(fig8, bbox_inches ='tight')
            plt.close(fig8)
            
            fig4 = eqmf(Holding)
            pdf.savefig(fig4, bbox_inches='tight')
            plt.close(fig4)
            
            fig6 = dmf(Holding)
            pdf.savefig(fig6, bbox_inches='tight')
            plt.close(fig6)

    except Exception as e:
        print(f"Error generating reports: {e}")