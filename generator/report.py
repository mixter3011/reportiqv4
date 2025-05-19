import io
import pandas as pd
import locale
try:
    locale.setlocale(locale.LC_ALL, 'en_IN.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, '')
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageTemplate, Frame, Image, BaseDocTemplate, PageBreak, Flowable, KeepTogether

def background(canvas, doc):
    logo_path = "/Users/sen/Desktop/reportiqv4/logo.png"
    logo_width = 1.5 * inch
    logo_height = 0.5 * inch
    x1 = doc.pagesize[0] - doc.rightMargin - logo_width + 0.9 * inch
    y1 = doc.pagesize[1] - doc.topMargin + 0.4 * inch
    
    canvas.saveState()
    canvas.setFillColor(white)  
    canvas.rect(0, 0, letter[0], letter[1], fill=1)
    canvas.restoreState()
    canvas.drawImage(logo_path, x1, y1, width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')

def cover (code, name):
    cover_page = f"{name}.pdf"
    doc = BaseDocTemplate(cover_page, pagesize=letter)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height + 0.8*inch)
    page = PageTemplate(id='FirstPage', frames=frame, onPage=background)
    doc.addPageTemplates([page])
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=HexColor('#3C3EA8')
    )
    
    header_style = ParagraphStyle(
        'CenteredHeaderStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    
    normal_style = ParagraphStyle(
        'CenteredNormalStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=14,
        alignment=TA_CENTER
    )

    content = []
    content.append(Paragraph("CUSTOMER STATEMENT", title_style))
    content.append(Spacer(1, 0.01*inch))    
    content.append(Spacer(1, 3*inch))
    content.append(Paragraph(f"{name}", header_style))
    content.append(Paragraph(f"{code}", normal_style))
    content.append(Spacer(1, 2*inch))
    
    return doc, content

def overview(direct_equity_market_value, etf_equity_market_value, debt_etf_market_value, 
             equity_mf_market_value, debt_mf_market_value, bond_market_value, df2, xirr_value=None):
        
    client_name = ""
    client_code = ""
    try:
        for col in df2.columns:
            if 'name' in col.lower():
                client_name = df2[col].iloc[0]
            if 'code' in col.lower() or 'id' in col.lower():
                client_code = df2[col].iloc[0]
    except (IndexError, ValueError, AttributeError):
        pass
    
    available_cash = 0
    try:
        balance_col = None
        for col in df2.columns:
            if col.upper() == 'BALANCE':
                balance_col = col
                break
                
        if balance_col is not None and len(df2) > 0:
            for value in df2[balance_col]:
                if pd.notna(value) and str(value).strip():
                    try:
                        available_cash = float(str(value).replace(',', ''))
                        break
                    except (ValueError, TypeError):
                        pass
                        
    except (IndexError, ValueError, AttributeError, KeyError) as e:
        print(f"Error accessing balance: {e}")
        available_cash = 0
    
    equity_total = direct_equity_market_value + etf_equity_market_value + equity_mf_market_value
    gold_total = bond_market_value
    debt_total = debt_etf_market_value + debt_mf_market_value
    
    cash_equivalent = debt_total + gold_total + available_cash
    
    total_portfolio_value = equity_total + debt_total + gold_total + available_cash
    
    cash_equivalent_percent = (cash_equivalent / total_portfolio_value * 100) if total_portfolio_value > 0 else 0
    equity_allocation_percent = (equity_total / total_portfolio_value * 100) if total_portfolio_value > 0 else 0

    styles = getSampleStyleSheet()
    
    heading = ParagraphStyle(
        'CenteredHeaderStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=18,
        alignment=TA_LEFT,
        spaceAfter=10,
        textColor=white,
        backColor=HexColor('#1e4388')
    )
    
    client_style = ParagraphStyle(
        'ClientStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        alignment=TA_CENTER
    )
    
    content = []
    content.append(PageBreak()) 
    
    class HorizontalLineFlowable(Flowable):
        def __init__(self, width=10, height=3):
            Flowable.__init__(self)
            self.width = width
            self.height = height
            
        def draw(self):
            self.canv.setStrokeColor(HexColor('#1e4388'))
            self.canv.setLineWidth(self.height)
            self.canv.line(5, 0, self.width, 0)
    
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        alignment=TA_CENTER,
        textColor=white,
        backColor=HexColor('#4d7cc3')
    )
    
    subheader_style = ParagraphStyle(
        'SubHeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        alignment=TA_LEFT,
        textColor=white,
        backColor=HexColor('#5f92d2')
    )
    
    if client_name and client_code:
        client_name_paragraph = Paragraph(f"{client_name}", client_style)
        client_code_paragraph = Paragraph(f"{client_code}", client_style)
        client_info = [[client_name_paragraph], [client_code_paragraph]]
        client_table = Table(client_info, colWidths=[8*inch]) 
        client_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        content.append(client_table)
    else:
        content.append(Paragraph("Holding Summary & Performance", heading))
    
    content.append(Spacer(1, 0.3*inch))
    
    composition_data = []
    
    composition_data.append([
        Paragraph("Portfolio Value", header_style), 
        Paragraph(ist(total_portfolio_value), header_style)
    ])
    
    if equity_total > 0:
        composition_data.append(["Equity", ist(equity_total)])
    if debt_total > 0:
        composition_data.append(["Debt", ist(debt_total)])
    if gold_total > 0:
        composition_data.append(["Gold", ist(gold_total)])
    if available_cash > 0:
        composition_data.append(["Available Cash", ist(available_cash)])
    
    composition_table = Table(composition_data, colWidths=[1.9*inch, 1.9*inch])
    composition_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#4d7cc3')),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, HexColor('#d6e1e8')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 0.2*inch),
        ('RIGHTPADDING', (1, 0), (1, -1), 0.2*inch),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('TOPPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, HexColor('#d6e1e8')),
    ]))
    
    cash_data = [
        ["Cash Equivalent:", ist(cash_equivalent)],
        ["Cash Equivalent %:", f"{cash_equivalent_percent:.0f}%"],
        ["Equity Allocation %:", f"{equity_allocation_percent:.0f}%"],
    ]
    
    if xirr_value is not None:
        cash_data.append(["XIRR:", f"{xirr_value:.0f}%"])
    
    cash_table = Table(cash_data, colWidths=[1.9*inch, 1.9*inch])
    cash_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 0.2*inch),
        ('RIGHTPADDING', (1, 0), (1, -1), 0.2*inch),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('TOPPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, HexColor('#d6e1e8')),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, HexColor('#d6e1e8')),
    ]))
    
    pie_title = Paragraph("Portfolio Composition", header_style)
    
    pie_title_table = Table([[pie_title]], colWidths=[8*inch])
    pie_title_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), HexColor('#4d7cc3')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('TOPPADDING', (0, 0), (-1, -1), 0.1*inch),
    ]))
    
    content.append(pie_title_table)
    
    try:
        labels = []
        sizes = []
        colors = []
        light_blue_shades = ['#70a3e0', '#81b4ed', '#92c5fa', '#a3d6ff', '#b4e7ff',
                           '#c5f8ff', '#d6ffff', '#e7ffff', '#f8ffff']
        color_index = 0
        
        if equity_total > 0:
            labels.append('Equity')
            sizes.append(equity_total)
            colors.append(light_blue_shades[color_index % len(light_blue_shades)])
            color_index += 1
        
        if debt_total > 0:
            labels.append('Debt')
            sizes.append(debt_total)
            colors.append(light_blue_shades[color_index % len(light_blue_shades)])
            color_index += 1
        
        if gold_total > 0:
            labels.append('Gold')
            sizes.append(gold_total)
            colors.append(light_blue_shades[color_index % len(light_blue_shades)])
            color_index += 1
        
        if available_cash > 0:
            labels.append('Available Cash')
            sizes.append(available_cash)
            colors.append(light_blue_shades[color_index % len(light_blue_shades)])
            color_index += 1
        
        if sizes:  
            plt.figure(figsize=(7, 6), facecolor='none')
            
            plt.pie(sizes, labels=None, colors=colors, 
                   autopct='%1.0f%%', startangle=90,
                   wedgeprops=dict(width=0.7, edgecolor='w'))
            
            center_circle = plt.Circle((0,0), 0.35, fc='#D6E1E8')
            plt.gca().add_patch(center_circle)
            
            legend = plt.legend(labels, loc="center right", bbox_to_anchor=(1.2, 0.5))
            
            plt.axis('equal')
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=300, bbox_inches='tight', transparent=True)
            buf.seek(0)
            
            img = Image(buf, width=6*inch, height=4*inch)
            chart_table = Table([[img]], colWidths=[8*inch])
            chart_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0.1*inch),
                ('TOPPADDING', (0, 0), (-1, -1), 0.1*inch),
            ]))
            content.append(chart_table)
            plt.close()
            
    except Exception as e:
        error_style = ParagraphStyle('ErrorStyle', parent=styles['Normal'], textColor=HexColor('#3C3EA8'))
        content.append(Paragraph(f"Unable to generate pie chart: {str(e)}", error_style))
    
    content.append(Spacer(1, 0.3*inch))

    tables_data = [[composition_table, cash_table]]
    tables_layout = Table(tables_data, colWidths=[4*inch, 4*inch])
    tables_layout.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.1*inch),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.1*inch),
    ]))
    
    content.append(tables_layout)
    
    return content

def deq(direct_equity, direct_equity_total, etf_equity, etf_equity_total, equity_mf, equity_mf_total):
    styles = getSampleStyleSheet()
    
    page_title_style = ParagraphStyle(
        'PageTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=HexColor('#3C3EA8')
    )
    
    section_title_style = ParagraphStyle(
        'SectionTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_LEFT,
        spaceAfter=6,
        leftIndent=-60,
        textColor=HexColor('#3C3EA8')
    )
    
    table_header_style = styles.add(ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        alignment=TA_CENTER,
        textColor=white
    ))
    
    def trim_etf_name(name, max_words=2):
        if not isinstance(name, str):
            return str(name)
        name = name.replace("NIPPON LIFE INDIA AM LTD#", "").strip()
        words = name.split()
        if len(words) <= max_words:
            return name
        return ' '.join(words[:max_words]) + '...'
    
    def trim_mf_name(name, max_words=3):
        if not isinstance(name, str):
            return str(name)
        words = name.split()
        if len(words) <= max_words:
            return name
        return ' '.join(words[:max_words]) + '...'
    
    def create_table(data, total_data, title_text, column_headers, is_mf=False):
        column_headers = [
            'Instrument Name', 
            'Quantity', 
            'Buy Price (Sum)', 
            'CMP (Sum)', 
            'P&L',  
            'Market Value'
        ]
        title = Paragraph(title_text, section_title_style)
        
        headers = [Paragraph(header, table_header_style) for header in column_headers]
        
        table_data = [headers]
        
        for _, row in data.iterrows():
            if is_mf:
                scheme_name = trim_mf_name(row['Unnamed: 1'])
                table_data.append([
                    scheme_name,  
                    ist(float(row['Unnamed: 2']), 0) if row['Unnamed: 2'] else '',
                    ist(float(row['Unnamed: 3']), 0).lstrip(',') if row['Unnamed: 3'] else '',
                    ist(float(row['Unnamed: 5']), 0).lstrip(',') if row['Unnamed: 5'] else '',
                    ist(float(row['Unnamed: 12']), 0).lstrip(',') if row['Unnamed: 12'] else '',
                    ist(float(row['Unnamed: 6']), 0) if row['Unnamed: 6'] else ''
                ])
            else:
                instrument_name = trim_etf_name(row['Unnamed: 0'])
                table_data.append([
                    instrument_name,  
                    ist(float(row['Unnamed: 1']), 0) if row['Unnamed: 1'] else '',
                    ist(float(row['Unnamed: 2']), 0).lstrip(',') if row['Unnamed: 2'] else '',
                    ist(float(row['Unnamed: 4']), 0).lstrip(',') if row['Unnamed: 4'] else '',
                    ist(float(row['Unnamed: 10']), 0).lstrip(',') if row['Unnamed: 10'] else '',
                    ist(float(row['Unnamed: 5']), 0) if row['Unnamed: 5'] else ''
                ])
        
        total_row = ['Total:'] + [ist(float(val), 0) if i > 0 and isinstance(val, (int, float)) else val for i, val in enumerate(total_data[1:])]
        table_data.append(total_row)
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#3C3EA8')),
            ('BACKGROUND', (0, -1), (-1, -1), HexColor('#3C3EA8')),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('TEXTCOLOR', (0, -1), (-1, -1), white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER')  
        ])
        
        column_widths = [
            2.2 * inch,  
            1 * inch,    
            1.2 * inch,  
            1.2 * inch,  
            1.2 * inch,  
            1.2 * inch   
        ]
        
        table = Table(table_data, colWidths=column_widths)
        table.setStyle(table_style)
        
        spacer = Spacer(1, 5)
        
        return [title, spacer, table]
    
    page_elements = [PageBreak()]
    
    page_elements.append(Paragraph("Detailed Holdings & Performance", page_title_style))
    page_elements.append(Spacer(1, 0.1*inch))
    
    sections = []
    if not direct_equity.empty:
        sections.append(("Direct Equity", direct_equity, direct_equity_total, 
                        ['Instrument Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], False))
    
    if not etf_equity.empty:
        sections.append(("Equity ETF", etf_equity, etf_equity_total, 
                        ['ETF Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], False))
    
    if not equity_mf.empty:
        sections.append(("Equity Mutual Funds", equity_mf, equity_mf_total, 
                        ['Scheme Name', 'Units', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], True))
    
    for i, (title, data, totals, headers, is_mf) in enumerate(sections):
        table_elements = create_table(data, totals, title, headers, is_mf)
        page_elements.extend(table_elements)
        
        if i < len(sections) - 1:
            page_elements.append(Spacer(1, 0.05*inch))  
    
    return page_elements

def deb(debt_etf, debt_etf_total, debt_mf, debt_mf_total, bond_data, bond_total):
    styles = getSampleStyleSheet()
    
    page_title_style = ParagraphStyle(
        'PageTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=HexColor('#3C3EA8')
    )
    
    section_title_style = ParagraphStyle(
        'SectionTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_LEFT,
        spaceAfter=6,
        leftIndent=-60,
        textColor=HexColor('#3C3EA8')
    )
    
    table_header_style = styles.add(ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        alignment=TA_CENTER,
        textColor=white
    ))
    
    def trim_name(name, max_words=3):
        if not isinstance(name, str):
            return str(name)
        words = name.split()
        if len(words) <= max_words:
            return name
        return ' '.join(words[:max_words]) + '...'
    
    def create_table(data, total_data, title_text, column_headers, is_mf=False, is_bond=False):
        title = Paragraph(title_text, section_title_style)
        
        headers = [Paragraph(header, table_header_style) for header in column_headers]
        
        table_data = [headers]
        
        for _, row in data.iterrows():
            if is_mf:
                scheme_name = trim_name(row['Unnamed: 1'])
                table_data.append([
                    scheme_name,  
                    ist(float(row['Unnamed: 2']), 0) if row['Unnamed: 2'] else '',
                    ist(float(row['Unnamed: 3']), 0).lstrip(',') if row['Unnamed: 3'] else '',
                    ist(float(row['Unnamed: 5']), 0).lstrip(',') if row['Unnamed: 5'] else '',
                    ist(float(row['Unnamed: 12']), 0).lstrip(',') if row['Unnamed: 12'] else '',
                    ist(float(row['Unnamed: 6']), 0) if row['Unnamed: 6'] else ''
                ])
            elif is_bond:
                bond_name = trim_name(row['Unnamed: 0'])
                table_data.append([
                    bond_name,  
                    ist(float(row['Unnamed: 1']), 0).lstrip(',') if row['Unnamed: 1'] else '',
                    ist(float(row['Unnamed: 2']), 0).lstrip(',') if row['Unnamed: 2'] else '',
                    ist(float(row['Unnamed: 4']), 0).lstrip(',') if row['Unnamed: 4'] else '',
                    ist(float(row['Unnamed: 10']), 0).lstrip(',') if row['Unnamed: 10'] else '',
                    ist(float(row['Unnamed: 5']), 0) if row['Unnamed: 5'] else ''
                ])
            else:
                instrument_name = trim_name(row['Unnamed: 0'])
                table_data.append([
                    instrument_name,  
                    ist(float(row['Unnamed: 1']), 0) if row['Unnamed: 1'] else '',
                    ist(float(row['Unnamed: 2']), 0).lstrip(',') if row['Unnamed: 2'] else '',
                    ist(float(row['Unnamed: 4']), 0).lstrip(',') if row['Unnamed: 4'] else '',
                    ist(float(row['Unnamed: 10']), 0).lstrip(',') if row['Unnamed: 10'] else '',
                    ist(float(row['Unnamed: 5']), 0) if row['Unnamed: 5'] else ''
                ])
        
        total_row = ['Total:'] + [ist(float(val), 0).lstrip(',') if i > 0 and isinstance(val, (int, float)) else val for i, val in enumerate(total_data[1:])]
        table_data.append(total_row)
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#3C3EA8')),
            ('BACKGROUND', (0, -1), (-1, -1), HexColor('#3C3EA8')),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('TEXTCOLOR', (0, -1), (-1, -1), white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER')  
        ])
        
        column_widths = [
            2.2 * inch,  
            1 * inch,    
            1.2 * inch,  
            1.2 * inch,  
            1.2 * inch,  
            1.2 * inch   
        ]
        
        table = Table(table_data, colWidths=column_widths)
        table.setStyle(table_style)
        
        spacer = Spacer(1, 10)
        
        return [title, spacer, table, spacer, spacer]
    
    page_elements = [PageBreak()]
    
    page_elements.append(Paragraph("Debt Holdings & Performance", page_title_style))
    page_elements.append(Spacer(1, 0.1*inch))
    
    sections = []
    if not debt_etf.empty:
        sections.append(("Debt ETF", debt_etf, debt_etf_total, 
                        ['ETF Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], False))
    
    if not debt_mf.empty:
        sections.append(("Debt Mutual Funds", debt_mf, debt_mf_total, 
                        ['Scheme Name', 'Units', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], True))
    
    if not bond_data.empty:
        sections.append(("Bonds", bond_data, bond_total, 
                        ['Bond Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L', 'Market Value'], False, True))
    
    for i, (title, data, totals, headers, is_mf, *args) in enumerate(sections):
        is_bond = args[0] if args else False
        table_elements = create_table(data, totals, title, headers, is_mf, is_bond)
        page_elements.extend(table_elements)
        
        if i < len(sections) - 1:
            page_elements.append(Spacer(1, 0.01*inch))
    
    return page_elements

def ist(number, decimal_places=0):
    try:
        num = float(number)
        num = round(num, decimal_places)
        
        negative = num < 0
        num = abs(num)
        
        if decimal_places == 0:
            num = int(num)
        
        int_part = str(int(num))
        
        result = ""
        if len(int_part) <= 3:
            result = int_part
        else:
            result = int_part[-3:]
            remaining = int_part[:-3]
            i = len(remaining)
            while i > 0:
                if i >= 2:
                    result = remaining[i-2:i] + "," + result
                else:
                    result = remaining[i-1:i] + "," + result
                i -= 2
        
        if decimal_places > 0:
            decimal_part = f"{num % 1:.{decimal_places}f}"[2:]
            result = result + "." + decimal_part
        
        if negative:
            result = "-" + result
            
        return result
    
    except Exception as e:
        return str(number)

def report_gen(df1, df2, df3=None, output_path=None):
    info_row = df1[df1['Unnamed: 0'] == 'Client Equity Code/UCID/Name'].index[0]
    c_info = str(df1.iloc[info_row, 1]).strip()
    parts = c_info.split('/')
    c_code = parts[0].strip()  
    c_name = parts[-1].strip()
    
    equity_row = df1[df1['Unnamed: 0'] == 'Equity:-'].index[0]
    mf_row = df1[df1['Unnamed: 0'] == 'Mutual Fund:-'].index[0]
    fno_row = df1[df1['Unnamed: 0'] == 'FnO:-'].index[0]
    bond_row = df1[df1['Unnamed: 0'] == 'Bond:-'].index[0]
    
    equity_header = df1.iloc[equity_row + 1].tolist()
    mf_header = df1.iloc[mf_row + 1].tolist()
    bond_header = df1.iloc[bond_row + 1].tolist()
    
    equity_end = mf_row - 4  
    mf_end = fno_row - 4
    
    bond_end = len(df1)
    for i in range(bond_row + 2, len(df1)):
        if i >= len(df1) or pd.isna(df1.iloc[i, 0]) or df1.iloc[i, 0] == '':
            bond_end = i
            break
    
    equity_data = df1.iloc[equity_row + 2:equity_end].copy()
    mf_data = df1.iloc[mf_row + 2:mf_end].copy()
    bond_data = df1.iloc[bond_row + 2:bond_end].copy()
    
    equity_data = equity_data[equity_data['Unnamed: 0'] != 'Total:']
    mf_data = mf_data[mf_data['Unnamed: 0'] != 'Total:']
    bond_data = bond_data[bond_data['Unnamed: 0'] != 'Total:']
    
    direct_equity = equity_data[~equity_data['Unnamed: 0'].str.contains('ETF', na=False)]
    direct_equity = direct_equity[~direct_equity['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]
    
    etf_equity = equity_data[equity_data['Unnamed: 0'].str.contains('ETF', na=False)]
    etf_equity = etf_equity[~etf_equity['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]
    
    debt_etf = equity_data[equity_data['Unnamed: 0'].str.contains('Nifty 1D Rate Liquid BeES', na=False)]
    
    equity_mf = mf_data[mf_data['Unnamed: 0'] == 'Equity']
    debt_mf = mf_data[mf_data['Unnamed: 0'] == 'Debt']
    
    def calculate_totals(data, cols_to_keep):
        totals = ['Total:']
        for col_idx, col in enumerate(cols_to_keep[1:], 1):  
            try:
                values = data[f'Unnamed: {col}']
                values = pd.to_numeric(values, errors='coerce')
                
                if not values.isna().all():
                    total = round(values.sum(), 2)
                    totals.append(total)
                else:
                    totals.append('')
            except (ValueError, TypeError):
                totals.append('')
        return totals
    
    equity_cols_to_keep = [0, 1, 2, 4, 5, 10]  
    mf_cols_to_keep = [1, 2, 3, 5, 6, 12]      
    bond_cols_to_keep = [0, 1, 2, 4, 5, 10]   
    
    direct_equity_total = calculate_totals(direct_equity, equity_cols_to_keep)
    etf_equity_total = calculate_totals(etf_equity, equity_cols_to_keep)
    debt_etf_total = calculate_totals(debt_etf, equity_cols_to_keep)
    equity_mf_total = calculate_totals(equity_mf, mf_cols_to_keep)
    debt_mf_total = calculate_totals(debt_mf, mf_cols_to_keep)
    bond_total = calculate_totals(bond_data, bond_cols_to_keep)
    
    direct_equity_market_value = direct_equity_total[4] if len(direct_equity_total) > 4 and direct_equity_total[4] != '' else 0
    etf_equity_market_value = etf_equity_total[4] if len(etf_equity_total) > 4 and etf_equity_total[4] != '' else 0
    debt_etf_market_value = debt_etf_total[4] if len(debt_etf_total) > 4 and debt_etf_total[4] != '' else 0
    equity_mf_market_value = equity_mf_total[4] if len(equity_mf_total) > 4 and equity_mf_total[4] != '' else 0
    debt_mf_market_value = debt_mf_total[4] if len(debt_mf_total) > 4 and debt_mf_total[4] != '' else 0
    bond_market_value = bond_total[4] if len(bond_total) > 4 and bond_total[4] != '' else 0
    
    try:
        df2['client_name'] = c_name
        df2['client_code'] = c_code
    except Exception:
        pass
    
    xirr_value = None
    if df3 is not None and not df3.empty:
        try:
            last_row = df3.iloc[-1]
            xirr_value = last_row.iloc[2] * 100  
        except (IndexError, ValueError, TypeError, AttributeError) as e:
            print(f"Error extracting XIRR value: {str(e)}")
    
    if output_path is None:
        output_path = f"{c_name}.pdf"
    
    doc, cover_page = cover(c_code, c_name)
    overview_page = overview(direct_equity_market_value, etf_equity_market_value, debt_etf_market_value, 
                            equity_mf_market_value, debt_mf_market_value, bond_market_value, df2, xirr_value)
    direct_equity_page = deq(direct_equity, direct_equity_total, etf_equity, etf_equity_total, equity_mf, equity_mf_total)
    debt_page = deb(debt_etf, debt_etf_total, debt_mf, debt_mf_total, bond_data, bond_total)
    pdf_content = cover_page + overview_page + direct_equity_page + debt_page
    
    doc = BaseDocTemplate(output_path, pagesize=letter)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height + 0.8*inch)
    page = PageTemplate(id='FirstPage', frames=frame, onPage=background)
    doc.addPageTemplates([page])
    
    doc.build(pdf_content)
    
    return output_path