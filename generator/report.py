import os
import io
import numpy as np
import pandas as pd
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as mpatches
import matplotlib.image as mpimg
from matplotlib.table import Table as MplTable

from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white, black, grey
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageTemplate, Frame, Image, BaseDocTemplate, NextPageTemplate, PageBreak, Flowable

def background(canvas, doc):
    logo_path = "/Users/sen/Desktop/reportiqv4/logo.png"
    footer_path = "/Users/sen/Desktop/reportiqv4/footer.png"
    logo_width = 1.5 * inch
    logo_height = 0.5 * inch
    footer_width = 5.2*inch
    footer_height = 0.3*inch
    x1 = doc.pagesize[0] - doc.rightMargin - logo_width + 0.9 * inch
    y1 = doc.pagesize[1] - doc.topMargin + 0.4 * inch
    x2 = doc.pagesize[0] - doc.rightMargin - footer_width + 1 * inch
    y2 = 0.4 * inch + footer_height + 10  
    
    x_left_text = x2 - 2.8 * inch  
    x_right_text = x2 - 3 * inch
    
    text_x = doc.leftMargin - 0.2 * inch
    text_y = 0.4 * inch
    
    bottom_text_1 = "This document is not valid without disclosure, please refer to the last page for the disclaimer. | Strictly Private & Confidential."
    bottom_text_2 = f"Incase of any query / feedback on the report, please write to query@motilaloswal.com. | Generated Date & Time : {datetime.now().strftime('%d-%b-%Y | %I:%M %p')}"
    
    canvas.saveState()
    canvas.setFillColor(HexColor('#D6E1E8'))  
    canvas.rect(0, 0, letter[0], letter[1], fill=1)
    canvas.restoreState()
    canvas.drawImage(logo_path, x1, y1, width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')
    canvas.drawImage(footer_path, x2, y2, width=footer_width, height=footer_height, preserveAspectRatio=True, mask='auto')
    canvas.drawString(x_left_text, y2 + footer_height / 2, "WINNING PORTFOLIOS")
    canvas.setFont("Helvetica", 8)
    canvas.drawString(text_x + 0.4*inch, text_y + 12, bottom_text_1)
    canvas.drawString(text_x, text_y, bottom_text_2)
    powered_text = "POWERED BY KNOWLEDGE"
    canvas.setFont("Helvetica-Bold", 12)

    text_width = canvas.stringWidth(powered_text, "Helvetica-Bold", 8)
    text_height = 12  

    text_x = x_right_text
    text_y = y2 + footer_height / 2 - 0.2 * inch

    canvas.setFillColor(HexColor("#990000"))
    canvas.rect(text_x, text_y - 2, text_width + 58, text_height, stroke=0, fill=1)

    canvas.setFillColor(white)
    canvas.drawString(text_x + 1, text_y, powered_text)
    
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
        textColor=HexColor('#990000')
    )
    
    header_style = ParagraphStyle(
        'CenteredHeaderStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    
    sub_style = ParagraphStyle(
        'NormalStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        alignment=TA_CENTER
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
    content.append(Paragraph(f"Report Level : Member | Report Period : Since Inception to {datetime.now().strftime('%d-%b-%Y')}", sub_style))    
    content.append(Spacer(1, 3*inch))
    content.append(Paragraph(f"{name}", header_style))
    content.append(Paragraph(f"UCID : {code}", normal_style))
    content.append(Spacer(1, 2*inch))
    footer_style = ParagraphStyle(
        'FooterStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        alignment=TA_LEFT
    )
    
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
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=HexColor('#3C3EA8')
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
            self.canv.setStrokeColor(HexColor('#3C3EA8'))
            self.canv.setLineWidth(self.height)
            self.canv.line(5, 0, self.width, 0)
            
    header_line = HorizontalLineFlowable(6*inch)  
    
    if client_name and client_code:
        client_name_paragraph = Paragraph(f"{client_name}", heading)
        client_code_paragraph = Paragraph(f"{client_code}", heading)
        client_info = [[client_name_paragraph], [client_code_paragraph]]
        client_table = Table(client_info, colWidths=[8*inch]) 
        client_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        content.append(client_table)
    else:
        content.append(Paragraph("Holding Summary & Performance", heading))
    
    content.append(Spacer(1, 0.1*inch))
    content.append(header_line)
    content.append(Spacer(1, 0.5*inch))
    
    composition_data = []
    composition_data.append(["Portfolio Value", f"{total_portfolio_value:,.0f}"])
    
    if equity_total > 0:
        composition_data.append(["Equity", f"{equity_total:,.0f}"])
    if debt_total > 0:
        composition_data.append(["Debt", f"{debt_total:,.0f}"])
    if gold_total > 0:
        composition_data.append(["Gold", f"{gold_total:,.0f}"])
    if available_cash > 0:
        composition_data.append(["Available Cash", f"{available_cash:,.0f}"])
    
    composition_table = Table(composition_data, colWidths=[2*inch, 2*inch])
    composition_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#D6E1E8')), 
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 14),
        ('GRID', (0, 0), (-1, -1), 1, HexColor('#D6E1E8')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 0.1*inch),
        ('RIGHTPADDING', (1, 0), (1, -1), 0.1*inch),
        ('FONTSIZE', (0, 0), (-1, -1), 14),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, HexColor('#D6E1E8')),  
    ]))
    
    cash_data = [
        ["Cash Equivalent:", f"{cash_equivalent:,.0f}"],
        ["Cash Equivalent %:", f"{cash_equivalent_percent:.2f}%"],
        ["Equity Allocation %:", f"{equity_allocation_percent:.2f}%"],
    ]
    
    # Add XIRR if available
    if xirr_value is not None:
        cash_data.append(["XIRR:", f"{xirr_value:.2f}%"])
    
    cash_table = Table(cash_data, colWidths=[2*inch, 2*inch])
    cash_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, -1), 0.1*inch),
        ('RIGHTPADDING', (1, 0), (1, -1), 0.1*inch),
        ('FONTSIZE', (0, 0), (-1, -1), 14),
    ]))
    
    left_data = [[composition_table], [Spacer(1, 0.5*inch)], [cash_table]]
    left_col = Table(left_data, colWidths=[4*inch])
    left_col.setStyle(TableStyle([
        ('VALIGN', (0, 0), (0, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (0, -1), 0),
        ('RIGHTPADDING', (0, 0), (0, -1), 0),
        ('TOPPADDING', (0, 0), (0, -1), 0),
        ('BOTTOMPADDING', (0, 0), (0, -1), 0),
    ]))
    
    class VerticalLineFlowable(Flowable):
        def __init__(self, height):
            Flowable.__init__(self)
            self.height = height
            
        def draw(self):
            self.canv.setStrokeColor(HexColor('#000000'))
            self.canv.line(15, 15, 15, self.height)
    
    divider = VerticalLineFlowable(3.3*inch)
    
    right_content = []
    pie_title = Paragraph("Portfolio Composition", title_style)
    right_content.append(pie_title)
    
    try:
        labels = []
        sizes = []
        colors = []
        
        if equity_total > 0:
            labels.append('Equity')
            sizes.append(equity_total)
            colors.append('#00FFFF')  
        
        if debt_total > 0:
            labels.append('Debt')
            sizes.append(debt_total)
            colors.append('#90EE90')  
        
        if gold_total > 0:
            labels.append('Gold')
            sizes.append(gold_total)
            colors.append('#FFCC99')  
        
        if available_cash > 0:
            labels.append('Available Cash')
            sizes.append(available_cash)
            colors.append('#FFFFC5')  
        
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
            
            img = Image(buf, width=4*inch, height=3*inch)
            right_content.append(img)
            plt.close()
            
    except Exception as e:
        error_style = ParagraphStyle('ErrorStyle', parent=styles['Normal'], textColor=HexColor('#990000'))
        right_content.append(Paragraph(f"Unable to generate pie chart: {str(e)}", error_style))
    
    right_col = []
    for item in right_content:
        right_col.append([item])
    
    right_table = Table(right_col, colWidths=[4*inch])
    right_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('VALIGN', (0, 0), (0, -1), 'TOP'),
    ]))
    
    main_table_data = [[left_col, divider, right_table]]
    main_table = Table(main_table_data, colWidths=[4*inch, 0.1*inch, 4*inch])
    main_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0.2*inch),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0.2*inch),
    ]))
    
    centered_table = Table([[main_table]], colWidths=[8.2*inch])  
    centered_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    content.append(centered_table)
    
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
        textColor=HexColor('#990000')
    )
    
    subtitle_style = ParagraphStyle(
        'NormalStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        alignment=TA_CENTER
    )
    
    section_title_style = ParagraphStyle(
        'SectionTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_LEFT,
        spaceAfter=6,
        leftIndent=-60,
        textColor=HexColor('#990000')
    )
    
    try:
        table_header_style = styles['TableHeader']
    except KeyError:
        table_header_style = styles.add(ParagraphStyle(
            'TableHeader',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=9,
            alignment=TA_CENTER
        ))
    
    def trim_etf_name(name, max_words=5):
        if not isinstance(name, str):
            return str(name)
        
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
        title = Paragraph(title_text, section_title_style)
        
        headers = [Paragraph(header, table_header_style) for header in column_headers]
        
        table_data = [headers]
        
        for _, row in data.iterrows():
            if is_mf:
                scheme_name = trim_mf_name(row['Unnamed: 1'])
                table_data.append([
                    scheme_name,  
                    str(row['Unnamed: 2']),  
                    str(row['Unnamed: 3']),  
                    str(row['Unnamed: 5']),  
                    str(row['Unnamed: 12']),  
                    str(row['Unnamed: 6']) 
                ])
            else:
                instrument_name = trim_etf_name(row['Unnamed: 0'])
                table_data.append([
                    instrument_name,  
                    str(row['Unnamed: 1']),  
                    str(row['Unnamed: 2']),  
                    str(row['Unnamed: 4']),  
                    str(row['Unnamed: 10']),  
                    str(row['Unnamed: 5'])  
                ])
        
        if is_mf:
            table_data.append([
                'Total:',
                total_data[1] if len(total_data) > 1 else '',
                total_data[2] if len(total_data) > 2 else '',
                total_data[3] if len(total_data) > 3 else '',
                total_data[5] if len(total_data) > 5 else '',
                total_data[4] if len(total_data) > 4 else ''   
            ])
        else:
            table_data.append([
                'Total:',
                total_data[1] if len(total_data) > 1 else '',
                total_data[2] if len(total_data) > 2 else '',
                total_data[3] if len(total_data) > 3 else '',
                total_data[5] if len(total_data) > 5 else '',  
                total_data[4] if len(total_data) > 4 else ''   
            ])
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),  
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),    
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
    
    page_elements.append(Paragraph("Detailed Holdings & Performance", page_title_style))
    page_elements.append(Spacer(1, 0.1*inch))
    page_elements.append(Paragraph(f"As on {datetime.now().strftime('%d-%b-%Y')}", subtitle_style))
    
    sections = []
    if not direct_equity.empty:
        sections.append(("Direct Equity", direct_equity, direct_equity_total, 
                        ['Instrument Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], False))
    
    if not etf_equity.empty:
        sections.append(("Equity ETF", etf_equity, etf_equity_total, 
                        ['ETF Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], False))
    
    if not equity_mf.empty:
        sections.append(("Equity Mutual Funds", equity_mf, equity_mf_total, 
                        ['Scheme Name', 'Units', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], True))
    
    for i, (title, data, totals, headers, is_mf) in enumerate(sections):
        table_elements = create_table(data, totals, title, headers, is_mf)
        page_elements.extend(table_elements)
        
        if i < len(sections) - 1:
            page_elements.append(Spacer(1, 0.01*inch))
    
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
        textColor=HexColor('#990000')
    )
    
    subtitle_style = ParagraphStyle(
        'NormalStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        alignment=TA_CENTER
    )
    
    section_title_style = ParagraphStyle(
        'SectionTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        alignment=TA_LEFT,
        spaceAfter=6,
        leftIndent=-60,
        textColor=HexColor('#990000')
    )
    
    try:
        table_header_style = styles['TableHeader']
    except KeyError:
        table_header_style = styles.add(ParagraphStyle(
            'TableHeader',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=9,
            alignment=TA_CENTER
        ))
    
    def trim_name(name, max_words=5):
        if not isinstance(name, str):
            return str(name)
        
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
    
    def create_table(data, total_data, title_text, column_headers, is_mf=False, is_bond=False):
        title = Paragraph(title_text, section_title_style)
        
        headers = [Paragraph(header, table_header_style) for header in column_headers]
        
        table_data = [headers]
        
        for _, row in data.iterrows():
            if is_mf:
                scheme_name = trim_mf_name(row['Unnamed: 1'])
                table_data.append([
                    scheme_name,  
                    str(row['Unnamed: 2']),  
                    str(row['Unnamed: 3']),  
                    str(row['Unnamed: 5']),  
                    str(row['Unnamed: 12']),  
                    str(row['Unnamed: 6']) 
                ])
            elif is_bond:
                bond_name = trim_name(row['Unnamed: 0'])
                table_data.append([
                    bond_name,  
                    str(row['Unnamed: 1']),  
                    str(row['Unnamed: 2']),  
                    str(row['Unnamed: 4']),  
                    str(row['Unnamed: 10']),  
                    str(row['Unnamed: 5'])  
                ])
            else:
                instrument_name = trim_name(row['Unnamed: 0'])
                table_data.append([
                    instrument_name,  
                    str(row['Unnamed: 1']),  
                    str(row['Unnamed: 2']),  
                    str(row['Unnamed: 4']),  
                    str(row['Unnamed: 10']),  
                    str(row['Unnamed: 5'])  
                ])
        
        if is_mf:
            table_data.append([
                'Total:',
                total_data[1] if len(total_data) > 1 else '',
                total_data[2] if len(total_data) > 2 else '',
                total_data[3] if len(total_data) > 3 else '',
                total_data[5] if len(total_data) > 5 else '',
                total_data[4] if len(total_data) > 4 else ''   
            ])
        else:
            table_data.append([
                'Total:',
                total_data[1] if len(total_data) > 1 else '',
                total_data[2] if len(total_data) > 2 else '',
                total_data[3] if len(total_data) > 3 else '',
                total_data[5] if len(total_data) > 5 else '',  
                total_data[4] if len(total_data) > 4 else ''   
            ])
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),  
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),    
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
    page_elements.append(Paragraph(f"As on {datetime.now().strftime('%d-%b-%Y')}", subtitle_style))
    
    sections = []
    if not debt_etf.empty:
        sections.append(("Debt ETF", debt_etf, debt_etf_total, 
                        ['ETF Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], False))
    
    if not debt_mf.empty:
        sections.append(("Debt Mutual Funds", debt_mf, debt_mf_total, 
                        ['Scheme Name', 'Units', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], True))
    
    if not bond_data.empty:
        sections.append(("Bonds", bond_data, bond_total, 
                        ['Bond Name', 'Quantity', 'Buy Price (Sum)', 'CMP (Sum)', 'P&L (Sum)', 'Market Value'], False, True))
    
    for i, (title, data, totals, headers, is_mf, *args) in enumerate(sections):
        is_bond = args[0] if args else False
        table_elements = create_table(data, totals, title, headers, is_mf, is_bond)
        page_elements.extend(table_elements)
        
        if i < len(sections) - 1:
            page_elements.append(Spacer(1, 0.01*inch))
    
    return page_elements

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