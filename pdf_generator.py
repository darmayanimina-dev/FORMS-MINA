"""
FORMS - FAB OPEX Reconciliation & Monitoring System
PDF Generator: Executive, formal PDF summary report via ReportLab.
Target file: Rekap Realisasi OPEX - [bulan apa].pdf
"""

import io
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas


class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas to dynamically compute and print total page numbers 'Halaman X dari Y'.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))

        # Footer divider line
        self.setStrokeColor(colors.HexColor("#E2E8F0"))
        self.setLineWidth(0.5)
        self.line(36, 36, A4[0] - 36, 36)

        footer_text = f"FORMS - FAB OPEX Reconciliation & Monitoring System | Halaman {self._pageNumber} dari {page_count}"
        self.drawString(36, 24, footer_text)
        
        timestamp_str = datetime.now().strftime("%d/%m/%Y %H:%M WIB")
        self.drawRightString(A4[0] - 36, 24, f"Dicetak: {timestamp_str}")
        self.restoreState()


def format_rupiah(value: float) -> str:
    """Format float into accounting currency string."""
    if value < 0:
        return f"-Rp {abs(value):,.2f}"
    return f"Rp {value:,.2f}"


def generate_rekap_pdf(df_rekap: pd.DataFrame, period: str, total_saldo: float) -> io.BytesIO:
    """
    Generates a formal, printable PDF summary report of OPEX realization.
    Columns: Code Account, Account Name, Total Saldo
    Final Row: TOTAL SEMUA
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=50
    )

    styles = getSampleStyleSheet()
    
    # Custom Typography
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=18,
        textColor=colors.HexColor('#0F172A'),
        alignment=0 # Left
    )

    subtitle_style = ParagraphStyle(
        'DocSub',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=0
    )

    meta_label_style = ParagraphStyle(
        'MetaLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#475569')
    )

    meta_val_style = ParagraphStyle(
        'MetaVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#0F172A')
    )

    cell_style = ParagraphStyle(
        'CellRegular',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#1E293B')
    )

    cell_bold_style = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#0F172A')
    )

    elements = []

    # 1. Header Banner & Branding
    elements.append(Paragraph("FAB OPEX RECONCILIATION & MONITORING SYSTEM (FORMS)", subtitle_style))
    elements.append(Spacer(1, 3))
    elements.append(Paragraph("LAPORAN REKAPITULASI REALISASI BIAYA OPERASIONAL", title_style))
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1E3A8A"), spaceAfter=10))

    # 2. Metadata Information Box
    print_date = datetime.now().strftime("%d %B %Y %H:%M")
    meta_table_data = [
        [
            Paragraph("<b>Periode Laporan:</b>", meta_label_style),
            Paragraph(f"<b>{period}</b>", meta_val_style),
            Paragraph("<b>Total Pos Akun:</b>", meta_label_style),
            Paragraph(f"{len(df_rekap)} Akun Unik", meta_val_style)
        ],
        [
            Paragraph("<b>Divisi/Unit Kerja:</b>", meta_label_style),
            Paragraph("SERVICE", meta_val_style),
            Paragraph("<b>Total Beban OPEX:</b>", meta_label_style),
            Paragraph(f"<b>{format_rupiah(total_saldo)}</b>", meta_val_style)
        ]
    ]

    meta_table = Table(meta_table_data, colWidths=[100, 160, 100, 160])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 14))

    # 3. Main Rekapitulasi Table
    # Columns required by user: Code Account, Account Name, Total Saldo
    # We include 'No' for formal numbered accounting format
    table_data = [
        [
            Paragraph("<b>NO</b>", ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)),
            Paragraph("<b>CODE ACCOUNT</b>", ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)),
            Paragraph("<b>ACCOUNT NAME</b>", ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=0)),
            Paragraph("<b>TOTAL SALDO</b>", ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=2)),
        ]
    ]

    for idx, row in df_rekap.iterrows():
        saldo_num = float(row.get("Total Saldo", 0.0) or 0.0)
        saldo_color = "#DC2626" if saldo_num < 0 else "#0F172A"
        saldo_p_style = ParagraphStyle(
            f'Saldo_{idx}',
            parent=styles['Normal'],
            fontName='Helvetica-Bold' if saldo_num < 0 else 'Helvetica',
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor(saldo_color),
            alignment=2 # Right align
        )

        table_data.append([
            Paragraph(str(idx + 1), ParagraphStyle('CenterNo', parent=cell_style, alignment=1)),
            Paragraph(str(row.get("Code Account", "") or ""), ParagraphStyle('CenterCode', parent=cell_style, alignment=1)),
            Paragraph(str(row.get("Account Name", "") or ""), cell_style),
            Paragraph(format_rupiah(saldo_num), saldo_p_style)
        ])

    # Final Total Row
    total_label_style = ParagraphStyle('TotalLabel', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#0F172A'), alignment=2)
    total_val_style = ParagraphStyle('TotalVal', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#0F172A'), alignment=2)

    table_data.append([
        Paragraph("", cell_style),
        Paragraph("<b>TOTAL SEMUA</b>", total_label_style),
        Paragraph("", cell_style),
        Paragraph(format_rupiah(total_saldo), total_val_style)
    ])

    # Table Widths: Total page width = 595 - 72 = 523 pt
    col_widths = [32, 110, 231, 150]
    rekap_table = Table(table_data, colWidths=col_widths, repeatRows=1)

    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]

    # Zebra Striping
    for r in range(1, len(table_data) - 1):
        if r % 2 == 0:
            t_style.append(('BACKGROUND', (0, r), (-1, r), colors.HexColor('#F8FAFC')))

    # Total Row Styling
    total_row_idx = len(table_data) - 1
    t_style.extend([
        ('SPAN', (1, total_row_idx), (2, total_row_idx)),
        ('BACKGROUND', (0, total_row_idx), (-1, total_row_idx), colors.HexColor('#E2E8F0')),
        ('LINEABOVE', (0, total_row_idx), (-1, total_row_idx), 1.5, colors.HexColor('#0F172A')),
        ('LINEBELOW', (0, total_row_idx), (-1, total_row_idx), 2, colors.HexColor('#0F172A')),
        ('TOPPADDING', (0, total_row_idx), (-1, total_row_idx), 6),
        ('BOTTOMPADDING', (0, total_row_idx), (-1, total_row_idx), 6),
    ])

    rekap_table.setStyle(TableStyle(t_style))
    elements.append(rekap_table)

    # 4. Sign-off / Verification Box (Kiri FAB, Kanan SUAH)
    elements.append(Spacer(1, 20))
    sign_table_data = [
        [
            Paragraph("<b>Dibuat Oleh:</b><br/><br/><br/><br/>( FAB )", ParagraphStyle('Sign1', fontName='Helvetica-Bold', fontSize=9, alignment=1)),
            Paragraph("", ParagraphStyle('SignSpace')),
            Paragraph("<b>Disetujui Oleh:</b><br/><br/><br/><br/>( SUAH )", ParagraphStyle('Sign2', fontName='Helvetica-Bold', fontSize=9, alignment=1))
        ]
    ]
    sign_table = Table(sign_table_data, colWidths=[200, 123, 200])
    sign_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    elements.append(KeepTogether([sign_table]))

    doc.build(elements, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer
