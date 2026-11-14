"""Deterministic, Unicode Arabic A4 landscape receipt; repeated headers/footers."""
from __future__ import annotations
from io import BytesIO
from pathlib import Path
import math
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors

ASSETS = Path(__file__).parent / 'assets'
pdfmetrics.registerFont(TTFont('Arabic', str(ASSETS / 'Arabic.ttf')))
PALETTE = ['#ffc7ce', '#f8cbad', '#ffeb9c', '#c6efce']


def shape(text):
    return get_display(arabic_reshaper.reshape(str(text)))


def wrap(text, width, size):
    lines, current = [], ''
    for word in str(text).split():
        candidate = (current + ' ' + word).strip()
        if pdfmetrics.stringWidth(shape(candidate), 'Arabic', size) <= width:
            current = candidate
        else:
            if current:
                lines.append(current)
            # Split unusually long codes instead of letting text escape its cell.
            current = ''
            for char in word:
                if pdfmetrics.stringWidth(shape(current + char), 'Arabic', size) > width:
                    lines.append(current)
                    current = char
                else:
                    current += char
    if current:
        lines.append(current)
    return lines or ['']


def render_receipt(receipt: dict) -> bytes:
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4), pageCompression=1)
    W, H = landscape(A4)
    c.setTitle(receipt['settings']['document_title'] + ' - ' + str(receipt['serial']))
    c.setAuthor(receipt['settings']['company'])
    margin, content = 22, W - 44
    widths = [24, 78, 180, 34, 34, 34, 66, 83, 90, content - 723]

    def text(x, y, value, size=8, align='right', color='#152c42'):
        c.setFillColor(colors.HexColor(color))
        c.setFont('Arabic', size)
        getattr(c, {'right': 'drawRightString', 'center': 'drawCentredString', 'left': 'drawString'}[align])(x, y, shape(value))

    def box(x, y, width, height, value='', fill='#ffffff', size=8, max_lines=2, ink='#152c42'):
        c.setFillColor(colors.HexColor(fill))
        c.setStrokeColor(colors.HexColor('#becbd5'))
        c.setLineWidth(.35)
        c.rect(x, y, width, height, fill=1, stroke=1)
        lines = wrap(value, width - 8, size)
        if len(lines) > max_lines:
            # Long notes are printed in a complete appendix, with a visible marker here.
            lines = lines[:max_lines - 1] + ['… انظر التفاصيل']
        baseline = y + height / 2 + (len(lines)-1) * (size + 1) / 2 - size * .35
        for line in lines:
            text(x + width / 2, baseline, line, size, 'center', ink)
            baseline -= size + 1

    items = receipt['items']
    pages = max(1, math.ceil(len(items) / 22))
    note_appendix = []
    for i, item in enumerate(items):
        # Always append full long values rather than silently truncating any detail.
        if any(len(wrap(value, width-8, 6.4))>2 for value,width in [(item['name'],widths[2]),(item.get('notes',''),widths[9]),(item.get('reason',''),widths[8])]):
            note_appendix.append((i+1, item))

    for page in range(pages):
        settings, customer, driver = receipt['settings'], receipt['customer'], receipt['driver']
        c.drawImage(str(ASSETS/'company.jpeg'), W - margin - 390, H - 83, width=390, height=67, preserveAspectRatio=True, mask='auto')
        text(margin + 8, H - 33, 'C.C.M. No. ' + settings.get('membership', ''), 8, 'left')
        text(margin + 8, H - 52, settings['company'], 7.5, 'left')
        text(W / 2, H - 99, settings['document_title'], 13, 'center')
        text(W - margin, H - 119, f"رقم السند: #{receipt['serial']}     رقم الإرجاع: {receipt.get('return_number', '')}     تاريخ الفاتورة: {receipt.get('invoice_date') or '—'}", 8)
        text(margin, H - 119, 'تاريخ السند: ' + receipt['created_at'][:19].replace('T', ' '), 8, 'left')
        summary = [f"العميل: {customer['name']}", f"الكود: {customer['code']}", f"المندوب: {customer.get('rep_code','')} - {customer.get('rep_name','')}", f"القطاع: {customer.get('sector','')}"]
        # Separate lines keep long customer names legible.
        text(W - margin, H - 135, summary[0], 8)
        text(W - margin, H - 149, summary[1] + '     ' + summary[2] + '     ' + summary[3], 8)
        # Logical RTL order from right edge.
        labels = ['م', 'كود الصنف', 'اسم الصنف / السلعة', 'كرتون', 'علبة', 'حبة', 'تاريخ الانتهاء', 'المخزن / الحالة', 'سبب الإرجاع', 'ملاحظات']
        top, row_h, head_h = H - 157, 16, 25
        x = W-margin
        for width, label in zip(widths, labels):
            x -= width
            box(x, top-head_h, width, head_h, label, '#143e68', 7, 2, '#ffffff')
        chunk = items[page*22:(page+1)*22]
        for r in range(22):
            item = chunk[r] if r < len(chunk) else None
            y = top-head_h-(r+1)*row_h
            vals = [page*22+r+1, item['product_code'], item['name'], item['cartons'], item['boxes'], item['pieces'], item['expiry'], item['status'], item.get('reason',''), item.get('notes','')] if item else ['']*10
            x = W-margin
            for k, (width, value) in enumerate(zip(widths, vals)):
                x -= width
                fill = PALETTE[item['status_index']] if item and k==7 else ('#f4f8fb' if r%2 else '#ffffff')
                box(x, y, width, row_h, value, fill, 6.4 if k in (2,8,9) else 6.7, 2)
        bottom = top-head_h-22*row_h
        totals = receipt['totals']
        text(W-margin, bottom-17, f"الإجمالي: {len(items)} صنف     كرتون: {totals['cartons']}     علبة: {totals['boxes']}     حبة: {totals['pieces']}", 9)
        text(W-margin, bottom-34, 'أمين المستودع: ' + settings.get('keeper',''), 8)
        text(W/2, bottom-34, 'المحاسب: ' + settings.get('accountant',''), 8, 'center')
        text(margin, bottom-34, 'السائق: ' + driver.get('code','') + ' - ' + driver.get('name',''), 8, 'left')
        text(W-margin, 17, 'أمين الصندوق: ' + receipt.get('operator',''), 7)
        text(margin, 17, f"صفحة {page+1} / {pages}" + (' • توجد تفاصيل إضافية' if note_appendix or receipt.get('notes') else ''), 7, 'left')
        c.showPage()

    if note_appendix or receipt.get('notes'):
        y = H-35
        text(W-margin, y, f"تفاصيل السند #{receipt['serial']}", 13)
        y -= 24
        values = [('ملاحظات السند', receipt.get('notes',''))]
        for i, item in note_appendix:
            values.append((f"الصنف {i} — {item['product_code']}", item['name']+' | '+item.get('reason','')+' | '+item.get('notes','')))
        for title, value in values:
            if not value:
                continue
            for line in wrap(title + ': ' + value, content - 10, 9):
                if y < 35:
                    c.showPage()
                    y = H-35
                text(W-margin, y, line, 9)
                y -= 16
            y -= 8
        c.showPage()
    c.save()
    return buffer.getvalue()
