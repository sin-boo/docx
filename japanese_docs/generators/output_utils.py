import os
from xml.sax.saxutils import escape

from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Paragraph as PdfParagraph
from reportlab.platypus import SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle


PDF_FONT_NAME = 'HeiseiKakuGo-W5'


def ensure_pdf_fonts():
    try:
        pdfmetrics.getFont(PDF_FONT_NAME)
    except KeyError:
        pdfmetrics.registerFont(UnicodeCIDFont(PDF_FONT_NAME))


def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError(f'Unsupported parent type: {type(parent)}')

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _cell_text(cell):
    parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    return '<br/>'.join(escape(part) for part in parts) if parts else ' '


def _doc_to_story(doc):
    ensure_pdf_fonts()
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        'WorksheetBody',
        parent=styles['BodyText'],
        fontName=PDF_FONT_NAME,
        fontSize=9,
        leading=11,
        spaceAfter=4,
    )
    table_cell_style = ParagraphStyle(
        'WorksheetCell',
        parent=body_style,
        fontName=PDF_FONT_NAME,
        fontSize=8,
        leading=10,
        spaceAfter=0,
    )
    story = []

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                story.append(PdfParagraph(escape(text).replace('\n', '<br/>'), body_style))
                story.append(Spacer(1, 0.04 * inch))
        elif isinstance(block, Table):
            data = []
            for row in block.rows:
                data.append([PdfParagraph(_cell_text(cell), table_cell_style) for cell in row.cells])

            if not data:
                continue

            pdf_table = PdfTable(data, repeatRows=1)
            style_cmds = [
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#DDEAF2')),
            ]
            pdf_table.setStyle(TableStyle(style_cmds))
            story.append(pdf_table)
            story.append(Spacer(1, 0.08 * inch))

    return story


def _draw_footer(canvas, _doc):
    ensure_pdf_fonts()
    canvas.saveState()
    canvas.setFont(PDF_FONT_NAME, 7)
    canvas.setFillColor(colors.HexColor('#6B7B8D'))
    canvas.drawRightString(7.85 * inch, 0.22 * inch, 'neuralforge.cc')
    canvas.restoreState()


def export_pdf(doc, pdf_path):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.35 * inch,
        rightMargin=0.35 * inch,
        topMargin=0.35 * inch,
        bottomMargin=0.45 * inch,
    )
    story = _doc_to_story(doc)
    pdf.build(story, onFirstPage=_draw_footer, onLaterPages=_draw_footer)


def get_output_paths(script_dir, filename, subdir=None):
    docs_dir = os.path.dirname(script_dir)
    docs_output_dir = os.path.join(docs_dir, 'output', 'docs')
    pdf_output_dir = os.path.join(docs_dir, 'output', 'pdf')

    if subdir:
        docs_output_dir = os.path.join(docs_output_dir, subdir)
        pdf_output_dir = os.path.join(pdf_output_dir, subdir)

    os.makedirs(docs_output_dir, exist_ok=True)
    os.makedirs(pdf_output_dir, exist_ok=True)

    docx_path = os.path.join(docs_output_dir, filename)
    pdf_path = os.path.join(pdf_output_dir, os.path.splitext(filename)[0] + '.pdf')
    return docx_path, pdf_path


def save_outputs(doc, script_dir, filename, subdir=None):
    docx_path, pdf_path = get_output_paths(script_dir, filename, subdir=subdir)
    doc.save(docx_path)
    export_pdf(doc, pdf_path)
    return docx_path, pdf_path
