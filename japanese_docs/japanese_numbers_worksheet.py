import os
import random
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Colour palette (matches greetings worksheet) ───────────────────────────────
NAVY  = RGBColor(0x1A, 0x3A, 0x5C)
TEAL  = RGBColor(0x00, 0x7A, 0x87)
GOLD  = RGBColor(0xE6, 0xA8, 0x17)
LIGHT = RGBColor(0xF0, 0xF6, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x1A, 0x1A, 0x2E)
GREY  = RGBColor(0x6B, 0x7B, 0x8D)
MINT  = RGBColor(0xD4, 0xF1, 0xF4)   # soft teal tint for alternating rows

# ── Number data ────────────────────────────────────────────────────────────────
# (english_label, kanji, hiragana, romaji, notes)
NUMBERS = [
    ("1  — One",        "一",     "いち",         "ichi",        ""),
    ("2  — Two",        "二",     "に",           "ni",          ""),
    ("3  — Three",      "三",     "さん",         "san",         ""),
    ("4  — Four",       "四",     "し / よん",    "shi / yon",   "Both are used; よん is safer in daily speech."),
    ("5  — Five",       "五",     "ご",           "go",          ""),
    ("6  — Six",        "六",     "ろく",         "roku",        ""),
    ("7  — Seven",      "七",     "しち / なな",  "shichi / nana","なな is more common in daily speech."),
    ("8  — Eight",      "八",     "はち",         "hachi",       ""),
    ("9  — Nine",       "九",     "く / きゅう",  "ku / kyuu",   "きゅう is preferred in most contexts."),
    ("10 — Ten",        "十",     "じゅう",       "juu",         ""),
    ("11 — Eleven",     "十一",   "じゅういち",   "juu-ichi",    "Ten + one."),
    ("12 — Twelve",     "十二",   "じゅうに",     "juu-ni",      "Ten + two."),
    ("20 — Twenty",     "二十",   "にじゅう",     "ni-juu",      "Two × ten."),
    ("30 — Thirty",     "三十",   "さんじゅう",   "san-juu",     "Three × ten."),
    ("100 — One hundred","百",    "ひゃく",       "hyaku",       "New counter word."),
    ("1000 — One thousand","千",  "せん",         "sen",         "New counter word."),
]

# Matching page: focus on 1–20 in a compact two-column layout
MATCH_NUMBERS = [
    *NUMBERS[:12],  # 1–12
    ("13 — Thirteen", "十三", "じゅうさん",   "juu-san",   ""),
    ("14 — Fourteen", "十四", "じゅうよん",   "juu-yon",   ""),
    ("15 — Fifteen",  "十五", "じゅうご",     "juu-go",    ""),
    ("16 — Sixteen",  "十六", "じゅうろく",   "juu-roku",  ""),
    ("17 — Seventeen","十七", "じゅうなな",   "juu-nana",  ""),
    ("18 — Eighteen", "十八", "じゅうはち",   "juu-hachi", ""),
    ("19 — Nineteen", "十九", "じゅうきゅう", "juu-kyuu",  ""),
    NUMBERS[12],     # 20
]


# ── XML / formatting helpers (same as greetings file) ─────────────────────────
def _rgb_hex(rgb):
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

def set_cell_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  _rgb_hex(rgb))
    tcPr.append(shd)

def set_para_bg(para, rgb):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  _rgb_hex(rgb))
    pPr.append(shd)

def add_run(para, text, bold=False, italic=False, size=12, color=DARK, font='Calibri'):
    r = para.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.size  = Pt(size)
    r.font.color.rgb = color
    r.font.name  = font
    return r

def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6 if level == 1 else 4)
    p.paragraph_format.space_after  = Pt(2)
    add_run(p, text, bold=True,
            size=14 if level == 1 else 12,
            color=NAVY if level == 1 else TEAL)
    return p

def section_banner(doc, text, bg=NAVY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Cm(0.3)
    add_run(p, text, bold=True, size=12, color=WHITE)
    set_para_bg(p, bg)
    return p

def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    p.add_run()._r.append(br)

def divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '4')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), '007A87')
    pBdr.append(bot)
    pPr.append(pBdr)


def get_matching_bank():
    """Return a deterministic answer bank and lookup for Part 2."""
    items = list(MATCH_NUMBERS)
    rng = random.Random(42)
    rng.shuffle(items)

    bank = []
    answer_lookup = {}
    for idx, item in enumerate(items):
        label = chr(ord('A') + idx)
        bank.append((label, item))
        answer_lookup[item[0]] = label
    return bank, answer_lookup


# ── Page builders ──────────────────────────────────────────────────────────────

def build_page1(doc):
    """Banner, info row, objectives, Part 1 — write the numbers."""

    # ── Title banner ──────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=2, cols=1)
    tbl.style = 'Table Grid'
    r0 = tbl.rows[0].cells[0]
    set_cell_bg(r0, NAVY)
    r0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = r0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before, p0.paragraph_format.space_after = Pt(5), Pt(2)
    add_run(p0, 'Beginner Japanese', bold=True, size=20, color=WHITE)
    add_run(p0, '\nNumbers 1–1000  ·  数字 (sūji)', bold=False, size=13, color=GOLD)
    r1 = tbl.rows[1].cells[0]
    set_cell_bg(r1, TEAL)
    p1 = r1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before, p1.paragraph_format.space_after = Pt(3), Pt(3)
    add_run(p1, 'Student Worksheet  ·  Page 1 of 3', bold=False, size=10, color=WHITE)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ── Student info row ──────────────────────────────────────────────────────
    info = doc.add_table(rows=1, cols=4)
    info.style = 'Table Grid'
    for i, (lbl, blank) in enumerate([
        ('Name',   '_________________________________'),
        ('Class',  '___________'),
        ('Date',   '__________________'),
        ('Period', '____'),
    ]):
        c = info.rows[0].cells[i]
        set_cell_bg(c, LIGHT)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        add_run(p, f'{lbl}: ', bold=True, size=10, color=NAVY)
        add_run(p, blank,      size=10, color=DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ── Objectives ────────────────────────────────────────────────────────────
    heading(doc, 'Learning Objectives')
    for obj in [
        'Read and write Japanese numbers 1–1,000 in kanji and hiragana.',
        'Learn the romaji (romanized) pronunciation for each number.',
        'Understand the building-block pattern: 十 (juu), 百 (hyaku), 千 (sen).',
    ]:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent  = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        add_run(p, obj, size=10)

    # ── Instructions ─────────────────────────────────────────────────────────
    heading(doc, 'Instructions — Part 1')
    for n, txt in [
        ('1.', 'Write the kanji (Chinese-origin numeral) on the first line.'),
        ('2.', 'Write the hiragana reading on the second line.'),
        ('3.', 'Write the romaji on the third line.  Hints are in parentheses.'),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent  = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        add_run(p, f'{n} ', bold=True, size=10, color=TEAL)
        add_run(p, txt, size=10)
    divider(doc)

    # ── Part 1 — Write the numbers ────────────────────────────────────────────
    section_banner(doc, 'Part 1 — Write the Numbers', bg=NAVY)

    for i, (eng, _kanji, _hira, hint, note) in enumerate(NUMBERS, 1):
        # Question label
        pq = doc.add_paragraph()
        pq.paragraph_format.space_before = Pt(3)
        pq.paragraph_format.space_after  = Pt(0)
        add_run(pq, f'{i:>2}.  ', bold=True, size=10, color=TEAL)
        add_run(pq, eng, bold=False, size=10, color=DARK)
        if note:
            add_run(pq, f'  ★ {note}', italic=True, size=8, color=GREY)

        # Three answer lines
        for label in ('Kanji:    ', 'Hiragana:', 'Romaji:  '):
            pl = doc.add_paragraph()
            pl.paragraph_format.left_indent  = Cm(1.0)
            pl.paragraph_format.space_before = Pt(0)
            pl.paragraph_format.space_after  = Pt(0)
            add_run(pl, label, bold=True, size=9, color=TEAL)
            add_run(pl, '_' * 35, size=9, color=DARK)
            if label.startswith('Romaji') and hint:
                add_run(pl, f'  ({hint})', italic=True, size=8, color=GREY)


def build_page2(doc):
    """Compact matching exercise with a lettered answer bank."""
    page_break(doc)

    # ── Page header ───────────────────────────────────────────────────────────
    ph = doc.add_paragraph()
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ph.paragraph_format.space_before, ph.paragraph_format.space_after = Pt(0), Pt(2)
    add_run(ph, 'Student Worksheet  ·  Page 2 of 3', bold=False, size=10, color=GREY)

    # ── Instructions ─────────────────────────────────────────────────────────
    heading(doc, 'Instructions — Part 2: Matching')
    inst_items = [
        ('1.', 'Look at the kanji and the hiragana reading in each box.'),
        ('2.', 'Find the matching English word in the answer bank.'),
        ('3.', 'Write the correct letter on the blank line for each number.'),
    ]
    for n, txt in inst_items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent  = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        add_run(p, f'{n} ', bold=True, size=10, color=TEAL)
        add_run(p, txt, size=10)

    tip_p = doc.add_paragraph()
    tip_p.paragraph_format.left_indent  = Cm(0.4)
    tip_p.paragraph_format.space_before = Pt(1)
    tip_p.paragraph_format.space_after  = Pt(3)
    add_run(tip_p, 'Tip: ', bold=True, size=10, color=GOLD)
    add_run(tip_p, 'Numbers 11-19 all start with 十, so spot the ones digit after it.', italic=True, size=10, color=GREY)
    divider(doc)

    section_banner(doc, 'Part 2 — Match the Number Meanings', bg=TEAL)

    intro = doc.add_paragraph()
    intro.paragraph_format.space_before = Pt(1)
    intro.paragraph_format.space_after = Pt(2)
    add_run(intro, 'Answer bank', bold=True, size=11, color=NAVY)
    add_run(intro, '  Write the correct letter for each number box below.', size=10, color=DARK)

    bank, _answer_lookup = get_matching_bank()
    bank_tbl = doc.add_table(rows=5, cols=4)
    bank_tbl.style = 'Table Grid'
    for idx, (label, item) in enumerate(bank):
        row_idx = idx // 4
        col_idx = idx % 4
        eng_display = item[0].split('—')[-1].strip()
        cell = bank_tbl.rows[row_idx].cells[col_idx]
        set_cell_bg(cell, LIGHT if row_idx % 2 == 0 else WHITE)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        add_run(p, f'{label}. ', bold=True, size=10, color=TEAL)
        add_run(p, eng_display, size=9, color=DARK)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    grid_tbl = doc.add_table(rows=10, cols=2)
    grid_tbl.style = 'Table Grid'
    for pair_idx in range(10):
        for col_idx in range(2):
            item_idx = pair_idx * 2 + col_idx
            eng, kanji, hira, _ro, _note = MATCH_NUMBERS[item_idx]
            cell = grid_tbl.rows[pair_idx].cells[col_idx]
            set_cell_bg(cell, MINT if pair_idx % 2 == 0 else WHITE)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            p_top = cell.paragraphs[0]
            p_top.paragraph_format.space_before = Pt(2)
            p_top.paragraph_format.space_after = Pt(0)
            add_run(p_top, f'{item_idx + 1}. ', bold=True, size=10, color=NAVY)
            add_run(p_top, 'Letter: ', bold=True, size=9, color=TEAL)
            add_run(p_top, '____', size=9, color=DARK)

            p_body = cell.add_paragraph()
            p_body.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_body.paragraph_format.space_before = Pt(1)
            p_body.paragraph_format.space_after = Pt(2)
            add_run(p_body, kanji, bold=True, size=18, color=NAVY)
            add_run(p_body, f'\n{hira}', size=10, color=TEAL)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    heading(doc, 'How the Numbers Work', level=2)
    for tip in [
        '1-10 are the base numbers you memorize first: 一 いち, 二 に, 三 さん, 四 し/よん, 五 ご, 六 ろく, 七 しち/なな, 八 はち, 九 く/きゅう, 十 じゅう.',
        '11-19 use 十 + the ones digit: 11 = 十一 (じゅういち), 14 = 十四 (じゅうよん), 19 = 十九 (じゅうきゅう).',
        '20 is 二十 (にじゅう). The same pattern keeps going: 21 = 二十一 (にじゅういち), 35 = 三十五 (さんじゅうご), 99 = 九十九 (きゅうじゅうきゅう), 100 = 百 (ひゃく), 1000 = 千 (せん).',
    ]:
        pt = doc.add_paragraph()
        pt.paragraph_format.left_indent = Cm(0.6)
        pt.paragraph_format.space_before = Pt(0)
        pt.paragraph_format.space_after = Pt(1)
        add_run(pt, '- ', bold=True, size=10, color=GOLD)
        add_run(pt, tip, size=8)


def build_page3(doc):
    """Answer key — teacher copy."""
    page_break(doc)

    # ── Answer key banner ─────────────────────────────────────────────────────
    tbl = doc.add_table(rows=2, cols=1)
    tbl.style = 'Table Grid'
    ra0 = tbl.rows[0].cells[0]
    set_cell_bg(ra0, DARK)
    ra0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    pa0 = ra0.paragraphs[0]
    pa0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pa0.paragraph_format.space_before, pa0.paragraph_format.space_after = Pt(5), Pt(2)
    add_run(pa0, 'Answer Key — Teacher Copy', bold=True, size=16, color=WHITE)
    add_run(pa0, '\nPage 3 of 3', bold=False, size=11, color=GOLD)
    ra1 = tbl.rows[1].cells[0]
    set_cell_bg(ra1, GOLD)
    pa1 = ra1.paragraphs[0]
    pa1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pa1.paragraph_format.space_before, pa1.paragraph_format.space_after = Pt(2), Pt(2)
    add_run(pa1, 'For teacher use only — do not distribute to students', bold=False, size=9, color=DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ── Part 1 answers ────────────────────────────────────────────────────────
    heading(doc, 'Part 1 — Write the Numbers (Answers)')

    atbl = doc.add_table(rows=1, cols=5)
    atbl.style = 'Table Grid'
    for i, col_label in enumerate(['#', 'English', 'Kanji', 'Hiragana', 'Romaji']):
        c = atbl.rows[0].cells[i]
        set_cell_bg(c, NAVY)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        add_run(p, col_label, bold=True, size=10, color=WHITE)

    for row_i, (eng, kanji, hira, romaji, note) in enumerate(NUMBERS):
        row = atbl.add_row()
        bg  = LIGHT if row_i % 2 == 0 else WHITE
        for ci, (val, clr, sz, bld) in enumerate([
            (str(row_i + 1), GREY,  9,  False),
            (eng,            DARK,  10, False),
            (kanji,          TEAL,  13, True),
            (hira,           NAVY,  10, False),
            (romaji,         DARK,  9,  False),
        ]):
            c = row.cells[ci]
            set_cell_bg(c, bg)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
            if ci == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(p, val, bold=bld, size=sz, color=clr)
        # Note in last cell if present
        if note:
            p.add_run('')  # keep existing cell
            c_note = row.cells[4]
            p_note = c_note.paragraphs[0]
            add_run(p_note, f'  ★ {note}', italic=True, size=8, color=GREY)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    divider(doc)

    # ── Part 2 matching answers ───────────────────────────────────────────────
    heading(doc, 'Part 2 — Matching Answers')

    _, answer_lookup = get_matching_bank()
    mtbl = doc.add_table(rows=1, cols=5)
    mtbl.style = 'Table Grid'
    for i, col_label in enumerate(['#', 'Kanji', 'Reading', 'Letter', 'English']):
        c = mtbl.rows[0].cells[i]
        set_cell_bg(c, TEAL)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        add_run(p, col_label, bold=True, size=10, color=WHITE)

    for row_i, (eng, kanji, hira, _ro, _note) in enumerate(MATCH_NUMBERS):
        row = mtbl.add_row()
        bg  = MINT if row_i % 2 == 0 else WHITE
        eng_display = eng.split('—')[-1].strip()
        for ci, (val, clr, sz, bld) in enumerate([
            (str(row_i + 1),         GREY,  9,  True),
            (kanji,                  TEAL, 14, True),
            (hira,                   NAVY, 10, False),
            (answer_lookup[eng],     TEAL, 10, True),
            (eng_display,            DARK, 10, False),
        ]):
            c = row.cells[ci]
            set_cell_bg(c, bg)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
            add_run(p, val, bold=bld, size=sz, color=clr)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    divider(doc)

    # ── Teacher activity suggestions ──────────────────────────────────────────
    heading(doc, "Teacher's Notes & Suggested Activities")
    for title, desc in [
        ('Counting drill (5 min)',       'Count aloud 1–20 as a class, then backwards in pairs.'),
        ('Flash cards (10 min)',          'Show a kanji card; students call out the English or hiragana reading.'),
        ('Pattern discovery (5 min)',    'Write 11–19 on board; ask students to spot the pattern (十 + digit).'),
        ('Matching check (pair work)',   'Partners compare answer letters and explain how they matched each number.'),
        ('Extension — bigger numbers',  'Introduce 万 (man = 10,000) for advanced students.'),
    ]:
        pa = doc.add_paragraph()
        pa.paragraph_format.left_indent  = Cm(0.6)
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after  = Pt(1)
        add_run(pa, f'{title}: ', bold=True,  size=10, color=NAVY)
        add_run(pa, desc,         bold=False, size=10, color=DARK)


# ── Entry point ────────────────────────────────────────────────────────────────
def main():
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    build_page1(doc)
    build_page2(doc)
    build_page3(doc)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'japanese_numbers_worksheet.docx')
    doc.save(out)
    print(f'Saved: {out}')


if __name__ == '__main__':
    main()
