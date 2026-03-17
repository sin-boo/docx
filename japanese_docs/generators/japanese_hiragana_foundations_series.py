import math
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Colour palette ─────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x1A, 0x3A, 0x5C)
TEAL  = RGBColor(0x00, 0x7A, 0x87)
GOLD  = RGBColor(0xE6, 0xA8, 0x17)
LIGHT = RGBColor(0xF0, 0xF6, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x1A, 0x1A, 0x2E)
GREY  = RGBColor(0x6B, 0x7B, 0x8D)
MINT  = RGBColor(0xD4, 0xF1, 0xF4)


WORKSHEETS = [
    {
        "slug": "hiragana_set_01_vowels_worksheet.docx",
        "title": "Hiragana Set 1 — Vowels",
        "subtitle": "あ い う え お",
        "focus": "Learn the five basic vowel sounds. These vowels appear in almost every Japanese word.",
        "kana": [
            ("あ", "a", "Open and clear, like a in father."),
            ("い", "i", "Short i, like ee in see."),
            ("う", "u", "Short u, with rounded lips."),
            ("え", "e", "Short e, like e in bed."),
            ("お", "o", "Short o, like o in open."),
        ],
        "pairs": [("あい", "a-i"), ("うえ", "u-e"), ("あお", "a-o"), ("いえ", "i-e")],
        "teaching_notes": [
            "These five vowels are the sound base for all the other hiragana rows.",
            "Say each vowel evenly without stretching it too much.",
            "Read two-kana practice by saying one sound, then the next: あい = a-i.",
        ],
    },
    {
        "slug": "hiragana_set_02_k_row_worksheet.docx",
        "title": "Hiragana Set 2 — K Row",
        "subtitle": "か き く け こ",
        "focus": "Keep the k sound the same and change only the vowel after it.",
        "kana": [
            ("か", "ka", "k + a"),
            ("き", "ki", "k + i"),
            ("く", "ku", "k + u"),
            ("け", "ke", "k + e"),
            ("こ", "ko", "k + o"),
        ],
        "pairs": [("かき", "ka-ki"), ("くけ", "ku-ke"), ("ここ", "ko-ko"), ("かこ", "ka-ko")],
        "teaching_notes": [
            "The consonant stays k all the way through the row.",
            "Only the vowel changes: a, i, u, e, o.",
            "Practice saying the row in order until it sounds smooth.",
        ],
    },
    {
        "slug": "hiragana_set_03_s_row_worksheet.docx",
        "title": "Hiragana Set 3 — S Row",
        "subtitle": "さ し す せ そ",
        "focus": "The s row introduces し, which sounds like shi instead of si.",
        "kana": [
            ("さ", "sa", "s + a"),
            ("し", "shi", "Sounds like shee."),
            ("す", "su", "s + u"),
            ("せ", "se", "s + e"),
            ("そ", "so", "s + o"),
        ],
        "pairs": [("さし", "sa-shi"), ("すし", "su-shi"), ("せそ", "se-so"), ("そさ", "so-sa")],
        "teaching_notes": [
            "し is one of the first kana that does not match a simple consonant + vowel pattern.",
            "Students often remember すし because sushi is already familiar in English.",
            "Keep the vowel steady even when the spelling changes.",
        ],
    },
    {
        "slug": "hiragana_set_04_t_row_worksheet.docx",
        "title": "Hiragana Set 4 — T Row",
        "subtitle": "た ち つ て と",
        "focus": "The t row has two special sounds: ち = chi and つ = tsu.",
        "kana": [
            ("た", "ta", "t + a"),
            ("ち", "chi", "Sounds like chee."),
            ("つ", "tsu", "Starts with ts sound."),
            ("て", "te", "t + e"),
            ("と", "to", "t + o"),
        ],
        "pairs": [("たち", "ta-chi"), ("つて", "tsu-te"), ("とた", "to-ta"), ("てと", "te-to")],
        "teaching_notes": [
            "ち and つ are common beginner trouble spots, so say them slowly first.",
            "つ begins with a small ts sound before the vowel u.",
            "Mixing this row with the s row is normal at first, so practice aloud often.",
        ],
    },
    {
        "slug": "hiragana_set_05_n_row_worksheet.docx",
        "title": "Hiragana Set 5 — N Row",
        "subtitle": "な に ぬ ね の",
        "focus": "This row is regular and is good for building reading confidence.",
        "kana": [
            ("な", "na", "n + a"),
            ("に", "ni", "n + i"),
            ("ぬ", "nu", "n + u"),
            ("ね", "ne", "n + e"),
            ("の", "no", "n + o"),
        ],
        "pairs": [("なに", "na-ni"), ("ぬね", "nu-ne"), ("のに", "no-ni"), ("なの", "na-no")],
        "teaching_notes": [
            "This row follows the sound pattern very clearly, which makes it a good review row.",
            "に appears in many basic words, so it is useful to memorize early.",
            "Read each pair evenly and do not rush the second syllable.",
        ],
    },
    {
        "slug": "hiragana_set_06_h_row_worksheet.docx",
        "title": "Hiragana Set 6 — H Row",
        "subtitle": "は ひ ふ へ ほ",
        "focus": "The h row is mostly regular, but ふ sounds like fu instead of hu.",
        "kana": [
            ("は", "ha", "h + a"),
            ("ひ", "hi", "h + i"),
            ("ふ", "fu", "Soft f sound with u."),
            ("へ", "he", "h + e"),
            ("ほ", "ho", "h + o"),
        ],
        "pairs": [("はひ", "ha-hi"), ("ふへ", "fu-he"), ("ほは", "ho-ha"), ("ひふ", "hi-fu")],
        "teaching_notes": [
            "ふ is pronounced more softly than an English full f sound.",
            "Later, は and へ can sound different in particles, but here use the regular row sound.",
            "This worksheet focuses only on the basic row reading.",
        ],
    },
    {
        "slug": "hiragana_set_07_m_row_worksheet.docx",
        "title": "Hiragana Set 7 — M Row",
        "subtitle": "ま み む め も",
        "focus": "The m row is regular and easy to blend into simple syllable pairs.",
        "kana": [
            ("ま", "ma", "m + a"),
            ("み", "mi", "m + i"),
            ("む", "mu", "m + u"),
            ("め", "me", "m + e"),
            ("も", "mo", "m + o"),
        ],
        "pairs": [("まみ", "ma-mi"), ("むめ", "mu-me"), ("もま", "mo-ma"), ("みも", "mi-mo")],
        "teaching_notes": [
            "This is another strong confidence-building row because the sounds are very regular.",
            "Encourage students to read pairs in one smooth rhythm instead of stopping between kana.",
            "Quick oral drills work well with this row.",
        ],
    },
    {
        "slug": "hiragana_set_08_y_row_worksheet.docx",
        "title": "Hiragana Set 8 — Y Row",
        "subtitle": "や ゆ よ",
        "focus": "The y row has only three kana, so students can focus on hearing the vowel clearly.",
        "kana": [
            ("や", "ya", "y + a"),
            ("ゆ", "yu", "y + u"),
            ("よ", "yo", "y + o"),
        ],
        "pairs": [("やゆ", "ya-yu"), ("ゆよ", "yu-yo"), ("よや", "yo-ya"), ("やよ", "ya-yo")],
        "teaching_notes": [
            "There is no yi or ye in the basic hiragana chart.",
            "Because the row is short, it works well as a quick review worksheet.",
            "Focus on keeping ya, yu, and yo clearly different from one another.",
        ],
    },
    {
        "slug": "hiragana_set_09_r_row_worksheet.docx",
        "title": "Hiragana Set 9 — R Row",
        "subtitle": "ら り る れ ろ",
        "focus": "The Japanese r sound is light, somewhere between English r, l, and d.",
        "kana": [
            ("ら", "ra", "Light tapped r + a"),
            ("り", "ri", "Light tapped r + i"),
            ("る", "ru", "Light tapped r + u"),
            ("れ", "re", "Light tapped r + e"),
            ("ろ", "ro", "Light tapped r + o"),
        ],
        "pairs": [("らり", "ra-ri"), ("るれ", "ru-re"), ("ろら", "ro-ra"), ("りろ", "ri-ro")],
        "teaching_notes": [
            "The Japanese r is lighter than the English r sound.",
            "Students should try a quick tap with the tongue instead of a long hard r.",
            "Say the row slowly first, then build speed after the sound feels natural.",
        ],
    },
    {
        "slug": "hiragana_set_10_w_n_worksheet.docx",
        "title": "Hiragana Set 10 — W Row and ん",
        "subtitle": "わ を ん",
        "focus": "This worksheet introduces the last basic kana in the beginner chart.",
        "kana": [
            ("わ", "wa", "w + a"),
            ("を", "o", "Usually pronounced o in modern Japanese."),
            ("ん", "n", "A stand-alone n sound."),
        ],
        "pairs": [("わを", "wa-o"), ("をん", "o-n"), ("わん", "wa-n"), ("んわ", "n-wa")],
        "teaching_notes": [
            "を is mainly used as a grammar particle, but beginners still learn it in the chart.",
            "ん is special because it is a sound on its own, not a consonant + vowel pair.",
            "This worksheet is a good review point for the full beginner hiragana chart.",
        ],
    },
]


def _rgb_hex(rgb):
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def set_cell_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), _rgb_hex(rgb))
    tcPr.append(shd)


def set_para_bg(para, rgb):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), _rgb_hex(rgb))
    pPr.append(shd)


def add_run(para, text, bold=False, italic=False, size=12, color=DARK, font='Calibri'):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = font
    return r


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6 if level == 1 else 4)
    p.paragraph_format.space_after = Pt(2)
    add_run(p, text, bold=True, size=14 if level == 1 else 12,
            color=NAVY if level == 1 else TEAL)
    return p


def section_banner(doc, text, bg=NAVY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.3)
    add_run(p, text, bold=True, size=12, color=WHITE)
    set_para_bg(p, bg)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    p.add_run()._r.append(br)


def add_page_label(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    add_run(p, text, size=10, color=GREY)


def add_info_row(doc):
    info = doc.add_table(rows=1, cols=4)
    info.style = 'Table Grid'
    for i, (lbl, blank) in enumerate([
        ('Name', '________________________'),
        ('Class', '___________'),
        ('Date', '______________'),
        ('Period', '____'),
    ]):
        c = info.rows[0].cells[i]
        set_cell_bg(c, LIGHT)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, f'{lbl}: ', bold=True, size=10, color=NAVY)
        add_run(p, blank, size=10, color=DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def build_sound_bank(kana_items):
    rotated = kana_items[1:] + kana_items[:1]
    bank = []
    answer_lookup = {}
    for idx, (_, sound, _tip) in enumerate(rotated):
        label = chr(ord('A') + idx)
        bank.append((label, sound))
    for original, (_, sound, _tip) in zip(kana_items, kana_items):
        pass
    for label, sound in bank:
        for kana, item_sound, _tip in kana_items:
            if item_sound == sound:
                answer_lookup[kana] = label
                break
    return bank, answer_lookup


def get_row_pattern_note(config):
    sounds = [sound for _, sound, _tip in config['kana']]
    if config['title'].endswith('Vowels'):
        return "These five vowels are the sound base of the whole hiragana chart: a, i, u, e, o."
    if len(sounds) == 5:
        return f"Most full hiragana rows follow the vowel order a, i, u, e, o. In this set that sounds like: {', '.join(sounds)}."
    return f"This is a short row, so it does not use all five vowel spots. Memorize the order as: {', '.join(sounds)}."


def get_special_sound_note(config):
    sounds = [sound for _, sound, _tip in config['kana']]
    special_map = {
        'shi': 'し = shi',
        'chi': 'ち = chi',
        'tsu': 'つ = tsu',
        'fu': 'ふ = fu',
    }
    specials = [text for sound, text in special_map.items() if sound in sounds]
    if 'W Row and ん' in config['title']:
        return "This set has two special basics: を is usually read as o, and ん is its own sound by itself."
    if specials:
        return f"Watch the special reading in this set: {', '.join(specials)}."
    return "This set is very regular, so you can focus on the letter shape and the changing vowel sound."


def build_header(doc, config):
    tbl = doc.add_table(rows=2, cols=1)
    tbl.style = 'Table Grid'
    r0 = tbl.rows[0].cells[0]
    set_cell_bg(r0, NAVY)
    r0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p0 = r0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(5)
    p0.paragraph_format.space_after = Pt(2)
    add_run(p0, 'Beginner Japanese', bold=True, size=20, color=WHITE)
    add_run(p0, f"\n{config['title']}", size=13, color=GOLD)

    r1 = tbl.rows[1].cells[0]
    set_cell_bg(r1, TEAL)
    p1 = r1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(3)
    p1.paragraph_format.space_after = Pt(3)
    add_run(p1, 'Student Worksheet  ·  Page 1 of 2', size=10, color=WHITE)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def build_page1(doc, config):
    build_header(doc, config)
    add_info_row(doc)

    heading(doc, 'What you are learning')
    for tip in [
        f"This worksheet focuses on the set: {config['subtitle']}.",
        "Each hiragana letter stands for one sound chunk or beat. Read one kana smoothly instead of spelling it out like separate English letters.",
        config['focus'],
        get_row_pattern_note(config),
        get_special_sound_note(config),
        "Say the sound, copy the kana, then match the kana to the correct sound hint.",
    ]:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        add_run(p, tip, size=10)

    section_banner(doc, 'Sound Guide', bg=TEAL)
    guide_tbl = doc.add_table(rows=len(config['kana']) + 1, cols=3)
    guide_tbl.style = 'Table Grid'
    for ci, label in enumerate(['Kana', 'Sound', 'Beginner tip']):
        c = guide_tbl.rows[0].cells[ci]
        set_cell_bg(c, NAVY)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, label, bold=True, size=10, color=WHITE)

    for row_i, (kana, sound, tip) in enumerate(config['kana'], 1):
        bg = LIGHT if row_i % 2 == 1 else WHITE
        row = guide_tbl.rows[row_i]
        for ci, (val, clr, sz, bld) in enumerate([
            (kana, NAVY, 14, True),
            (sound, TEAL, 11, True),
            (tip, DARK, 9, False),
        ]):
            c = row.cells[ci]
            set_cell_bg(c, bg)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if ci < 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(p, val, bold=bld, size=sz, color=clr)

    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    section_banner(doc, 'Part 1 — Copy the Kana', bg=NAVY)

    copy_intro = doc.add_paragraph()
    copy_intro.paragraph_format.space_before = Pt(1)
    copy_intro.paragraph_format.space_after = Pt(2)
    add_run(copy_intro, 'Look at each kana. ', size=10, color=DARK)
    add_run(copy_intro, 'Say the sound out loud first', bold=True, size=10, color=TEAL)
    add_run(copy_intro, ' and focus on ', size=10, color=DARK)
    add_run(copy_intro, 'how it sounds', bold=True, size=10, color=NAVY)
    add_run(copy_intro, ', not on describing the symbol. ', size=10, color=DARK)
    add_run(copy_intro, 'Then copy the kana itself', bold=True, size=10, color=TEAL)
    add_run(copy_intro, ' on both blank lines.', size=10, color=DARK)

    copy_tbl = doc.add_table(rows=math.ceil(len(config['kana']) / 2), cols=2)
    copy_tbl.style = 'Table Grid'
    copy_tbl.autofit = False
    for row in copy_tbl.rows:
        for cell in row.cells:
            cell.width = Inches(4.0)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for idx, (kana, _sound, _tip) in enumerate(config['kana']):
        row_idx = idx // 2
        col_idx = idx % 2
        cell = copy_tbl.rows[row_idx].cells[col_idx]
        set_cell_bg(cell, MINT if row_idx % 2 == 0 else WHITE)
        p0 = cell.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(0)
        add_run(p0, kana, bold=True, size=20, color=NAVY)

        for line_no in range(1, 3):
            pl = cell.add_paragraph()
            pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pl.paragraph_format.space_before = Pt(0)
            pl.paragraph_format.space_after = Pt(0)
            add_run(pl, f'Copy {line_no}: ', bold=True, size=8, color=TEAL)
            add_run(pl, '_' * 15, size=9, color=DARK)

    if len(config['kana']) % 2 == 1:
        copy_tbl.rows[-1].cells[-1].text = ''

    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    section_banner(doc, 'Part 2 — Match the Sound', bg=TEAL)
    bank, _answer_lookup = build_sound_bank(config['kana'])

    bank_tbl = doc.add_table(rows=1, cols=len(bank))
    bank_tbl.style = 'Table Grid'
    for idx, (label, sound) in enumerate(bank):
        cell = bank_tbl.rows[0].cells[idx]
        set_cell_bg(cell, LIGHT)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, f'{label}. ', bold=True, size=10, color=TEAL)
        add_run(p, sound, size=9, color=DARK)

    match_tbl = doc.add_table(rows=len(config['kana']), cols=2)
    match_tbl.style = 'Table Grid'
    for idx, (kana, _sound, _tip) in enumerate(config['kana']):
        bg = LIGHT if idx % 2 == 0 else WHITE
        row = match_tbl.rows[idx]
        left = row.cells[0]
        right = row.cells[1]
        set_cell_bg(left, bg)
        set_cell_bg(right, bg)
        left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_left = left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_left.paragraph_format.space_before = Pt(2)
        p_left.paragraph_format.space_after = Pt(2)
        add_run(p_left, kana, bold=True, size=14, color=NAVY)

        p_right = right.paragraphs[0]
        p_right.paragraph_format.space_before = Pt(2)
        p_right.paragraph_format.space_after = Pt(2)
        add_run(p_right, 'Letter: ', bold=True, size=10, color=TEAL)
        add_run(p_right, '____', size=10, color=DARK)


def build_page2(doc, config):
    page_break(doc)
    add_page_label(doc, 'Reference & Answer Key  ·  Page 2 of 2')
    section_banner(doc, 'Part 3 — Read and Check', bg=DARK)

    intro = doc.add_paragraph()
    intro.paragraph_format.space_before = Pt(1)
    intro.paragraph_format.space_after = Pt(2)
    add_run(intro, 'Use this page as a reference after students finish page 1. Read each kana left to right and treat each one as one clean sound beat.', size=10, color=DARK)

    pair_tbl = doc.add_table(rows=len(config['pairs']) + 1, cols=2)
    pair_tbl.style = 'Table Grid'
    for ci, label in enumerate(['Practice pair', 'Reading']):
        c = pair_tbl.rows[0].cells[ci]
        set_cell_bg(c, NAVY)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, label, bold=True, size=10, color=WHITE)

    for idx, (pair, reading) in enumerate(config['pairs'], 1):
        bg = LIGHT if idx % 2 == 1 else WHITE
        row = pair_tbl.rows[idx]
        for ci, (val, clr, sz, bld) in enumerate([
            (pair, NAVY, 12, True),
            (reading, DARK, 10, False),
        ]):
            c = row.cells[ci]
            set_cell_bg(c, bg)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            add_run(p, val, bold=bld, size=sz, color=clr)

    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    section_banner(doc, 'Part 2 Answer Key', bg=TEAL)
    bank, answer_lookup = build_sound_bank(config['kana'])

    answer_tbl = doc.add_table(rows=len(config['kana']) + 1, cols=3)
    answer_tbl.style = 'Table Grid'
    for ci, label in enumerate(['Kana', 'Correct letter', 'Sound']):
        c = answer_tbl.rows[0].cells[ci]
        set_cell_bg(c, TEAL)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_run(p, label, bold=True, size=10, color=WHITE)

    for idx, (kana, sound, _tip) in enumerate(config['kana'], 1):
        bg = MINT if idx % 2 == 1 else WHITE
        row = answer_tbl.rows[idx]
        for ci, (val, clr, sz, bld) in enumerate([
            (kana, NAVY, 13, True),
            (answer_lookup[kana], TEAL, 10, True),
            (sound, DARK, 10, False),
        ]):
            c = row.cells[ci]
            set_cell_bg(c, bg)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            add_run(p, val, bold=bld, size=sz, color=clr)

    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    heading(doc, 'How to study this set', level=2)
    for tip in [
        "Point to one letter at a time and say the full sound in one step, such as ka or shi.",
        "When you read a practice pair, do not stop too long between the two kana.",
        "After checking the answer key, cover the sound column and quiz yourself again from memory.",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        add_run(p, '▸ ', bold=True, size=10, color=GOLD)
        add_run(p, tip, size=10, color=DARK)

    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    heading(doc, 'Beginner notes')
    for tip in config['teaching_notes']:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        add_run(p, '▸ ', bold=True, size=10, color=GOLD)
        add_run(p, tip, size=10, color=DARK)


def build_doc(config):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.2)
        section.bottom_margin = Inches(0.2)
        section.left_margin = Inches(0.2)
        section.right_margin = Inches(0.2)

    build_page1(doc, config)
    build_page2(doc, config)
    return doc


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    docs_dir = os.path.dirname(script_dir)
    output_dir = os.path.join(docs_dir, 'output', 'hiragana_series')
    os.makedirs(output_dir, exist_ok=True)

    for config in WORKSHEETS:
        doc = build_doc(config)
        out = os.path.join(output_dir, config['slug'])
        doc.save(out)
        print(f'Saved: {out}')


if __name__ == '__main__':
    main()
