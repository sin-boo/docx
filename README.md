# reusre-maker

Python project for generating DOCX documents. Includes Japanese worksheet generators for greetings, dialogues, and numbers.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

- **Simple docx:** `python create_docx.py` → creates `output.docx`
- **Japanese greetings worksheet:** `python3 japanese_docs/generators/japanese_greetings_worksheet.py` → creates `japanese_docs/output/japanese_greetings_worksheet.docx`
- **Japanese dialogue worksheet:** `python3 japanese_docs/generators/japanese_dialogue_worksheet.py` → creates `japanese_docs/output/japanese_dialogue_worksheet.docx`
- **Japanese numbers worksheet:** `python3 japanese_docs/generators/japanese_numbers_worksheet.py` → creates `japanese_docs/output/japanese_numbers_worksheet.docx`
- **Japanese hiragana foundations series:** `python3 japanese_docs/generators/japanese_hiragana_foundations_series.py` → creates 10 worksheets in `japanese_docs/output/hiragana_series/`

## Japanese docs folder structure

- `japanese_docs/generators/` — Python worksheet generator scripts
- `japanese_docs/output/` — generated `.docx` files
- `japanese_docs/output/hiragana_series/` — 10 early-beginner hiragana worksheets
- `japanese_docs/output/legacy/` — older generated files kept for reference
- `japanese_docs/ai_instructions.json` — project-specific editing notes

## Layout

- Page 1: Student worksheet (banner, info, objectives, instructions, Part 1 questions)
- Page 2: Part 2 roleplay, pronunciation tips
- Page 3: Answer key (teacher copy), notes, printing tips
