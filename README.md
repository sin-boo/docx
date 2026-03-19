# reusre-maker

Python project for generating DOCX documents and matching PDF exports. Includes Japanese worksheet generators for greetings, dialogues, hiragana, and numbers.

## Setup

```bash
./setup.sh
```

Manual setup:

```bash
python3 -m pip install -r requirements.txt
sudo apt-get update
sudo apt-get install -y libreoffice-core libreoffice-writer
```

## Usage

- **Simple docx:** `python create_docx.py` → creates `output.docx`
- **Japanese greetings worksheet:** `python3 japanese_docs/generators/japanese_greetings_worksheet.py` → creates `japanese_docs/output/docs/japanese_greetings_worksheet.docx` and `japanese_docs/output/pdf/japanese_greetings_worksheet.pdf`
- **Japanese dialogue worksheet:** `python3 japanese_docs/generators/japanese_dialogue_worksheet.py` → creates `japanese_docs/output/docs/japanese_dialogue_worksheet.docx` and `japanese_docs/output/pdf/japanese_dialogue_worksheet.pdf`
- **Japanese numbers worksheet:** `python3 japanese_docs/generators/japanese_numbers_worksheet.py` → creates `japanese_docs/output/docs/japanese_numbers_worksheet.docx` and `japanese_docs/output/pdf/japanese_numbers_worksheet.pdf`
- **Japanese hiragana foundations series:** `python3 japanese_docs/generators/japanese_hiragana_foundations_series.py` → creates 10 DOCX files in `japanese_docs/output/docs/hiragana_series/` and 10 PDFs in `japanese_docs/output/pdf/hiragana_series/`

## Japanese docs folder structure

- `japanese_docs/generators/` — Python worksheet generator scripts
- `japanese_docs/output/docs/` — generated `.docx` files
- `japanese_docs/output/pdf/` — generated `.pdf` files
- `japanese_docs/output/docs/hiragana_series/` — 10 early-beginner hiragana DOCX worksheets
- `japanese_docs/output/pdf/hiragana_series/` — 10 early-beginner hiragana PDF worksheets
- `japanese_docs/output/docs/legacy/` — older generated files kept for reference
- `japanese_docs/ai_instructions.json` — project-specific editing notes

## PDF generation

- PDF export now requires **LibreOffice / soffice**
- there is **no fallback PDF generator**
- if `soffice` is missing, worksheet export fails with an explicit error

## Layout

- Page 1: Student worksheet (banner, info, objectives, instructions, Part 1 questions)
- Page 2: Part 2 roleplay, pronunciation tips
- Page 3: Answer key (teacher copy), notes, printing tips
