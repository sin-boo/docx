# reusre-maker

Python project for generating DOCX documents and matching PDF exports. Includes Japanese worksheet generators.

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
- **Japanese greetings worksheet:** `cd japanese_docs && python3 japanese_greetings_worksheet.py` → creates `japanese_docs/output/docs/japanese_greetings_worksheet.docx` and `japanese_docs/output/pdf/japanese_greetings_worksheet.pdf`
- **Japanese numbers worksheet:** `cd japanese_docs && python3 japanese_numbers_worksheet.py` → creates `japanese_docs/output/docs/japanese_numbers_worksheet.docx` and `japanese_docs/output/pdf/japanese_numbers_worksheet.pdf`

## Japanese docs output structure

- `japanese_docs/output/docs/` — generated `.docx` files
- `japanese_docs/output/pdf/` — generated `.pdf` files from LibreOffice
- `japanese_docs/output/docs/legacy/` — older generated `.docx` files kept for reference

## PDF generation

- PDF export now requires **LibreOffice / soffice**
- there is **no fallback PDF generator**
- if `soffice` is missing, worksheet export will fail with an explicit error

## Layout

- Page 1: Student worksheet (banner, info, objectives, instructions, Part 1 questions)
- Page 2: Part 2 roleplay, pronunciation tips
- Page 3: Answer key (teacher copy), notes, printing tips
