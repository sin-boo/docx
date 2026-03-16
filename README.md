# reusre-maker

Python project for generating DOCX documents. Includes a Japanese greetings worksheet (student + answer key).

## Setup

```bash
pip install -r requirements.txt
```

## Usage

- **Simple docx:** `python create_docx.py` → creates `output.docx`
- **Japanese greetings worksheet:** `cd japanese_docs && python japanese_greetings_worksheet.py` → creates `japanese_docs/japanese_greetings_worksheet.docx`

## Layout

- Page 1: Student worksheet (banner, info, objectives, instructions, Part 1 questions)
- Page 2: Part 2 roleplay, pronunciation tips
- Page 3: Answer key (teacher copy), notes, printing tips
