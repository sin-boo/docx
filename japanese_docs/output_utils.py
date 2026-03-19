import os
import shutil
import subprocess


def require_soffice():
    for name in ('soffice', 'libreoffice'):
        path = shutil.which(name)
        if path:
            return path
    raise RuntimeError(
        'LibreOffice is required for PDF export. Install it and make sure '
        '"soffice" is available on PATH.'
    )


def get_output_paths(script_dir, filename, subdir=None):
    docs_dir = os.path.join(script_dir, 'output', 'docs')
    pdf_dir = os.path.join(script_dir, 'output', 'pdf')
    if subdir:
        docs_dir = os.path.join(docs_dir, subdir)
        pdf_dir = os.path.join(pdf_dir, subdir)

    os.makedirs(docs_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)

    docx_path = os.path.join(docs_dir, filename)
    pdf_path = os.path.join(pdf_dir, os.path.splitext(filename)[0] + '.pdf')
    return docx_path, pdf_path


def convert_docx_to_pdf(docx_path, pdf_path):
    soffice = require_soffice()
    pdf_dir = os.path.dirname(pdf_path)
    cmd = [
        soffice,
        '--headless',
        '--convert-to',
        'pdf:writer_pdf_Export',
        '--outdir',
        pdf_dir,
        docx_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f'LibreOffice PDF conversion failed for {docx_path}\n'
            f'stdout:\n{result.stdout}\n\nstderr:\n{result.stderr}'
        )

    generated_pdf = os.path.join(
        pdf_dir,
        os.path.splitext(os.path.basename(docx_path))[0] + '.pdf',
    )
    if not os.path.exists(generated_pdf):
        raise RuntimeError(
            f'LibreOffice reported success but no PDF was created for {docx_path}.'
        )
    if os.path.abspath(generated_pdf) != os.path.abspath(pdf_path):
        os.replace(generated_pdf, pdf_path)


def save_docx_and_pdf(doc, script_dir, filename, subdir=None):
    docx_path, pdf_path = get_output_paths(script_dir, filename, subdir=subdir)
    doc.save(docx_path)
    convert_docx_to_pdf(docx_path, pdf_path)
    return docx_path, pdf_path
