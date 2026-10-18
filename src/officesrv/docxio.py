from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def new_doc():
    return Document()


def open_doc(path):
    return Document(path)


def add_para(doc, text, size=None, bold=False, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    if size:
        run.font.size = Pt(size)
    run.bold = bold
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return p


def add_table(doc, rows, header=False, widths=None):
    tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
    if header:
        for cell in tbl.rows[0].cells:
            for par in cell.paragraphs:
                for run in par.runs:
                    run.bold = True
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            tbl.cell(r, c).text = str(val)
    if widths:
        for c, w in enumerate(widths):
            for row in tbl.rows:
                row.cells[c].width = Cm(w)
    return tbl


def doc_text(doc, max_chars=20000):
    chunks = [p.text for p in doc.paragraphs if p.text.strip()]
    for tbl in doc.tables:
        for row in tbl.rows:
            chunks.append(" | ".join(c.text for c in row.cells))
    return "\n".join(chunks)[:max_chars]


def replace_text(doc, old, new):
    hits = 0
    for p in doc.paragraphs:
        if old in p.text:
            for run in p.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
                    hits += 1
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old in p.text:
                        for run in p.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, new)
                                hits += 1
    return hits
