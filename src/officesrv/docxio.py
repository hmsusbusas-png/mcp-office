import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

aligns = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
paper = {"A4": (21.0, 29.7), "A3": (29.7, 42.0), "A5": (14.8, 21.0), "Letter": (21.59, 27.94)}
inline = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")


def open_doc(path):
    if not Path(path).is_file():
        raise FileNotFoundError(f"no such file: {path}")
    try:
        return Document(path)
    except Exception as exc:
        raise ValueError(f"cannot open {path} as docx: {exc}") from exc


def color_of(hex_color):
    text = hex_color.lstrip("#")
    if len(text) != 6:
        raise ValueError(f"color must look like #RRGGBB, got {hex_color!r}")
    return RGBColor.from_string(text.upper())


def styled_runs(p, text, size=None, bold=False, italic=False, color=None, font=None):
    for chunk in inline.split(text):
        if not chunk:
            continue
        piece = p.add_run()
        if chunk.startswith("**") and chunk.endswith("**"):
            piece.text, piece.bold = chunk[2:-2], True
        elif chunk.startswith("`") and chunk.endswith("`"):
            piece.text = chunk[1:-1]
            piece.font.name = "Consolas"
        elif chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            piece.text, piece.italic = chunk[1:-1], True
        else:
            piece.text = chunk
        if bold:
            piece.bold = True
        if italic:
            piece.italic = True
        if size:
            piece.font.size = Pt(size)
        if color:
            piece.font.color.rgb = color_of(color)
        if font:
            piece.font.name = font


def add_block(doc, spec):
    if not isinstance(spec, dict):
        spec = {"text": str(spec)}
    if spec.get("page_break"):
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        return
    if spec.get("image"):
        add_image(doc, spec["image"], spec.get("width_cm"), spec.get("align", "center"))
        return
    if spec.get("table"):
        add_table(doc, spec["table"], spec.get("header", True), spec.get("widths_cm"))
        return
    style = spec.get("style") or ""
    level = re.fullmatch(r"(?i)heading\s*(\d)", style)
    if level:
        p = doc.add_heading(level=int(level.group(1)))
    elif style.lower() == "title":
        p = doc.add_heading(level=0)
    else:
        try:
            p = doc.add_paragraph(style=style or None)
        except KeyError as exc:
            raise ValueError(f"unknown style {style!r}, try Normal, List Bullet, List Number, Quote, Heading 1..9") from exc
    styled_runs(p, spec.get("text", ""), spec.get("size"), spec.get("bold"), spec.get("italic"), spec.get("color"), spec.get("font"))
    if spec.get("align"):
        if spec["align"] not in aligns:
            raise ValueError(f"align is one of {', '.join(aligns)}")
        p.alignment = aligns[spec["align"]]


def add_image(doc, image, width_cm=None, align="center"):
    if not Path(image).is_file():
        raise FileNotFoundError(f"no such image: {image}")
    doc.add_picture(str(image), width=Cm(width_cm) if width_cm else None)
    doc.paragraphs[-1].alignment = aligns.get(align, WD_ALIGN_PARAGRAPH.CENTER)


def add_table(doc, grid, header=True, widths=None):
    if not grid or not any(grid):
        raise ValueError("table needs at least one non empty row")
    ncols = max(len(r) for r in grid)
    tbl = doc.add_table(rows=len(grid), cols=ncols)
    tbl.style = "Table Grid"
    for r, row in enumerate(grid):
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.text = "" if c >= len(row) or row[c] is None else str(row[c])
            if header and r == 0:
                for piece in cell.paragraphs[0].runs:
                    piece.bold = True
    if widths:
        for c, w in enumerate(widths[:ncols]):
            for row in tbl.rows:
                row.cells[c].width = Cm(w)
    return tbl


def from_markdown(doc, text):
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("|"):
            grid = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                    grid.append(cells)
                i += 1
            add_table(doc, grid)
            continue
        heading = re.match(r"(#{1,6})\s+(.*)", stripped)
        bullet = re.match(r"[-*+]\s+(.*)", stripped)
        number = re.match(r"\d+[.)]\s+(.*)", stripped)
        picture = re.match(r"!\[[^\]]*\]\(([^)]+)\)", stripped)
        if heading:
            add_block(doc, {"text": heading.group(2), "style": f"Heading {len(heading.group(1))}"})
        elif stripped in ("---", "***", "\\pagebreak"):
            add_block(doc, {"page_break": True})
        elif picture:
            add_image(doc, picture.group(1), 15)
        elif bullet:
            add_block(doc, {"text": bullet.group(1), "style": "List Bullet"})
        elif number:
            add_block(doc, {"text": number.group(1), "style": "List Number"})
        elif stripped.startswith(">"):
            add_block(doc, {"text": stripped.lstrip("> "), "style": "Quote"})
        else:
            para = [stripped]
            while i + 1 < len(lines) and lines[i + 1].strip() and not re.match(r"(#|[-*+]\s|\d+[.)]\s|\||>|!\[)", lines[i + 1].strip()):
                i += 1
                para.append(lines[i].strip())
            add_block(doc, {"text": " ".join(para)})
        i += 1


def page_number_field(paragraph):
    piece = paragraph.add_run()
    for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            node = OxmlElement("w:fldChar")
            node.set(qn("w:fldCharType"), kind)
        else:
            node = OxmlElement("w:instrText")
            node.set(qn("xml:space"), "preserve")
            node.text = text
        piece._r.append(node)


def header_footer(doc, header, footer, page_numbers):
    for section in doc.sections:
        if header is not None:
            section.header.paragraphs[0].text = header
        if footer is not None or page_numbers:
            p = section.footer.paragraphs[0]
            p.text = footer or ""
            if page_numbers:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if footer:
                    p.add_run("  ")
                page_number_field(p)


def page_setup(doc, size, landscape, margins_cm):
    for section in doc.sections:
        if size:
            if size not in paper:
                raise ValueError(f"size is one of {', '.join(paper)}")
            w, h = paper[size]
            section.page_width, section.page_height = Cm(w), Cm(h)
        if landscape is not None:
            w, h = section.page_width, section.page_height
            section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
            if (landscape and w < h) or (not landscape and w > h):
                section.page_width, section.page_height = h, w
        if margins_cm is not None:
            section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(margins_cm)
    s = doc.sections[0]
    return {"width_cm": round(s.page_width.cm, 2), "height_cm": round(s.page_height.cm, 2),
            "landscape": s.orientation == WD_ORIENT.LANDSCAPE, "margin_cm": round(s.left_margin.cm, 2)}


def doc_text(doc, max_chars):
    chunks = []
    tables = iter(doc.tables)
    paras = iter(doc.paragraphs)
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = next(paras)
            if not p.text.strip():
                continue
            name = p.style.name if p.style is not None else ""
            level = re.fullmatch(r"Heading (\d)", name)
            prefix = "#" * int(level.group(1)) + " " if level else "- " if "List" in name else ""
            chunks.append(prefix + p.text)
        elif tag == "tbl":
            rows = [[c.text.replace("\n", " ") for c in row.cells] for row in next(tables).rows]
            if rows:
                chunks.append("| " + " | ".join(rows[0]) + " |")
                chunks.append("|" + "---|" * len(rows[0]))
                chunks.extend("| " + " | ".join(r) + " |" for r in rows[1:])
    text = "\n".join(chunks)
    return text if len(text) <= max_chars else text[:max_chars] + f"\n[cut at {max_chars} of {len(text)} chars]"


def stats(doc):
    words = sum(len(p.text.split()) for p in doc.paragraphs)
    return {"paragraphs": len(doc.paragraphs), "tables": len(doc.tables), "images": len(doc.inline_shapes),
            "sections": len(doc.sections), "words": words,
            "headings": [p.text for p in doc.paragraphs if p.style is not None and p.style.name.startswith("Heading")][:50]}


def all_paragraphs(doc):
    yield from doc.paragraphs
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in doc.sections:
        yield from section.header.paragraphs
        yield from section.footer.paragraphs


def replace_text(doc, old, new):
    if not old:
        raise ValueError("text to replace is empty")
    hits = 0
    for p in all_paragraphs(doc):
        if old not in p.text:
            continue
        hits += p.text.count(old)
        if any(old in piece.text for piece in p.runs):
            for piece in p.runs:
                piece.text = piece.text.replace(old, new)
        else:
            # match spans several runs, keep the first run formatting
            merged = p.text.replace(old, new)
            for piece in p.runs[1:]:
                piece.text = ""
            p.runs[0].text = merged
    return hits


def delete_paragraphs(doc, containing):
    gone = 0
    for p in list(doc.paragraphs):
        if containing in p.text:
            p._element.getparent().remove(p._element)
            gone += 1
    return gone
