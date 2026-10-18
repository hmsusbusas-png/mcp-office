import logging
import sys
import time
from functools import wraps
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from . import comoffice, docxio, pptxio, xlsxio

logging.basicConfig(stream=sys.stderr, level=logging.INFO, format="%(levelname)s %(name)s %(message)s")
log = logging.getLogger("office")

app = FastMCP("office", instructions=(
    "Word, Excel and PowerPoint files. docx, xlsx and pptx are written directly without Office. "
    "office_to_pdf and xlsx_recalculate start a hidden Microsoft Office instance. Paths are absolute Windows paths. "
    "Editing tools change the file in place. Formulas written with openpyxl have no cached values until xlsx_recalculate runs."
))


def tool(description, read_only=False, destructive=False):
    hints = ToolAnnotations(readOnlyHint=read_only, destructiveHint=destructive, idempotentHint=read_only, openWorldHint=False)

    def register(fn):
        @wraps(fn)
        def timed(*args, **kwargs):
            start = time.monotonic()
            try:
                return fn(*args, **kwargs)
            finally:
                log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
        app.tool(description=description, annotations=hints)(timed)
        return fn
    return register


def new_target(path, suffix):
    target = Path(path).expanduser().resolve()
    if target.suffix.lower() != suffix:
        raise ValueError(f"path must end with {suffix}")
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


block_help = ("blocks: [{text, style (Title, Heading 1..9, List Bullet, List Number, Quote, Normal), size, bold, italic, "
              "color #RRGGBB, font, align}] or {image, width_cm} or {table: rows} or {page_break: true}. "
              "Inline **bold**, *italic* and `code` work in text")


@tool(f"Create a docx from blocks. {block_help}")
def docx_create(path: str, blocks: list[dict[str, Any]]) -> dict:
    target = new_target(path, ".docx")
    doc = docxio.Document()
    for spec in blocks:
        docxio.add_block(doc, spec)
    doc.save(target)
    return {"path": str(target), **docxio.stats(doc)}


@tool("Create a docx from Markdown: # headings, - and 1. lists, | tables |, > quotes, ![](image path), --- page break, **bold** *italic* `code`")
def docx_from_markdown(path: str, markdown: str) -> dict:
    target = new_target(path, ".docx")
    doc = docxio.Document()
    docxio.from_markdown(doc, markdown)
    doc.save(target)
    return {"path": str(target), **docxio.stats(doc)}


@tool("Read a docx as Markdown in document order: headings, lists, tables. Also returns counts and the heading outline", read_only=True)
def docx_read(path: str, max_chars: int = 30000) -> dict:
    doc = docxio.open_doc(path)
    return {"text": docxio.doc_text(doc, max_chars), **docxio.stats(doc)}


@tool(f"Edit a docx in place: replace [[old, new]] in body, tables, headers and footers, delete paragraphs containing text, append blocks or markdown. {block_help}", destructive=True)
def docx_edit(path: str, replace: list[list[str]] | None = None, delete_containing: str = "",
              append: list[dict[str, Any]] | None = None, append_markdown: str = "") -> dict:
    for pair in replace or []:
        if len(pair) != 2:
            raise ValueError("each replace item is [old, new]")
    doc = docxio.open_doc(path)
    hits = sum(docxio.replace_text(doc, old, new) for old, new in replace or [])
    gone = docxio.delete_paragraphs(doc, delete_containing) if delete_containing else 0
    for spec in append or []:
        docxio.add_block(doc, spec)
    if append_markdown:
        docxio.from_markdown(doc, append_markdown)
    doc.save(path)
    return {"replaced": hits, "deleted": gone, "appended": len(append or []) + bool(append_markdown)}


@tool("Append a table to a docx, first row bold when header is true, optional column widths in cm")
def docx_table(path: str, rows: list[list[Any]], header: bool = True, widths_cm: list[float] | None = None) -> str:
    doc = docxio.open_doc(path)
    tbl = docxio.add_table(doc, rows, header, widths_cm)
    doc.save(path)
    return f"table {len(tbl.rows)}x{len(tbl.columns)} added"


@tool("Append an image (png, jpg, gif, bmp, tiff) to a docx, optional width in cm")
def docx_image(path: str, image: str, width_cm: float | None = None, align: str = "center") -> str:
    doc = docxio.open_doc(path)
    docxio.add_image(doc, image, width_cm, align)
    doc.save(path)
    return f"image added, {len(doc.inline_shapes)} images in document"


@tool("Page layout of every section: size A4, A3, A5, Letter, landscape, margins in cm, header and footer text, page numbers in footer")
def docx_page(path: str, size: str = "", landscape: bool | None = None, margins_cm: float | None = None,
              header: str | None = None, footer: str | None = None, page_numbers: bool = False) -> dict:
    doc = docxio.open_doc(path)
    layout = docxio.page_setup(doc, size, landscape, margins_cm)
    docxio.header_footer(doc, header, footer, page_numbers)
    doc.save(path)
    return layout


@tool("Create an xlsx: sheets is {name: rows}. Strings starting with = are formulas. First row is styled as a header, columns are sized")
def xlsx_create(path: str, sheets: dict[str, list[list[Any]]], freeze_header: bool = True) -> str:
    if not sheets:
        raise ValueError("give at least one sheet")
    target = new_target(path, ".xlsx")
    book = xlsxio.new_book()
    for name, grid in sheets.items():
        ws = book.create_sheet(name[:31])
        xlsxio.write_rows(ws, grid)
        if grid:
            xlsxio.style_header(ws)
            xlsxio.fit_widths(ws)
            if freeze_header:
                ws.freeze_panes = "A2"
    book.save(target)
    return str(target)


@tool("Read a sheet or a range like A1:D20. formulas=true returns formula text, otherwise values cached by Excel", read_only=True)
def xlsx_read(path: str, sheet: str = "", cell_range: str = "", max_cells: int = 10000, formulas: bool = False) -> dict:
    book = xlsxio.open_book(path, values_only=not formulas)
    ws = xlsxio.sheet_of(book, sheet)
    return {"sheet": ws.title, "sheets": book.sheetnames, "dimensions": ws.dimensions,
            "rows": xlsxio.read_rows(ws, cell_range, max_cells)}


@tool("Write rows into a sheet. mode append adds below data, replace clears the sheet, at writes from start_cell like C5", destructive=True)
def xlsx_write(path: str, sheet: str, rows: list[list[Any]], mode: str = "append", start_cell: str = "A1") -> str:
    if mode not in ("append", "replace", "at"):
        raise ValueError("mode is append, replace or at")
    book = xlsxio.open_book(path)
    ws = book[sheet] if sheet in book.sheetnames else book.create_sheet(sheet[:31])
    if mode == "replace":
        ws.delete_rows(1, ws.max_row)
    if mode == "at":
        anchor = ws[start_cell]
        row, col = anchor.row, anchor.column
    else:
        has_data = ws.max_row > 1 or ws.cell(1, 1).value is not None
        row, col = (ws.max_row + 1 if has_data else 1), 1
    xlsxio.write_rows(ws, rows, row, col)
    book.save(path)
    return f"{len(rows)} rows written to {ws.title} from {ws.cell(row, col).coordinate}"


@tool("Set individual cells {B2: 5, C2: '=B2*2'} with optional number formats {C2: '0.00%'}", destructive=True)
def xlsx_cells(path: str, sheet: str, cells: dict[str, Any], number_formats: dict[str, str] | None = None) -> str:
    book = xlsxio.open_book(path)
    ws = book[sheet] if sheet in book.sheetnames else book.create_sheet(sheet[:31])
    for ref, val in cells.items():
        ws[ref] = val
    for ref, fmt in (number_formats or {}).items():
        ws[ref].number_format = fmt
    book.save(path)
    return f"{len(cells)} cells set on {ws.title}"


@tool("Format a sheet: header style and color, borders, column widths, freeze panes, autofilter, number formats per column {B: '#,##0.00'}")
def xlsx_format(path: str, sheet: str, header_rows: int = 1, header_color: str = "#4472C4", borders: bool = True,
                freeze: str = "A2", autofilter: bool = True, column_formats: dict[str, str] | None = None) -> str:
    book = xlsxio.open_book(path)
    ws = xlsxio.sheet_of(book, sheet)
    if header_rows:
        xlsxio.style_header(ws, header_rows, header_color)
    if borders:
        xlsxio.grid_borders(ws)
    if column_formats:
        xlsxio.column_formats(ws, column_formats, header_rows)
    xlsxio.fit_widths(ws)
    ws.freeze_panes = freeze or None
    if autofilter and ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    book.save(path)
    return f"{ws.title} formatted, {ws.dimensions}"


@tool("Add a native Excel chart: kind bar, column, line, pie, scatter. label_col is the category column letter, value_cols the series letters, first_row holds the series names")
def xlsx_chart(path: str, sheet: str, kind: str, label_col: str, value_cols: list[str], first_row: int = 1,
               last_row: int = 0, title: str = "", anchor: str = "", y_title: str = "", x_title: str = "") -> str:
    book = xlsxio.open_book(path)
    ws = xlsxio.sheet_of(book, sheet)
    count = xlsxio.chart(ws, kind, label_col.upper(), [c.upper() for c in value_cols], first_row, last_row, title, anchor, y_title, x_title)
    book.save(path)
    return f"{kind} chart added, {count} charts on {ws.title}"


@tool("Sheets: action list, add, rename, delete or copy. Name is the sheet, new_name the target for rename and copy", destructive=True)
def xlsx_sheets(path: str, action: str = "list", name: str = "", new_name: str = "") -> list[str]:
    book = xlsxio.open_book(path)
    if action == "add":
        book.create_sheet(name[:31])
    elif action == "rename":
        xlsxio.sheet_of(book, name).title = new_name[:31]
    elif action == "delete":
        if len(book.sheetnames) == 1:
            raise ValueError("a workbook needs at least one sheet")
        book.remove(xlsxio.sheet_of(book, name))
    elif action == "copy":
        book.copy_worksheet(xlsxio.sheet_of(book, name)).title = (new_name or name + " copy")[:31]
    elif action != "list":
        raise ValueError("action is list, add, rename, delete or copy")
    if action != "list":
        book.save(path)
    return book.sheetnames


@tool("Import a csv into a sheet (creates the xlsx if missing). Delimiter is detected, numbers are converted")
def xlsx_import_csv(csv_path: str, xlsx_path: str, sheet: str = "data", delimiter: str = "", encoding: str = "utf-8-sig") -> dict:
    target = Path(xlsx_path).expanduser().resolve()
    book = xlsxio.open_book(target) if target.is_file() else xlsxio.new_book()
    if sheet in book.sheetnames:
        book.remove(book[sheet])
    ws = book.create_sheet(sheet[:31])
    rows, sep = xlsxio.import_csv(ws, csv_path, delimiter, encoding)
    if rows:
        xlsxio.style_header(ws)
        xlsxio.fit_widths(ws)
        ws.freeze_panes = "A2"
    target.parent.mkdir(parents=True, exist_ok=True)
    book.save(target)
    return {"path": str(target), "sheet": ws.title, "rows": rows, "delimiter": sep}


@tool("Export a sheet to csv with cached values, utf-8 with BOM by default so Excel opens Cyrillic correctly")
def xlsx_export_csv(xlsx_path: str, csv_path: str, sheet: str = "", delimiter: str = ";", encoding: str = "utf-8-sig") -> dict:
    ws = xlsxio.sheet_of(xlsxio.open_book(xlsx_path, values_only=True), sheet)
    return {"path": csv_path, "rows": xlsxio.export_csv(ws, csv_path, delimiter, encoding)}


@tool("Recalculate all formulas in Excel and save, so cached values become readable. Needs Microsoft Excel", destructive=True)
def xlsx_recalculate(path: str) -> dict:
    return comoffice.recalculate(path)


@tool("Create a pptx. slides: [{title, subtitle, bullets: [text or {text, level}], image, image_width_cm, table: rows, notes, layout: title|content|section|two|title_only|blank, font_size}]")
def pptx_create(path: str, slides: list[dict[str, Any]], widescreen: bool = True) -> dict:
    target = new_target(path, ".pptx")
    deck = pptxio.Presentation()
    if widescreen:
        deck.slide_width, deck.slide_height = pptxio.Cm(33.867), pptxio.Cm(19.05)
    for spec in slides:
        pptxio.add_slide(deck, spec)
    deck.save(target)
    return {"path": str(target), "slides": len(deck.slides)}


@tool("Append slides to an existing pptx, same slide format as pptx_create")
def pptx_add(path: str, slides: list[dict[str, Any]]) -> dict:
    deck = pptxio.open_deck(path)
    for spec in slides:
        pptxio.add_slide(deck, spec)
    deck.save(path)
    return {"path": path, "slides": len(deck.slides)}


@tool("Read text, tables and speaker notes of every slide", read_only=True)
def pptx_read(path: str, max_chars: int = 30000) -> dict:
    deck = pptxio.open_deck(path)
    return {"slides": len(deck.slides), "text": pptxio.deck_text(deck, max_chars)}


@tool("Convert docx, xlsx, pptx and older Office formats to PDF via Word, Excel or PowerPoint")
def office_to_pdf(path: str, out_path: str = "") -> dict:
    return comoffice.to_pdf(path, out_path)


def run():
    app.run(transport="stdio")
