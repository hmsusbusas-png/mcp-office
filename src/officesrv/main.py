import logging
import sys
import time
from functools import wraps
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import docxio, pdfout, xlsxio

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("office")

app = FastMCP("office")


def timed(fn):
    @wraps(fn)
    def wrap(*args, **kwargs):
        start = time.monotonic()
        try:
            return fn(*args, **kwargs)
        finally:
            log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
    return wrap


def fresh_target(path, suffix):
    target = Path(path).resolve()
    if target.suffix.lower() != suffix:
        raise ValueError(f"path must end with {suffix}")
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def existing(path):
    if not Path(path).is_file():
        raise FileNotFoundError(f"no such file: {path}")
    return path


@app.tool(description="Create a docx. paragraphs: [{text, style (Heading 1..3 or Word style), size, bold, italic, align}], table: rows of cells")
@timed
def docx_create(path: str, paragraphs: list[dict[str, Any]], table: list[list[Any]] | None = None) -> str:
    target = fresh_target(path, ".docx")
    doc = docxio.Document()
    for spec in paragraphs:
        docxio.add_para(doc, spec)
    if table:
        docxio.add_table(doc, table, header=True)
    doc.save(target)
    return str(target)


@app.tool(description="Read docx text in document order, tables as rows joined by |")
@timed
def docx_read(path: str, max_chars: int = 20000) -> str:
    return docxio.doc_text(docxio.open_doc(existing(path)), max_chars)


@app.tool(description="Append paragraphs and replace text in an existing docx, replace: [[old, new], ...]")
@timed
def docx_edit(path: str, append: list[dict[str, Any]] | None = None, replace: list[list[str]] | None = None) -> dict:
    for pair in replace or []:
        if len(pair) != 2:
            raise ValueError("each replace item is [old, new]")
    doc = docxio.open_doc(existing(path))
    hits = sum(docxio.replace_text(doc, old, new) for old, new in replace or [])
    for spec in append or []:
        docxio.add_para(doc, spec)
    doc.save(path)
    return {"appended": len(append or []), "replaced": hits}


@app.tool(description="Append a table to a docx, first row bold when header is true, optional column widths in cm")
@timed
def docx_table(path: str, rows: list[list[Any]], header: bool = True, widths_cm: list[float] | None = None) -> str:
    doc = docxio.open_doc(existing(path))
    tbl = docxio.add_table(doc, rows, header, widths_cm)
    doc.save(path)
    return f"table {len(tbl.rows)}x{len(tbl.columns)} added"


@app.tool(description="Create an xlsx: sheets is {sheet name: rows}, first row is styled as header")
@timed
def xlsx_create(path: str, sheets: dict[str, list[list[Any]]]) -> str:
    if not sheets:
        raise ValueError("give at least one sheet")
    target = fresh_target(path, ".xlsx")
    book = xlsxio.new_book()
    for name, grid in sheets.items():
        ws = book.create_sheet(name[:31])
        xlsxio.write_rows(ws, grid)
        if grid:
            xlsxio.style_header(ws)
            xlsxio.fit_widths(ws)
    book.save(target)
    return str(target)


@app.tool(description="Read rows from a sheet, formulas=false returns values cached by Excel")
@timed
def xlsx_read(path: str, sheet: str = "", max_cells: int = 5000, formulas: bool = False) -> dict:
    book = xlsxio.open_book(existing(path), values_only=not formulas)
    ws = xlsxio.sheet_of(book, sheet) if sheet else book.worksheets[0]
    grid = xlsxio.read_rows(ws, max_cells)
    return {"sheet": ws.title, "sheets": book.sheetnames, "rows": grid}


@app.tool(description="Write rows into a sheet. mode append adds below existing data, replace clears the sheet first")
@timed
def xlsx_write(path: str, sheet: str, rows: list[list[Any]], mode: str = "append") -> str:
    if mode not in ("append", "replace"):
        raise ValueError("mode is append or replace")
    book = xlsxio.open_book(existing(path))
    ws = book[sheet] if sheet in book.sheetnames else book.create_sheet(sheet[:31])
    if mode == "replace":
        ws.delete_rows(1, ws.max_row)
    has_data = ws.max_row > 1 or ws.cell(1, 1).value is not None
    start = ws.max_row + 1 if has_data else 1
    xlsxio.write_rows(ws, rows, start_row=start)
    book.save(path)
    return f"{len(rows)} rows written to {ws.title} from row {start}"


@app.tool(description="Format a sheet: styled header rows, thin borders, column widths")
@timed
def xlsx_table(path: str, sheet: str, header_rows: int = 1) -> str:
    book = xlsxio.open_book(existing(path))
    ws = xlsxio.sheet_of(book, sheet)
    xlsxio.style_header(ws, header_rows)
    xlsxio.grid_borders(ws)
    xlsxio.fit_widths(ws)
    book.save(path)
    return f"{sheet} formatted, {ws.max_row}x{ws.max_column}"


@app.tool(description="Convert docx/doc/rtf via Word or xlsx/xls via Excel to PDF, needs Microsoft Office")
@timed
def office_to_pdf(path: str, out_path: str = "") -> dict:
    return pdfout.to_pdf(path, out_path)


def run():
    app.run(transport="stdio")
