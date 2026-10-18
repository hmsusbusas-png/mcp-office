import logging
import sys
import time
from functools import wraps
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from . import docxio, xlsxio

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("office")

app = FastMCP("office")


def quiet(fn):
    @wraps(fn)
    def wrap(*args, **kwargs):
        start = time.monotonic()
        try:
            return fn(*args, **kwargs)
        finally:
            log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
    return wrap


def _ensure(path):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


@app.tool(description="Create a docx with paragraphs and optional table")
@quiet
def docx_create(path: str, paragraphs: list[dict], table: list[list[str]] | None = None) -> str:
    doc = docxio.new_doc()
    for p in paragraphs:
        docxio.add_para(doc, p.get("text", ""), p.get("size"), p.get("bold", False), p.get("align"))
    if table:
        docxio.add_table(doc, table, header=True)
    target = _ensure(path)
    doc.save(target)
    return str(target)


@app.tool(description="Read text and tables from a docx")
@quiet
def docx_read(path: str, max_chars: int = 20000) -> str:
    return docxio.doc_text(docxio.open_doc(path), max_chars)


@app.tool(description="Append paragraphs or replace text in an existing docx")
@quiet
def docx_edit(path: str, append: list[dict] | None = None, replace: list[list[str]] | None = None) -> str:
    doc = docxio.open_doc(path)
    appended = 0
    for p in append or []:
        docxio.add_para(doc, p.get("text", ""), p.get("size"), p.get("bold", False), p.get("align"))
        appended += 1
    hits = 0
    for pair in replace or []:
        hits += docxio.replace_text(doc, pair[0], pair[1])
    doc.save(path)
    return f"appended {appended}, replaced {hits}"


@app.tool(description="Add a formatted table to a docx")
@quiet
def docx_table(path: str, rows: list[list[str]], header: bool = True, widths_cm: list[float] | None = None) -> str:
    doc = docxio.open_doc(path)
    docxio.add_table(doc, rows, header, widths_cm)
    doc.save(path)
    return f"table {len(rows)}x{len(rows[0])} added to {path}"


@app.tool(description="Create an xlsx workbook with sheets and rows")
@quiet
def xlsx_create(path: str, sheets: dict[str, list[list]]) -> str:
    book = xlsxio.new_book()
    book.remove(book.active)
    for name, rows in sheets.items():
        ws = xlsxio.write_sheet(book, name, rows)
        xlsxio.format_header(ws, len(rows[0]))
        xlsxio.auto_width(ws)
    target = _ensure(path)
    book.save(target)
    return str(target)


@app.tool(description="Read rows from a sheet")
@quiet
def xlsx_read(path: str, sheet: str, max_cells: int = 5000) -> list[list]:
    return xlsxio.read_sheet(xlsxio.open_book(path), sheet, max_cells)


@app.tool(description="Write rows into a sheet, append below existing content by default")
@quiet
def xlsx_write(path: str, sheet: str, rows: list[list], mode: str = "append") -> str:
    book = xlsxio.open_book(path)
    ws = book[sheet] if sheet in book.sheetnames else book.create_sheet(sheet)
    start = ws.max_row + 1 if mode == "append" else 1
    xlsxio.write_sheet(book, sheet, rows, start_row=start)
    book.save(path)
    return f"{len(rows)} rows written to {sheet} from row {start}"


@app.tool(description="Format sheet: bold header, borders, widths")
@quiet
def xlsx_table(path: str, sheet: str, header_rows: int = 1) -> str:
    book = xlsxio.open_book(path)
    ws = book[sheet]
    xlsxio.format_header(ws, ws.max_column, header_rows)
    xlsxio.add_table_style(ws, ws.max_row, ws.max_column)
    xlsxio.auto_width(ws)
    book.save(path)
    return f"{sheet} formatted"


def run():
    app.run(transport="stdio")
