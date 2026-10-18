from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def new_book():
    return Workbook()


def open_book(path):
    return load_workbook(path)


def write_sheet(book, sheet, rows, start_row=1, start_col=1):
    ws = book[sheet] if sheet in book.sheetnames else book.create_sheet(sheet)
    for r, row in enumerate(rows, start=start_row):
        for c, val in enumerate(row, start=start_col):
            ws.cell(r, c, val)
    return ws


def read_sheet(book, sheet, max_cells=5000):
    ws = book[sheet]
    out = []
    for row in ws.iter_rows(values_only=True):
        out.append(list(row))
        if len(out) * len(out[-1] or [1]) > max_cells:
            break
    return out


def format_header(ws, ncols, nrows=1, fill="4472C4"):
    thin = Side(style="thin")
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(horizontal="center")
            cell.border = Border(bottom=thin)


def add_table_style(ws, rows, cols):
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.border = Border(*[Side(style="thin")] * 4)


def auto_width(ws):
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
