import csv
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

thin = Side(style="thin")
charts = {"bar": BarChart, "column": BarChart, "line": LineChart, "pie": PieChart, "scatter": ScatterChart}


def open_book(path, values_only=False):
    if not Path(path).is_file():
        raise FileNotFoundError(f"no such file: {path}")
    try:
        return load_workbook(path, data_only=values_only)
    except Exception as exc:
        raise ValueError(f"cannot open {path} as xlsx: {exc}") from exc


def sheet_of(book, name):
    if not name:
        return book.worksheets[0]
    if name not in book.sheetnames:
        raise ValueError(f"no sheet {name!r}, sheets are {book.sheetnames}")
    return book[name]


def write_rows(ws, grid, start_row=1, start_col=1):
    for r, row in enumerate(grid, start=start_row):
        for c, val in enumerate(row, start=start_col):
            ws.cell(r, c, val)


def plain(val):
    if isinstance(val, (datetime, date, time)):
        return val.isoformat()
    return val


def read_rows(ws, cell_range, max_cells):
    rows = ws[cell_range] if cell_range else ws.iter_rows()
    if cell_range and not isinstance(rows, tuple):
        rows = ((rows,),)
    elif cell_range and rows and not isinstance(rows[0], tuple):
        rows = (rows,)
    out, seen = [], 0
    for row in rows:
        cells = [plain(c.value) for c in row]
        while cells and cells[-1] is None:
            cells.pop()
        seen += max(1, len(cells))
        if seen > max_cells:
            break
        out.append(cells)
    while out and not out[-1]:
        out.pop()
    return out


def style_header(ws, nrows=1, fill="4472C4"):
    for r in range(1, nrows + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=fill.lstrip("#"))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def grid_borders(ws):
    edge = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = edge


def fit_widths(ws):
    for c in range(1, ws.max_column + 1):
        letter = get_column_letter(c)
        longest = max((len(str(cell.value)) for cell in ws[letter] if cell.value is not None), default=0)
        ws.column_dimensions[letter].width = min(max(longest + 2, 8), 60)


def column_formats(ws, formats, header_rows):
    for letter, fmt in formats.items():
        for cell in ws[letter.upper()][header_rows:]:
            cell.number_format = fmt


def new_book():
    book = Workbook()
    book.remove(book.active)
    return book


def chart(ws, kind, label_col, value_cols, first_row, last_row, title, anchor, y_title, x_title):
    if kind not in charts:
        raise ValueError(f"kind is one of {', '.join(charts)}")
    last_row = last_row or ws.max_row
    plot = charts[kind]()
    if kind == "bar":
        plot.type = "bar"
    elif kind == "column":
        plot.type = "col"
    plot.title = title or None
    labels = Reference(ws, min_col=column_index_from_string(label_col), min_row=first_row + 1, max_row=last_row)
    if kind == "scatter":
        for letter in value_cols:
            col = column_index_from_string(letter)
            values = Reference(ws, min_col=col, min_row=first_row, max_row=last_row)
            plot.series.append(Series(values, labels, title_from_data=True))
    else:
        for letter in value_cols:
            col = column_index_from_string(letter)
            plot.add_data(Reference(ws, min_col=col, min_row=first_row, max_row=last_row), titles_from_data=True)
        plot.set_categories(labels)
    if kind != "pie":
        plot.y_axis.title = y_title or None
        plot.x_axis.title = x_title or None
        # openpyxl hides axes by default, new Excel then draws no labels
        plot.x_axis.delete = False
        plot.y_axis.delete = False
    for line in plot.series:
        line.smooth = False
    plot.width, plot.height = 15, 8
    ws.add_chart(plot, anchor or f"{get_column_letter(ws.max_column + 2)}2")
    return len(ws._charts)


def import_csv(ws, csv_path, delimiter, encoding):
    if not Path(csv_path).is_file():
        raise FileNotFoundError(f"no such file: {csv_path}")
    with open(csv_path, newline="", encoding=encoding) as handle:
        sample = handle.read(4096)
        handle.seek(0)
        sep = delimiter or csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        rows = 0
        for row in csv.reader(handle, delimiter=sep):
            ws.append([number_or_text(v) for v in row])
            rows += 1
    return rows, sep


def number_or_text(text):
    probe = text.strip().replace(" ", "").replace(" ", "")
    if probe and probe.lstrip("-").replace(".", "", 1).replace(",", "", 1).isdigit():
        try:
            return float(probe.replace(",", ".")) if ("." in probe or "," in probe) else int(probe)
        except ValueError:
            return text
    return text


def export_csv(ws, csv_path, delimiter, encoding):
    Path(csv_path).parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding=encoding) as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        rows = 0
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if v is None else plain(v) for v in row])
            rows += 1
    return rows
