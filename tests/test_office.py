import asyncio
import struct
import zlib

import pytest
from openpyxl import load_workbook

from officesrv import main
from officesrv.main import app


def tiny_png(path):
    def chunk(kind, body):
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
    rows = b"".join(b"\x00" + b"\xff\x00\x00" * 4 for _ in range(4))
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 4, 8, 2, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))
    return path


def test_every_tool_is_registered():
    names = {t.name for t in asyncio.run(app.list_tools())}
    assert len(names) == 21
    assert {"docx_from_markdown", "xlsx_chart", "pptx_create", "office_to_pdf"} <= names


def test_markdown_document(tmp_path):
    picture = tiny_png(tmp_path / "dot.png")
    md = f"""# Отчёт

Абзац с **жирным** и *курсивом*.

- пункт один
- пункт два

| a | b |
|---|---|
| 1 | 2 |

![]({picture})

---

## Хвост
"""
    made = main.docx_from_markdown(str(tmp_path / "md.docx"), md)
    assert made["tables"] == 1 and made["images"] == 1 and made["headings"] == ["Отчёт", "Хвост"]
    text = main.docx_read(made["path"])["text"]
    assert "# Отчёт" in text and "- пункт один" in text and "| 1 | 2 |" in text


def test_docx_blocks_edit_and_layout(tmp_path):
    path = str(tmp_path / "blocks.docx")
    main.docx_create(path, [{"text": "Титул", "style": "Title"}, {"text": "старое слово", "bold": True, "color": "#C00000"},
                            {"table": [["x", "y"], [1, 2]]}, {"page_break": True}])
    edit = main.docx_edit(path, replace=[["старое", "новое"]], delete_containing="Титул", append_markdown="## Добавка")
    assert edit == {"replaced": 1, "deleted": 1, "appended": 1}
    main.docx_table(path, [["h1", "h2"], ["v1"]], widths_cm=[3, 4])
    main.docx_image(path, str(tiny_png(tmp_path / "p.png")), 2)
    layout = main.docx_page(path, size="A4", landscape=True, margins_cm=1.5, header="шапка", footer="стр.", page_numbers=True)
    assert layout == {"width_cm": 29.7, "height_cm": 21.0, "landscape": True, "margin_cm": 1.5}
    read = main.docx_read(path)
    assert "новое слово" in read["text"] and "Титул" not in read["text"] and read["tables"] == 2 and read["images"] == 1
    with pytest.raises(ValueError):
        main.docx_create(str(tmp_path / "bad.docx"), [{"text": "x", "style": "NoSuchStyle"}])
    with pytest.raises(ValueError):
        main.docx_edit(path, replace=[["only one"]])


def test_xlsx_workflow(tmp_path):
    path = str(tmp_path / "book.xlsx")
    main.xlsx_create(path, {"Продажи": [["Месяц", "План", "Факт"], ["Янв", 100, 90], ["Фев", 120, 130]]})
    main.xlsx_cells(path, "Продажи", {"A4": "Итого", "B4": "=SUM(B2:B3)", "D2": "=C2/B2"}, {"D2": "0.0%"})
    main.xlsx_format(path, "Продажи", column_formats={"B": "#,##0"})
    assert "chart" in main.xlsx_chart(path, "Продажи", "column", "A", ["B", "C"], last_row=3, title="План")
    main.xlsx_chart(path, "Продажи", "line", "A", ["B"], last_row=3, anchor="H20")
    main.xlsx_write(path, "Продажи", [["Мар", 110, 100]], mode="at", start_cell="A5")
    assert main.xlsx_read(path, "Продажи", "B4", formulas=True)["rows"] == [["=SUM(B2:B3)"]]
    assert main.xlsx_read(path, "Продажи", "A5:C5", formulas=True)["rows"] == [["Мар", 110, 100]]
    book = load_workbook(path)
    sheet = book["Продажи"]
    assert len(sheet._charts) == 2 and sheet.freeze_panes == "A2" and sheet.auto_filter.ref
    assert sheet["D2"].number_format == "0.0%" and sheet["A1"].font.bold
    chart = sheet._charts[1]
    assert chart.x_axis.delete is False and all(s.smooth is False for s in chart.series)
    main.xlsx_write(path, "Продажи", [["заново"]], mode="replace")
    assert main.xlsx_read(path, "Продажи", formulas=True)["rows"] == [["заново"]]
    assert main.xlsx_sheets(path, "copy", "Продажи", "Копия") == ["Продажи", "Копия"]
    assert main.xlsx_sheets(path, "rename", "Копия", "Архив") == ["Продажи", "Архив"]
    assert main.xlsx_sheets(path, "delete", "Архив") == ["Продажи"]
    with pytest.raises(ValueError):
        main.xlsx_sheets(path, "delete", "Продажи")
    with pytest.raises(ValueError):
        main.xlsx_chart(path, "Продажи", "radar", "A", ["B"])


def test_csv_round_trip(tmp_path):
    source = tmp_path / "in.csv"
    source.write_text("Имя;Сумма\nа;1,5\nб;1\N{NO-BREAK SPACE}200\nв;-3\n", encoding="utf-8-sig")
    book = str(tmp_path / "csv.xlsx")
    imported = main.xlsx_import_csv(str(source), book, "CSV")
    assert imported["rows"] == 4 and imported["delimiter"] == ";"
    assert main.xlsx_read(book, "CSV")["rows"][1:] == [["а", 1.5], ["б", 1200], ["в", -3]]
    out = tmp_path / "out.csv"
    main.xlsx_export_csv(book, str(out), "CSV")
    assert out.read_text(encoding="utf-8-sig").splitlines()[1] == "а;1.5"


def test_pptx(tmp_path):
    path = str(tmp_path / "deck.pptx")
    picture = str(tiny_png(tmp_path / "s.png"))
    made = main.pptx_create(path, [{"title": "Кварц", "subtitle": "отчёт"},
                                   {"title": "Итоги", "bullets": ["P3_121", {"text": "a = 4.913", "level": 1}], "notes": "заметка"},
                                   {"title": "Картинка", "image": picture, "layout": "title_only"},
                                   {"title": "Таблица", "table": [["a", "c"], ["4.913", "5.404"]], "layout": "title_only"}])
    assert made["slides"] == 4
    assert main.pptx_add(path, [{"title": "Конец", "layout": "section"}])["slides"] == 5
    text = main.pptx_read(path)["text"]
    assert "a = 4.913" in text and "notes: заметка" in text and "[image]" in text and "4.913 | 5.404" in text


@pytest.mark.live
def test_pdf_and_recalculate_through_office(tmp_path):
    doc = main.docx_from_markdown(str(tmp_path / "r.docx"), "# PDF\n\nтекст")["path"]
    book = str(tmp_path / "f.xlsx")
    main.xlsx_create(book, {"S": [["a", "b"], [2, "=A2*21"]]})
    main.pptx_create(str(tmp_path / "d.pptx"), [{"title": "PDF"}])
    for source in (doc, book, str(tmp_path / "d.pptx")):
        assert main.office_to_pdf(source)["bytes"] > 1000
    assert main.xlsx_recalculate(book)["formula_errors"] == 0
    assert main.xlsx_read(book, "S")["rows"][1] == [2, 42]
