# mcp-office

MCP-сервер для документов Word, таблиц Excel и презентаций PowerPoint в Claude Desktop. Файлы docx, xlsx и pptx пишутся напрямую, Office для этого не нужен. Экспорт в PDF и пересчёт формул идут через установленный Microsoft Office.

## Инструменты

**Word**

| Инструмент | Что делает |
|---|---|
| `docx_create` | документ из блоков: заголовки, абзацы со стилями и цветом, списки, таблицы, картинки, разрывы страниц |
| `docx_from_markdown` | документ из Markdown: заголовки, списки, таблицы, цитаты, картинки, жирный и курсив |
| `docx_read` | текст в виде Markdown по порядку документа, число абзацев, таблиц, картинок, слов, оглавление |
| `docx_edit` | замена текста в теле, таблицах и колонтитулах, удаление абзацев, дописывание блоков или Markdown |
| `docx_table` | таблица с жирной шапкой и шириной колонок |
| `docx_image` | вставка картинки |
| `docx_page` | формат листа, ориентация, поля, колонтитулы, номера страниц |

**Excel**

| Инструмент | Что делает |
|---|---|
| `xlsx_create` | книга из листов, формулы через `=`, оформленная шапка, закреплённая первая строка |
| `xlsx_read` | лист или диапазон `A1:D20`, значения или текст формул |
| `xlsx_write` | строки в конец, с заменой листа или с заданной ячейки |
| `xlsx_cells` | отдельные ячейки и их числовой формат |
| `xlsx_format` | шапка, рамки, ширина колонок, закрепление, автофильтр, форматы колонок |
| `xlsx_chart` | настоящая диаграмма Excel: столбцы, линии, круговая, точечная |
| `xlsx_sheets` | список, добавление, переименование, удаление и копирование листов |
| `xlsx_import_csv` | импорт CSV с определением разделителя и чисел |
| `xlsx_export_csv` | выгрузка листа в CSV (UTF-8 с BOM, кириллица в Excel открывается правильно) |
| `xlsx_recalculate` | пересчёт формул в Excel, после него `xlsx_read` видит значения |

**PowerPoint и PDF**

| Инструмент | Что делает |
|---|---|
| `pptx_create` | презентация: титульный слайд, пункты с уровнями, картинки, таблицы, заметки докладчика |
| `pptx_add` | дописать слайды в существующую презентацию |
| `pptx_read` | текст, таблицы и заметки всех слайдов |
| `office_to_pdf` | PDF из docx, doc, rtf, odt, xlsx, xls, csv, pptx, ppt через Word, Excel или PowerPoint |

## Установка

Нужны Windows, [uv](https://docs.astral.sh/uv/) и Claude Desktop. Для `office_to_pdf` и `xlsx_recalculate` нужен Microsoft Office.

```powershell
git clone https://github.com/hmsusbusas-png/mcp-office C:\mcp\mcp-office
cd C:\mcp\mcp-office
uv sync
```

## Подключение к Claude Desktop

Settings, Developer, Edit Config. Путь к uv нужен полный, его покажет `where.exe uv`.

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\Users\\<имя>\\.local\\bin\\uv.exe",
      "args": ["run", "--directory", "C:\\mcp\\mcp-office", "python", "-m", "officesrv"]
    }
  }
}
```

После правки полностью закройте Claude Desktop из трея и запустите снова.

## Проверка

```powershell
uv run pytest
uv run pytest -m live
```

Вторая команда дополнительно проверяет экспорт в PDF и пересчёт через установленный Office. Каждый такой вызов запускает отдельный скрытый экземпляр Word, Excel или PowerPoint и закрывает его после работы.
