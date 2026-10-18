import threading
import winreg
from pathlib import Path

word_kinds = (".docx", ".doc", ".rtf", ".odt")
excel_kinds = (".xlsx", ".xlsm", ".xls", ".csv", ".ods")
slide_kinds = (".pptx", ".ppt", ".odp")


def registered(prog_id):
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + r"\CLSID"))
        return True
    except OSError:
        return False


def in_sta(job, timeout):
    failure, box = [], {}

    def work():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            box["value"] = job()
        except Exception as exc:
            failure.append(exc)
        finally:
            pythoncom.CoUninitialize()

    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        raise TimeoutError(f"Office did not finish in {timeout}s, it may be waiting on a dialog")
    if failure:
        raise RuntimeError(f"Office failed: {failure[0]}")
    return box.get("value")


def office_app(prog_id):
    from win32com.client import DispatchEx
    if not registered(prog_id):
        raise RuntimeError(f"{prog_id} is not installed, this needs Microsoft Office")
    app = DispatchEx(prog_id)
    app.DisplayAlerts = 0 if prog_id != "Excel.Application" else False
    return app


def export_pdf(source, target):
    kind = source.suffix.lower()
    if kind in word_kinds:
        word = office_app("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
            doc.ExportAsFixedFormat(str(target), 17)
            doc.Close(False)
        finally:
            word.Quit()
    elif kind in excel_kinds:
        excel = office_app("Excel.Application")
        excel.Visible = False
        try:
            book = excel.Workbooks.Open(str(source), ReadOnly=True)
            book.ExportAsFixedFormat(0, str(target))
            book.Close(False)
        finally:
            excel.Quit()
    else:
        slides = office_app("PowerPoint.Application")
        try:
            deck = slides.Presentations.Open(str(source), True, False, False)
            deck.SaveAs(str(target), 32)
            deck.Close()
        finally:
            slides.Quit()


def to_pdf(path, out_path, timeout=180):
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"no such file: {path}")
    if source.suffix.lower() not in word_kinds + excel_kinds + slide_kinds:
        raise ValueError("convertible: docx doc rtf odt, xlsx xlsm xls csv ods, pptx ppt odp")
    target = Path(out_path).resolve() if out_path else source.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)
    in_sta(lambda: export_pdf(source, target), timeout)
    if not target.is_file():
        raise RuntimeError(f"Office reported success but {target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}


def recalculate(path, timeout=180):
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"no such file: {path}")

    def job():
        excel = office_app("Excel.Application")
        excel.Visible = False
        try:
            book = excel.Workbooks.Open(str(source))
            excel.CalculateFull()
            errors = 0
            for sheet in book.Worksheets:
                try:
                    errors += sheet.UsedRange.SpecialCells(-4123, 16).Count
                except Exception:
                    pass
            book.Save()
            book.Close(False)
            return errors
        finally:
            excel.Quit()

    errors = in_sta(job, timeout)
    return {"path": str(source), "formula_errors": errors}
