import threading
import winreg
from pathlib import Path

word_pdf = 17
excel_pdf = 0


def registered(prog_id):
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + r"\CLSID"))
        return True
    except OSError:
        return False


def convert(source, target):
    import pythoncom
    from win32com.client import DispatchEx

    pythoncom.CoInitialize()
    try:
        if source.suffix.lower() in (".docx", ".doc", ".rtf"):
            word = DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(str(target), word_pdf)
                doc.Close(False)
            finally:
                word.Quit()
        else:
            excel = DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                book = excel.Workbooks.Open(str(source), ReadOnly=True)
                book.ExportAsFixedFormat(excel_pdf, str(target))
                book.Close(False)
            finally:
                excel.Quit()
    finally:
        pythoncom.CoUninitialize()


def to_pdf(path, out_path, timeout=120):
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"no such file: {path}")
    kind = source.suffix.lower()
    if kind not in (".docx", ".doc", ".rtf", ".xlsx", ".xls"):
        raise ValueError("only docx, doc, rtf, xlsx, xls can be converted")
    prog_id = "Word.Application" if kind in (".docx", ".doc", ".rtf") else "Excel.Application"
    if not registered(prog_id):
        raise RuntimeError(f"{prog_id} is not installed, PDF export needs Microsoft Office")
    target = Path(out_path).resolve() if out_path else source.with_suffix(".pdf")
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    failure = []

    def work():
        try:
            convert(source, target)
        except Exception as exc:
            failure.append(exc)

    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        raise TimeoutError(f"Office did not finish in {timeout}s, it may be waiting on a dialog")
    if failure:
        raise RuntimeError(f"Office export failed: {failure[0]}")
    if not target.is_file():
        raise RuntimeError(f"Office reported success but {target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}
