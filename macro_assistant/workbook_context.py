import json
import sys
from pathlib import Path


def _read_with_openpyxl(path: Path) -> list[dict]:
    import openpyxl

    # openpyxl read_only mode is extremely fast and low memory
    # data_only=True ensures we see cached values or headers if formulas are used
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        sheets = []
        for name in wb.sheetnames[:100]:
            sheet = wb[name]
            dimensions = sheet.dimensions or "A1"
            headers = []
            header_row = 1

            for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                # Look for first row with non-empty cells
                non_empty = [c for c in row if c is not None and str(c).strip() != ""]
                if non_empty or row_idx == 1:
                    header_row = row_idx
                    headers = [str(val)[:200] if val is not None else "" for val in row[:50]]
                    break

            sheets.append(
                {
                    "name": str(name)[:100],
                    "used_range": str(dimensions)[:100],
                    "header_row": header_row,
                    "headers": headers,
                }
            )
        return sheets
    finally:
        wb.close()


def _read_with_com(path: Path) -> list[dict]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        sheets = []
        for index in range(1, min(workbook.Worksheets.Count, 100) + 1):
            sheet = workbook.Worksheets(index)
            used = sheet.UsedRange
            first_row = int(used.Row)
            first_column = int(used.Column)
            last_column = first_column + int(used.Columns.Count) - 1
            headers = []
            header_range = sheet.Range(sheet.Cells(first_row, first_column), sheet.Cells(first_row, min(last_column, first_column + 49)))
            values = header_range.Value
            if values is not None:
                row_values = values[0] if isinstance(values, tuple) and values and isinstance(values[0], tuple) else values
                if not isinstance(row_values, tuple):
                    row_values = (row_values,)
                headers = [str(value)[:200] if value is not None else "" for value in row_values]
            sheets.append(
                {
                    "name": str(sheet.Name)[:100],
                    "used_range": str(used.Address)[:100],
                    "header_row": first_row,
                    "headers": headers,
                }
            )
        return sheets
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def read_workbook_context(path: str | Path) -> str:
    """Read sheet metadata and the first used row only; never read body data."""
    target_path = Path(path).resolve()
    if not target_path.is_file():
        raise FileNotFoundError(f"Workbook not found: {target_path}")

    # First attempt pure-Python openpyxl (works on Windows, macOS, Linux without Office)
    try:
        sheets = _read_with_openpyxl(target_path)
        return json.dumps({"sheets": sheets}, ensure_ascii=True)
    except Exception as openpyxl_error:
        # Fall back to COM on Windows if openpyxl encountered an unusual format
        if sys.platform == "win32":
            try:
                sheets = _read_with_com(target_path)
                return json.dumps({"sheets": sheets}, ensure_ascii=True)
            except Exception:
                pass
        raise RuntimeError(f"Could not read workbook: {openpyxl_error}") from openpyxl_error

