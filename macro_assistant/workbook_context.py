import json
from pathlib import Path


def read_workbook_context(path: str | Path) -> str:
    """Read sheet metadata and the first used row only; never read body data."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise RuntimeError("This feature requires Windows and desktop Microsoft Excel.") from error

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        workbook = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True, AddToMru=False)
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
        return json.dumps({"sheets": sheets}, ensure_ascii=True)
    except Exception as error:
        raise RuntimeError("Excel could not read this workbook. Close it in Excel and try again.") from error
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        finally:
            sheet = None
            used = None
            header_range = None
            workbook = None
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            excel = None
            pythoncom.CoUninitialize()
