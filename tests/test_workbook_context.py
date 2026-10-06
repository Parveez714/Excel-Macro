import json
from pathlib import Path
import openpyxl
from macro_assistant.workbook_context import read_workbook_context


def test_read_workbook_context_openpyxl(tmp_path: Path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "SalesData"
    sheet.append(["Region", "Revenue", "Cost", "Margin"])
    sheet.append(["North", 1000, 400, 600])

    test_file = tmp_path / "sales.xlsx"
    wb.save(test_file)
    wb.close()

    context_json = read_workbook_context(test_file)
    data = json.loads(context_json)

    assert "sheets" in data
    assert len(data["sheets"]) == 1
    sheet_info = data["sheets"][0]
    assert sheet_info["name"] == "SalesData"
    assert sheet_info["header_row"] == 1
    assert sheet_info["headers"] == ["Region", "Revenue", "Cost", "Margin"]
