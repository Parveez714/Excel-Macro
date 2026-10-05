import sys
from pathlib import Path

import pytest

from macro_assistant.excel_injector import inject_macro, is_vbom_enabled


@pytest.mark.skipif(sys.platform != "win32", reason="Excel automation is Windows-only")
def test_inject_macro_into_new_workbook(tmp_path: Path):
    win32_client = pytest.importorskip("win32com.client")
    pytest.importorskip("pythoncom")
    if not is_vbom_enabled():
        pytest.skip("Trust access to the VBA project object model is disabled")

    try:
        excel = win32_client.DispatchEx("Excel.Application")
    except Exception:
        pytest.skip("Desktop Excel is not installed or could not be started")
    source = tmp_path / "integration.xlsx"
    try:
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(str(source), FileFormat=51)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()

    result = inject_macro(source, "Public Sub TestMacro()\nEnd Sub")
    assert result.output_path == tmp_path / "Macro_integration.xlsm"
    assert result.output_path.exists()
