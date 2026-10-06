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


def test_export_cross_platform(tmp_path: Path, monkeypatch):
    import openpyxl
    from macro_assistant.excel_injector import _export_cross_platform, inject_macro

    # Create dummy workbook using openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = "Test"
    source = tmp_path / "sample.xlsx"
    wb.save(source)
    wb.close()

    # Simulate non-Windows environment
    monkeypatch.setattr(sys, "platform", "linux")
    result = inject_macro(source, "Public Sub FormatReport()\n    Range(\"A1\").Font.Bold = True\nEnd Sub")

    assert result.output_path == tmp_path / "Macro_sample.xlsx"
    assert result.output_path.exists()
    assert result.macro_name == "FormatReport"
    assert result.applied_in_excel is False
    assert result.module_path is not None
    assert result.module_path.exists()
    assert result.module_path.name == "Macro_sample.bas"
    assert 'Attribute VB_Name = "FormatReport"' in result.module_path.read_text(encoding="utf-8")
    assert 'Range("A1").Font.Bold = True' in result.module_path.read_text(encoding="utf-8")

