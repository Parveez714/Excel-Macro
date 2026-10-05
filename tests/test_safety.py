import pytest

from macro_assistant.safety import scan_vba


@pytest.mark.parametrize(
    "code, expected",
    [
        ('Sub A()\nShell "calc.exe"\nEnd Sub', "Shell command execution"),
        ('CreateObject("WScript.Shell")', "Windows Script Host access"),
        ('Kill "C:\\temp\\*.*"', "File deletion"),
        ('Set fso = CreateObject("Scripting.FileSystemObject")\nfso.DeleteFile path', "FileSystemObject deletion"),
        ('URLDownloadToFile 0, url, path, 0, 0', "URL download or external network access"),
        ('CreateObject("MSXML2.XMLHTTP")', "URL download or external network access"),
        ('ActiveSheet.FollowHyperlink Address:=url', "URL download or external network access"),
        ('SaveSetting "app", "section", "key", "value"', "Registry modification"),
        ('RegSetValueExA hKey, valueName, 0, REG_SZ, data, Len(data)', "Registry modification"),
    ],
)
def test_scanner_flags_risky_operations(code, expected):
    assert expected in scan_vba(code)


def test_scanner_accepts_simple_macro():
    assert scan_vba("Public Sub ColorRows()\n    Range(\"A1\").Interior.Color = vbRed\nEnd Sub") == []
