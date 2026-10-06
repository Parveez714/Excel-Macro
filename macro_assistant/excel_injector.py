import re
import subprocess
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


EXCEL_TIMEOUT_SECONDS = 90


class TrustAccessError(RuntimeError):
    pass


class CompileCheckError(RuntimeError):
    pass


@dataclass(frozen=True)
class InjectionResult:
    output_path: Path
    macro_name: str
    compile_checked: bool
    module_path: Path | None = None
    applied_in_excel: bool = True


def _office_versions() -> list[str]:
    if sys.platform != "win32":
        return []
    try:
        import winreg

        office_root = r"Software\Microsoft\Office"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, office_root) as key:
            versions = []
            for index in range(winreg.QueryInfoKey(key)[0]):
                version = winreg.EnumKey(key, index)
                if re.fullmatch(r"\d+\.\d+", version):
                    versions.append(version)
            return sorted(versions, key=lambda value: tuple(map(int, value.split("."))), reverse=True)
    except (OSError, ImportError):
        return ["16.0"]


def is_vbom_enabled() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg

        for version in _office_versions() or ["16.0"]:
            policy_path = rf"Software\Policies\Microsoft\Office\{version}\Excel\Security"
            for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(root, policy_path) as policy:
                        value, _ = winreg.QueryValueEx(policy, "AccessVBOM")
                        return value == 1
                except OSError:
                    continue
            key_path = rf"Software\Microsoft\Office\{version}\Excel\Security"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                    if value == 1:
                        return True
            except OSError:
                continue
    except (OSError, ImportError):
        return False
    return False


def enable_vbom_for_current_user() -> None:
    if sys.platform != "win32":
        raise TrustAccessError("This setting is only applicable to Microsoft Excel on Windows.")
    try:
        import winreg

        version = (_office_versions() or ["16.0"])[0]
        policy_path = rf"Software\Policies\Microsoft\Office\{version}\Excel\Security"
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(root, policy_path) as policy:
                    winreg.QueryValueEx(policy, "AccessVBOM")
                raise TrustAccessError("Your organization manages this setting. Ask IT to enable programmatic access to the VBA project object model.")
            except FileNotFoundError:
                continue
            except OSError:
                continue
        key_path = rf"Software\Microsoft\Office\{version}\Excel\Security"
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        except OSError:
            raise TrustAccessError("Windows could not change this setting. It may be managed by your organization; ask IT for help.") from None
    except ImportError:
        raise TrustAccessError("Windows Registry access is unavailable.") from None


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


def _macro_name(code: str) -> str:
    match = re.search(r"(?im)^\s*Public\s+Sub\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s*\(|\s*$)", code)
    if not match:
        match = re.search(r"(?im)^\s*Sub\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s*\(|\s*$)", code)
    return match.group(1) if match else "GeneratedMacro"


def _try_compile(excel) -> tuple[bool, str | None]:
    try:
        controls = excel.VBE.CommandBars("Menu Bar").Controls
        debug_menu = next((controls.Item(index) for index in range(1, controls.Count + 1) if "debug" in str(controls.Item(index).Caption).lower()), None)
        if debug_menu is None:
            return False, None
        debug_controls = debug_menu.Controls
        compile_item = next(
            (debug_controls.Item(index) for index in range(1, debug_controls.Count + 1) if "compile" in str(debug_controls.Item(index).Caption).lower()),
            None,
        )
        if compile_item is None:
            return False, None
        compile_item.Execute()
        return True, None
    except Exception:
        return True, "Excel could not verify that the VBA compiles. Review the code in Excel's VBA editor before running it."


def _export_cross_platform(source: Path, code: str, macro_name: str) -> InjectionResult:
    """Save workbook copy and VBA .bas module file for macOS/Linux or systems without COM."""
    import shutil

    # Create destination workbook copy (.xlsm or .xlsx)
    ext = ".xlsm" if source.suffix.lower() == ".xlsm" else source.suffix.lower()
    output_path = source.with_name(f"Macro_{source.stem}{ext}")
    counter = 2
    while output_path.exists():
        output_path = source.with_name(f"Macro_{source.stem} ({counter}){ext}")
        counter += 1

    shutil.copy2(source, output_path)

    # Save VBA .bas module alongside it
    module_path = output_path.with_suffix(".bas")
    bas_counter = 2
    while module_path.exists():
        module_path = output_path.with_name(f"{output_path.stem} ({bas_counter}).bas")
        bas_counter += 1

    # Standard VBA module header for clean import in Excel (Windows & Mac)
    bas_content = f'Attribute VB_Name = "{macro_name}"\n\n{code.strip()}\n'
    module_path.write_text(bas_content, encoding="utf-8")

    return InjectionResult(
        output_path=output_path,
        macro_name=macro_name,
        compile_checked=False,
        module_path=module_path,
        applied_in_excel=False,
    )


def inject_macro(source_path: str | Path, code: str) -> InjectionResult:
    source = Path(source_path).resolve()
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Choose an .xlsx or .xlsm workbook.")

    macro_name = _macro_name(code)

    # On macOS, Linux, or environments without win32com/Excel, export the ready workbook copy + .bas module
    if sys.platform != "win32":
        return _export_cross_platform(source, code, macro_name)

    # On Windows: Check if win32com and Excel are available
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        # If pywin32 is not installed, gracefully fall back to cross-platform export
        return _export_cross_platform(source, code, macro_name)

    if not is_vbom_enabled():
        raise TrustAccessError("Excel's programmatic VBA access is turned off. Enable it using the button in the app, then try again.")

    output_path = source.with_name(f"Macro_{source.stem}.xlsm")
    counter = 2
    while output_path.exists():
        output_path = source.with_name(f"Macro_{source.stem} ({counter}).xlsm")
        counter += 1

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    component = None
    module = None
    compile_checked = False
    watchdog = None
    timed_out = threading.Event()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1
        try:
            import win32process

            _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)

            def _kill_excel():
                timed_out.set()
                subprocess.run(["taskkill", "/F", "/PID", str(excel_pid)], capture_output=True)

            watchdog = threading.Timer(EXCEL_TIMEOUT_SECONDS, _kill_excel)
            watchdog.daemon = True
            watchdog.start()
        except Exception:
            watchdog = None
        workbook = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        component = workbook.VBProject.VBComponents.Add(1)
        component.Name = f"GeneratedMacro{datetime.now():%H%M%S}"
        module = component.CodeModule
        module.AddFromString(code)
        compile_checked, compile_error = _try_compile(excel)
        if compile_error:
            raise CompileCheckError(compile_error)
        if macro_name:
            excel.Run(f"'{workbook.Name}'!{macro_name}")
        workbook.SaveAs(str(output_path), FileFormat=52)
        workbook.Close(SaveChanges=True)
        workbook = None
        return InjectionResult(output_path, macro_name, compile_checked, module_path=None, applied_in_excel=True)
    except (TrustAccessError, CompileCheckError):
        raise
    except Exception as error:
        if timed_out.is_set():
            raise RuntimeError(
                f"Excel stopped responding for {EXCEL_TIMEOUT_SECONDS} seconds, usually because the VBA hit a compile or runtime error dialog. Excel was closed."
            ) from error
        raise RuntimeError(f"Excel error while adding or running macro: {error}") from error
    finally:
        if watchdog is not None:
            watchdog.cancel()
        module = None
        component = None
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

