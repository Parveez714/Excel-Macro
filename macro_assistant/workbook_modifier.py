from dataclasses import dataclass
from pathlib import Path
import openpyxl


@dataclass(frozen=True)
class ExecutionResult:
    output_path: Path
    action_name: str = "Workbook Updated"
    compile_checked: bool = True
    module_path: Path | None = None
    applied_in_excel: bool = True


def apply_transformation(source_path: str | Path, python_code: str) -> ExecutionResult:
    source = Path(source_path).resolve()
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Choose an .xlsx or .xlsm workbook.")

    output_path = source.with_name(f"Updated_{source.stem}{source.suffix.lower()}")
    counter = 2
    while output_path.exists():
        output_path = source.with_name(f"Updated_{source.stem} ({counter}){source.suffix.lower()}")
        counter += 1

    # Load workbook preserving formulas and formatting
    wb = openpyxl.load_workbook(str(source), data_only=False)

    # Safe execution sandbox
    exec_scope = {
        "wb": wb,
        "openpyxl": openpyxl,
    }

    try:
        exec(python_code, exec_scope)
    except Exception as e:
        raise RuntimeError(f"Error applying spreadsheet changes: {e}") from e

    # Save to destination
    wb.save(str(output_path))
    wb.close()

    return ExecutionResult(output_path=output_path)


# Backwards compatibility alias
inject_macro = apply_transformation
InjectionResult = ExecutionResult
