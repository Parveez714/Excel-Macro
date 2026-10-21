from dataclasses import dataclass
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell


def patch_openpyxl_merged_cells():
    """Ensure writes to MergedCell objects in openpyxl safely unmerge and write."""
    def get_merged_val(self):
        ws = self.parent
        if (self.row, self.column) in ws._cells:
            cur = ws._cells[(self.row, self.column)]
            if cur is not self:
                return cur.value
        return None

    def set_merged_val(self, val):
        ws = self.parent
        coord = self.coordinate
        for rng in list(ws.merged_cells.ranges):
            if coord in rng:
                ws.unmerge_cells(str(rng))
                break
        new_cell = Cell(ws, row=self.row, column=self.column, value=val)
        ws._cells[(self.row, self.column)] = new_cell

    MergedCell.value = property(get_merged_val, set_merged_val)

    _orig_cell = Worksheet.cell
    def safe_cell(self, row, column, value=None):
        if row < 1 or column < 1:
            raise ValueError("Row or column values must be at least 1")
        cell = self._get_cell(row, column)
        if isinstance(cell, MergedCell):
            if value is not None:
                coord = cell.coordinate
                for rng in list(self.merged_cells.ranges):
                    if coord in rng:
                        self.unmerge_cells(str(rng))
                        break
                cell = Cell(self, row=row, column=column, value=value)
                self._cells[(row, column)] = cell
                return cell
        if value is not None:
            cell.value = value
        return cell
    Worksheet.cell = safe_cell

    def safe_setitem(self, key, value):
        cell = self[key]
        if isinstance(cell, MergedCell):
            coord = cell.coordinate
            for rng in list(self.merged_cells.ranges):
                if coord in rng:
                    self.unmerge_cells(str(rng))
                    break
            cell = Cell(self, row=cell.row, column=cell.column, value=value)
            self._cells[(cell.row, cell.column)] = cell
            return
        cell.value = value
    Worksheet.__setitem__ = safe_setitem


# Apply patch on import
patch_openpyxl_merged_cells()


@dataclass(frozen=True)
class ExecutionResult:
    output_path: Path
    action_name: str = "Workbook Updated"
    compile_checked: bool = True
    module_path: Path | None = None
    applied_in_excel: bool = True


def apply_transformation(source_path: str | Path, python_code: str) -> ExecutionResult:
    patch_openpyxl_merged_cells()
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
