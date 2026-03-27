"""Read Excel data with heuristic range detection."""
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


def get_sheet_names(filepath: str) -> list[str]:
    """Return list of sheet names from an Excel workbook."""
    wb = load_workbook(filepath, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def read_sheet_data(filepath: str, sheet_name: str, cell_range: str | None = None) -> list[list]:
    """Read data from a specific sheet, optionally within a cell range.

    Returns a 2D array of cell values (preserving types).
    """
    wb = load_workbook(filepath, read_only=False, data_only=True)
    ws = wb[sheet_name]

    if cell_range:
        min_col, min_row, max_col, max_row = parse_range(cell_range)
    else:
        min_row = ws.min_row or 1
        max_row = ws.max_row or 1
        min_col = ws.min_column or 1
        max_col = ws.max_column or 1

    data = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        row_data = []
        for cell in row:
            row_data.append(cell.value)
        data.append(row_data)

    wb.close()
    return data


def read_full_sheet(filepath: str, sheet_name: str) -> tuple[list[list], int, int, int, int]:
    """Read the entire used range of a sheet.

    Returns (data, min_row, min_col, max_row, max_col).
    """
    wb = load_workbook(filepath, read_only=False, data_only=True)
    ws = wb[sheet_name]

    min_row = ws.min_row or 1
    max_row = ws.max_row or 1
    min_col = ws.min_column or 1
    max_col = ws.max_column or 1

    data = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        row_data = []
        for cell in row:
            row_data.append(cell.value)
        data.append(row_data)

    wb.close()
    return data, min_row, min_col, max_row, max_col


def parse_range(cell_range: str) -> tuple[int, int, int, int]:
    """Parse an Excel range string like 'B5:K30' into (min_col, min_row, max_col, max_row)."""
    match = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', cell_range.strip().upper())
    if not match:
        raise ValueError(f"Invalid range format: {cell_range}. Use Excel notation like B5:K30.")

    min_col = column_index_from_string(match.group(1))
    min_row = int(match.group(2))
    max_col = column_index_from_string(match.group(3))
    max_row = int(match.group(4))

    return min_col, min_row, max_col, max_row


def range_to_string(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    """Convert numeric range bounds back to Excel notation like 'B5:K30'."""
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def get_sheet_dimensions(filepath: str, sheet_name: str) -> tuple[int, int]:
    """Get the max row and max column of a sheet."""
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb[sheet_name]
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    wb.close()
    return max_row, max_col
