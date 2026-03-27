"""Extract and parse tables from .docx files."""
from docx import Document
from docx.table import Table
from typing import Optional


def extract_tables(filepath: str) -> list[list[list[str]]]:
    """Extract all tables from a Word document.

    Returns a list of tables, where each table is a 2D array of cell strings.
    Merged cells are expanded into their constituent grid positions.
    """
    doc = Document(filepath)
    tables = []

    for table in doc.tables:
        grid = _table_to_grid(table)
        if grid and any(any(cell.strip() for cell in row) for row in grid):
            tables.append(grid)

    return tables


def _table_to_grid(table: Table) -> list[list[str]]:
    """Convert a python-docx Table to a 2D grid of strings, handling merged cells."""
    rows = table.rows
    if not rows:
        return []

    num_rows = len(rows)
    num_cols = max(len(row.cells) for row in rows) if rows else 0

    if num_cols == 0:
        return []

    # Build grid, handling merged cells
    grid: list[list[Optional[str]]] = [[None] * num_cols for _ in range(num_rows)]

    for r_idx, row in enumerate(rows):
        cells = row.cells
        for c_idx in range(min(len(cells), num_cols)):
            cell = cells[c_idx]
            text = cell.text.strip()
            if grid[r_idx][c_idx] is None:
                grid[r_idx][c_idx] = text

    # Fill any remaining None values with empty string
    for r_idx in range(num_rows):
        for c_idx in range(num_cols):
            if grid[r_idx][c_idx] is None:
                grid[r_idx][c_idx] = ""

    return grid


def get_table_label(table_data: list[list[str]], index: int) -> str:
    """Generate a label for a table based on its content."""
    # Try to use the first row's first cell as a label hint
    if table_data and table_data[0]:
        first_cell = table_data[0][0].strip()
        if first_cell and len(first_cell) < 80:
            return f"Table {index + 1}: {first_cell}"
    return f"Table {index + 1}"
