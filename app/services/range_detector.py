"""Heuristic engine to find data table boundaries in Excel sheets."""
from app.services.excel_reader import read_full_sheet, range_to_string


def detect_range(filepath: str, sheet_name: str) -> dict:
    """Detect the data table range within an Excel sheet.

    Returns a dict with:
        - range: str (e.g., "B5:K30")
        - data: list[list] (the extracted 2D data)
        - has_header: bool
        - min_row, min_col, max_row, max_col: int (absolute sheet coordinates)
    """
    data, sheet_min_row, sheet_min_col, sheet_max_row, sheet_max_col = read_full_sheet(
        filepath, sheet_name
    )

    if not data or (sheet_max_row - sheet_min_row + 1) == 0:
        return {
            "range": None,
            "data": [],
            "has_header": False,
            "error": f"No data found on sheet {sheet_name}",
        }

    num_rows = len(data)
    num_cols = len(data[0]) if data else 0

    if num_rows == 0 or num_cols == 0:
        return {
            "range": None,
            "data": [],
            "has_header": False,
            "error": f"No data found on sheet {sheet_name}",
        }

    # Pass 1: Build occupancy grid
    occupancy = _build_occupancy_grid(data, num_rows, num_cols)

    # Pass 2: Row and column density
    row_density = _calc_row_density(occupancy, num_rows, num_cols)
    col_density = _calc_col_density(occupancy, num_rows, num_cols)

    # Pass 3: Find dense core
    core = _find_dense_core(occupancy, row_density, col_density, num_rows, num_cols)
    if core is None:
        return {
            "range": None,
            "data": [],
            "has_header": False,
            "error": f"No data found on sheet {sheet_name}",
        }

    r_start, c_start, r_end, c_end = core

    # Pass 4: Trim edges
    r_start, c_start, r_end, c_end = _trim_edges(data, r_start, c_start, r_end, c_end)

    # Pass 5: Header detection
    has_header = _detect_header(data, r_start, c_start, c_end)

    # Convert back to absolute sheet coordinates
    abs_min_row = sheet_min_row + r_start
    abs_min_col = sheet_min_col + c_start
    abs_max_row = sheet_min_row + r_end
    abs_max_col = sheet_min_col + c_end

    range_str = range_to_string(abs_min_col, abs_min_row, abs_max_col, abs_max_row)

    # Extract the detected region
    detected_data = []
    for r in range(r_start, r_end + 1):
        row = []
        for c in range(c_start, c_end + 1):
            row.append(data[r][c])
        detected_data.append(row)

    return {
        "range": range_str,
        "data": detected_data,
        "has_header": has_header,
        "min_row": abs_min_row,
        "min_col": abs_min_col,
        "max_row": abs_max_row,
        "max_col": abs_max_col,
    }


def _build_occupancy_grid(data: list[list], num_rows: int, num_cols: int) -> list[list[bool]]:
    """Pass 1: Build boolean occupancy grid."""
    grid = []
    for r in range(num_rows):
        row = []
        for c in range(num_cols):
            val = data[r][c]
            occupied = val is not None and str(val).strip() != ""
            row.append(occupied)
        grid.append(row)
    return grid


def _calc_row_density(occupancy: list[list[bool]], num_rows: int, num_cols: int) -> list[float]:
    """Pass 2a: Calculate row density."""
    densities = []
    for r in range(num_rows):
        count = sum(1 for c in range(num_cols) if occupancy[r][c])
        densities.append(count / num_cols if num_cols > 0 else 0)
    return densities


def _calc_col_density(occupancy: list[list[bool]], num_rows: int, num_cols: int) -> list[float]:
    """Pass 2b: Calculate column density."""
    densities = []
    for c in range(num_cols):
        count = sum(1 for r in range(num_rows) if occupancy[r][c])
        densities.append(count / num_rows if num_rows > 0 else 0)
    return densities


def _find_dense_core(
    occupancy: list[list[bool]],
    row_density: list[float],
    col_density: list[float],
    num_rows: int,
    num_cols: int,
) -> tuple[int, int, int, int] | None:
    """Pass 3: Find the dense core region."""
    # Find non-sparse rows (density >= 20%)
    non_sparse_rows = [r for r in range(num_rows) if row_density[r] >= 0.20]
    non_sparse_cols = [c for c in range(num_cols) if col_density[c] >= 0.20]

    if not non_sparse_rows or not non_sparse_cols:
        return None

    # Start from the first non-sparse row/col
    r_start = non_sparse_rows[0]
    r_end = non_sparse_rows[-1]
    c_start = non_sparse_cols[0]
    c_end = non_sparse_cols[-1]

    # Refine: within this region, find contiguous rows with >= 50% occupancy
    # and columns with >= 30% occupancy
    region_rows = []
    for r in range(r_start, r_end + 1):
        region_cols_count = sum(1 for c in range(c_start, c_end + 1) if occupancy[r][c])
        region_width = c_end - c_start + 1
        if region_cols_count / region_width >= 0.50 if region_width > 0 else False:
            region_rows.append(r)

    if not region_rows:
        return r_start, c_start, r_end, c_end

    # Find largest contiguous block of qualifying rows
    best_start = region_rows[0]
    best_end = region_rows[0]
    curr_start = region_rows[0]

    for i in range(1, len(region_rows)):
        if region_rows[i] == region_rows[i - 1] + 1:
            # Contiguous
            if region_rows[i] - curr_start > best_end - best_start:
                best_start = curr_start
                best_end = region_rows[i]
        else:
            curr_start = region_rows[i]

    # Check last sequence
    if region_rows[-1] - curr_start > best_end - best_start:
        best_start = curr_start
        best_end = region_rows[-1]

    r_start = best_start
    r_end = best_end

    # Similarly refine columns within the row range
    region_cols = []
    for c in range(c_start, c_end + 1):
        region_rows_count = sum(1 for r in range(r_start, r_end + 1) if occupancy[r][c])
        region_height = r_end - r_start + 1
        if region_rows_count / region_height >= 0.30 if region_height > 0 else False:
            region_cols.append(c)

    if region_cols:
        c_start = region_cols[0]
        c_end = region_cols[-1]

    return r_start, c_start, r_end, c_end


def _trim_edges(
    data: list[list], r_start: int, c_start: int, r_end: int, c_end: int
) -> tuple[int, int, int, int]:
    """Pass 4: Trim entirely empty leading/trailing rows and columns."""
    # Trim leading empty rows
    while r_start <= r_end:
        if all(
            data[r_start][c] is None or str(data[r_start][c]).strip() == ""
            for c in range(c_start, c_end + 1)
        ):
            r_start += 1
        else:
            break

    # Trim trailing empty rows
    while r_end >= r_start:
        if all(
            data[r_end][c] is None or str(data[r_end][c]).strip() == ""
            for c in range(c_start, c_end + 1)
        ):
            r_end -= 1
        else:
            break

    # Trim leading empty columns
    while c_start <= c_end:
        if all(
            data[r][c_start] is None or str(data[r][c_start]).strip() == ""
            for r in range(r_start, r_end + 1)
        ):
            c_start += 1
        else:
            break

    # Trim trailing empty columns
    while c_end >= c_start:
        if all(
            data[r][c_end] is None or str(data[r][c_end]).strip() == ""
            for r in range(r_start, r_end + 1)
        ):
            c_end -= 1
        else:
            break

    return r_start, c_start, r_end, c_end


def _detect_header(data: list[list], r_start: int, c_start: int, c_end: int) -> bool:
    """Pass 5: Check if the first row is a header (all non-empty cells are text, no numbers)."""
    for c in range(c_start, c_end + 1):
        val = data[r_start][c]
        if val is None or str(val).strip() == "":
            continue
        if isinstance(val, (int, float)):
            return False
        # Try parsing as number
        try:
            text = str(val).strip().replace(",", "").replace("$", "").replace("%", "")
            text = text.strip("()")
            if text:
                float(text)
                return False  # It's numeric
        except (ValueError, TypeError):
            pass  # It's text, continue
    return True
