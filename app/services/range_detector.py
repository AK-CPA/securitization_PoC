"""Heuristic engine to find data table boundaries in Excel sheets.

Supports detecting multiple distinct data regions on a single sheet.
"""
from app.services.excel_reader import read_full_sheet, range_to_string


def detect_range(filepath: str, sheet_name: str) -> dict:
    """Detect the primary data table range within an Excel sheet.

    Returns a dict with:
        - range: str (e.g., "B5:K30")
        - data: list[list] (the extracted 2D data)
        - has_header: bool
        - min_row, min_col, max_row, max_col: int (absolute sheet coordinates)
    """
    results = detect_ranges(filepath, sheet_name)
    if not results:
        return {
            "range": None,
            "data": [],
            "has_header": False,
            "error": f"No data found on sheet {sheet_name}",
        }
    # Return the largest region as the primary
    return results[0]


def detect_ranges(filepath: str, sheet_name: str, gap_tolerance: int = 1) -> list[dict]:
    """Detect all distinct data regions on a sheet.

    Regions are separated by blank rows. A gap of up to `gap_tolerance` blank
    rows is allowed within a region before it splits into two.

    Returns a list of dicts (sorted by size descending), each with:
        - range: str (e.g., "B5:K30")
        - data: list[list]
        - has_header: bool
        - min_row, min_col, max_row, max_col: int (absolute sheet coordinates)
    """
    data, sheet_min_row, sheet_min_col, sheet_max_row, sheet_max_col = read_full_sheet(
        filepath, sheet_name
    )

    if not data or len(data) == 0:
        return []

    num_rows = len(data)
    num_cols = len(data[0]) if data else 0

    if num_rows == 0 or num_cols == 0:
        return []

    # Build occupancy grid
    occupancy = _build_occupancy_grid(data, num_rows, num_cols)

    # Find row bands: groups of rows that have data, allowing gap_tolerance blank rows
    row_bands = _find_row_bands(occupancy, num_rows, num_cols, gap_tolerance)

    if not row_bands:
        return []

    # For each row band, find the column extent and build a region
    regions = []
    for band_start, band_end in row_bands:
        # Find column boundaries within this band
        col_start, col_end = _find_col_extent(occupancy, band_start, band_end, num_cols)
        if col_start is None:
            continue

        # Trim edges
        r_start, c_start, r_end, c_end = _trim_edges(
            data, band_start, col_start, band_end, col_end
        )

        if r_start > r_end or c_start > c_end:
            continue

        # Header detection
        has_header = _detect_header(data, r_start, c_start, c_end)

        # Convert to absolute sheet coordinates
        abs_min_row = sheet_min_row + r_start
        abs_min_col = sheet_min_col + c_start
        abs_max_row = sheet_min_row + r_end
        abs_max_col = sheet_min_col + c_end

        range_str = range_to_string(abs_min_col, abs_min_row, abs_max_col, abs_max_row)

        # Extract the detected region data
        detected_data = []
        for r in range(r_start, r_end + 1):
            row = []
            for c in range(c_start, c_end + 1):
                row.append(data[r][c])
            detected_data.append(row)

        region_size = (r_end - r_start + 1) * (c_end - c_start + 1)

        regions.append({
            "range": range_str,
            "data": detected_data,
            "has_header": has_header,
            "min_row": abs_min_row,
            "min_col": abs_min_col,
            "max_row": abs_max_row,
            "max_col": abs_max_col,
            "_size": region_size,
        })

    # Sort by size descending (largest region first)
    regions.sort(key=lambda r: r["_size"], reverse=True)

    # Remove internal _size key
    for r in regions:
        del r["_size"]

    return regions


def _find_row_bands(
    occupancy: list[list[bool]],
    num_rows: int,
    num_cols: int,
    gap_tolerance: int,
) -> list[tuple[int, int]]:
    """Find bands of rows containing data, splitting on gaps > gap_tolerance blank rows.

    A row is considered "occupied" if at least 15% of its cells are non-empty.
    """
    # Classify each row
    occupied_rows = []
    for r in range(num_rows):
        count = sum(1 for c in range(num_cols) if occupancy[r][c])
        density = count / num_cols if num_cols > 0 else 0
        occupied_rows.append(density >= 0.15)

    # Group into bands with gap tolerance
    bands = []
    band_start = None
    gap_count = 0

    for r in range(num_rows):
        if occupied_rows[r]:
            if band_start is None:
                band_start = r
            gap_count = 0
        else:
            if band_start is not None:
                gap_count += 1
                if gap_count > gap_tolerance:
                    # Close the current band (end before the gap)
                    band_end = r - gap_count
                    if band_end >= band_start:
                        bands.append((band_start, band_end))
                    band_start = None
                    gap_count = 0

    # Close final band
    if band_start is not None:
        band_end = num_rows - 1
        # Trim trailing empty rows
        while band_end > band_start and not occupied_rows[band_end]:
            band_end -= 1
        if band_end >= band_start:
            bands.append((band_start, band_end))

    # Filter out tiny bands (less than 2 rows)
    bands = [(s, e) for s, e in bands if (e - s + 1) >= 2]

    return bands


def _find_col_extent(
    occupancy: list[list[bool]],
    r_start: int,
    r_end: int,
    num_cols: int,
) -> tuple[int | None, int | None]:
    """Find the column start and end for a row band based on density."""
    col_counts = []
    band_height = r_end - r_start + 1
    for c in range(num_cols):
        count = sum(1 for r in range(r_start, r_end + 1) if occupancy[r][c])
        col_counts.append(count)

    # Find columns with at least 20% occupancy in this band
    threshold = max(1, band_height * 0.20)
    active_cols = [c for c in range(num_cols) if col_counts[c] >= threshold]

    if not active_cols:
        return None, None

    return active_cols[0], active_cols[-1]


def _build_occupancy_grid(data: list[list], num_rows: int, num_cols: int) -> list[list[bool]]:
    """Build boolean occupancy grid."""
    grid = []
    for r in range(num_rows):
        row = []
        for c in range(num_cols):
            val = data[r][c]
            occupied = val is not None and str(val).strip() != ""
            row.append(occupied)
        grid.append(row)
    return grid


def _trim_edges(
    data: list[list], r_start: int, c_start: int, r_end: int, c_end: int
) -> tuple[int, int, int, int]:
    """Trim entirely empty leading/trailing rows and columns."""
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
    """Check if the first row is a header (all non-empty cells are text, no numbers)."""
    for c in range(c_start, c_end + 1):
        val = data[r_start][c]
        if val is None or str(val).strip() == "":
            continue
        if isinstance(val, (int, float)):
            return False
        try:
            text = str(val).strip().replace(",", "").replace("$", "").replace("%", "")
            text = text.strip("()")
            if text:
                float(text)
                return False
        except (ValueError, TypeError):
            pass
    return True
