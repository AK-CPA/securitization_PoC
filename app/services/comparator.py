"""Core comparison engine: precision detection, X-Y calculation."""
import re
from typing import Any

# Regex to identify numeric cells (with optional $, commas, parens for negatives, %)
NUMERIC_PATTERN = re.compile(r'^[\s$]*\(?[\d,]+\.?\d*\)?[%]?\s*$')


def is_numeric_string(value: str) -> bool:
    """Check if a string represents a numeric value."""
    if not value or not value.strip():
        return False
    cleaned = value.strip()
    return bool(NUMERIC_PATTERN.match(cleaned))


def parse_numeric_string(value: str) -> float:
    """Parse a formatted numeric string to a float.

    Handles: $1,234.56, (1,234.56), 12.3456%, etc.
    """
    text = str(value).strip()
    is_negative = "(" in text and ")" in text
    is_percent = text.endswith("%")

    # Strip everything except digits, decimal point
    cleaned = text.replace("$", "").replace(",", "").replace("%", "")
    cleaned = cleaned.replace("(", "").replace(")", "").strip()

    if not cleaned:
        raise ValueError(f"Cannot parse '{value}' as numeric")

    result = float(cleaned)
    if is_negative:
        result = -result
    if is_percent:
        result = result / 100.0

    return result


def detect_precision(cell_text: str) -> int:
    """Detect the decimal precision of a numeric string.

    Counts digits after the decimal point in the original string representation.
    """
    text = str(cell_text).strip()
    # Remove % but keep the decimal structure
    text_for_precision = text.replace("%", "").strip()
    text_for_precision = text_for_precision.replace("(", "").replace(")", "").strip()

    if "." in text_for_precision:
        decimal_part = text_for_precision.split(".")[-1]
        # Remove any trailing non-digit characters
        decimal_digits = ""
        for ch in decimal_part:
            if ch.isdigit():
                decimal_digits += ch
            else:
                break
        return len(decimal_digits)
    return 0


def detect_row_precision(row: list[str]) -> int:
    """Detect precision for a row: max precision across all numeric cells.

    If no numeric cells found, defaults to 2.
    """
    max_prec = -1
    for cell in row:
        if cell and is_numeric_string(str(cell)):
            prec = detect_precision(str(cell))
            max_prec = max(max_prec, prec)

    return max_prec if max_prec >= 0 else 2


def detect_table_precision(word_table: list[list[str]]) -> list[int]:
    """Detect precision for each row of a Word table."""
    return [detect_row_precision(row) for row in word_table]


def compare_tables(
    word_table: list[list[str]],
    excel_data: list[list[Any]],
    row_precisions: list[int],
) -> dict:
    """Compare a Word table against Excel data positionally.

    Returns a dict with:
        - diff_grid: 2D array of differences (None for skipped/text cells)
        - status_grid: 2D array of statuses ("match", "mismatch", "skipped", "type_mismatch", "unmatched")
        - match_count: int
        - mismatch_count: int
        - total_cells: int
        - word_rows, word_cols, excel_rows, excel_cols: dimensions
    """
    word_rows = len(word_table)
    word_cols = max((len(r) for r in word_table), default=0)
    excel_rows = len(excel_data)
    excel_cols = max((len(r) for r in excel_data), default=0)

    # Compare over the overlapping region
    compare_rows = min(word_rows, excel_rows)
    compare_cols = min(word_cols, excel_cols)

    diff_grid = []
    status_grid = []
    match_count = 0
    mismatch_count = 0
    total_cells = 0

    for r in range(compare_rows):
        diff_row = []
        status_row = []
        precision = row_precisions[r] if r < len(row_precisions) else 2

        for c in range(compare_cols):
            word_val = word_table[r][c] if c < len(word_table[r]) else ""
            excel_val = excel_data[r][c] if c < len(excel_data[r]) else None

            word_str = str(word_val).strip() if word_val else ""
            word_is_num = is_numeric_string(word_str)
            excel_is_num = isinstance(excel_val, (int, float)) and not isinstance(excel_val, bool)

            # Also check if excel_val is a numeric string
            if not excel_is_num and excel_val is not None:
                excel_str = str(excel_val).strip()
                if is_numeric_string(excel_str):
                    excel_is_num = True

            if not word_str and (excel_val is None or str(excel_val).strip() == ""):
                # Both empty
                diff_row.append(None)
                status_row.append("skipped")
            elif not word_is_num and not excel_is_num:
                # Both non-numeric
                diff_row.append(None)
                status_row.append("skipped")
            elif word_is_num and not excel_is_num:
                diff_row.append(None)
                status_row.append("type_mismatch")
                mismatch_count += 1
                total_cells += 1
            elif not word_is_num and excel_is_num:
                diff_row.append(None)
                status_row.append("type_mismatch")
                mismatch_count += 1
                total_cells += 1
            else:
                # Both numeric - do the comparison
                try:
                    word_float = parse_numeric_string(word_str)
                    if isinstance(excel_val, (int, float)):
                        excel_float = float(excel_val)
                    else:
                        excel_float = parse_numeric_string(str(excel_val))

                    rounded_word = round(word_float, precision)
                    rounded_excel = round(excel_float, precision)
                    diff = round(rounded_word - rounded_excel, precision)

                    diff_row.append(diff)
                    total_cells += 1
                    if diff == 0:
                        status_row.append("match")
                        match_count += 1
                    else:
                        status_row.append("mismatch")
                        mismatch_count += 1
                except (ValueError, TypeError):
                    diff_row.append(None)
                    status_row.append("skipped")

        # Pad if word has more columns
        diff_row.extend([None] * (compare_cols - len(diff_row)))
        status_row.extend(["skipped"] * (compare_cols - len(status_row)))

        diff_grid.append(diff_row)
        status_grid.append(status_row)

    # Mark unmatched rows/cols
    for r in range(compare_rows, max(word_rows, excel_rows)):
        diff_grid.append([None] * compare_cols)
        status_grid.append(["unmatched"] * compare_cols)

    return {
        "diff_grid": diff_grid,
        "status_grid": status_grid,
        "match_count": match_count,
        "mismatch_count": mismatch_count,
        "total_cells": total_cells,
        "word_rows": word_rows,
        "word_cols": word_cols,
        "excel_rows": excel_rows,
        "excel_cols": excel_cols,
    }
