"""Claude API integration for loose language tie-out.

Uses Claude to:
1. Extract numeric assertions from matched sentences
2. Determine what Excel data to compare against
3. Perform the comparison
"""
import json
import logging
from anthropic import Anthropic

from app.config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)


def get_client() -> Anthropic:
    """Get an Anthropic client. Raises if no API key configured."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError(
            "ANTHROPIC_API_KEY environment variable is not set. "
            "Please set it to use loose language tie-out features."
        )
    return Anthropic(api_key=ANTHROPIC_API_KEY)


def extract_and_compare(
    matched_sentence: str,
    excel_data: list[list],
    sheet_name: str,
) -> dict:
    """Use Claude to extract values from sentence and compare against Excel data.

    Args:
        matched_sentence: The sentence from the offering document
        excel_data: 2D array of Excel data (with headers in first row if present)
        sheet_name: Name of the Excel sheet for context

    Returns:
        dict with:
            - document_values: list of {label, value} extracted from sentence
            - excel_values: list of {label, value} extracted from Excel
            - comparisons: list of {label, doc_value, excel_value, difference, status}
            - summary: str description of what was compared
            - status: "pass" or "fail"
    """
    client = get_client()

    # Format Excel data as a readable table for Claude
    excel_text = _format_excel_for_prompt(excel_data, sheet_name)

    prompt = f"""You are a financial document reconciliation assistant. Your job is to compare numeric data assertions in an offering document sentence against source data in an Excel spreadsheet.

## Sentence from Offering Document
"{matched_sentence}"

## Source Excel Data (Sheet: {sheet_name})
{excel_text}

## Your Task
1. **Identify the assertion**: What numeric claim is the sentence making? (e.g., "states with concentration > 5%", "top 10 obligors represent X%", "weighted average coupon is Y%")

2. **Extract document values**: Pull out every specific number/percentage mentioned in the sentence along with its label/context.

3. **Find corresponding Excel values**: Based on the assertion logic, find the matching values in the Excel data. Apply the same criteria described in the sentence (e.g., filter for > 5%, sum values, compute weighted averages).

4. **Compare**: For each document value, find the corresponding Excel value and compute the difference.

## Response Format
Respond with ONLY valid JSON (no markdown, no explanation outside JSON):
{{
    "summary": "Brief description of what assertion was compared",
    "document_values": [
        {{"label": "descriptive label", "value": numeric_value}},
        ...
    ],
    "excel_values": [
        {{"label": "descriptive label", "value": numeric_value}},
        ...
    ],
    "comparisons": [
        {{
            "label": "what is being compared",
            "doc_value": numeric_value_from_sentence,
            "excel_value": numeric_value_from_excel,
            "difference": doc_minus_excel,
            "status": "match" or "mismatch"
        }},
        ...
    ]
}}

Rules:
- Values should be raw numbers (not strings). Percentages should be in percent form (e.g., 5.25 not 0.0525).
- "match" means the difference rounds to 0 at the precision shown in the document.
- If you cannot find a corresponding Excel value for a document value, set excel_value to null and status to "unmatched".
- If the sentence contains no numeric assertions, return empty arrays.
- Compare values at the precision displayed in the sentence."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )

        response_text = response.content[0].text.strip()

        # Parse JSON — handle potential markdown wrapping
        if response_text.startswith("```"):
            response_text = response_text.split("```")[1]
            if response_text.startswith("json"):
                response_text = response_text[4:]
            response_text = response_text.strip()

        result = json.loads(response_text)

        # Determine overall status
        comparisons = result.get("comparisons", [])
        if not comparisons:
            result["status"] = "pass"
        elif any(c.get("status") in ("mismatch", "unmatched") for c in comparisons):
            result["status"] = "fail"
        else:
            result["status"] = "pass"

        return result

    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse Claude response as JSON: {e}")
        return {
            "summary": "Error parsing Claude response",
            "document_values": [],
            "excel_values": [],
            "comparisons": [],
            "status": "error",
            "error": f"Failed to parse response: {e}",
            "raw_response": response_text if 'response_text' in dir() else "",
        }
    except Exception as e:
        logger.error(f"Claude API error: {e}")
        return {
            "summary": "API Error",
            "document_values": [],
            "excel_values": [],
            "comparisons": [],
            "status": "error",
            "error": str(e),
        }


def _format_excel_for_prompt(data: list[list], sheet_name: str) -> str:
    """Format Excel data as a readable text table for the prompt."""
    if not data:
        return "(empty sheet)"

    lines = []
    for r_idx, row in enumerate(data):
        cells = []
        for val in row:
            if val is None:
                cells.append("")
            else:
                cells.append(str(val))
        lines.append(" | ".join(cells))

        # Add a separator after the first row (assumed header)
        if r_idx == 0:
            lines.append("-" * 60)

    return "\n".join(lines)


def build_loose_output_data(items: list[dict]) -> list[dict]:
    """Build structured output data for the Excel output builder.

    Args:
        items: list of LooseComparisonItem-like dicts

    Returns:
        list of comparison dicts compatible with the output builder format
    """
    output = []
    for item in items:
        comparisons = item.get("comparison_result", {}).get("comparisons", [])
        if not comparisons:
            continue

        # Build a "word_table" equivalent: rows of [label, doc_value]
        word_rows = []
        excel_rows = []
        diff_rows = []
        status_rows = []

        for comp in comparisons:
            label = comp.get("label", "")
            doc_val = comp.get("doc_value")
            excel_val = comp.get("excel_value")
            diff = comp.get("difference")
            status = comp.get("status", "skipped")

            word_rows.append([label, doc_val if doc_val is not None else ""])
            excel_rows.append([label, excel_val if excel_val is not None else ""])
            diff_rows.append([None, diff if diff is not None else None])

            if status == "match":
                status_rows.append(["skipped", "match"])
            elif status == "mismatch":
                status_rows.append(["skipped", "mismatch"])
            elif status == "unmatched":
                status_rows.append(["skipped", "unmatched"])
            else:
                status_rows.append(["skipped", "skipped"])

        output.append({
            "table_label": item.get("summary", "Loose Language"),
            "word_filename": item.get("word_filename", ""),
            "excel_filename": item.get("excel_filename", ""),
            "excel_tab_name": item.get("sheet_name", ""),
            "excel_range": "",
            "word_table": word_rows,
            "excel_data": excel_rows,
            "diff_grid": diff_rows,
            "status_grid": status_rows,
            "row_precisions": [2] * len(word_rows),
        })

    return output
