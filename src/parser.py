"""
Excel View Parser (Steps 1-3)

Reads an xlsx file, extracts structured data from a specified tab,
and maps cell fill colors to status values.

Usage:
    from parser import parse_workbook, parse_sheet
    
    # Get all tabs
    tabs = parse_workbook("data/sot.xlsx")
    
    # Parse a specific tab
    data = parse_sheet("data/sot.xlsx", "MSA Premium Detail")
"""

import openpyxl
from typing import Optional


# Color-to-status mapping.
# Keys are hex color codes (uppercase, no #). Values are status strings.
# "no_fill" is a special key for cells with no background color.
COLOR_STATUS_MAP = {
    "92D050": "Parity",
    "00B050": "Parity",
    "FFFF00": "By Design",
    "FFC000": "In Progress",
    "FF8C00": "In Progress",
    "no_fill": "Gap",
}

# Columns where color encodes status
COLOR_STATUS_COLUMNS = ["Desktop Timeline", "Mobile Timeline"]

# Default tolerance for color matching (0-255 per channel)
COLOR_TOLERANCE = 30


def parse_workbook(filepath: str) -> list[str]:
    """Return the list of tab (sheet) names in the workbook."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def parse_sheet(filepath: str, sheet_name: str = "MSA Premium Detail") -> list[dict]:
    """
    Parse a single sheet from the workbook into a list of row dicts.
    
    Each row dict has:
    - One key per column header, with the cell value
    - For color-status columns (Desktop/Mobile Timeline):
      - "{column}_status" key with the mapped status string
      - "{column}_color" key with the raw hex color (for debugging)
    
    Skips empty rows and rows where all data cells are empty.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    
    ws = wb[sheet_name]
    
    # Step 1: Read headers from row 1
    headers = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val is not None:
            headers.append(str(val).strip())
        else:
            headers.append(f"Column_{col}")
    
    # Identify which column indices are color-status columns
    color_col_indices = []
    for i, h in enumerate(headers):
        if h in COLOR_STATUS_COLUMNS:
            color_col_indices.append(i)
    
    # Steps 2-3: Read data rows, extract values and colors
    rows = []
    for row_idx in range(2, ws.max_row + 1):
        row_data = {}
        all_empty = True
        
        for col_idx, header in enumerate(headers):
            cell = ws.cell(row=row_idx, column=col_idx + 1)
            value = cell.value
            
            if value is not None:
                all_empty = False
            
            # Store the cell value
            row_data[header] = value
            
            # For color-status columns, extract fill color and map to status
            if col_idx in color_col_indices:
                color_hex = _extract_fill_color(cell)
                status = _map_color_to_status(color_hex)
                row_data[f"{header}_status"] = status
                row_data[f"{header}_color"] = color_hex
        
        # Skip empty rows
        if all_empty:
            continue
        
        rows.append(row_data)
    
    wb.close()
    return rows


def get_schema(filepath: str, sheet_name: str = "MSA Premium Detail") -> dict:
    """
    Return the schema of a sheet: column names, which are color-status columns,
    and the color mapping in use.
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    
    ws = wb[sheet_name]
    
    headers = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val is not None:
            headers.append(str(val).strip())
    
    wb.close()
    
    return {
        "sheet_name": sheet_name,
        "columns": headers,
        "color_status_columns": [h for h in headers if h in COLOR_STATUS_COLUMNS],
        "color_mapping": {v: k for k, v in COLOR_STATUS_MAP.items() if k != "no_fill"},
        "row_count": None,  # Filled after parsing
    }


def _extract_fill_color(cell) -> str:
    """
    Extract the background fill color from a cell.
    Returns the hex color string (e.g., "92D050") or "no_fill".
    """
    fill = cell.fill
    
    if fill is None or fill.fill_type is None or fill.fill_type == "none":
        return "no_fill"
    
    if fill.start_color and fill.start_color.rgb:
        rgb = str(fill.start_color.rgb)
        # openpyxl sometimes returns "00000000" for no fill
        if rgb == "00000000":
            return "no_fill"
        # Strip alpha channel if present (AARRGGBB -> RRGGBB)
        if len(rgb) == 8:
            rgb = rgb[2:]
        return rgb.upper()
    
    return "no_fill"


def _map_color_to_status(color_hex: str) -> str:
    """
    Map a hex color to a status string.
    Uses exact match first, then falls back to nearest color within tolerance.
    """
    if color_hex == "no_fill":
        return COLOR_STATUS_MAP["no_fill"]
    
    # Exact match
    if color_hex in COLOR_STATUS_MAP:
        return COLOR_STATUS_MAP[color_hex]
    
    # Nearest match within tolerance
    best_match = None
    best_distance = float("inf")
    
    for known_hex, status in COLOR_STATUS_MAP.items():
        if known_hex == "no_fill":
            continue
        distance = _color_distance(color_hex, known_hex)
        if distance < best_distance:
            best_distance = distance
            best_match = status
    
    if best_distance <= COLOR_TOLERANCE * 3:  # Sum of RGB channel tolerances
        return best_match
    
    # No match found
    return f"Unknown ({color_hex})"


def _color_distance(hex1: str, hex2: str) -> int:
    """Calculate Manhattan distance between two hex colors."""
    try:
        r1, g1, b1 = int(hex1[0:2], 16), int(hex1[2:4], 16), int(hex1[4:6], 16)
        r2, g2, b2 = int(hex2[0:2], 16), int(hex2[2:4], 16), int(hex2[4:6], 16)
        return abs(r1 - r2) + abs(g1 - g2) + abs(b1 - b2)
    except (ValueError, IndexError):
        return float("inf")


# --- Quick test when run directly ---
if __name__ == "__main__":
    import sys
    
    filepath = sys.argv[1] if len(sys.argv) > 1 else "data/test_sot.xlsx"
    
    print("=" * 60)
    print(f"FILE: {filepath}")
    print("=" * 60)
    
    # Step 1: List tabs
    tabs = parse_workbook(filepath)
    print(f"\nTabs: {tabs}")
    
    # Step 1: Show schema
    schema = get_schema(filepath)
    print(f"\nColumns: {schema['columns']}")
    print(f"Color-status columns: {schema['color_status_columns']}")
    
    # Steps 2-3: Parse data
    data = parse_sheet(filepath)
    print(f"\nRows parsed: {len(data)}")
    
    # Show first 5 rows
    print(f"\nFirst 5 rows:")
    for row in data[:5]:
        feature = row.get("Feature", "")
        priority = row.get("Priority", "")
        dri = row.get("DRI", "")
        dt_status = row.get("Desktop Timeline_status", "")
        mt_status = row.get("Mobile Timeline_status", "")
        print(f"  {priority} | {feature:<40} | {dri:<20} | Desktop: {dt_status:<12} | Mobile: {mt_status}")
    
    # Show color distribution
    print(f"\nDesktop status distribution:")
    from collections import Counter
    dt_counts = Counter(r.get("Desktop Timeline_status") for r in data)
    for status, count in dt_counts.most_common():
        print(f"  {status}: {count}")
    
    mt_counts = Counter(r.get("Mobile Timeline_status") for r in data)
    print(f"\nMobile status distribution:")
    for status, count in mt_counts.most_common():
        print(f"  {status}: {count}")
