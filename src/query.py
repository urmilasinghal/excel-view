"""
Excel View Query Engine (Step 4)

Provides filtering, sorting, and rollup functions over parsed Excel data.
Works with the list-of-dicts output from parser.parse_sheet().

Usage:
    from parser import parse_sheet
    from query import filter_data, sort_data, rollup, search
    
    data = parse_sheet("data/sot.xlsx")
    
    # Filter
    p0_gaps = filter_data(data, Priority="P0", Desktop_Timeline_status="Gap")
    
    # Sort
    sorted_data = sort_data(data, by="Priority")
    
    # Rollup
    by_status = rollup(data, group_by="Desktop Timeline_status")
"""

from typing import Optional
from collections import Counter


def filter_data(data: list[dict], **filters) -> list[dict]:
    """
    Filter rows by column values. Supports multiple filters (AND logic).
    
    Column names use underscores for spaces:
        filter_data(data, Priority="P0", Desktop_Timeline_status="Gap")
    matches rows where Priority == "P0" AND Desktop Timeline_status == "Gap".
    
    Values are case-insensitive.
    Use a list for OR within a column:
        filter_data(data, Priority=["P0", "P1"])
    """
    results = data
    
    for key, value in filters.items():
        # Convert underscores back to spaces for column lookup
        col_name = key.replace("_", " ")
        
        # Also try with underscores (for derived columns like "Desktop Timeline_status")
        # Find the actual column name in the data
        actual_key = _resolve_column_name(data, key, col_name)
        if actual_key is None:
            continue
        
        if isinstance(value, list):
            # OR logic: match any value in the list
            lower_values = [str(v).lower() for v in value]
            results = [r for r in results if str(r.get(actual_key, "")).lower() in lower_values]
        else:
            lower_value = str(value).lower()
            results = [r for r in results if str(r.get(actual_key, "")).lower() == lower_value]
    
    return results


def filter_not(data: list[dict], **filters) -> list[dict]:
    """
    Filter rows that do NOT match the given values (exclusion filter).
    Same syntax as filter_data but inverts the match.
    
        filter_not(data, Desktop_Timeline_status="Parity")
    returns rows where Desktop Timeline status is NOT Parity.
    """
    results = data
    
    for key, value in filters.items():
        col_name = key.replace("_", " ")
        actual_key = _resolve_column_name(data, key, col_name)
        if actual_key is None:
            continue
        
        if isinstance(value, list):
            lower_values = [str(v).lower() for v in value]
            results = [r for r in results if str(r.get(actual_key, "")).lower() not in lower_values]
        else:
            lower_value = str(value).lower()
            results = [r for r in results if str(r.get(actual_key, "")).lower() != lower_value]
    
    return results


def sort_data(data: list[dict], by: str, descending: bool = False) -> list[dict]:
    """
    Sort rows by a column.
    
        sort_data(data, by="Priority")
        sort_data(data, by="DRI", descending=True)
    """
    actual_key = _resolve_column_name(data, by, by.replace("_", " "))
    if actual_key is None:
        return data
    
    return sorted(data, key=lambda r: str(r.get(actual_key, "")), reverse=descending)


def select_columns(data: list[dict], columns: list[str]) -> list[dict]:
    """
    Return only the specified columns from each row.
    
        select_columns(data, ["Feature", "DRI", "MSA Premium Notes"])
    """
    results = []
    for row in data:
        selected = {}
        for col in columns:
            actual_key = _resolve_column_name(data, col, col.replace("_", " "))
            if actual_key and actual_key in row:
                selected[actual_key] = row[actual_key]
        results.append(selected)
    return results


def rollup(data: list[dict], group_by: str) -> dict[str, int]:
    """
    Count rows grouped by a column value.
    
        rollup(data, group_by="Desktop Timeline_status")
        # Returns: {"Parity": 20, "Gap": 5, "By Design": 3, "In Progress": 1}
    """
    actual_key = _resolve_column_name(data, group_by, group_by.replace("_", " "))
    if actual_key is None:
        return {}
    
    return dict(Counter(str(r.get(actual_key, "None")) for r in data))


def cross_rollup(data: list[dict], row_group: str, col_group: str) -> dict[str, dict[str, int]]:
    """
    Cross-tabulation: count rows by two dimensions.
    
        cross_rollup(data, row_group="Priority", col_group="Desktop Timeline_status")
        # Returns: {"P0": {"Parity": 15, "Gap": 2}, "P1": {"Parity": 3, "Gap": 1}, ...}
    """
    row_key = _resolve_column_name(data, row_group, row_group.replace("_", " "))
    col_key = _resolve_column_name(data, col_group, col_group.replace("_", " "))
    
    if row_key is None or col_key is None:
        return {}
    
    result = {}
    for row in data:
        rv = str(row.get(row_key, "None"))
        cv = str(row.get(col_key, "None"))
        if rv not in result:
            result[rv] = {}
        result[rv][cv] = result[rv].get(cv, 0) + 1
    
    return result


def search(data: list[dict], text: str, columns: Optional[list[str]] = None) -> list[dict]:
    """
    Search for text across all columns (or specified columns).
    Case-insensitive substring match.
    
        search(data, "agent")
        search(data, "agent", columns=["Feature", "MSA Premium Notes"])
    """
    text_lower = text.lower()
    results = []
    
    for row in data:
        search_cols = columns if columns else list(row.keys())
        for col in search_cols:
            actual_key = _resolve_column_name(data, col, col.replace("_", " "))
            if actual_key and actual_key in row:
                val = row.get(actual_key, "")
                if val is not None and text_lower in str(val).lower():
                    results.append(row)
                    break
    
    return results


def has_value(data: list[dict], column: str) -> list[dict]:
    """Return rows where the column has a non-empty value."""
    actual_key = _resolve_column_name(data, column, column.replace("_", " "))
    if actual_key is None:
        return data
    return [r for r in data if r.get(actual_key) is not None and str(r.get(actual_key, "")).strip() != ""]


def has_no_value(data: list[dict], column: str) -> list[dict]:
    """Return rows where the column is empty or None."""
    actual_key = _resolve_column_name(data, column, column.replace("_", " "))
    if actual_key is None:
        return data
    return [r for r in data if r.get(actual_key) is None or str(r.get(actual_key, "")).strip() == ""]


def _resolve_column_name(data: list[dict], *candidates: str) -> Optional[str]:
    """
    Find the actual column name in the data, trying multiple candidates.
    Handles mixed spaces/underscores (e.g., "Desktop_Timeline_status" 
    matches "Desktop Timeline_status").
    """
    if not data:
        return candidates[0] if candidates else None
    
    available = set(data[0].keys())
    
    for candidate in candidates:
        # Exact match
        if candidate in available:
            return candidate
        # Try with underscores replaced by spaces
        spaced = candidate.replace("_", " ")
        if spaced in available:
            return spaced
    
    # Normalized match: compare with all separators standardized
    def normalize(s):
        return s.lower().replace("_", " ").replace("  ", " ").strip()
    
    for candidate in candidates:
        norm_candidate = normalize(candidate)
        for key in available:
            if normalize(key) == norm_candidate:
                return key
    
    return None


def format_results(data: list[dict], columns: Optional[list[str]] = None, max_rows: int = 50) -> str:
    """
    Format query results as a readable string.
    Useful for agent output.
    """
    if not data:
        return "No matching rows found."
    
    # Determine columns to show
    if columns:
        show_cols = columns
    else:
        # Default: skip internal color columns
        show_cols = [k for k in data[0].keys() if not k.endswith("_color")]
    
    lines = []
    lines.append(f"Found {len(data)} rows.")
    lines.append("")
    
    # Header
    header = " | ".join(str(c) for c in show_cols)
    lines.append(header)
    lines.append("-" * len(header))
    
    # Rows
    display_data = data[:max_rows]
    for row in display_data:
        values = []
        for col in show_cols:
            actual = _resolve_column_name(data, col, col.replace("_", " "))
            val = row.get(actual, "") if actual else ""
            values.append(str(val) if val is not None else "")
        lines.append(" | ".join(values))
    
    if len(data) > max_rows:
        lines.append(f"... and {len(data) - max_rows} more rows.")
    
    return "\n".join(lines)


# --- Quick test when run directly ---
if __name__ == "__main__":
    import sys
    sys.path.insert(0, ".")
    from parser import parse_sheet
    
    filepath = sys.argv[1] if len(sys.argv) > 1 else "data/test_sot.xlsx"
    data = parse_sheet(filepath)
    
    print("=" * 60)
    print("QUERY ENGINE TEST")
    print("=" * 60)
    
    # Test 1: Filter P0 features
    p0 = filter_data(data, Priority="P0")
    print(f"\nP0 features: {len(p0)}")
    
    # Test 2: P0 gaps on desktop
    p0_gaps = filter_data(data, Priority="P0", Desktop_Timeline_status="Gap")
    print(f"P0 desktop gaps: {len(p0_gaps)}")
    for r in p0_gaps:
        print(f"  {r['Feature']} (DRI: {r['DRI']})")
    
    # Test 3: Available on mobile but not desktop
    mobile_parity = filter_data(data, Mobile_Timeline_status="Parity")
    mobile_not_desktop = filter_not(mobile_parity, Desktop_Timeline_status="Parity")
    print(f"\nParity on mobile but NOT desktop: {len(mobile_not_desktop)}")
    for r in mobile_not_desktop:
        print(f"  {r['Feature']} | Desktop: {r['Desktop Timeline_status']} | Mobile: {r['Mobile Timeline_status']}")
    
    # Test 4: Rollup by desktop status
    print(f"\nDesktop status rollup:")
    for status, count in rollup(data, "Desktop Timeline_status").items():
        print(f"  {status}: {count}")
    
    # Test 5: Cross rollup
    print(f"\nPriority x Desktop status:")
    cross = cross_rollup(data, "Priority", "Desktop Timeline_status")
    for priority, statuses in sorted(cross.items()):
        print(f"  {priority}: {statuses}")
    
    # Test 6: Search
    agent_rows = search(data, "agent")
    print(f"\nSearch 'agent': {len(agent_rows)} rows")
    for r in agent_rows:
        print(f"  {r['Feature']}")
    
    # Test 7: Her actual question
    print(f"\n{'=' * 60}")
    print("Urmila's query: P0 features on mobile but not desktop, no timeline")
    print("=" * 60)
    p0_data = filter_data(data, Priority="P0")
    mobile_yes = filter_data(p0_data, Mobile_Timeline_status="Parity")
    desktop_no = filter_not(mobile_yes, Desktop_Timeline_status="Parity")
    no_timeline = has_no_value(desktop_no, "Desktop Timeline")
    result = select_columns(no_timeline, ["Feature", "DRI", "MSA Premium Notes"])
    print(format_results(no_timeline, columns=["Feature", "DRI", "MSA Premium Notes"]))
