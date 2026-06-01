"""
Excel View Tools (Step 5 specs -> Step 6 implementation)

Two halves:
1. TOOLS  -- the descriptions Claude reads to decide which tool to call.
2. The functions that actually run, built on parser.py + query.py.

See docs/tools.md for the full spec.
"""

import re
from collections import Counter
from datetime import datetime

from parser import parse_sheet
from query import filter_data, search

# Where the SOT lives. Change this one line to point at a different file.
SOT_PATH = r"C:\Users\ursingha\Downloads\ParitySOT.xlsx"

# Current focus priority. Sorts to the top of every list. Change as focus shifts.
FOCUS_PRIORITY = "P0.5"

# Year for timeline dates (cells store only M/D).
TIMELINE_YEAR = 2026

# Sort orders
_BASE_PRIORITY_ORDER = ["P0", "P0.5", "P1", "P2"]
_STATUS_ORDER = ["Gap", "In Progress", "Under Review", "By Design", "Parity"]

# Columns
_TIMELINE_COLS = ["Desktop Timeline", "Mobile Timeline"]
_NOTE_COLS = ["Feature", "MSA Premium Notes", "Internal Notes"]
_DISPLAY = {"status": "Status", "DRI": "Owner"}

# Map a group_by name to the actual row key
_GROUP_KEYS = {
    "area": "Area", "section": "Section", "priority": "Priority",
    "status": "status", "owner": "DRI",
}


# ----------------------------------------------------------------------------
# Half 1: tool descriptions Claude reads
# ----------------------------------------------------------------------------

TOOLS = [
    {
        "name": "find_features",
        "description": (
            "Find and list features in the parity SOT that match filters. "
            "Use for list-style questions like 'all P0 gaps', 'what is Ian "
            "working on', or 'in-progress features with no timeline'. All "
            "filters are optional and combine with AND."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "area": {"type": "string", "description": "Area, e.g. Left Nav, Header"},
                "section": {"type": "string", "description": "Section within an area"},
                "priority": {"type": "string", "enum": ["P0", "P0.5", "P1", "P2"]},
                "status": {
                    "type": "string",
                    "enum": ["Parity", "In Progress", "Gap", "By Design", "Under Review"],
                    "description": "Web parity status",
                },
                "owner": {"type": "string", "description": "DRI name. Matches shared-ownership rows too"},
                "platform": {"type": "string", "enum": ["desktop", "mobile"]},
                "has_timeline": {
                    "type": "boolean",
                    "description": "True for features with a timeline, false for those without. Scoped by platform if given",
                },
                "keyword": {"type": "string", "description": "Substring search across feature name and notes"},
                "include_notes": {"type": "boolean", "description": "Include the notes column in output"},
            },
        },
    },
    {
        "name": "get_feature",
        "description": (
            "Get the full detail for one feature: status, owner, priority, "
            "timelines, blockers, and notes. Use when the user asks about a "
            "specific feature by name. Matches partial names."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "feature": {"type": "string", "description": "Feature name or part of it"},
            },
            "required": ["feature"],
        },
    },
    {
        "name": "summarize",
        "description": (
            "Count features grouped by a dimension, for rollups and reporting. "
            "Optionally cross-tab by a second dimension (e.g. priority by "
            "status). Accepts the same filters as find_features to summarize a "
            "subset."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "group_by": {"type": "string", "enum": ["area", "section", "priority", "status", "owner"]},
                "then_by": {"type": "string", "enum": ["area", "section", "priority", "status", "owner"],
                            "description": "Optional second dimension for a cross-tab"},
                "area": {"type": "string"},
                "section": {"type": "string"},
                "priority": {"type": "string", "enum": ["P0", "P0.5", "P1", "P2"]},
                "status": {"type": "string", "enum": ["Parity", "In Progress", "Gap", "By Design", "Under Review"]},
                "owner": {"type": "string"},
            },
            "required": ["group_by"],
        },
    },
    {
        "name": "schedule",
        "description": (
            "Show what stages (SDF) or ships (WW) and when, by platform, sorted "
            "soonest first. Use for time questions like 'what ships before "
            "June 15'. Forward-looking: already-shipped items are excluded."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "milestone": {"type": "string", "enum": ["SDF", "WW"], "description": "SDF=staging, WW=shipping. Default WW"},
                "platform": {"type": "string", "enum": ["desktop", "mobile"]},
                "before": {"type": "string", "description": "Date bound, M/D (e.g. 6/15)"},
                "after": {"type": "string", "description": "Date bound, M/D"},
                "priority": {"type": "string", "enum": ["P0", "P0.5", "P1", "P2"]},
                "status": {"type": "string", "enum": ["Parity", "In Progress", "Gap", "By Design", "Under Review"]},
                "owner": {"type": "string"},
            },
        },
    },
]


# ----------------------------------------------------------------------------
# Shared filtering (used by find_features, summarize, schedule)
# ----------------------------------------------------------------------------

def _apply_filters(data, area=None, section=None, priority=None, status=None,
                   owner=None, platform=None, has_timeline=None, keyword=None):
    if area:
        data = filter_data(data, Area=area)
    if section:
        data = filter_data(data, Section=section)
    if priority:
        data = filter_data(data, Priority=priority)
    if status:
        data = filter_data(data, status=status)
    if owner:
        data = [r for r in data if _owner_matches(r, owner)]
    if keyword:
        data = search(data, keyword, columns=_NOTE_COLS)
    if has_timeline is not None:
        data = [r for r in data if _has_timeline(r, platform) == has_timeline]
    return data


def _owner_matches(row, owner):
    """Owner filter that includes shared ownership (multi-DRI cells)."""
    return owner.strip().lower() in str(row.get("DRI") or "").lower()


def _has_timeline(row, platform):
    desktop = bool(row.get("Desktop Timeline"))
    mobile = bool(row.get("Mobile Timeline"))
    if platform == "desktop":
        return desktop
    if platform == "mobile":
        return mobile
    return desktop or mobile


# ----------------------------------------------------------------------------
# Tool 1: find_features
# ----------------------------------------------------------------------------

def find_features(area=None, section=None, priority=None, status=None,
                  owner=None, platform=None, has_timeline=None,
                  keyword=None, include_notes=False):
    data = _apply_filters(parse_sheet(SOT_PATH), area, section, priority, status,
                          owner, platform, has_timeline, keyword)
    data = sorted(data, key=_sort_key)
    columns = _choose_columns(area, section, priority, platform, has_timeline, include_notes)
    return _format_table(data, columns)


def _sort_key(row):
    return (_priority_rank(str(row.get("Priority") or "")),
            _status_rank(str(row.get("status") or "")),
            str(row.get("Feature") or "").lower())


def _priority_rank(priority):
    order = [FOCUS_PRIORITY] + [p for p in _BASE_PRIORITY_ORDER if p != FOCUS_PRIORITY]
    return order.index(priority) if priority in order else len(order)


def _status_rank(status):
    return _STATUS_ORDER.index(status) if status in _STATUS_ORDER else len(_STATUS_ORDER)


def _choose_columns(area, section, priority, platform, has_timeline, include_notes):
    columns = ["Feature", "Priority", "status", "DRI"]
    if area or section:
        columns = ["Area", "Section"] + columns
    if platform or has_timeline is not None or priority == "P0.5":
        columns += _TIMELINE_COLS
    if include_notes:
        columns.append("MSA Premium Notes")
    return columns


def _format_table(data, columns):
    if not data:
        return "No matching features found."
    headers = [_DISPLAY.get(c, c) for c in columns]
    rows = [[(_format_owner(r.get(c)) if c == "DRI" else _cell(r.get(c))) for c in columns] for r in data]
    widths = [len(h) for h in headers]
    for row in rows:
        for i, val in enumerate(row):
            widths[i] = max(widths[i], len(val))

    def line(cells):
        return " | ".join(val.ljust(widths[i]) for i, val in enumerate(cells))

    noun = "feature" if len(data) == 1 else "features"
    out = [f"{len(data)} {noun}:", line(headers), "-" * (sum(widths) + 3 * (len(widths) - 1))]
    out += [line(row) for row in rows]
    return "\n".join(out)


def _cell(value):
    if value is None or str(value).strip() == "":
        return "Not set"
    return str(value).replace("\n", " ")


# ----------------------------------------------------------------------------
# Tool 2: get_feature
# ----------------------------------------------------------------------------

def get_feature(feature):
    data = parse_sheet(SOT_PATH)
    matches = [r for r in data if feature.strip().lower() in str(r.get("Feature") or "").lower()]

    if not matches:
        return f"No feature found matching '{feature}'."
    if len(matches) > 1:
        names = "\n".join(f"  - {r.get('Feature')}" for r in matches)
        return f"{len(matches)} features match '{feature}'. Which one?\n{names}"

    return _feature_detail(matches[0])


def _feature_detail(r):
    notes = r.get("MSA Premium Notes")
    marker = _notes_marker(notes)
    note_label = f"Notes ({marker})" if marker else "Notes"
    lines = [
        str(r.get("Feature") or "Not set"),
        f"Area / Section : {_cell(r.get('Area'))} / {_cell(r.get('Section'))}",
        f"Priority       : {_cell(r.get('Priority'))}",
        f"Status (web)   : {_cell(r.get('status'))}",
        f"Owner          : {_format_owner(r.get('DRI'))}",
        f"Desktop        : {_format_timeline(r.get('Desktop Timeline'))}",
        f"Mobile         : {_format_timeline(r.get('Mobile Timeline'))}",
        f"Blockers       : {_extract_blockers(notes) or 'Not set'}",
        f"{note_label}{' ' * max(0, 15 - len(note_label))}: {_cell(notes)}",
        f"Issue          : {_cell(_issue(r))}",
    ]
    return "\n".join(lines)


def _issue(row):
    for k, v in row.items():
        if str(k).lower().startswith("issue"):
            return v
    return None


def _notes_marker(notes):
    m = re.search(r"\((\d{1,2}\.\d{1,2})\)", str(notes or ""))
    return m.group(1) if m else None


def _extract_blockers(notes):
    if not notes:
        return None
    found = []
    for line in str(notes).split("\n"):
        if re.search(r"blocked:|dependent on", line, re.I):
            found.append(line.strip(" *•\t"))
    return "; ".join(found) if found else None


def _format_owner(value):
    if value is None or str(value).strip() == "":
        return "Not set"
    parts = re.split(r"[;\n]", str(value))
    return ", ".join(p.strip() for p in parts if p.strip())


# ----------------------------------------------------------------------------
# Timeline parsing (Option B): SDF = staging, WW = shipping
# ----------------------------------------------------------------------------

def _parse_milestones(cell):
    """Return {'SDF': value, 'WW': value} where value is a datetime or a state
    string ('Shipped'/'Available'). Unknown labels (e.g. MSIT) are ignored."""
    result = {}
    if not cell:
        return result
    s = str(cell)
    # Format A: "LABEL: value"
    for label, val in re.findall(r"(SDF|WW|MSIT)\s*:\s*([^\n,]+)", s, re.I):
        if label.upper() in ("SDF", "WW"):
            result[label.upper()] = _parse_value(val.strip())
    # Format B: "M/D (LABEL)"
    for date_str, label in re.findall(r"(\d{1,2}/\d{1,2})\s*\(\s*(SDF|WW)\s*\)", s, re.I):
        result[label.upper()] = _parse_date(date_str)
    # Bare state, e.g. "Shipped"
    if not result and re.search(r"shipped|available", s, re.I):
        result["WW"] = s.strip()
    return result


def _parse_value(val):
    d = _parse_date(val)
    return d if d else val  # date, or a state word like "Available"


def _parse_date(text):
    m = re.search(r"(\d{1,2})/(\d{1,2})", str(text))
    if not m:
        return None
    return datetime(TIMELINE_YEAR, int(m.group(1)), int(m.group(2)))


def _format_timeline(cell):
    ms = _parse_milestones(cell)
    if not ms:
        return "Not set"
    parts = []
    for label in ("SDF", "WW"):
        if label in ms:
            v = ms[label]
            shown = f"{v.month}/{v.day}" if isinstance(v, datetime) else v
            parts.append(f"{label} {shown}")
    return ", ".join(parts) if parts else "Not set"


# ----------------------------------------------------------------------------
# Tool 3: summarize
# ----------------------------------------------------------------------------

def summarize(group_by, then_by=None, area=None, section=None, priority=None,
              status=None, owner=None):
    data = _apply_filters(parse_sheet(SOT_PATH), area, section, priority, status, owner)
    row_key = _GROUP_KEYS[group_by]

    if not then_by:
        counts = Counter(_group_val(r, row_key) for r in data)
        return _format_single(group_by, counts)

    col_key = _GROUP_KEYS[then_by]
    matrix = {}
    col_values = set()
    for r in data:
        rv, cv = _group_val(r, row_key), _group_val(r, col_key)
        matrix.setdefault(rv, Counter())[cv] += 1
        col_values.add(cv)
    return _format_matrix(group_by, then_by, matrix, col_values)


def _group_val(row, key):
    v = row.get(key)
    return str(v).replace("\n", " ") if v not in (None, "") else "Not set"


def _ordered(values, key):
    if key == "Priority":
        return sorted(values, key=_priority_rank)
    if key == "status":
        return sorted(values, key=_status_rank)
    return sorted(values)


def _format_single(group_by, counts):
    rows = _ordered(counts.keys(), _GROUP_KEYS[group_by])
    width = max([len(group_by)] + [len(r) for r in rows])
    out = [f"{group_by.capitalize()}".ljust(width) + " | Count"]
    out.append("-" * (width + 8))
    for r in rows:
        out.append(r.ljust(width) + f" | {counts[r]}")
    out.append("Total".ljust(width) + f" | {sum(counts.values())}")
    return "\n".join(out)


def _format_matrix(group_by, then_by, matrix, col_values):
    cols = _ordered(col_values, _GROUP_KEYS[then_by])
    rows = _ordered(matrix.keys(), _GROUP_KEYS[group_by])
    headers = [group_by] + cols + ["Total"]
    table = []
    for r in rows:
        line = [r] + [str(matrix[r].get(c, 0)) for c in cols] + [str(sum(matrix[r].values()))]
        table.append(line)
    totals = ["Total"] + [str(sum(matrix[r].get(c, 0) for r in rows)) for c in cols]
    totals.append(str(sum(sum(matrix[r].values()) for r in rows)))
    table.append(totals)

    widths = [len(h) for h in headers]
    for line in table:
        for i, v in enumerate(line):
            widths[i] = max(widths[i], len(v))

    def fmt(cells):
        return " | ".join(v.ljust(widths[i]) for i, v in enumerate(cells))

    out = [f"{group_by} x {then_by}", fmt(headers), "-" * (sum(widths) + 3 * (len(widths) - 1))]
    out += [fmt(line) for line in table]
    return "\n".join(out)


# ----------------------------------------------------------------------------
# Tool 4: schedule
# ----------------------------------------------------------------------------

def schedule(milestone="WW", platform=None, before=None, after=None,
             priority=None, status=None, owner=None):
    milestone = milestone.upper()
    data = _apply_filters(parse_sheet(SOT_PATH), priority=priority, status=status, owner=owner)
    platforms = [platform] if platform in ("desktop", "mobile") else ["desktop", "mobile"]
    col = {"desktop": "Desktop Timeline", "mobile": "Mobile Timeline"}
    before_d, after_d = _parse_date(before), _parse_date(after)

    entries = []
    for r in data:
        for p in platforms:
            ms = _parse_milestones(r.get(col[p]))
            d = ms.get(milestone)
            if not isinstance(d, datetime):  # exclude states (Shipped/Available) and missing
                continue
            if before_d and d > before_d:
                continue
            if after_d and d < after_d:
                continue
            entries.append((d, r.get("Feature"), p, r.get("status")))

    entries.sort(key=lambda e: e[0])
    if not entries:
        return f"No features with a {milestone} date matching those bounds."

    label = "Shipping (WW)" if milestone == "WW" else "Staging (SDF)"
    rows = [[f, p, f"{d.month}/{d.day}", str(s or "Not set")] for d, f, p, s in entries]
    headers = ["Feature", "Platform", milestone, "Status"]
    widths = [len(h) for h in headers]
    for row in rows:
        for i, v in enumerate(row):
            widths[i] = max(widths[i], len(v))

    def line(cells):
        return " | ".join(v.ljust(widths[i]) for i, v in enumerate(cells))

    out = [f"{label}, sorted soonest first:", line(headers),
           "-" * (sum(widths) + 3 * (len(widths) - 1))]
    out += [line(row) for row in rows]
    return "\n".join(out)


# Lookup so the agent can run whichever tool Claude picks
DISPATCH = {
    "find_features": find_features,
    "get_feature": get_feature,
    "summarize": summarize,
    "schedule": schedule,
}


# --- Quick test when run directly ---
if __name__ == "__main__":
    print("=== get_feature('Auto') ===")
    print(get_feature("Auto"))
    print("\n=== get_feature('Search') (ambiguous) ===")
    print(get_feature("Search"))
    print("\n=== summarize(priority x status) ===")
    print(summarize("priority", "status"))
    print("\n=== summarize(status for P0.5) ===")
    print(summarize("status", priority="P0.5"))
    print("\n=== schedule(WW) ===")
    print(schedule(milestone="WW"))
