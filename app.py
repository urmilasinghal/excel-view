"""
Excel View - Local Web Viewer

Run this file to start the website on your computer.
Open your browser to http://localhost:5000 to see your data.

Usage:
    python app.py
    python app.py --file data/your_file.xlsx --sheet "Sheet Name"
"""

import sys
import os

# Allow importing from src/ folder
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from flask import Flask, render_template, request
from parser import parse_sheet, parse_workbook, get_schema
from query import filter_data, sort_data, rollup, search

app = Flask(__name__)

# --- Configuration ---
DATA_FILE = os.path.join(os.path.dirname(__file__), "data", "test_sot.xlsx")
SHEET_NAME = "MSA Premium Detail"

# --- Load data once at startup ---
data = parse_sheet(DATA_FILE, SHEET_NAME)
schema = get_schema(DATA_FILE, SHEET_NAME)

# Columns to show filters for
FILTER_COLUMNS = [
    "Area",
    "Priority",
    "DRI",
    "Desktop Timeline_status",
    "Mobile Timeline_status",
]

# Columns to display in the table
DISPLAY_COLUMNS = [
    "Area",
    "Section",
    "Feature",
    "Priority",
    "DRI",
    "Desktop Timeline",
    "Desktop Timeline_status",
    "Mobile Timeline",
    "Mobile Timeline_status",
    "Issue",
    "MSA Premium Notes",
]

# Status colors for the HTML (background colors for status cells)
STATUS_COLORS = {
    "Parity": "#E2EFDA",
    "By Design": "#FFF2CC",
    "In Progress": "#FCE4D6",
    "Gap": "#FADBD8",
}

STATUS_PILL_COLORS = {
    "Parity": "#92D050",
    "By Design": "#FFD700",
    "In Progress": "#FFA500",
    "Gap": "#E74C3C",
}


def load_data():
    """Load data from the Excel file."""
    global data, schema
    data = parse_sheet(DATA_FILE, SHEET_NAME)
    schema = get_schema(DATA_FILE, SHEET_NAME)
    print(f"Loaded {len(data)} features from {SHEET_NAME}")


@app.route("/")
def index():
    # --- Apply filters from URL parameters ---
    filtered = list(data)
    active_filters = {}

    for col in FILTER_COLUMNS:
        param_name = col.replace(" ", "_")
        val = request.args.get(param_name, "All")
        if val and val != "All":
            active_filters[col] = val
            filtered = [
                r for r in filtered
                if str(r.get(col, "")).lower() == val.lower()
            ]

    # --- Search ---
    search_text = request.args.get("search", "").strip()
    if search_text:
        filtered = search(filtered, search_text)

    # --- Sort ---
    sort_col = request.args.get("sort", "")
    sort_desc = request.args.get("desc", "false") == "true"
    if sort_col:
        filtered = sort_data(filtered, by=sort_col, descending=sort_desc)

    # --- Rollups ---
    desktop_rollup = rollup(filtered, "Desktop Timeline_status")
    mobile_rollup = rollup(filtered, "Mobile Timeline_status")
    priority_rollup = rollup(filtered, "Priority")

    # --- Unique values for filter dropdowns ---
    unique_values = {}
    for col in FILTER_COLUMNS:
        values = set()
        for r in data:
            v = r.get(col)
            if v is not None and str(v).strip():
                values.add(str(v))
        unique_values[col] = sorted(values)

    return render_template(
        "index.html",
        data=filtered,
        total=len(data),
        filtered_count=len(filtered),
        active_filters=active_filters,
        search_text=search_text,
        sort_col=sort_col,
        sort_desc=sort_desc,
        unique_values=unique_values,
        filter_columns=FILTER_COLUMNS,
        display_columns=DISPLAY_COLUMNS,
        desktop_rollup=desktop_rollup,
        mobile_rollup=mobile_rollup,
        priority_rollup=priority_rollup,
        status_colors=STATUS_COLORS,
        status_pill_colors=STATUS_PILL_COLORS,
        sheet_name=SHEET_NAME,
    )


@app.route("/reload")
def reload_data():
    """Reload data from the Excel file (manual refresh)."""
    load_data()
    return '<script>window.location="/";</script>'


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Excel View - Local Web Viewer")
    parser.add_argument("--file", default=DATA_FILE, help="Path to xlsx file")
    parser.add_argument("--sheet", default=SHEET_NAME, help="Sheet name to display")
    parser.add_argument("--port", type=int, default=5000, help="Port number")
    args = parser.parse_args()

    DATA_FILE = args.file
    SHEET_NAME = args.sheet

    load_data()

    print()
    print("=" * 50)
    print("  Excel View is running!")
    print(f"  Open your browser to: http://localhost:{args.port}")
    print("  Press Ctrl+C to stop")
    print("=" * 50)
    print()

    app.run(debug=True, port=args.port)
