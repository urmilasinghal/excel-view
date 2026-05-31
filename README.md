# Excel View

**Excel query agent.** Reads an xlsx file, understands its structure and color-coded status, and shows it in a browser. Starting with the Entra-MSA Feature Parity SOT. Future: Power Apps + SharePoint, then a conversational agent and MCP server.

## Quick Start

```bash
# First time setup
python -m venv venv
venv\Scripts\activate          # Windows
pip install -r requirements.txt

# Run the website
python app.py
```

Open your browser to **http://localhost:5000**. That's it.

## What You See

- **Rollup bars** at the top: counts by Desktop status, Mobile status, and Priority
- **Filter dropdowns**: Area, Priority, DRI, Desktop status, Mobile status
- **Search box**: find features by name or notes
- **Sortable columns**: click any column header to sort
- **Color-coded status cells**: Parity (green), By Design (yellow), In Progress (orange), Gap (red)
- **Reload button**: re-reads the xlsx file without restarting

## Using Your Own Data

```bash
python app.py --file "C:\path\to\your\file.xlsx" --sheet "Sheet Name"
```

## Orientation

- [CLAUDE.md](CLAUDE.md) -- operator manual for Claude Code sessions
- [NOW.md](NOW.md) -- current priorities and what's next
- [docs/design.md](docs/design.md) -- full design: final state, build plan, data schema, decisions

## Repo Layout

```
excel-view/
├── CLAUDE.md              Operator manual for Claude Code
├── NOW.md                 Current priorities
├── README.md              This file
├── requirements.txt       Python dependencies (openpyxl, flask)
├── app.py                 Flask web server (run this)
├── templates/
│   └── index.html         Web page template
├── data/
│   └── test_sot.xlsx      Test data (32 rows, SOT schema)
├── docs/
│   └── design.md          Full design document
└── src/
    ├── parser.py           Reads xlsx, extracts data, maps cell colors to status
    └── query.py            Filter, sort, search, rollup functions
```

## Data Schema

The Feature Parity SOT has 9 columns. Status is encoded in cell fill colors.

| Column | Type |
|--------|------|
| Area | Text |
| Section | Text |
| Feature | Text |
| Priority | P0, P1, P2 |
| DRI | Text |
| Desktop Timeline | Date + color = status |
| Mobile Timeline | Date + color = status |
| Issue | Text |
| MSA Premium Notes | Text |

Color mapping: green = Parity, yellow = By Design, orange = In Progress, no fill = Gap.

## Build Path

| Step | What | Status |
|------|------|--------|
| 1-3 | Parse xlsx, extract data, read cell colors | Done |
| 4 | Query engine (filter, sort, rollup) | Done |
| Flask | Local website with filters and rollups | Done |
| Power Apps | Connect to SharePoint, Microsoft tools | Next |
| 5-8 | Tool definitions, agent, conversational loop, MCP | Future |
