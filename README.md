# Excel View

**Excel query agent.** Reads an xlsx file, understands its structure and color-coded status, and answers natural language questions. Starting with the Entra-MSA Feature Parity SOT, then generalizing to any Excel file. End state: an MCP server that any AI assistant can plug into.

## Architecture

```
xlsx file (color-coded status in cell fills)
    |
Parser (openpyxl, reads values + cell formatting)
    |
Query Engine (filter, sort, search, rollup)
    |
Tool Definitions (structured function specs)
    |
Agent (Claude API + tool use)
    |
MCP Server (expose tools for any AI client)
```

Eight build steps, bottom to top. See [docs/design.md](docs/design.md) for the full plan.

## Orientation

- [CLAUDE.md](CLAUDE.md) -- operator manual for Claude Code sessions
- [NOW.md](NOW.md) -- current step and what's next
- [docs/design.md](docs/design.md) -- full design: final state, data schema, color mapping, decisions log

## Repo Layout

```
excel-view/
├── CLAUDE.md              Operator manual for Claude Code
├── NOW.md                 Current priorities
├── README.md              This file
├── requirements.txt       Python dependencies
├── data/
│   └── test_sot.xlsx      Test data (32 rows, SOT schema, color-coded)
├── docs/
│   └── design.md          Full design document
└── src/
    ├── parser.py           Steps 1-3: read xlsx, extract data, map colors
    ├── query.py            Step 4: filter, sort, search, rollup
    ├── tools.py            Step 5: tool definitions (pending)
    ├── agent.py            Steps 6-7: agent + conversational loop (pending)
    └── mcp_server.py       Step 8: MCP server (pending)
```

## Setup

```bash
python -m venv venv
source venv/bin/activate      # Linux/Mac
venv\Scripts\activate          # Windows
pip install -r requirements.txt
```

## Quick Test

```bash
cd src
python parser.py ../data/test_sot.xlsx
python query.py ../data/test_sot.xlsx
```

## Current Status

Steps 1-4 complete (parser + query engine). Step 5 (tool definitions) is next. See [NOW.md](NOW.md).

## Data Schema

The Feature Parity SOT has 9 columns. Status is encoded in cell fill colors, not in a text column.

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
