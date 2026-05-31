# NOW

*What to work on next. Updated as steps complete.*

---

**Last Updated:** May 31, 2026

## Status

Steps 1-4 complete (parser + query engine). Flask web viewer built and tested.

Local website runs at localhost:5000. Shows SOT data in a table with:
- Rollup bars (Desktop status, Mobile status, Priority counts)
- Dropdown filters (Area, Priority, DRI, Desktop status, Mobile status)
- Text search
- Sortable column headers (click to sort)
- Color-coded status cells
- Reload button to refresh data from the xlsx

Working against test data (32 rows). Real SOT file not yet integrated.

## Active

**Strategy change:** Three-step approach before agent work.
1. **Flask local website (done).** Parse SOT, view in browser, filter, sort, rollup.
2. **Power Apps + SharePoint.** Read SOT directly from SharePoint. Learn Microsoft tools.
3. **Discuss SharePoint hosting.** Internal hosting for team/leadership access.

Next immediate step: Urmila runs the Flask app on her machine for the first time.

## Blocked

- Real SOT file needed for color calibration (exact hex values).
- "Issue" column: full header name and content type to confirm.

## Next (after Flask validation)

- Urmila tests the Flask app locally with real SOT data
- Step 2: Power Apps connected to SharePoint
- Step 5: Tool definitions (agent work, later)

## Deferred

- Agent (Steps 5-8): tool definitions, Claude API, conversational loop, MCP
- Generalize to any xlsx file
- Write-back to xlsx
