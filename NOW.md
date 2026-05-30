# NOW

*What to work on next. Updated as steps complete.*

---

**Last Updated:** May 30, 2026

## Status

Steps 1-4 complete. Parser reads xlsx files, extracts structured data, maps cell fill colors to status. Query engine supports filtering, sorting, searching, and rollups. Tested against 32-row test dataset matching the real SOT schema.

Working against test data. Real SOT file not yet integrated (Urmila updating the xlsx: removing old columns F, I, and LINE BREAK rows).

## Active

1. **Step 5: Define tool specifications.** Urmila drives this. What questions should the agent answer? What parameters matter? What should the output look like? This is the product design step.

## Blocked

- Real SOT file needed for color calibration (exact hex values for green/yellow/orange).
- "Issue" column: full header name and content type to confirm.

## Next (after Step 5)

- Step 6: Build the agent (Claude API + tool use)
- Step 7: Conversational loop
- Step 8: MCP server

## Deferred

- Generalize to any xlsx file
- Write-back to xlsx
- Web viewer (optional visual layer for team/leadership)
