# Excel View: Design Document

**Project:** Excel query agent
**Phase 1 Target:** Entra-MSA Feature Parity SOT
**Created:** May 29, 2026
**Status:** Design complete. Ready to build.

---

## Final State

An AI agent that can read any Excel file, understand its structure, and let users interact in natural language. "Show me all P0 gaps." "Who owns the most in-progress items?" "Mark this feature as Parity and set the date to today." It reads, views, filters, rolls up, and writes back. Exposed as an MCP so any AI assistant can plug in.

---

## Build Path: 8 Steps from Zero to Agent

Agent-first approach. No web viewer on the critical path. Each step teaches one thing, builds one thing, and works on its own.

### Steps 1-4: Parser and Query Layer (Claude builds, Urmila validates)

Engineering work. Claude builds the code. Urmila reviews the output to confirm data and status mapping are correct.

| Step | Build | Learn |
|------|-------|-------|
| 1 | Open the xlsx with Python. Print column headers. | How Python reads Excel files. |
| 2 | Read every row from MSA Premium Detail tab. Print as structured data. | How to pull structured data from a spreadsheet. |
| 3 | Extract cell fill colors from Desktop/Mobile Timeline. Map to status. | How to read cell formatting, not just values. |
| 4 | Write query functions: get all P0 features, get all gaps, filter by DRI. | How to make data queryable with code. |

### Steps 5-8: Agent and MCP (Urmila drives)

Product and platform work. This is where the PM skills live.

| Step | Build | Learn |
|------|-------|-------|
| 5 | Define tool specifications: names, descriptions, parameters, output shapes. | What a "tool" means in AI. Product design for agents. |
| 6 | Send a question + tool definitions to Claude API. Claude picks the tool. Code runs it. Claude answers. | How agents actually work under the hood. |
| 7 | Add a loop. Ask follow-up questions. The agent keeps context. | Multi-turn tool use. Interaction design. |
| 8 | Wrap tool functions as an MCP server. Connect to Claude.ai. | How MCP works. Platform thinking. Exposing tools for any AI. |

### After Step 8: Future Phases

| Phase | What It Adds |
|-------|-------------|
| Generalize | Same agent, any xlsx file (not just Feature Parity SOT) |
| Write-back | Agent can update the xlsx. "Mark this as Parity." |
| Web viewer (optional) | Visual dashboard layer on top of the agent for team/leadership audiences |

---

## Why This Order

- Steps 1-4 are engineering. Claude builds the parser. Urmila validates output. PM doesn't need to write parsing code but must understand the data model it produces, because the data model shapes what tools she can define.
- Step 5 is the most important step. This is where Urmila decides what questions the agent can answer. What filters matter. What output looks like. Pure product design.
- Step 6 teaches what happens under the hood: LLM + tools + execution loop. Understanding this makes MCP (Step 8) make sense instead of feeling like magic.
- Step 7 adds conversation. Follow-ups, context, interaction patterns.
- Step 8 wraps everything as MCP so Claude.ai or any AI client can use the tools directly. No script to run each time.

---

## Key Design Decisions (Evolution)

The design went through three iterations in one session:

**Iteration 1: Web viewer with features.**
Named views, shareable URLs, rollups, auto-sync, audience staging (1A/1B/1C). Nine architecture decisions. Over-engineered.

**Iteration 2: Work backward from final state.**
Final state locked (agent + MCP + any Excel + read/write + NL). Tried working backward. Got tangled. Two paths evaluated: Path A (no web viewer) vs Path B (web viewer included). Path B chosen, then questioned.

**Iteration 3: Agent-first, bite-sized steps.**
Urmila asked: "Do we really need a web viewer?" Then described her actual use case: "Give you the file, ask a question, get an answer." That's the agent, not a web viewer. Collapsed to 8 build steps. Web viewer moved to optional future phase.

---

## Data Schema: MSA Premium Detail

### Columns (post-update)

| # | Column | Type | Notes |
|---|--------|------|-------|
| 1 | Area | Text | Grouping level 1 (Left Nav, Header, etc.) |
| 2 | Section | Text | Grouping level 2 (Navigation, Links & Sections, etc.) |
| 3 | Feature | Text | Feature name |
| 4 | Priority | Text | P0, P1, P2 |
| 5 | DRI | Text | Owner name |
| 6 | Desktop Timeline | Date + Color | Date value with cell fill color encoding status |
| 7 | Mobile Timeline | Date + Color | Date value with cell fill color encoding status |
| 8 | Issue | Text | TBD: confirm full header name and content type |
| 9 | MSA Premium Notes | Text | Free-text notes |

### Color-to-Status Mapping

| Cell Fill Color | Status |
|-----------------|--------|
| Green | Parity |
| Yellow | By Design |
| Orange | In Progress |
| No fill | Gap |

**Note:** Exact hex color values need confirmation from the actual xlsx file. Parser will use tolerance-based color matching.

### Parsing Rules

- Status derived from cell fill color in columns 6 and 7, not from a text column
- Parser reads cell formatting (openpyxl fill properties), not just cell values
- LINE BREAK separator rows removed from source file by Urmila
- Color matching: nearest-match with tolerance for shade variations
- Empty rows or non-data rows skipped

---

## Workbook Tabs

| Tab | Step 1-4 Target | Notes |
|-----|-----------------|-------|
| Summary | No | Future: generalize phase |
| MSA Premium Detail | Yes | Primary data tab |
| MSA Helix Migration | No | Future: generalize phase |
| Legend | No | TBD: confirm tab exists |

---

## Open Items

- [ ] Get actual xlsx file for hex color extraction
- [ ] "Issue" column: confirm full header name and content type
- [ ] Confirm all tabs in the workbook
- [ ] Urmila to update xlsx (remove columns F, I, LINE BREAK rows)

---

## Decisions Log

| Date | Decision | Context |
|------|----------|---------|
| 2026-05-29 | Final state = Agent + MCP + any Excel + read/write + NL | End goal locked |
| 2026-05-29 | Agent-first, not web-viewer-first | Urmila's use case is NL queries, not UI filters |
| 2026-05-29 | 8 build steps from zero to agent | Bite-sized build + learn progression |
| 2026-05-29 | Steps 1-4 = Claude builds, Steps 5-8 = Urmila drives | PM focuses on tool design, agent, MCP |
| 2026-05-29 | MCP is Step 8, separate from agent (Step 6) | Step 6 teaches the mechanics; Step 8 wraps as platform |
| 2026-05-29 | Web viewer = optional future phase | Not on critical path to agent |
| 2026-05-29 | Phase 1 file = Feature Parity SOT | Updated format |
| 2026-05-29 | xlsx is the SOT | Agent reads only (write-back is future phase) |
| 2026-05-29 | Status encoded in cell color, not text column | Original columns F and I removed |
| 2026-05-29 | No fill = Gap | Color mapping confirmed |
| 2026-05-29 | Path B evaluated then superseded | Agent-first made web viewer optional |
