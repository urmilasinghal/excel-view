# Excel View: Tool Specifications (Step 5)

The tools the agent can call. This is the product-design layer: it defines what
questions the agent can answer, what parameters matter, and what output looks like.

**Status:** complete. All four tools defined: `find_features`, `get_feature`,
`summarize`, `schedule`. Ready for Step 6 (build the agent).

---

## Shared conventions (apply to every tool)

- **Audience:** anyone interested in Feature Parity, not just the PM. Names and
  output must read clearly without prior context.
- **Output:** tight and consistent formatting, no extra spacing or blank-line
  padding. Tables for lists, short narrative for a single or few results.
- **Missing data:** shown as `Not set`.
- **Scope:** MSA Premium Detail tab only for now. Issues and Helix tabs later.
- **Status is web-only** (`MCA Web MSA Gap`). Mobile parity status is pending
  from the mobile team. Cross-platform parity questions are timeline-only today.
- **Timelines:** Desktop and Mobile columns are canonical. SDF = staging,
  WW = shipping. Year assumed 2026. Unknown labels ignored. Populated only for
  P0.5 features today. Legacy single `Timeline` column is ignored.

---

## Tool 1: find_features

**Purpose:** find and list features matching any combination of filters. The
workhorse tool, answers most list-style questions.

**Parameters (all optional, combine with AND):**

| Parameter | Meaning |
|-----------|---------|
| `area` | Filter by Area (e.g. Left Nav) |
| `section` | Filter by Section |
| `priority` | P0, P0.5, P1, P2 |
| `status` | Parity, In Progress, Gap, By Design, Under Review |
| `owner` | DRI name. Includes shared-ownership rows (multi-owner cells split on `;` or newline) |
| `platform` | desktop or mobile. Scopes timeline-related filters |
| `has_timeline` | true / false. With no platform: timeline on either platform. With platform: that platform only |
| `keyword` | substring search across Feature + MSA Premium Notes + Internal Notes |
| `include_notes` | default false. When true (or when the question is about blockers), notes are returned |

**Output:**
- A count line ("12 features:") then the rows.
- **Smart default columns:**
  - Always: Feature, Priority, Status, Owner
  - Add Desktop/Mobile when the query touches timeline, platform, or P0.5
  - Add Area/Section when the query groups or filters by them
- Empty fields shown as `Not set`.

**Default sort:**
1. Focus priority first (configurable, default `P0.5`), then P0, P1, P2
2. Status, attention-first: Gap, In Progress, Under Review, By Design, Parity
3. Feature name A to Z

---

## Tool 2: summarize

**Purpose:** counts and rollups for reporting. Feeds status reports.

**Parameters:**

| Parameter | Meaning |
|-----------|---------|
| `group_by` | dimension to count by: area, section, priority, status, owner |
| `then_by` | optional second dimension for a cross-tab (e.g. priority x status) |
| filters | same filters as `find_features`, to summarize a subset |

**Output:** a count table with totals. Single dimension is a list with a total;
two dimensions is a matrix with row and column totals. No percentages.

```
Priority x Status
Priority | Parity | In Progress | Gap | By Design | Total
P0       | 80     | 1           | 0   | 6         | 87
P0.5     | 1      | 10          | 2   | 0         | 13
P1       | 32     | 104         | 14  | 14        | 164
P2       | 41     | 1           | 31  | 15        | 88
Total    | 154    | 116         | 47  | 35        | 352
```

## Tool 3: get_feature

**Purpose:** the deep dive. Return everything about one named feature in a single
clean view.

**Parameter:**

| Parameter | Meaning |
|-----------|---------|
| `feature` | Feature name or partial. Case-insensitive partial match |

**Matching behavior:**
- One match: return full detail.
- Multiple matches: return a short "which one did you mean?" list. Do not guess.
- No match: "No feature found matching 'X'."

**Output:** narrative labeled block (not a table, since it's one feature). Tight,
no extra spacing. Notes and extracted blockers included by default.

```
Auto Suggestions
Area / Section : Input Box / Compose
Priority       : P0.5
Status (web)   : In Progress
Owner          : Sudheender Srinivasan, Che-bin Liu
Desktop        : SDF 6/4, WW 6/11
Mobile         : Not set
Blockers       : Historical Prompts blocked pending Storage team
Notes (5.29)   : Curated prompts working in internal testing, Bing prompts blocked on infra
Issue          : Not set
```

Fields: Feature, Area/Section, Priority, Status (web), Owner(s), Desktop
timeline, Mobile timeline, Blockers (extracted from notes), Notes (with freshness
marker), Issue link.

## Tool 4: schedule

**Purpose:** the time-based view. Answer "what ships or stages when, by platform."
Uses the SDF/WW date parsing.

**Parameters:**

| Parameter | Meaning |
|-----------|---------|
| `milestone` | SDF (staging) or WW (shipping). Default WW |
| `platform` | desktop, mobile, or both (default both) |
| `before` / `after` | date bounds (e.g. ships before 6/15, after 6/1) |
| filters | same filters as `find_features`, to scope |

**Output:** features with their milestone date, sorted **soonest-first**.

```
Shipping (WW) by 6/13, desktop:
Feature           | WW Ship | Status
Auto Suggestions  | 6/11    | In Progress
```

**Forward-looking only.** Cells with `Shipped` or `Available` (no real date) are
**excluded** from schedule results. To see those, use `get_feature` or
`find_features`.

Notes: year assumed 2026. Timelines populated only for P0.5 today.
