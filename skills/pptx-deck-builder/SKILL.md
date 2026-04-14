---
name: pptx-deck-builder
description: >
  Builds branded performance presentation decks (.pptx) from a user-provided PowerPoint template.
  The skill extracts the template's exact visual chrome (header/footer bar images, background images,
  fonts, colors, slide dimensions) and builds fresh data-filled slides that match the brand perfectly.
  Use this skill whenever the user wants to: create a performance deck or report from their own .pptx
  template, build slides from Intentwise account data or a CSV file, generate KPI scorecards / trend
  charts / data tables / insight slides in a branded format, or says anything like "use my template",
  "build a deck", "create slides", "make a report", "build a presentation" with a .pptx file involved.
  Trigger even if the user only mentions a template file path without explicitly saying "skill".
---

# PPTX Deck Builder

Build branded performance decks by combining a user's .pptx template with live data from Intentwise
or a CSV file. The output slides inherit the exact visual identity of the template — not just colors,
but the actual chrome images (header bars, footer bars, backgrounds) placed at pixel-perfect positions.

## Prerequisites
```bash
pip install python-pptx pillow lxml
```

---

## Workflow

### Step 1 — Receive the template

The user will provide a `.pptx` file path. Accept it from the message or ask:
> "What's the path to your PowerPoint template?"

Then run the **template inspector** to build a profile:

```bash
python3 ~/.claude/skills/pptx-deck-builder/build_deck.py inspect "/path/to/template.pptx"
```

This prints a JSON profile with:
- `slide_w_in`, `slide_h_in` — slide dimensions in inches
- `layouts` — list of layout names (e.g. `["TITLE", "DEFAULT"]`)
- `chrome.title_slides` — background image positions/sizes on TITLE-type slides
- `chrome.default_slides` — header bar + footer bar positions/sizes on DEFAULT slides
- `fonts.heading`, `fonts.body` — detected font names
- `colors.title_text`, `colors.header_text`, `colors.accent` — detected hex colors
- `content_area` — `{top, bottom, left, right}` in inches (safe zone for content)

Understand the profile before planning slides. If inspection fails, read the error and fix the path or ask the user.

---

### Step 2 — Plan the slides

Ask the user what they want in the deck. Guide them with this menu of slide types:

| Type | What it shows |
|------|--------------|
| `cover` | Title, client name, period, presenter |
| `section` | Section divider with number, title, subtitle |
| `kpi` | Grid of metric cards with value + delta vs prior period |
| `chart` | Column, bar, or line trend chart |
| `table` | Sortable data table (campaigns, ASINs, keywords, etc.) |
| `insights` | WIN / RISK / ACTION / INFO tagged bullets + optional callout metric |

For each slide ask:
1. Slide type
2. Slide title
3. What metrics/data to show
4. Data source: **Intentwise** (pull live) or **CSV** (file path)

Present the full plan back to the user as a numbered list and confirm before building.

**Example plan:**
```
1. cover       — "Amazon Performance Overview" | Wiser Drones | Mar 15–Apr 13 2026
2. section     — 01 Advertising Performance
3. kpi         — Advertising Summary (ROAS, ACOS, Spend, Sales, Orders, Impressions, Clicks, CTR)
4. chart       — Weekly Spend vs Ad Sales (column, 4 weeks)
5. table       — Top 10 Campaigns by Ad Sales
6. section     — 02 Retail Performance
7. kpi         — Retail Summary (Ordered Revenue, Shipped Revenue, Glance Views, TACOS)
8. insights    — Key Insights & Recommendations
```

---

### Step 3 — Source the data

#### Path A — Intentwise live data

Use the Intentwise MCP tools in this order:

1. `get_organization` → get `organization_id`
2. `get_intentwise_accounts(organization_id)` → find the right account, note `account_id` and available `data_source_type` values
3. `search_schema(text, channel, data_source)` → find the right table/entity for each metric set
4. `get_insights(text, organization_id, account_id, channel, data_source, entity)` → get the data

Pull **current period** and **prior period** in parallel when computing deltas.

Parse the returned `rows` into metric dicts like:
```python
{"name": "ROAS", "value": 3.51, "metric_type": "multiplier", "delta_pct": -6.4}
```

**Metric types:** `currency` · `multiplier` · `percent` · `number` · `ratio` · `days` · `rank`

**Delta sign convention:**
- Higher = better metrics (revenue, ROAS, orders, sessions, CVR): positive delta → GREEN ▲
- Cost/efficiency metrics (ACOS, TACOS, CPC, spend, stranded): positive delta → RED ▲
- Rank metrics (BSR, keyword rank): numerically lower = better, invert sign

#### Path B — CSV upload

User provides a file path. Read with Python:
```python
import csv
with open(path) as f:
    rows = list(csv.DictReader(f))
```

Ask the user which columns map to which metrics, then build metric dicts.

---

### Step 4 — Build the deck

Write a Python script that:
1. Imports `DeckBuilder` from `~/.claude/skills/pptx-deck-builder/build_deck.py`
2. Instantiates it with the template path and desired output path
3. Calls the appropriate `add_*` methods in slide order
4. Calls `save()`

Then run it with `python3 /tmp/build_<client>_deck.py`.

**Template script:**
```python
import sys
sys.path.insert(0, "/Users/intentwiseks/.claude/skills/pptx-deck-builder")
from build_deck import DeckBuilder

builder = DeckBuilder(
    template_path="/path/to/template.pptx",
    output_path="/path/to/output.pptx",
)

builder.add_cover(
    title="Amazon Performance Overview",
    client_name="Wiser Drones",
    period="Mar 15 – Apr 13, 2026",
    report_type="Advertising · Vendor Retail",
    presenter="Intentwise",
    date="April 14, 2026",
)

builder.add_section("Advertising Performance", subtitle="Sponsored Products", number="01")

builder.add_kpi(
    slide_title="Advertising Summary",
    period_label="Mar 15–Apr 13 vs Feb 13–Mar 14, 2026",
    metrics=[
        {"name": "Ad Sales",  "value": 83818, "metric_type": "currency",   "delta_pct": 78.6},
        {"name": "ROAS",      "value": 3.51,  "metric_type": "multiplier", "delta_pct": -6.4},
        {"name": "ACOS",      "value": 28.5,  "metric_type": "percent",    "delta_pct": 7.0},
        {"name": "Spend",     "value": 23884, "metric_type": "currency",   "delta_pct": 91.1},
        {"name": "Orders",    "value": 5775,  "metric_type": "number",     "delta_pct": 36.6},
        {"name": "Impressions","value": 4195183,"metric_type":"number",    "delta_pct": 81.4},
        {"name": "Clicks",    "value": 23270, "metric_type": "number",     "delta_pct": 39.1},
        {"name": "CTR",       "value": 0.55,  "metric_type": "percent",    "delta_pct": -23.6},
    ]
)

builder.add_chart(
    slide_title="Weekly Spend vs Ad Sales",
    chart_type="column",   # column | bar | line | bar_stacked
    subtitle="Mar 16 – Apr 6, 2026 | Complete Weeks",
    chart_data={
        "categories": ["Wk Mar 16", "Wk Mar 23", "Wk Mar 30", "Wk Apr 6"],
        "series": [
            {"name": "Ad Spend", "values": [4873, 6000, 6347, 5297]},
            {"name": "Ad Sales", "values": [19143, 21204, 21109, 17555]},
        ]
    }
)

builder.add_table(
    slide_title="Top 10 Campaigns by Ad Sales",
    subtitle="Sorted by Ad Sales",
    column_headers=["Campaign", "Type", "Ad Sales", "Spend", "ROAS", "ACOS", "Orders"],
    rows=[
        ["Campaign A", "SP A", "$13,783", "$2,312", "5.96x", "16.8%", "471"],
        # ... more rows
    ],
    highlight_col=2,
)

builder.add_insights(
    slide_title="Key Insights & Recommendations",
    bullets=[
        {"lead": "Ad Sales up +79%", "body": "...", "tag": "WIN"},
        {"lead": "ROAS slipped -6%", "body": "...", "tag": "RISK"},
        {"lead": "Increase WISER-055 budget", "body": "...", "tag": "ACTION"},
    ],
    highlight_metric={"label": "Revenue Opportunity", "value": "$12K+", "caption": "..."},
)

path = builder.save()
print(f"✅ Saved: {path}  ({builder.slide_count} slides)")
```

---

### Step 5 — Confirm output

After the script runs successfully:
- Report the output file path and total slide count
- Briefly describe each section
- Offer to adjust: slide order, font sizes, card layout, chart type, color overrides

---

## Critical build rules

These are the hard-won lessons from building real decks — skipping any of them causes broken or ugly output.

### 1. Strip template slides correctly
The naive approach of removing IDs from `_sldIdLst` leaves image blobs in the PPTX zip, causing duplicate-name warnings and wrong slides appearing on open. The correct approach:
```python
for sldId in list(prs.slides._sldIdLst):
    rId = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    if rId:
        try: prs.part.drop_rel(rId)
        except: pass
    prs.slides._sldIdLst.remove(sldId)
```
`DeckBuilder.__init__` does this automatically.

### 2. Use actual chrome images, not rectangles
Every template has real PNG/JPEG images for header bars, footer bars, and backgrounds. Extract their blobs with `shape.image.blob` and re-insert with `slide.shapes.add_picture(io.BytesIO(blob), ...)`.
Never try to simulate these with colored rectangles — the result will never match.

### 3. Never assume slide dimensions
Templates vary wildly. The SQP Webinar template is 20"×11.25", not the standard 13.33"×7.5". Always read `prs.slide_width` and `prs.slide_height` and derive all positions as fractions of these values.

### 4. Derive content area from chrome
Don't hardcode content positions. Calculate:
```
content_top    = max(header.bottom for header in chrome) + padding
content_bottom = min(footer.top for footer in chrome) - padding
left_margin    = min(header.left for header in chrome)
right_margin   = slide_width - left_margin
content_width  = right_margin - left_margin
```

### 5. Do not use generate_slides.py
That library has Intentwise-red hardcoded into its `THEME` global. Patching it is fragile. `build_deck.py` is fully self-contained and uses only the extracted template colors.

---

## Slide type reference

### KPI metrics dict
```python
{
    "name": "ROAS",           # Display label
    "value": 3.51,            # Numeric value
    "metric_type": "multiplier",  # currency | multiplier | percent | number | ratio | days | rank
    "delta_pct": -6.4,        # Optional. % change vs prior period (omit key if no comparison)
}
```

### Insight bullet dict
```python
{
    "lead": "Short bold headline",
    "body": "Supporting detail sentence.",
    "tag": "WIN",   # WIN | RISK | ACTION | INFO
}
```

### Highlight metric dict (optional callout in insights slide)
```python
{
    "label": "Revenue Opportunity",
    "value": "$12K+",
    "caption": "estimated from reallocation",
}
```
