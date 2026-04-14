# ppt-template-builder

A Claude Code skill plugin that builds branded performance presentation decks (`.pptx`) from any PowerPoint template.

## What it does

Give Claude a `.pptx` template and tell it what data to include. The skill:

- **Extracts the exact visual chrome** — header bar images, footer bar images, background images — from your template and re-inserts them at pixel-perfect positions on every new slide
- **Pulls live data** from Intentwise (ROAS, ACOS, Spend, Ad Sales, Orders, etc.) or reads from a CSV you provide
- **Builds branded slides** — KPI scorecards, trend charts, data tables, section dividers, insights — that look like they came from your template

## Slide types

| Type | What it produces |
|------|-----------------|
| `cover` | Title, client name, period, presenter |
| `section` | Section divider with number, title, subtitle |
| `kpi` | Grid of metric cards with value + delta vs prior period |
| `chart` | Column, bar, or line trend chart |
| `table` | Data table (campaigns, ASINs, keywords, etc.) |
| `insights` | WIN / RISK / ACTION / INFO tagged bullets + optional callout |

## Installation

Install via Claude Code:

```
/install-plugin https://github.com/kenton-web/ppt-template-builder
```

## Prerequisites

```bash
pip install python-pptx pillow lxml
```

## Usage

Just mention your template file and what you want:

> "Use my template at ~/Downloads/MyBrand.pptx to build a performance deck for Acme Co. Pull last 30 days of advertising data from Intentwise — I want a KPI scorecard, weekly spend chart, top campaigns table, and insights."

Or with CSV data:

> "Build a deck using MyTemplate.pptx. Here's my data: [paste CSV]. Add a cover slide and an insights slide."

## How it works

The skill uses `build_deck.py` — a self-contained Python library that:

1. **Inspects** your template to extract chrome images, fonts, colors, and slide dimensions
2. **Strips** the template slides cleanly (avoiding the common orphan-image bug)
3. **Rebuilds** each slide type using the extracted chrome, so new slides inherit your exact brand identity

```bash
# Inspect any template
python3 build_deck.py inspect /path/to/template.pptx
```

```python
from build_deck import DeckBuilder

builder = DeckBuilder(template_path="MyTemplate.pptx", output_path="output.pptx")
builder.add_cover(title="Q1 Review", client_name="Acme Co", period="Q1 2026", ...)
builder.add_kpi(slide_title="Ad Summary", metrics=[...])
builder.add_chart(slide_title="Weekly Trend", chart_data={...})
builder.add_table(slide_title="Top Campaigns", column_headers=[...], rows=[...])
builder.add_insights(slide_title="Key Insights", bullets=[...])
path = builder.save()
```
