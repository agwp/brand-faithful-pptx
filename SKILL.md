---
name: brand-faithful-pptx
description: Create, edit, repair, and automate professional PowerPoint .pptx decks with python-pptx while preserving corporate branding, slide masters, template slide layouts, placeholders, theme fonts, theme colors, and editable native PowerPoint charts backed by embedded Excel workbooks. Use when a task mentions corporate PowerPoint templates, brand-compliant decks, slide masters, existing branding, template-compliant layouts, native/editable charts, Excel-editable chart data, replacing chart data, or avoiding image/SVG chart output.
---

# Brand Faithful PPTX

## Core Rule

Build from the user's PowerPoint template (check `/references` folder) or existing branded deck whenever one is available or provided by the user:

```python
from pptx import Presentation

prs = Presentation(template_path)
```

Do not recreate the visual system from scratch. Use the template's slide masters, slide layouts, placeholders, theme fonts, theme colors, and existing chart styling. Only use `Presentation()` with no path when the user has not provided a template and no existing deck is being edited.

## Workflow

1. Inspect the deck before editing.
   - Run `scripts/inspect_template.py TEMPLATE.pptx --format markdown` to list layouts and placeholder `idx` values.
   - Choose layouts by name/placeholder structure, not by hard-coded index unless the template audit proves the index.

2. Add content through placeholders first.
   - Set titles via `slide.shapes.title.text` when present.
   - Access known placeholders with `slide.placeholders[idx]`; `idx` is stable for that layout.
   - Prefer chart placeholders for charts. `ChartPlaceholder.insert_chart()` preserves the placeholder's position and size.

3. Use native PowerPoint charts.
   - Create charts with `ChartPlaceholder.insert_chart()` or `slide.shapes.add_chart()`.
   - Use `CategoryChartData` for category charts, `XyChartData` for scatter charts, and `BubbleChartData` for bubble charts.
   - Use `Chart.replace_data(chart_data)` to update existing editable charts.
   - Do not render charts as SVG, PNG, screenshots, canvas, or pasted pictures unless the user explicitly asks for non-editable artwork.

4. Verify the output.
   - Run `scripts/verify_native_charts.py OUTPUT.pptx` for decks that contain charts.
   - Inspect/render the final deck when visual fidelity matters.

## References

- Read `references/python-pptx-patterns.md` for template, layout, placeholder, and chart-placement patterns.
- Read `references/native-chart-recipes.md` for copyable native-chart examples.
- Read `references/template-audit.md` before mapping a new corporate template.

## Script Quick Start

```bash
python scripts/inspect_template.py template.pptx --format markdown
python scripts/example_native_chart_deck.py --template template.pptx --output sample-native-chart.pptx
python scripts/verify_native_charts.py sample-native-chart.pptx
```

When a template has no chart placeholder, use the closest content layout and explicit coordinates from the audited placeholder geometry. Keep those coordinates derived from the template, not invented as a new design system.
