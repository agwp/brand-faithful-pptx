# Template Audit

Use this reference when a template is unfamiliar or the requested slide type is ambiguous.

## Audit steps

1. Run:

   ```bash
   python scripts/inspect_template.py template.pptx --format markdown
   ```

2. Identify layouts with:
   - A title placeholder.
   - A chart placeholder for chart-heavy slides.
   - Content placeholders with useful geometry when no chart placeholder exists.
   - Section divider, agenda, comparison, and appendix layouts if the task needs them.

3. Record the selected layout names and placeholder `idx` values in your working notes before generating slides.

## Mapping content to layouts

Use the most specific layout that matches the requested content:

- Title slide: title/subtitle/date/client placeholders.
- Executive summary: title plus body/content placeholders.
- Chart slide: title plus chart placeholder, or title plus content placeholder geometry.
- Comparison: layouts with two content placeholders.
- Table-heavy slide: table placeholder or a wide content placeholder.

## Fallback when no chart placeholder exists

Use audited placeholder geometry:

```python
content = slide.placeholders[1]
graphic_frame = slide.shapes.add_chart(
    chart_type,
    content.left,
    content.top,
    content.width,
    content.height,
    chart_data,
)
chart = graphic_frame.chart
```

This still produces a native chart. The fallback only gives up placeholder insertion, not chart editability.

## Quality checks

- Confirm every slide was created from a template layout.
- Confirm titles are in title placeholders where possible.
- Confirm charts are graphic frames with `has_chart == True`.
- Confirm `verify_native_charts.py` reports embedded workbook relationships.
- Confirm no chart request was satisfied with a picture unless the user asked for a static image.
