# python-pptx Patterns

## Source-of-truth hierarchy

1. Existing user deck when editing.
2. User-provided `.pptx` template or example deck when creating.
3. Built-in `Presentation()` default only when no template exists.

`python-pptx` treats a presentation as a package of parts. Most design information lives in the theme, slide masters, and slide layouts that remain even when a deck has no slides. Starting from `Presentation(template_path)` preserves those parts.

## Layout selection

Audit the template before adding slides:

```python
for i, layout in enumerate(prs.slide_layouts):
    print(i, layout.name)
    for ph in layout.placeholders:
        pf = ph.placeholder_format
        print(" ", pf.idx, pf.type, ph.name, ph.left, ph.top, ph.width, ph.height)
```

Use layout names and placeholder structures to pick slides. Avoid assuming layout `1` is always "Title and Content"; corporate templates vary.

## Placeholder use

Prefer placeholders over new shapes because they inherit the template's position, size, and formatting.

```python
slide = prs.slides.add_slide(layout)

if slide.shapes.title:
    slide.shapes.title.text = "Quarterly performance"

body = slide.placeholders[2]
body.text = "Key message"
```

Placeholder lookup is by `idx`, not sequence position. The title placeholder is commonly `idx == 0`; user-created layout placeholders often start at `idx == 10`.

## Native chart insertion

Charts are not shapes. `add_chart()` and `insert_chart()` return a graphic frame; access the editable chart via `.chart`.

```python
graphic_frame = slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data)
chart = graphic_frame.chart
```

If the layout has a chart placeholder:

```python
graphic_frame = slide.placeholders[10].insert_chart(chart_type, chart_data)
chart = graphic_frame.chart
```

After rich-content insertion, the original placeholder reference is invalid; keep the returned graphic frame.

## Existing chart updates

Use `chart.replace_data(chart_data)` to update both the chart XML and the embedded workbook. This keeps the chart editable through PowerPoint's "Edit Data" workflow.

```python
for shape in slide.shapes:
    if getattr(shape, "has_chart", False):
        shape.chart.replace_data(chart_data)
```

## Styling posture

Use the template's chart style unless the task requires a specific style. Light-touch chart properties such as legend position, data labels, axis titles, and number formats are appropriate; rebuilding colors by hand is usually a sign that the template is not being used well.
