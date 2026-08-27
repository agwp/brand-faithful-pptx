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

The plain fallback catalog at `references/plain-fallback-100-layouts.pptx` is opt-in only. Do not use, merge, or adapt its layouts unless the user explicitly approves fallback layout use after the official template has been audited. Official user-provided masters and layouts always take precedence.

## Content Discipline

A placeholder is an available destination, not a content requirement. Placeholder presence does not authorize new content. Never invent text merely to populate a placeholder, balance a slide, or exercise a layout. If the source material does not justify an optional text element, leave it empty, remove only the slide-level instance when that is safe, or select a better approved layout. Do not alter the slide master or layout to hide the problem.

Use this priority order:

1. Source fidelity.
2. Comprehension.
3. Visual hierarchy.
4. Decoration.

Decoration should normally be absent. Content may come from one of these origins:

- `source`: copied directly from user-supplied or approved material.
- `transformed_source`: a faithful shortening, restructuring, or summary when editing is authorized.
- `functional_label`: a necessary date, unit, source note, axis label, table heading, or section name derived directly from the source.
- `generated_with_user_authorization`: new copy the user explicitly asked the agent to write.

Unless supplied, functionally required, or explicitly requested, do not add:

- Eyebrow headings, kickers, overlines, or small contextual headers above a title.
- Taglines, straplines, slogans, campaign language, idioms, metaphors, inspirational statements, or rhetorical flourishes.
- Generic subtitles or framing copy that repeats the title.
- Image captions that do not identify, explain, credit, or source the image.
- Takeaway panels, callouts, or conclusions unsupported by the supplied evidence.
- New running headers, footers, confidentiality labels, section markers, or administrative copy.
- Any text written only because the selected layout contains a placeholder for it.

Preserve fixed logos, rules, slide numbers, and approved text already embedded in a master or layout. Do not treat preserving fixed brand furniture as permission to add more header or footer copy.

Slide titles must be literal, concise, and specific to the slide's actual content. Use a conclusion-led title only when the conclusion is supported by the source. Captions are optional and must be limited to factual identification, necessary explanatory context, evidence notes, or source and credit information.

## Workflow

1. Plan the supplied content before selecting layouts.
   - Track each planned text element with one of the content origins above and, where useful, a source reference.
   - Select the layout whose primary regions match the content that actually exists. Do not choose a layout that requires invented supporting copy.
   - Treat subtitle, caption, quote, attribution, section-label, and callout placeholders as optional unless the audited template explicitly records otherwise.

2. Inspect the deck before editing.
   - Run `scripts/inspect_template.py TEMPLATE.pptx --format markdown` to list layouts and placeholder `idx` values.
   - Choose layouts by name/placeholder structure, not by hard-coded index unless the template audit proves the index.

### Binary-hostile agent fallback

Some agents, chat clients, and review tools corrupt or ignore binary `.pptx` attachments. A `.pptx` is a ZIP archive, so when binary handling is unreliable, ask the user to provide either the original `.pptx` or an unpacked package directory/ZIP made from it. Use the helper script to round-trip safely:

```bash
python scripts/pptx_package.py unpack template.pptx references/template-package --force
python scripts/pptx_package.py pack references/template-package rebuilt-template.pptx --force
```

When working from an unpacked package, treat XML and relationship files as source material for inspection or repair, then repack before running `python-pptx` tools. Never copy `.pptx` bytes through text channels, JSON strings, or markdown fences; use filesystem paths, attachments, or the unpacked package workflow instead.

3. For content that exists, use an appropriate placeholder first.
   - Set titles via `slide.shapes.title.text` when present.
   - Access known placeholders with `slide.placeholders[idx]`; `idx` is stable for that layout.
   - Prefer chart placeholders for charts. `ChartPlaceholder.insert_chart()` preserves the placeholder's position and size.
   - Leave an unused optional placeholder empty. Do not add a freeform text box to compensate for an unsuitable layout when another approved layout fits.

4. Use native PowerPoint charts.
   - Create charts with `ChartPlaceholder.insert_chart()` or `slide.shapes.add_chart()`.
   - Use `CategoryChartData` for category charts, `XyChartData` for scatter charts, and `BubbleChartData` for bubble charts.
   - Use `Chart.replace_data(chart_data)` to update existing editable charts.
   - Do not render charts as SVG, PNG, screenshots, canvas, or pasted pictures unless the user explicitly asks for non-editable artwork.

5. Verify the output.
   - Run `scripts/verify_native_charts.py OUTPUT.pptx` for decks that contain charts.
   - Inspect/render the final deck when visual fidelity matters.
   - Do not treat an empty optional placeholder as a QA failure.
   - Audit every text element not copied directly from a source. Confirm it is an authorized faithful transformation, a necessary functional label, or explicitly authorized new copy. Delete text that is redundant, vague, unsupported, decorative, or present only to fill space.
   - Inspect freeform text boxes as well as placeholders; placeholder auditing alone cannot detect unsupported additions.

## References

- Read `references/python-pptx-patterns.md` for template, layout, placeholder, and chart-placement patterns.
- Read `references/native-chart-recipes.md` for copyable native-chart examples.
- Read `references/template-audit.md` before mapping a new corporate template.

## Script Quick Start

```bash
python scripts/inspect_template.py template.pptx --format markdown
python scripts/pptx_package.py unpack template.pptx references/template-package --force
python scripts/pptx_package.py pack references/template-package rebuilt-template.pptx --force
python scripts/example_native_chart_deck.py --template template.pptx --output sample-native-chart.pptx
python scripts/verify_native_charts.py sample-native-chart.pptx
```

Build and verify the repository's opt-in fallback catalog with:

```bash
python scripts/build_plain_fallback_deck.py
python scripts/verify_plain_fallback_deck.py references/plain-fallback-100-layouts.pptx
```

When a template has no chart placeholder, use the closest content layout and explicit coordinates from the audited placeholder geometry. Keep those coordinates derived from the template, not invented as a new design system.

These content-discipline rules also apply to any user-approved fallback layouts. More available layouts or placeholders never broaden the supplied content and never authorize decorative copy.
