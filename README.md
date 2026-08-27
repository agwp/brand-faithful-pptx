# Why This Skill Exists

This skill was made for professional PowerPoint work where the goal is to match an existing corporate style, not invent a new visual system.

Many AI tools create slides by drawing new shapes, choosing new colors, or exporting charts as images. That can produce a deck that looks plausible in isolation but fails in a corporate setting because it ignores the company's slide masters, layouts, typography, spacing, chart styles, and brand rules.

`brand-faithful-pptx` takes the opposite approach. It starts from an existing `.pptx` template or branded deck, uses the slide masters and layouts already inside that file, places content into placeholders where possible, and keeps charts as native PowerPoint charts backed by embedded Excel workbook data.

The result should be a deck that feels like it belongs to the organization that owns the template. A future editor should be able to open the file in PowerPoint, reuse normal slide layouts, and right-click a chart to edit its data in Excel instead of being stuck with a static SVG, PNG, or screenshot.

## Content before layout

The skill does not treat every placeholder as a slot that must be filled. It selects a layout for the title, text, data, and images that actually exist, then leaves unused optional placeholders empty. It does not invent kickers, straplines, generic subtitles, unsupported takeaways, or unnecessary image captions to make a layout look complete.

Fixed master and layout elements such as logos, footer rules, slide numbers, and approved labels remain intact. New header, footer, confidentiality, caption, or callout copy is added only when supplied, directly supported, functionally required, or explicitly requested.

User-approved fallback layouts follow the same rules. Having more layouts available does not broaden the source material, and chart layouts remain plain: chart furniture and commentary are included only when they make the supplied data easier to understand.

## Opt-in fallback layout catalog

[`references/plain-fallback-100-layouts.pptx`](references/plain-fallback-100-layouts.pptx) contains one genuine PowerPoint slide master, 100 named reusable slide layouts, and 100 example slides—one example for every layout requested in issue #3. It is a plain starting point for cases where an audited corporate template has no suitable official layout. It must not be used, merged, or restyled without explicit user approval; official template layouts remain the default.

The chart examples use restrained native PowerPoint charts with embedded Excel workbooks where `python-pptx` supports the chart type. They omit chart titles, axis titles, gridlines, and redundant legends. Treemap, heatmap, and geographic examples use editable PowerPoint shapes or tables because the corresponding Office extended charts cannot be created safely by `python-pptx`.

Rebuild and verify the catalog with:

```bash
python3 scripts/build_plain_fallback_deck.py
python3 scripts/verify_plain_fallback_deck.py references/plain-fallback-100-layouts.pptx
python3 scripts/verify_native_charts.py references/plain-fallback-100-layouts.pptx
python3 scripts/inspect_template.py references/plain-fallback-100-layouts.pptx \
  --format markdown \
  --role-overrides references/plain-fallback-placeholder-roles.json
```

## Recommended set up for local use

1. Download this repo to use locally.
2. Add your approved PowerPoint deck into the `/references` folder.
3. Direct your local AI app to this folder to install the skill with your PowerPoint template.

## Attribution

This skill is built around the Python `python-pptx` library, created and maintained by Steve Canny and contributors. Project source: https://github.com/scanny/python-pptx
