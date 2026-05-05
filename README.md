# Why This Skill Exists

This skill was made for professional PowerPoint work where the goal is to match an existing corporate style, not invent a new visual system.

Many AI tools create slides by drawing new shapes, choosing new colors, or exporting charts as images. That can produce a deck that looks plausible in isolation but fails in a corporate setting because it ignores the company's slide masters, layouts, typography, spacing, chart styles, and brand rules.

`brand-faithful-pptx` takes the opposite approach. It starts from an existing `.pptx` template or branded deck, uses the slide masters and layouts already inside that file, places content into placeholders where possible, and keeps charts as native PowerPoint charts backed by embedded Excel workbook data.

The result should be a deck that feels like it belongs to the organization that owns the template. A future editor should be able to open the file in PowerPoint, reuse normal slide layouts, and right-click a chart to edit its data in Excel instead of being stuck with a static SVG, PNG, or screenshot.

## Attribution

This skill is built around the Python `python-pptx` library, created and maintained by Steve Canny and contributors. Project source: https://github.com/scanny/python-pptx
