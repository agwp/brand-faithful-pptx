#!/usr/bin/env python3
"""Inspect PowerPoint template layouts and placeholders."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from pptx import Presentation


EMU_PER_INCH = 914400


def emu_to_inches(value: int | None) -> float | None:
    if value is None:
        return None
    return round(value / EMU_PER_INCH, 3)


def placeholder_info(shape: Any) -> dict[str, Any]:
    fmt = shape.placeholder_format
    return {
        "idx": fmt.idx,
        "type": str(fmt.type),
        "name": shape.name,
        "left_in": emu_to_inches(shape.left),
        "top_in": emu_to_inches(shape.top),
        "width_in": emu_to_inches(shape.width),
        "height_in": emu_to_inches(shape.height),
    }


def inspect(path: Path) -> dict[str, Any]:
    prs = Presentation(str(path))
    layouts = []
    for index, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for shape in layout.placeholders:
            placeholders.append(placeholder_info(shape))
        layouts.append(
            {
                "index": index,
                "name": layout.name,
                "placeholder_count": len(placeholders),
                "placeholders": placeholders,
            }
        )
    return {
        "path": str(path),
        "slide_width_in": emu_to_inches(prs.slide_width),
        "slide_height_in": emu_to_inches(prs.slide_height),
        "layout_count": len(layouts),
        "layouts": layouts,
    }


def to_markdown(report: dict[str, Any]) -> str:
    lines = [
        f"# Template Audit: {report['path']}",
        "",
        f"- Size: {report['slide_width_in']} x {report['slide_height_in']} in",
        f"- Layouts: {report['layout_count']}",
        "",
    ]
    for layout in report["layouts"]:
        lines.append(f"## {layout['index']}: {layout['name']}")
        if not layout["placeholders"]:
            lines.append("")
            lines.append("_No placeholders._")
            lines.append("")
            continue
        lines.append("")
        lines.append("| idx | type | name | left | top | width | height |")
        lines.append("| ---: | --- | --- | ---: | ---: | ---: | ---: |")
        for ph in layout["placeholders"]:
            lines.append(
                "| {idx} | {type} | {name} | {left_in} | {top_in} | {width_in} | {height_in} |".format(
                    **ph
                )
            )
        lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path, help="Path to .pptx template/deck")
    parser.add_argument(
        "--format",
        choices=("json", "markdown"),
        default="json",
        help="Output format",
    )
    args = parser.parse_args()

    report = inspect(args.template)
    if args.format == "markdown":
        print(to_markdown(report))
    else:
        print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
