#!/usr/bin/env python3
"""Verify a .pptx contains native PowerPoint charts with embedded workbooks."""

from __future__ import annotations

import argparse
import json
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
CHART_NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
PACKAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"


def rels_path_for(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", filename + ".rels")


def normalize_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def inspect_pptx(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        chart_parts = sorted(name for name in names if name.startswith("ppt/charts/chart") and name.endswith(".xml"))
        media_parts = sorted(name for name in names if name.startswith("ppt/media/"))
        charts = []

        for chart_part in chart_parts:
            chart_xml = ET.fromstring(zf.read(chart_part))
            has_external_data = chart_xml.find(".//c:externalData", CHART_NS) is not None
            rels_name = rels_path_for(chart_part)
            workbook_targets = []

            if rels_name in names:
                rels_xml = ET.fromstring(zf.read(rels_name))
                for rel in rels_xml.findall("rel:Relationship", REL_NS):
                    if rel.attrib.get("Type") == PACKAGE_REL_TYPE:
                        target = rel.attrib.get("Target", "")
                        workbook_targets.append(normalize_target(chart_part, target))

            charts.append(
                {
                    "chart_part": chart_part,
                    "has_external_data": has_external_data,
                    "embedded_workbooks": workbook_targets,
                    "workbooks_exist": all(target in names for target in workbook_targets),
                }
            )

        missing_workbooks = [
            chart
            for chart in charts
            if not chart["has_external_data"]
            or not chart["embedded_workbooks"]
            or not chart["workbooks_exist"]
        ]

        return {
            "path": str(path),
            "chart_count": len(charts),
            "charts": charts,
            "media_part_count": len(media_parts),
            "media_parts": media_parts,
            "ok": bool(charts) and not missing_workbooks,
        }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="Path to .pptx to verify")
    parser.add_argument("--warn-only", action="store_true", help="Always exit 0")
    args = parser.parse_args()

    report = inspect_pptx(args.pptx)
    print(json.dumps(report, indent=2))
    return 0 if report["ok"] or args.warn_only else 1


if __name__ == "__main__":
    raise SystemExit(main())
