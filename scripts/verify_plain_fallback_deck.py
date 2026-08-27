#!/usr/bin/env python3
"""Verify the generated 100-layout plain fallback PowerPoint deck.

This verifier checks both the reusable layout catalog and its 100 example
slides.  It deliberately inspects the OOXML package rather than relying only
on python-pptx loading, because layout/master relationships and embedded chart
workbooks are package-level contracts.
"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable

from lxml import etree
from pptx import Presentation

try:  # Direct script execution places this directory first on sys.path.
    from build_plain_fallback_deck import LAYOUT_ROWS
except ModuleNotFoundError:  # Importing as scripts.verify_plain_fallback_deck.
    from scripts.build_plain_fallback_deck import LAYOUT_ROWS


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {"p": P_NS, "a": A_NS, "c": C_NS, "r": R_NS}
REL = {"rel": REL_NS}
CT = {"ct": CT_NS}

SLIDE_MASTER_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
SLIDE_LAYOUT_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
PACKAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"

SLIDE_LAYOUT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
)
CHART_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)
WORKBOOK_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

EXPECTED_LAYOUT_NAMES = tuple(name for _, name, _ in LAYOUT_ROWS)
NATIVE_CHART_LAYOUTS = frozenset(range(49, 68)) | {72}
EDITABLE_DATA_LAYOUTS = frozenset({68, 69, 70, 71})

# Expected plot structure produced by build_plain_fallback_deck.py.  The first
# value is the plot tag and the optional following values are barDir/grouping.
EXPECTED_PLOTS: dict[int, tuple[tuple[str, str | None, str | None], ...]] = {
    49: (("barChart", "col", "clustered"),),
    50: (("barChart", "col", "clustered"),),
    51: (("barChart", "col", "clustered"),),
    52: (("barChart", "col", "stacked"),),
    53: (("barChart", "col", "percentStacked"),),
    54: (("barChart", "bar", "clustered"),),
    55: (("barChart", "bar", "clustered"),),
    56: (("barChart", "bar", "stacked"),),
    57: (("barChart", "bar", "percentStacked"),),
    58: (("barChart", "bar", "clustered"),),
    59: (("pieChart", None, None),),
    60: (("doughnutChart", None, None),),
    61: (("lineChart", None, "standard"),),
    62: (("areaChart", None, "stacked"),),
    63: (
        ("barChart", "col", "clustered"),
        ("lineChart", None, "standard"),
    ),
    64: (("barChart", "col", "stacked"),),
    65: (("scatterChart", None, None),),
    66: (("bubbleChart", None, None),),
    67: (("barChart", "bar", "clustered"),),
    72: (("barChart", "col", "clustered"),),
}

PROHIBITED_COPY_PATTERNS = (
    ("where-insight-meets-impact", re.compile(r"\bwhere\s+insight\s+meets\s+impact\b", re.I)),
    ("smarter-path-forward", re.compile(r"\ba\s+smarter\s+path\s+forward\b", re.I)),
    ("vision-to-value", re.compile(r"\bfrom\s+vision\s+to\s+value\b", re.I)),
    (
        "unlocking-generic-value",
        re.compile(r"\bunlock(?:ing)?\s+(?:the\s+)?(?:power|potential|value)\b", re.I),
    ),
    (
        "where-x-meets-y",
        re.compile(r"\bwhere\s+[a-z][a-z -]{1,30}\s+meets\s+[a-z][a-z -]{1,30}\b", re.I),
    ),
)


def expected_layout_names() -> tuple[str, ...]:
    """Return the exact authoritative layout names in catalog order."""

    return EXPECTED_LAYOUT_NAMES


def _rels_name(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", filename + ".rels")


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _is_physical_part(name: str, directory: str, stem: str) -> bool:
    return bool(re.fullmatch(rf"{re.escape(directory)}/{re.escape(stem)}\d+\.xml", name))


def _relationship_map(root: etree._Element | None) -> dict[str, etree._Element]:
    if root is None:
        return {}
    return {rel.get("Id", ""): rel for rel in root.findall("rel:Relationship", REL)}


def _placeholder_type(ph: etree._Element) -> str:
    # The OOXML default placeholder type is obj when @type is absent.
    return ph.get("type", "obj")


def _part_content_type(
    part_name: str,
    overrides: dict[str, str],
    defaults: dict[str, str],
) -> str | None:
    if part_name in overrides:
        return overrides[part_name]
    suffix = posixpath.basename(part_name).rsplit(".", 1)
    extension = suffix[1].lower() if len(suffix) == 2 else ""
    return defaults.get(extension)


def _parse_xml(
    parts: dict[str, bytes],
    part_name: str,
    errors: list[str],
) -> etree._Element | None:
    payload = parts.get(part_name)
    if payload is None:
        errors.append(f"Missing required part: {part_name}")
        return None
    try:
        return etree.fromstring(payload)
    except (etree.XMLSyntaxError, ValueError) as exc:
        errors.append(f"Invalid XML in {part_name}: {exc}")
        return None


def _relationship_target(
    source_part: str,
    rel: etree._Element,
    parts: dict[str, bytes],
    errors: list[str],
) -> str | None:
    target = rel.get("Target")
    if not target:
        errors.append(f"Relationship {rel.get('Id')} in {source_part} has no target")
        return None
    if rel.get("TargetMode") == "External":
        errors.append(
            f"Relationship {rel.get('Id')} in {source_part} unexpectedly has an external target"
        )
        return None
    resolved = _resolve_target(source_part, target)
    if resolved not in parts:
        errors.append(
            f"Relationship {rel.get('Id')} in {source_part} resolves to missing part {resolved}"
        )
    return resolved


def _check_plot_contract(
    slide_number: int,
    chart_part: str,
    chart_root: etree._Element,
    errors: list[str],
) -> None:
    expected = EXPECTED_PLOTS[slide_number]
    supported_tags = {
        "barChart",
        "pieChart",
        "doughnutChart",
        "lineChart",
        "areaChart",
        "scatterChart",
        "bubbleChart",
    }
    actual: list[tuple[str, str | None, str | None]] = []
    for tag in supported_tags:
        for plot in chart_root.findall(f".//c:{tag}", NS):
            bar_dir = plot.find("c:barDir", NS)
            grouping = plot.find("c:grouping", NS)
            actual.append(
                (
                    tag,
                    bar_dir.get("val") if bar_dir is not None else None,
                    grouping.get("val") if grouping is not None else None,
                )
            )
    if sorted(actual) != sorted(expected):
        errors.append(
            f"Slide {slide_number} chart {chart_part} has plot contract {actual!r}; "
            f"expected {list(expected)!r}"
        )


def _check_chart_part(
    slide_number: int,
    chart_part: str,
    parts: dict[str, bytes],
    overrides: dict[str, str],
    defaults: dict[str, str],
    errors: list[str],
) -> str | None:
    chart_root = _parse_xml(parts, chart_part, errors)
    if chart_root is None:
        return None

    content_type = _part_content_type(chart_part, overrides, defaults)
    if content_type != CHART_CONTENT_TYPE:
        errors.append(
            f"Chart part {chart_part} has content type {content_type!r}, "
            f"expected {CHART_CONTENT_TYPE!r}"
        )

    if chart_root.findall(".//c:title", NS):
        errors.append(f"Slide {slide_number} chart {chart_part} contains a chart or axis title")
    if chart_root.findall(".//c:majorGridlines", NS):
        errors.append(f"Slide {slide_number} chart {chart_part} contains major gridlines")
    if chart_root.findall(".//c:minorGridlines", NS):
        errors.append(f"Slide {slide_number} chart {chart_part} contains minor gridlines")
    if chart_root.findall(".//a:blip", NS):
        errors.append(f"Slide {slide_number} chart {chart_part} contains an embedded image")

    _check_plot_contract(slide_number, chart_part, chart_root, errors)

    external_data = chart_root.findall(".//c:externalData", NS)
    if len(external_data) != 1:
        errors.append(
            f"Slide {slide_number} chart {chart_part} has {len(external_data)} externalData "
            "elements; expected one embedded workbook"
        )
        return None

    external_rid = external_data[0].get(f"{{{R_NS}}}id")
    chart_rels_name = _rels_name(chart_part)
    chart_rels_root = _parse_xml(parts, chart_rels_name, errors)
    chart_rels = _relationship_map(chart_rels_root)
    workbook_rel = chart_rels.get(external_rid or "")
    if workbook_rel is None:
        errors.append(
            f"Slide {slide_number} chart {chart_part} externalData relationship "
            f"{external_rid!r} is missing"
        )
        return None
    if workbook_rel.get("Type") != PACKAGE_REL:
        errors.append(
            f"Slide {slide_number} chart {chart_part} externalData relationship is "
            f"{workbook_rel.get('Type')!r}, expected embedded package"
        )
    workbook_part = _relationship_target(chart_part, workbook_rel, parts, errors)
    if workbook_part is None:
        return None
    if not workbook_part.lower().endswith(".xlsx"):
        errors.append(
            f"Slide {slide_number} chart {chart_part} workbook target is not .xlsx: "
            f"{workbook_part}"
        )
    workbook_type = _part_content_type(workbook_part, overrides, defaults)
    if workbook_type != WORKBOOK_CONTENT_TYPE:
        errors.append(
            f"Embedded workbook {workbook_part} has content type {workbook_type!r}, "
            f"expected {WORKBOOK_CONTENT_TYPE!r}"
        )
    return workbook_part


def _check_prohibited_copy(
    slide_number: int,
    slide_root: etree._Element,
    errors: list[str],
) -> None:
    for text_element in slide_root.findall(".//a:t", NS):
        text = " ".join((text_element.text or "").split())
        if not text:
            continue
        for label, pattern in PROHIBITED_COPY_PATTERNS:
            if pattern.search(text):
                errors.append(
                    f"Slide {slide_number} contains prohibited ornamental copy "
                    f"({label}): {text!r}"
                )


def inspect_deck(path: Path) -> dict[str, Any]:
    """Return a machine-readable acceptance report for *path*."""

    path = Path(path)
    errors: list[str] = []
    warnings: list[str] = []
    report: dict[str, Any] = {
        "path": str(path),
        "master_count": 0,
        "layout_count": 0,
        "slide_count": 0,
        "native_chart_slide_count": 0,
        "native_chart_part_count": 0,
        "embedded_workbook_count": 0,
        "editable_shape_data_slide_count": 0,
        "errors": errors,
        "warnings": warnings,
        "ok": False,
    }

    if not path.is_file():
        errors.append(f"Deck not found: {path}")
        return report

    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if len(names) != len(set(names)):
                errors.append("PPTX contains duplicate ZIP member names")
            corrupt = zf.testzip()
            if corrupt:
                errors.append(f"PPTX ZIP CRC failure in {corrupt}")
            parts = {name: zf.read(name) for name in names}
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        errors.append(f"Cannot read PPTX package: {exc}")
        return report

    content_types = _parse_xml(parts, "[Content_Types].xml", errors)
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    if content_types is not None:
        overrides = {
            item.get("PartName", "").lstrip("/"): item.get("ContentType", "")
            for item in content_types.findall("ct:Override", CT)
        }
        defaults = {
            item.get("Extension", "").lower(): item.get("ContentType", "")
            for item in content_types.findall("ct:Default", CT)
        }

    physical_masters = sorted(
        name
        for name in parts
        if _is_physical_part(name, "ppt/slideMasters", "slideMaster")
    )
    physical_layouts = sorted(
        name
        for name in parts
        if _is_physical_part(name, "ppt/slideLayouts", "slideLayout")
    )
    physical_slides = sorted(
        name for name in parts if _is_physical_part(name, "ppt/slides", "slide")
    )
    report["master_count"] = len(physical_masters)
    if len(physical_masters) != 1:
        errors.append(f"Expected exactly one physical slide master, found {len(physical_masters)}")
    if len(physical_layouts) != 100:
        errors.append(f"Expected exactly 100 physical slide layouts, found {len(physical_layouts)}")
    if len(physical_slides) != 100:
        errors.append(f"Expected exactly 100 physical slides, found {len(physical_slides)}")

    presentation_part = "ppt/presentation.xml"
    presentation = _parse_xml(parts, presentation_part, errors)
    presentation_rels = _parse_xml(parts, _rels_name(presentation_part), errors)
    presentation_rel_map = _relationship_map(presentation_rels)
    if presentation is None:
        return report

    master_ids = presentation.findall("p:sldMasterIdLst/p:sldMasterId", NS)
    if len(master_ids) != 1:
        errors.append(f"Expected presentation to reference one slide master, found {len(master_ids)}")
    master_part: str | None = None
    if master_ids:
        master_rid = master_ids[0].get(f"{{{R_NS}}}id")
        master_rel = presentation_rel_map.get(master_rid or "")
        if master_rel is None:
            errors.append(f"Presentation slide-master relationship {master_rid!r} is missing")
        else:
            if master_rel.get("Type") != SLIDE_MASTER_REL:
                errors.append(
                    f"Presentation relationship {master_rid!r} is not a slide-master relationship"
                )
            master_part = _relationship_target(presentation_part, master_rel, parts, errors)
    if master_part and physical_masters and master_part not in physical_masters:
        errors.append(f"Referenced master {master_part} is not the sole physical master")

    layout_parts: list[str] = []
    if master_part:
        master = _parse_xml(parts, master_part, errors)
        master_rels = _parse_xml(parts, _rels_name(master_part), errors)
        master_rel_map = _relationship_map(master_rels)
        layout_ids = master.findall("p:sldLayoutIdLst/p:sldLayoutId", NS) if master is not None else []
        report["layout_count"] = len(layout_ids)
        if len(layout_ids) != 100:
            errors.append(f"Expected master to reference 100 layouts, found {len(layout_ids)}")

        seen_layout_ids: set[str] = set()
        for index, layout_id in enumerate(layout_ids, start=1):
            numeric_id = layout_id.get("id", "")
            if numeric_id in seen_layout_ids:
                errors.append(f"Duplicate master layout ID {numeric_id!r}")
            seen_layout_ids.add(numeric_id)
            rid = layout_id.get(f"{{{R_NS}}}id")
            rel = master_rel_map.get(rid or "")
            if rel is None:
                errors.append(f"Master layout {index} relationship {rid!r} is missing")
                continue
            if rel.get("Type") != SLIDE_LAYOUT_REL:
                errors.append(f"Master relationship {rid!r} for layout {index} has wrong type")
            layout_part = _relationship_target(master_part, rel, parts, errors)
            if layout_part is None:
                continue
            layout_parts.append(layout_part)
            if _part_content_type(layout_part, overrides, defaults) != SLIDE_LAYOUT_CONTENT_TYPE:
                errors.append(
                    f"Layout {index} part {layout_part} lacks the slide-layout content type"
                )

            layout_root = _parse_xml(parts, layout_part, errors)
            if layout_root is None:
                continue
            expected_name = EXPECTED_LAYOUT_NAMES[index - 1] if index <= 100 else None
            c_sld = layout_root.find("p:cSld", NS)
            actual_name = c_sld.get("name") if c_sld is not None else None
            if actual_name != expected_name:
                errors.append(
                    f"Layout {index} is named {actual_name!r}; expected {expected_name!r}"
                )
            if layout_root.get("matchingName") != expected_name:
                errors.append(
                    f"Layout {index} matchingName is {layout_root.get('matchingName')!r}; "
                    f"expected {expected_name!r}"
                )

            placeholders = layout_root.findall(".//p:ph", NS)
            placeholder_indices = [ph.get("idx", "0") for ph in placeholders]
            if len(placeholder_indices) != len(set(placeholder_indices)):
                errors.append(f"Layout {index} contains duplicate placeholder indices")
            chart_placeholders = [ph for ph in placeholders if _placeholder_type(ph) == "chart"]
            if index in NATIVE_CHART_LAYOUTS:
                if len(chart_placeholders) != 1:
                    errors.append(
                        f"Layout {index} must contain exactly one chart placeholder; "
                        f"found {len(chart_placeholders)}"
                    )
            elif chart_placeholders:
                errors.append(
                    f"Non-native-chart layout {index} unexpectedly contains a chart placeholder"
                )
            if index in (68, 69, 70):
                object_placeholders = [
                    ph
                    for ph in placeholders
                    if _placeholder_type(ph) == "obj" and ph.get("idx") == "10"
                ]
                if len(object_placeholders) != 1:
                    errors.append(
                        f"Layout {index} must contain its editable object placeholder at idx 10"
                    )
            if index == 71:
                kpi_indices = {
                    ph.get("idx")
                    for ph in placeholders
                    if _placeholder_type(ph) == "body"
                }
                if not {"10", "11", "12", "13"}.issubset(kpi_indices):
                    errors.append("Layout 71 must contain four editable KPI body placeholders")
            if index == 100 and placeholders:
                errors.append(
                    f"Layout 100 must be a blank branded canvas with no placeholders; "
                    f"found {len(placeholders)}"
                )

            layout_rels = _parse_xml(parts, _rels_name(layout_part), errors)
            back_rels = (
                [
                    rel_item
                    for rel_item in layout_rels.findall("rel:Relationship", REL)
                    if rel_item.get("Type") == SLIDE_MASTER_REL
                ]
                if layout_rels is not None
                else []
            )
            if len(back_rels) != 1:
                errors.append(
                    f"Layout {index} must have exactly one back-relationship to the master; "
                    f"found {len(back_rels)}"
                )
            elif _relationship_target(layout_part, back_rels[0], parts, errors) != master_part:
                errors.append(f"Layout {index} back-relationship does not target {master_part}")

        if len(set(layout_parts)) != len(layout_parts):
            errors.append("Master references one or more layout parts more than once")
        if set(layout_parts) != set(physical_layouts):
            missing = sorted(set(physical_layouts) - set(layout_parts))
            extra = sorted(set(layout_parts) - set(physical_layouts))
            errors.append(
                f"Master/layout package mismatch; unreferenced={missing!r}, nonphysical={extra!r}"
            )

    slide_ids = presentation.findall("p:sldIdLst/p:sldId", NS)
    report["slide_count"] = len(slide_ids)
    if len(slide_ids) != 100:
        errors.append(f"Expected presentation to reference 100 slides, found {len(slide_ids)}")

    chart_targets: list[str] = []
    workbook_targets: list[str] = []
    editable_data_slides = 0
    for index, slide_id in enumerate(slide_ids, start=1):
        slide_rid = slide_id.get(f"{{{R_NS}}}id")
        slide_rel = presentation_rel_map.get(slide_rid or "")
        if slide_rel is None:
            errors.append(f"Presentation slide {index} relationship {slide_rid!r} is missing")
            continue
        if slide_rel.get("Type") != SLIDE_REL:
            errors.append(f"Presentation relationship for slide {index} has wrong type")
        slide_part = _relationship_target(presentation_part, slide_rel, parts, errors)
        if slide_part is None:
            continue
        slide_root = _parse_xml(parts, slide_part, errors)
        slide_rels_root = _parse_xml(parts, _rels_name(slide_part), errors)
        if slide_root is None or slide_rels_root is None:
            continue
        slide_rels = slide_rels_root.findall("rel:Relationship", REL)
        layout_rels = [rel for rel in slide_rels if rel.get("Type") == SLIDE_LAYOUT_REL]
        if len(layout_rels) != 1:
            errors.append(
                f"Slide {index} must have exactly one layout relationship; found {len(layout_rels)}"
            )
        elif index <= len(layout_parts):
            actual_layout = _relationship_target(slide_part, layout_rels[0], parts, errors)
            if actual_layout != layout_parts[index - 1]:
                errors.append(
                    f"Slide {index} uses {actual_layout}, expected {layout_parts[index - 1]}"
                )

        _check_prohibited_copy(index, slide_root, errors)

        chart_rels = [rel for rel in slide_rels if rel.get("Type") == CHART_REL]
        chart_refs = slide_root.findall(".//c:chart", NS)
        image_rels = [rel for rel in slide_rels if rel.get("Type") == IMAGE_REL]
        picture_count = len(slide_root.findall(".//p:pic", NS))
        if index in NATIVE_CHART_LAYOUTS:
            if len(chart_rels) != 1 or len(chart_refs) != 1:
                errors.append(
                    f"Native chart slide {index} must contain one chart relationship and one "
                    f"chart frame; found {len(chart_rels)} and {len(chart_refs)}"
                )
            if image_rels or picture_count:
                errors.append(
                    f"Native chart slide {index} contains image content and may be using an "
                    "image as a chart"
                )
            if len(chart_rels) == 1:
                chart_rid = chart_refs[0].get(f"{{{R_NS}}}id") if chart_refs else None
                if chart_rid != chart_rels[0].get("Id"):
                    errors.append(
                        f"Native chart slide {index} chart frame relationship {chart_rid!r} "
                        f"does not match {chart_rels[0].get('Id')!r}"
                    )
                chart_part = _relationship_target(slide_part, chart_rels[0], parts, errors)
                if chart_part:
                    chart_targets.append(chart_part)
                    workbook = _check_chart_part(
                        index,
                        chart_part,
                        parts,
                        overrides,
                        defaults,
                        errors,
                    )
                    if workbook:
                        workbook_targets.append(workbook)
        elif chart_rels or chart_refs:
            errors.append(
                f"Slide {index} is not a supported native-chart example but contains chart content"
            )

        if index in EDITABLE_DATA_LAYOUTS:
            if image_rels or picture_count:
                errors.append(
                    f"Editable data slide {index} contains a picture instead of editable content"
                )
            if index in (68, 69, 70):
                non_placeholder_shapes = [
                    shape
                    for shape in slide_root.findall(".//p:sp", NS)
                    if shape.find("p:nvSpPr/p:nvPr/p:ph", NS) is None
                ]
                minimum_shapes = {68: 6, 69: 24, 70: 5}[index]
                if len(non_placeholder_shapes) < minimum_shapes:
                    errors.append(
                        f"Editable data slide {index} has {len(non_placeholder_shapes)} native "
                        f"shapes; expected at least {minimum_shapes}"
                    )
                else:
                    editable_data_slides += 1
            else:
                body_shapes = []
                for shape in slide_root.findall(".//p:sp", NS):
                    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
                    if ph is None or _placeholder_type(ph) != "body":
                        continue
                    if ph.get("idx") in {"10", "11", "12", "13"}:
                        text = "".join(shape.xpath(".//a:t/text()", namespaces=NS)).strip()
                        if text:
                            body_shapes.append(shape)
                if len(body_shapes) != 4:
                    errors.append(
                        f"Editable KPI slide 71 has {len(body_shapes)} populated KPI placeholders; "
                        "expected four"
                    )
                else:
                    editable_data_slides += 1

        if index == 100 and slide_root.findall(".//p:ph", NS):
            errors.append("Slide 100 contains placeholders despite using the blank canvas layout")

    report["native_chart_slide_count"] = len(NATIVE_CHART_LAYOUTS)
    report["native_chart_part_count"] = len(chart_targets)
    report["embedded_workbook_count"] = len(workbook_targets)
    report["editable_shape_data_slide_count"] = editable_data_slides
    if len(chart_targets) != len(NATIVE_CHART_LAYOUTS):
        errors.append(
            f"Expected {len(NATIVE_CHART_LAYOUTS)} native chart parts, found {len(chart_targets)}"
        )
    if len(set(chart_targets)) != len(chart_targets):
        errors.append("Native chart slides do not each use a distinct chart part")
    if len(workbook_targets) != len(NATIVE_CHART_LAYOUTS):
        errors.append(
            f"Expected {len(NATIVE_CHART_LAYOUTS)} embedded workbooks, "
            f"found {len(workbook_targets)}"
        )
    if len(set(workbook_targets)) != len(workbook_targets):
        errors.append("Native charts do not each use a distinct embedded workbook")
    if editable_data_slides != len(EDITABLE_DATA_LAYOUTS):
        errors.append(
            f"Expected {len(EDITABLE_DATA_LAYOUTS)} editable shape/KPI data examples, "
            f"validated {editable_data_slides}"
        )

    # A final consumer-level parse supplements, but does not replace, the
    # package checks above.
    try:
        prs = Presentation(str(path))
        if len(prs.slide_masters) != 1:
            errors.append(
                f"python-pptx sees {len(prs.slide_masters)} masters; expected exactly one"
            )
        if len(prs.slide_layouts) != 100:
            errors.append(
                f"python-pptx sees {len(prs.slide_layouts)} layouts; expected exactly 100"
            )
        else:
            names = tuple(layout.name for layout in prs.slide_layouts)
            if names != EXPECTED_LAYOUT_NAMES:
                errors.append("python-pptx layout names/order do not match the catalog")
        if len(prs.slides) != 100:
            errors.append(f"python-pptx sees {len(prs.slides)} slides; expected exactly 100")
    except Exception as exc:  # pragma: no cover - exact package errors vary by dependency version.
        errors.append(f"python-pptx could not load the deck: {type(exc).__name__}: {exc}")

    report["ok"] = not errors
    return report


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "pptx",
        nargs="?",
        type=Path,
        default=Path("references/plain-fallback-100-layouts.pptx"),
        help="Generated fallback deck to verify",
    )
    parser.add_argument(
        "--warn-only",
        action="store_true",
        help="Print failures but exit successfully",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)
    report = inspect_deck(args.pptx)
    print(json.dumps(report, indent=2))
    return 0 if report["ok"] or args.warn_only else 1


if __name__ == "__main__":
    raise SystemExit(main())
