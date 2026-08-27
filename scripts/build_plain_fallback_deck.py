#!/usr/bin/env python3
"""Build the opt-in plain fallback deck with 100 named slide layouts.

The generator deliberately uses a two-stage workflow:

1. Create one PowerPoint master and inject 100 genuine slide-layout OOXML parts.
2. Reopen that template with python-pptx and populate one example slide per layout.

Charts supported by python-pptx remain native PowerPoint charts with embedded
Excel workbooks. Treemap, heatmap, and geographic examples use editable native
PowerPoint shapes/tables because those Office extended chart types are not
supported by python-pptx.
"""

from __future__ import annotations

import argparse
import copy
import posixpath
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_LABEL_POSITION,
    XL_LEGEND_POSITION,
)
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P_NS, "a": A_NS, "r": R_NS}
NSMAP = {"p": P_NS, "a": A_NS, "r": R_NS}

SLIDE_LAYOUT_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
SLIDE_MASTER_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
SLIDE_LAYOUT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
)

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
EMU_PER_INCH = 914400

NAVY = RGBColor(0x17, 0x32, 0x4D)
BLUE = RGBColor(0x2F, 0x80, 0xED)
SLATE = RGBColor(0x53, 0x62, 0x73)
MID = RGBColor(0xA7, 0xB3, 0xBF)
PALE = RGBColor(0xEE, 0xF2, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
AMBER = RGBColor(0xD9, 0x7B, 0x29)
RED = RGBColor(0xC6, 0x3C, 0x43)
SERIES_COLORS = (BLUE, NAVY, GREEN, AMBER, RED)
DESCENDING_BAR_LAYOUTS = frozenset({54, 55, 56, 57, 58, 67})


@dataclass(frozen=True)
class Slot:
    """A genuine placeholder on a slide layout."""

    role: str
    placeholder_type: str
    left: float
    top: float
    width: float
    height: float
    idx: int


@dataclass(frozen=True)
class LayoutSpec:
    number: int
    name: str
    category: str
    slots: tuple[Slot, ...]


LAYOUT_ROWS = (
    (1, "Standard title slide", "Opening and navigation"),
    (2, "Hero-image title slide", "Opening and navigation"),
    (3, "Split-panel title slide", "Opening and navigation"),
    (4, "Minimal title slide", "Opening and navigation"),
    (5, "Event or webinar title slide", "Opening and navigation"),
    (6, "Agenda slide", "Opening and navigation"),
    (7, "Agenda with progress indicator", "Opening and navigation"),
    (8, "Standard section divider", "Opening and navigation"),
    (9, "Image-led section divider", "Opening and navigation"),
    (10, "Numbered chapter divider", "Opening and navigation"),
    (11, "Single-speaker slide", "People and organization"),
    (12, "Two-speaker slide", "People and organization"),
    (13, "Panel-speaker slide", "People and organization"),
    (14, "Team grid slide", "People and organization"),
    (15, "Executive biography slide", "People and organization"),
    (16, "Title-only slide", "Core content"),
    (17, "Title and content slide", "Core content"),
    (18, "Title, subtitle and content slide", "Core content"),
    (19, "Title and two columns", "Core content"),
    (20, "Title with left third and right two-thirds", "Core content"),
    (21, "Title with left two-thirds and right third", "Core content"),
    (22, "Title and three columns", "Core content"),
    (23, "Title and four-card grid", "Core content"),
    (24, "50:50 stat and description", "Core content"),
    (25, "Big-statement slide", "Core content"),
    (26, "Single callout-stat slide", "Core content"),
    (27, "Three-stat row", "Core content"),
    (28, "Definition or explainer slide", "Core content"),
    (29, "Quotation slide", "Core content"),
    (30, "Quotation with portrait", "Core content"),
    (31, "Key-takeaways slide", "Core content"),
    (32, "Executive-summary slide", "Core content"),
    (33, "Recommendation and rationale slide", "Core content"),
    (34, "Side-by-side comparison slide", "Core content"),
    (35, "Pros and cons slide", "Core content"),
    (36, "Before-and-after slide", "Core content"),
    (37, "Problem-and-solution slide", "Core content"),
    (38, "Challenge, approach and result slide", "Core content"),
    (39, "Customer-proof slide", "Core content"),
    (40, "Frequently asked questions slide", "Core content"),
    (41, "Question or discussion-prompt slide", "Visual, product and media"),
    (42, "Left image and right text", "Visual, product and media"),
    (43, "Left text and right image", "Visual, product and media"),
    (44, "Full-bleed image with text overlay", "Visual, product and media"),
    (45, "Image-gallery slide", "Visual, product and media"),
    (46, "Product-screenshot slide", "Visual, product and media"),
    (47, "Annotated-screenshot slide", "Visual, product and media"),
    (48, "Video or live-demo slide", "Visual, product and media"),
    (49, "One-series column chart", "Chart and data"),
    (50, "Two-series column chart", "Chart and data"),
    (51, "Multi-series clustered column chart", "Chart and data"),
    (52, "Stacked column chart", "Chart and data"),
    (53, "100% stacked column chart", "Chart and data"),
    (54, "One-series horizontal bar chart", "Chart and data"),
    (55, "Two-series horizontal bar chart", "Chart and data"),
    (56, "Stacked horizontal bar chart", "Chart and data"),
    (57, "100% stacked horizontal bar chart", "Chart and data"),
    (58, "Ranked bar chart", "Chart and data"),
    (59, "Pie chart", "Chart and data"),
    (60, "Donut chart", "Chart and data"),
    (61, "Line chart", "Chart and data"),
    (62, "Area or stacked-area chart", "Chart and data"),
    (63, "Combination chart", "Chart and data"),
    (64, "Waterfall chart", "Chart and data"),
    (65, "Scatter-plot slide", "Chart and data"),
    (66, "Bubble-chart slide", "Chart and data"),
    (67, "Funnel-chart slide", "Chart and data"),
    (68, "Treemap slide", "Chart and data"),
    (69, "Heatmap slide", "Chart and data"),
    (70, "Geographic-map slide", "Chart and data"),
    (71, "KPI scorecard or dashboard", "Chart and data"),
    (72, "Chart with takeaway panel", "Chart and data"),
    (73, "Timeline slide", "Process and strategy"),
    (74, "Roadmap slide", "Process and strategy"),
    (75, "Horizontal-process slide", "Process and strategy"),
    (76, "Vertical-process slide", "Process and strategy"),
    (77, "Chevron-process slide", "Process and strategy"),
    (78, "Circular-process slide", "Process and strategy"),
    (79, "Process-funnel slide", "Process and strategy"),
    (80, "Swimlane-process slide", "Process and strategy"),
    (81, "Workflow or decision-flow slide", "Process and strategy"),
    (82, "Customer-journey slide", "Process and strategy"),
    (83, "Lifecycle slide", "Process and strategy"),
    (84, "Organizational-chart slide", "Process and strategy"),
    (85, "Ecosystem or network slide", "Process and strategy"),
    (86, "Hierarchy or pyramid slide", "Process and strategy"),
    (87, "Matrix slide", "Process and strategy"),
    (88, "Basic table slide", "Table and planning"),
    (89, "Comparison-table slide", "Table and planning"),
    (90, "Financial-table slide", "Table and planning"),
    (91, "Project-plan or Gantt slide", "Table and planning"),
    (92, "Calendar or schedule slide", "Table and planning"),
    (93, "RACI slide", "Table and planning"),
    (94, "Risk-register slide", "Table and planning"),
    (95, "Issue, action and status tracker", "Table and planning"),
    (96, "Next-steps slide", "Closing and administrative"),
    (97, "Q&A slide", "Closing and administrative"),
    (98, "Thank-you and contact slide", "Closing and administrative"),
    (99, "Sources, notes and legal slide", "Closing and administrative"),
    (100, "Blank branded canvas", "Closing and administrative"),
)


def slot(
    role: str,
    placeholder_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    idx: int,
) -> Slot:
    return Slot(role, placeholder_type, left, top, width, height, idx)


def title_slot(
    *, left: float = 0.75, top: float = 0.4, width: float = 11.85, height: float = 0.7
) -> Slot:
    return slot("Title", "title", left, top, width, height, 0)


def body_slot(
    idx: int, left: float, top: float, width: float, height: float, role: str = "Body content"
) -> Slot:
    return slot(role, "body", left, top, width, height, idx)


def picture_slot(idx: int, left: float, top: float, width: float, height: float) -> Slot:
    return slot("Image", "pic", left, top, width, height, idx)


def object_slot(idx: int, left: float, top: float, width: float, height: float) -> Slot:
    return slot("Editable content", "obj", left, top, width, height, idx)


def chart_slot(idx: int, left: float, top: float, width: float, height: float) -> Slot:
    return slot("Chart", "chart", left, top, width, height, idx)


def table_slot(idx: int, left: float, top: float, width: float, height: float) -> Slot:
    return slot("Table", "tbl", left, top, width, height, idx)


def media_slot(idx: int, left: float, top: float, width: float, height: float) -> Slot:
    return slot("Media", "media", left, top, width, height, idx)


def horizontal_slots(
    count: int,
    top: float,
    height: float,
    *,
    role: str = "Body content",
    start_idx: int = 10,
) -> tuple[Slot, ...]:
    gap = 0.2
    left = 0.75
    width = (11.83 - gap * (count - 1)) / count
    return tuple(
        body_slot(start_idx + i, left + i * (width + gap), top, width, height, role)
        for i in range(count)
    )


def vertical_slots(count: int, left: float, width: float, top: float, height: float) -> tuple[Slot, ...]:
    gap = 0.14
    item_height = (height - gap * (count - 1)) / count
    return tuple(
        body_slot(10 + i, left, top + i * (item_height + gap), width, item_height)
        for i in range(count)
    )


def layout_slots(number: int) -> tuple[Slot, ...]:
    """Return only the content regions essential to the layout's purpose."""

    if number == 1:
        return (title_slot(left=0.9, top=2.05, width=10.8, height=1.25), body_slot(10, 0.92, 3.55, 6.8, 1.0))
    if number == 2:
        return (
            picture_slot(10, 0.0, 0.0, 13.333, 7.5),
            title_slot(left=0.85, top=4.45, width=8.8, height=1.05),
            body_slot(11, 0.88, 5.65, 5.5, 0.7),
        )
    if number == 3:
        return (
            picture_slot(10, 0.0, 0.0, 6.0, 7.5),
            title_slot(left=6.7, top=2.0, width=5.75, height=1.25),
            body_slot(11, 6.72, 3.55, 5.2, 0.9),
        )
    if number == 4:
        return (title_slot(left=0.9, top=2.35, width=11.2, height=1.15), body_slot(10, 0.92, 3.8, 7.0, 0.65))
    if number == 5:
        return (title_slot(left=0.85, top=1.45, width=11.5, height=1.05), body_slot(10, 0.88, 2.8, 7.8, 2.0))
    if number == 6:
        return (title_slot(), body_slot(10, 0.95, 1.5, 10.9, 4.9))
    if number == 7:
        return (
            title_slot(),
            body_slot(10, 0.95, 1.55, 10.8, 4.55),
            body_slot(11, 0.95, 6.35, 10.8, 0.45, "Section label"),
        )
    if number == 8:
        return (title_slot(left=0.9, top=2.6, width=11.4, height=1.2),)
    if number == 9:
        return (
            picture_slot(10, 0.0, 0.0, 13.333, 7.5),
            title_slot(left=0.85, top=2.8, width=10.8, height=1.25),
        )
    if number == 10:
        return (
            body_slot(10, 0.9, 2.15, 1.45, 1.1, "Section label"),
            title_slot(left=2.55, top=2.2, width=9.8, height=1.1),
        )
    if number == 11:
        return (title_slot(), picture_slot(10, 0.85, 1.5, 4.0, 5.2), body_slot(11, 5.25, 1.65, 6.8, 4.6))
    if number == 12:
        return (
            title_slot(),
            picture_slot(10, 1.0, 1.55, 4.2, 3.55),
            body_slot(11, 1.0, 5.25, 4.2, 1.0),
            picture_slot(12, 7.1, 1.55, 4.2, 3.55),
            body_slot(13, 7.1, 5.25, 4.2, 1.0),
        )
    if number in (13, 14):
        count = 4 if number == 13 else 6
        cols = 4 if count == 4 else 3
        rows = 1 if count == 4 else 2
        slots: list[Slot] = [title_slot()]
        card_w = 2.75 if cols == 4 else 3.55
        x_gap = 0.25
        y_gap = 0.3
        for i in range(count):
            row, col = divmod(i, cols)
            left = 0.75 + col * (card_w + x_gap)
            top = 1.45 + row * (2.45 + y_gap)
            slots.append(picture_slot(10 + i * 2, left, top, card_w, 1.55))
            slots.append(body_slot(11 + i * 2, left, top + 1.65, card_w, 0.72))
        return tuple(slots)
    if number == 15:
        return (title_slot(), picture_slot(10, 0.85, 1.5, 3.8, 5.15), body_slot(11, 5.05, 1.55, 7.05, 4.95))
    if number == 16:
        return (title_slot(),)
    if number == 17:
        return (title_slot(), object_slot(10, 0.78, 1.45, 11.75, 5.35))
    if number == 18:
        return (
            title_slot(),
            slot("Subtitle", "subTitle", 0.78, 1.1, 11.7, 0.48, 10),
            object_slot(11, 0.78, 1.72, 11.75, 5.05),
        )
    if number == 19:
        return (title_slot(),) + horizontal_slots(2, 1.55, 5.05)
    if number == 20:
        return (title_slot(), body_slot(10, 0.75, 1.5, 3.55, 5.1), object_slot(11, 4.55, 1.5, 7.98, 5.1))
    if number == 21:
        return (title_slot(), object_slot(10, 0.75, 1.5, 7.98, 5.1), body_slot(11, 8.98, 1.5, 3.55, 5.1))
    if number == 22:
        return (title_slot(),) + horizontal_slots(3, 1.55, 5.0)
    if number == 23:
        return (title_slot(),) + horizontal_slots(4, 1.7, 4.75)
    if number == 24:
        return (title_slot(), body_slot(10, 0.8, 1.55, 5.35, 5.0), body_slot(11, 6.4, 1.55, 6.0, 5.0))
    if number == 25:
        return (title_slot(left=1.05, top=2.15, width=11.0, height=2.45),)
    if number == 26:
        return (title_slot(left=2.0, top=1.55, width=9.3, height=2.25), body_slot(10, 2.25, 4.15, 8.8, 1.25))
    if number == 27:
        return (title_slot(),) + horizontal_slots(3, 1.75, 4.6)
    if number == 28:
        return (title_slot(), body_slot(10, 0.8, 1.55, 3.5, 4.95), body_slot(11, 4.55, 1.55, 7.95, 4.95))
    if number == 29:
        return (
            body_slot(10, 1.15, 1.65, 11.0, 3.8, "Quote"),
            body_slot(11, 1.18, 5.65, 9.0, 0.8, "Attribution"),
        )
    if number == 30:
        return (
            picture_slot(10, 0.9, 1.25, 4.1, 5.3),
            body_slot(11, 5.45, 1.55, 6.7, 3.55, "Quote"),
            body_slot(12, 5.48, 5.35, 5.8, 0.8, "Attribution"),
        )
    if number == 31:
        return (title_slot(), body_slot(10, 0.9, 1.5, 11.0, 5.1))
    if number == 32:
        return (title_slot(), body_slot(10, 0.78, 1.5, 11.75, 1.35),) + horizontal_slots(
            3, 3.05, 3.35, start_idx=11
        )
    if number == 33:
        return (title_slot(), body_slot(10, 0.8, 1.55, 4.15, 4.95), body_slot(11, 5.2, 1.55, 7.3, 4.95))
    if number in (34, 35, 37):
        return (title_slot(),) + horizontal_slots(2, 1.55, 5.0)
    if number == 36:
        return (
            title_slot(),
            picture_slot(10, 0.8, 1.55, 5.65, 3.5),
            body_slot(11, 0.8, 5.2, 5.65, 1.15),
            picture_slot(12, 6.85, 1.55, 5.65, 3.5),
            body_slot(13, 6.85, 5.2, 5.65, 1.15),
        )
    if number == 38:
        return (title_slot(),) + horizontal_slots(3, 1.55, 5.0)
    if number == 39:
        return (title_slot(), picture_slot(10, 0.8, 1.55, 4.15, 4.95), body_slot(11, 5.2, 1.55, 7.3, 4.95))
    if number == 40:
        return (title_slot(), body_slot(10, 0.8, 1.55, 5.7, 5.0), body_slot(11, 6.75, 1.55, 5.7, 5.0))
    if number == 41:
        return (title_slot(left=1.15, top=2.05, width=11.0, height=2.2),)
    if number == 42:
        return (title_slot(), picture_slot(10, 0.75, 1.45, 6.0, 5.35), body_slot(11, 7.05, 1.55, 5.4, 5.0))
    if number == 43:
        return (title_slot(), body_slot(10, 0.78, 1.55, 5.4, 5.0), picture_slot(11, 6.5, 1.45, 6.03, 5.35))
    if number == 44:
        return (picture_slot(10, 0.0, 0.0, 13.333, 7.5), title_slot(left=0.9, top=4.75, width=9.5, height=1.4))
    if number == 45:
        return (
            title_slot(),
            picture_slot(10, 0.75, 1.5, 5.7, 2.25),
            picture_slot(11, 6.75, 1.5, 5.8, 2.25),
            picture_slot(12, 0.75, 4.0, 5.7, 2.25),
            picture_slot(13, 6.75, 4.0, 5.8, 2.25),
        )
    if number == 46:
        return (title_slot(), picture_slot(10, 0.8, 1.45, 11.75, 5.4))
    if number == 47:
        return (title_slot(), picture_slot(10, 0.75, 1.45, 8.25, 5.35), body_slot(11, 9.25, 1.55, 3.25, 5.05, "Callout"))
    if number == 48:
        return (title_slot(), media_slot(10, 0.85, 1.55, 11.55, 5.1))
    if 49 <= number <= 67:
        return (title_slot(), chart_slot(10, 0.8, 1.45, 11.75, 5.4))
    if number in (68, 69, 70):
        return (title_slot(), object_slot(10, 0.8, 1.45, 11.75, 5.4))
    if number == 71:
        return (title_slot(),) + horizontal_slots(4, 1.8, 4.5)
    if number == 72:
        return (
            title_slot(),
            chart_slot(10, 0.8, 1.5, 8.0, 5.2),
            body_slot(11, 9.05, 1.5, 3.45, 5.2, "Callout"),
        )
    if number in (73, 75, 77, 82, 83):
        return (title_slot(),) + horizontal_slots(5, 2.05, 3.75)
    if number == 74:
        return (title_slot(),) + horizontal_slots(3, 1.6, 4.85)
    if number == 76:
        return (title_slot(),) + vertical_slots(4, 1.1, 11.1, 1.45, 5.25)
    if number == 78:
        return (
            title_slot(),
            body_slot(10, 5.1, 2.75, 3.15, 1.4),
            body_slot(11, 1.2, 1.65, 3.0, 1.25),
            body_slot(12, 9.15, 1.65, 3.0, 1.25),
            body_slot(13, 1.2, 4.8, 3.0, 1.25),
            body_slot(14, 9.15, 4.8, 3.0, 1.25),
        )
    if number == 79:
        return (title_slot(),) + vertical_slots(4, 2.0, 9.3, 1.45, 5.25)
    if number == 80:
        return (title_slot(),) + vertical_slots(3, 0.9, 11.55, 1.45, 5.25)
    if number == 81:
        return (
            title_slot(),
            body_slot(10, 0.9, 2.65, 2.25, 1.35),
            body_slot(11, 3.75, 2.65, 2.25, 1.35),
            body_slot(12, 6.6, 2.65, 2.25, 1.35),
            body_slot(13, 9.45, 1.65, 2.35, 1.35),
            body_slot(14, 9.45, 4.55, 2.35, 1.35),
        )
    if number == 84:
        return (
            title_slot(),
            body_slot(10, 5.1, 1.5, 3.1, 1.15),
            body_slot(11, 1.35, 3.2, 3.1, 1.15),
            body_slot(12, 5.1, 3.2, 3.1, 1.15),
            body_slot(13, 8.85, 3.2, 3.1, 1.15),
            body_slot(14, 3.2, 5.0, 3.1, 1.15),
            body_slot(15, 7.05, 5.0, 3.1, 1.15),
        )
    if number == 85:
        return (
            title_slot(),
            body_slot(10, 5.0, 2.65, 3.3, 1.45),
            body_slot(11, 0.9, 1.45, 3.0, 1.2),
            body_slot(12, 9.4, 1.45, 3.0, 1.2),
            body_slot(13, 0.9, 5.05, 3.0, 1.2),
            body_slot(14, 9.4, 5.05, 3.0, 1.2),
        )
    if number == 86:
        return (title_slot(),) + vertical_slots(4, 2.0, 9.3, 1.45, 5.25)
    if number == 87:
        return (
            title_slot(),
            body_slot(10, 0.95, 1.55, 5.5, 2.35),
            body_slot(11, 6.85, 1.55, 5.5, 2.35),
            body_slot(12, 0.95, 4.2, 5.5, 2.35),
            body_slot(13, 6.85, 4.2, 5.5, 2.35),
        )
    if 88 <= number <= 95:
        return (title_slot(), table_slot(10, 0.75, 1.45, 11.83, 5.35))
    if number == 96:
        return (title_slot(),) + horizontal_slots(4, 1.75, 4.7)
    if number == 97:
        return (title_slot(left=1.0, top=2.3, width=11.3, height=1.8),)
    if number == 98:
        return (title_slot(left=0.95, top=2.0, width=10.8, height=1.3), body_slot(10, 0.98, 3.65, 7.5, 1.2))
    if number == 99:
        return (title_slot(), body_slot(10, 0.85, 1.5, 11.5, 5.1))
    if number == 100:
        return ()
    raise ValueError(f"No slot definition for layout {number}")


def layout_specs() -> tuple[LayoutSpec, ...]:
    return tuple(LayoutSpec(number, name, category, layout_slots(number)) for number, name, category in LAYOUT_ROWS)


def emu(inches: float) -> str:
    return str(round(inches * EMU_PER_INCH))


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def add_xfrm(parent: etree._Element, slot_spec: Slot) -> None:
    xfrm = etree.SubElement(parent, qname(A_NS, "xfrm"))
    etree.SubElement(xfrm, qname(A_NS, "off"), x=emu(slot_spec.left), y=emu(slot_spec.top))
    etree.SubElement(xfrm, qname(A_NS, "ext"), cx=emu(slot_spec.width), cy=emu(slot_spec.height))


def add_layout_placeholder(sp_tree: etree._Element, slot_spec: Slot, shape_id: int) -> None:
    shape = etree.SubElement(sp_tree, qname(P_NS, "sp"))
    nv_sp_pr = etree.SubElement(shape, qname(P_NS, "nvSpPr"))
    etree.SubElement(nv_sp_pr, qname(P_NS, "cNvPr"), id=str(shape_id), name=f"{slot_spec.role} {slot_spec.idx}")
    c_nv_sp_pr = etree.SubElement(nv_sp_pr, qname(P_NS, "cNvSpPr"))
    etree.SubElement(c_nv_sp_pr, qname(A_NS, "spLocks"), noGrp="1")
    nv_pr = etree.SubElement(nv_sp_pr, qname(P_NS, "nvPr"))
    ph_attrs = {"idx": str(slot_spec.idx)}
    if slot_spec.placeholder_type != "obj":
        ph_attrs["type"] = slot_spec.placeholder_type
    etree.SubElement(nv_pr, qname(P_NS, "ph"), **ph_attrs)

    sp_pr = etree.SubElement(shape, qname(P_NS, "spPr"))
    add_xfrm(sp_pr, slot_spec)
    geom = etree.SubElement(sp_pr, qname(A_NS, "prstGeom"), prst="rect")
    etree.SubElement(geom, qname(A_NS, "avLst"))

    tx_body = etree.SubElement(shape, qname(P_NS, "txBody"))
    body_pr = etree.SubElement(tx_body, qname(A_NS, "bodyPr"), lIns="91440", rIns="91440", tIns="45720", bIns="45720")
    if slot_spec.placeholder_type in ("title", "ctrTitle"):
        body_pr.set("anchor", "ctr")
    etree.SubElement(tx_body, qname(A_NS, "lstStyle"))
    paragraph = etree.SubElement(tx_body, qname(A_NS, "p"))
    run = etree.SubElement(paragraph, qname(A_NS, "r"))
    etree.SubElement(run, qname(A_NS, "rPr"), lang="en-GB")
    etree.SubElement(run, qname(A_NS, "t")).text = f"Click to add {slot_spec.role.lower()}"
    etree.SubElement(paragraph, qname(A_NS, "endParaRPr"), lang="en-GB")


def add_fixed_shape(
    sp_tree: etree._Element,
    shape_id: int,
    name: str,
    left: float,
    top: float,
    width: float,
    height: float,
    color: str,
) -> None:
    shape = etree.SubElement(sp_tree, qname(P_NS, "sp"))
    nv_sp_pr = etree.SubElement(shape, qname(P_NS, "nvSpPr"))
    etree.SubElement(nv_sp_pr, qname(P_NS, "cNvPr"), id=str(shape_id), name=name)
    etree.SubElement(nv_sp_pr, qname(P_NS, "cNvSpPr"))
    etree.SubElement(nv_sp_pr, qname(P_NS, "nvPr"))
    sp_pr = etree.SubElement(shape, qname(P_NS, "spPr"))
    add_xfrm(sp_pr, Slot(name, "obj", left, top, width, height, 999))
    geom = etree.SubElement(sp_pr, qname(A_NS, "prstGeom"), prst="rect")
    etree.SubElement(geom, qname(A_NS, "avLst"))
    solid = etree.SubElement(sp_pr, qname(A_NS, "solidFill"))
    etree.SubElement(solid, qname(A_NS, "srgbClr"), val=color)
    line = etree.SubElement(sp_pr, qname(A_NS, "ln"))
    etree.SubElement(line, qname(A_NS, "noFill"))
    tx_body = etree.SubElement(shape, qname(P_NS, "txBody"))
    etree.SubElement(tx_body, qname(A_NS, "bodyPr"))
    etree.SubElement(tx_body, qname(A_NS, "lstStyle"))
    etree.SubElement(tx_body, qname(A_NS, "p"))


def build_layout_xml(spec: LayoutSpec) -> bytes:
    root = etree.Element(
        qname(P_NS, "sldLayout"),
        nsmap=NSMAP,
        type="cust",
        preserve="1",
        matchingName=spec.name,
    )
    c_sld = etree.SubElement(root, qname(P_NS, "cSld"), name=spec.name)
    sp_tree = etree.SubElement(c_sld, qname(P_NS, "spTree"))
    nv_grp = etree.SubElement(sp_tree, qname(P_NS, "nvGrpSpPr"))
    etree.SubElement(nv_grp, qname(P_NS, "cNvPr"), id="1", name="")
    etree.SubElement(nv_grp, qname(P_NS, "cNvGrpSpPr"))
    etree.SubElement(nv_grp, qname(P_NS, "nvPr"))
    grp_sp_pr = etree.SubElement(sp_tree, qname(P_NS, "grpSpPr"))
    grp_xfrm = etree.SubElement(grp_sp_pr, qname(A_NS, "xfrm"))
    for tag, attrs in (
        ("off", {"x": "0", "y": "0"}),
        ("ext", {"cx": "0", "cy": "0"}),
        ("chOff", {"x": "0", "y": "0"}),
        ("chExt", {"cx": "0", "cy": "0"}),
    ):
        etree.SubElement(grp_xfrm, qname(A_NS, tag), **attrs)

    image_led = spec.number in (2, 3, 9, 44, 100)
    divider = spec.number in (8, 10)
    if not image_led:
        if divider:
            add_fixed_shape(sp_tree, 2, "Accent rule", 0.85, 2.0, 0.08, 2.25, "2F80ED")
        else:
            add_fixed_shape(sp_tree, 2, "Accent rule", 0.75, 1.22, 1.1, 0.06, "2F80ED")

    for offset, slot_spec in enumerate(spec.slots, start=10):
        add_layout_placeholder(sp_tree, slot_spec, offset)

    clr_map = etree.SubElement(root, qname(P_NS, "clrMapOvr"))
    etree.SubElement(clr_map, qname(A_NS, "masterClrMapping"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_layout_rels() -> bytes:
    root = etree.Element(qname(REL_NS, "Relationships"), nsmap={None: REL_NS})
    etree.SubElement(
        root,
        qname(REL_NS, "Relationship"),
        Id="rId1",
        Type=SLIDE_MASTER_REL,
        Target="../slideMasters/slideMaster1.xml",
    )
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_theme(parts: dict[str, bytes]) -> None:
    theme_name = "ppt/theme/theme1.xml"
    root = etree.fromstring(parts[theme_name])
    root.set("name", "Plain fallback")
    major_latin = root.find(".//a:themeElements/a:fontScheme/a:majorFont/a:latin", NS)
    minor_latin = root.find(".//a:themeElements/a:fontScheme/a:minorFont/a:latin", NS)
    if major_latin is not None:
        major_latin.set("typeface", "Aptos Display")
    if minor_latin is not None:
        minor_latin.set("typeface", "Aptos")

    colors = {
        "dk1": "17324D",
        "lt1": "FFFFFF",
        "dk2": "536273",
        "lt2": "EEF2F5",
        "accent1": "2F80ED",
        "accent2": "17324D",
        "accent3": "2E8B57",
        "accent4": "D97B29",
        "accent5": "C63C43",
        "accent6": "A7B3BF",
    }
    clr_scheme = root.find(".//a:themeElements/a:clrScheme", NS)
    if clr_scheme is not None:
        clr_scheme.set("name", "Plain fallback")
        for child in clr_scheme:
            local = etree.QName(child).localname
            if local not in colors:
                continue
            for existing in list(child):
                child.remove(existing)
            etree.SubElement(child, qname(A_NS, "srgbClr"), val=colors[local])
    parts[theme_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def replace_layout_parts(seed: Path, output: Path, specs: tuple[LayoutSpec, ...]) -> None:
    with zipfile.ZipFile(seed) as zf:
        parts = {name: zf.read(name) for name in zf.namelist()}

    for name in list(parts):
        if name.startswith("ppt/slideLayouts/"):
            del parts[name]

    master_name = "ppt/slideMasters/slideMaster1.xml"
    master_rels_name = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
    master = etree.fromstring(parts[master_name])
    master_rels = etree.fromstring(parts[master_rels_name])
    content_types = etree.fromstring(parts["[Content_Types].xml"])

    layout_list = master.find("p:sldLayoutIdLst", NS)
    if layout_list is None:
        c_sld = master.find("p:cSld", NS)
        layout_list = etree.Element(qname(P_NS, "sldLayoutIdLst"))
        if c_sld is None:
            raise RuntimeError("slide master has no cSld element")
        c_sld.addnext(layout_list)
    else:
        for child in list(layout_list):
            layout_list.remove(child)

    for rel in list(master_rels):
        if rel.get("Type") == SLIDE_LAYOUT_REL:
            master_rels.remove(rel)

    for override in list(content_types):
        if override.get("PartName", "").startswith("/ppt/slideLayouts/"):
            content_types.remove(override)

    # Administrative placeholders are intentionally absent from the fallback master.
    sp_tree = master.find("p:cSld/p:spTree", NS)
    if sp_tree is not None:
        for shape in list(sp_tree):
            ph = shape.find(".//p:ph", NS)
            if ph is not None and ph.get("type") in {"dt", "ftr", "hdr", "sldNum"}:
                sp_tree.remove(shape)

    for index, spec in enumerate(specs, start=1):
        rel_id = f"rIdLayout{index}"
        etree.SubElement(
            master_rels,
            qname(REL_NS, "Relationship"),
            Id=rel_id,
            Type=SLIDE_LAYOUT_REL,
            Target=f"../slideLayouts/slideLayout{index}.xml",
        )
        layout_id = etree.SubElement(
            layout_list,
            qname(P_NS, "sldLayoutId"),
            id=str(2147483648 + index),
        )
        layout_id.set(qname(R_NS, "id"), rel_id)
        etree.SubElement(
            content_types,
            qname(CT_NS, "Override"),
            PartName=f"/ppt/slideLayouts/slideLayout{index}.xml",
            ContentType=SLIDE_LAYOUT_CONTENT_TYPE,
        )
        parts[f"ppt/slideLayouts/slideLayout{index}.xml"] = build_layout_xml(spec)
        parts[f"ppt/slideLayouts/_rels/slideLayout{index}.xml.rels"] = build_layout_rels()

    parts[master_name] = etree.tostring(master, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts[master_rels_name] = etree.tostring(
        master_rels, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    parts["[Content_Types].xml"] = etree.tostring(
        content_types, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    patch_theme(parts)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(parts):
            zf.writestr(name, parts[name])


def create_placeholder_image(path: Path) -> None:
    width, height = 1600, 900
    image = Image.new("RGB", (width, height), (238, 242, 245))
    draw = ImageDraw.Draw(image)
    for y in range(height):
        ratio = y / max(height - 1, 1)
        color = (
            round(47 + (23 - 47) * ratio),
            round(128 + (50 - 128) * ratio),
            round(237 + (77 - 237) * ratio),
        )
        draw.line((0, y, width, y), fill=color)
    draw.polygon(((0, 650), (500, 360), (900, 720), (1260, 430), (1600, 700), (1600, 900), (0, 900)), fill=(23, 50, 77))
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path)


def set_shape_text(shape, text: str, *, title: bool = False, panel: bool = False) -> None:
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.14)
    text_frame.margin_right = Inches(0.14)
    text_frame.margin_top = Inches(0.08)
    text_frame.margin_bottom = Inches(0.08)
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text.split("\n")
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.text = line
        p_pr = paragraph._p.get_or_add_pPr()
        for child in list(p_pr):
            if child.tag in {
                qn("a:buAutoNum"),
                qn("a:buBlip"),
                qn("a:buChar"),
                qn("a:buNone"),
            }:
                p_pr.remove(child)
        p_pr.insert(0, OxmlElement("a:buNone"))
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_after = Pt(5 if index == 0 else 2)
        if paragraph.runs:
            run = paragraph.runs[0]
            run.font.name = "Aptos"
            run.font.color.rgb = NAVY if title or index == 0 else SLATE
            run.font.size = Pt(28 if title else (18 if index == 0 else 13))
            run.font.bold = bool(title or index == 0)
    if title:
        shape.fill.background()
        shape.line.fill.background()
    elif panel:
        shape.fill.solid()
        shape.fill.fore_color.rgb = PALE
        shape.line.color.rgb = MID
        shape.line.width = Pt(0.75)
    else:
        shape.fill.background()
        shape.line.fill.background()


def body_example(spec: LayoutSpec, slot_spec: Slot, ordinal: int) -> str:
    number = spec.number
    if number == 1:
        return "100 reusable layouts\nOpt-in fallback only"
    if number in (3, 4, 5):
        return "Presenter name · Role\n27 August 2026"
    if number in (6, 7):
        return "1  Context\n2  Evidence\n3  Decision\n4  Actions" if slot_spec.role != "Section label" else "Current section: Evidence"
    if number == 10:
        return "03"
    if 11 <= number <= 15:
        names = ("Alex Morgan", "Taylor Chen", "Jordan Singh", "Sam Rivera", "Riley Okafor", "Casey Brown")
        name = names[ordinal % len(names)]
        return f"{name}\nRole · Organization"
    if number == 18 and slot_spec.role == "Subtitle":
        return "Supporting context when supplied"
    if number in (24, 26, 27, 71):
        values = ("42%", "£3.2m", "18 days", "+12%")
        labels = ("Adoption rate", "Annual value", "Cycle time", "Year-on-year")
        return f"{values[ordinal % 4]}\n{labels[ordinal % 4]}"
    if number in (29, 30):
        if slot_spec.role == "Attribution":
            return "Name · Role · Organization"
        return "Source-grounded quotation"
    if number == 31:
        return "1  Finding\n2  Implication\n3  Action"
    if number == 32:
        labels = ("Recommendation\nDecision required", "Finding\nEvidence", "Impact\nMeasured effect", "Next step\nOwner")
        return labels[ordinal % len(labels)]
    if number == 33:
        return "Recommendation\nSpecific proposed action" if ordinal == 0 else "Rationale\nEvidence, impact and risk"
    if number == 34:
        return ("Option A\nCriteria and result", "Option B\nCriteria and result")[ordinal % 2]
    if number == 35:
        return ("Advantages\nEvidence-backed benefit", "Limitations\nKnown risk or cost")[ordinal % 2]
    if number == 37:
        return ("Problem\nEvidence and consequence", "Solution\nCapability and expected benefit")[ordinal % 2]
    if number == 38:
        return ("Challenge\nStarting condition", "Approach\nWork completed", "Result\nMeasured outcome")[ordinal % 3]
    if number == 39:
        return "Customer outcome\nMeasured result and use case"
    if number == 40:
        return ("Question 1\nPlain answer", "Question 2\nPlain answer")[ordinal % 2]
    if number in (42, 43):
        return "Heading\nFactual explanation of the supplied image"
    if number == 47:
        return "1  Feature\n2  Action\n3  Result"
    if number == 72:
        return "Supported conclusion\nAction justified by the data"
    if 73 <= number <= 87:
        return f"{ordinal + 1:02d}  Stage\nOwner or output"
    if number == 96:
        return f"{ordinal + 1:02d}\nAction · Owner · Date"
    if number == 98:
        return "Presenter name · Role\nemail@example.com"
    if number == 99:
        return "Sources\n1. Author, title, date\n\nMethodology or legal text when required"
    if number == 2:
        return "Presenter name · Role"
    return f"Item {ordinal + 1}\nFunctional supporting detail"


def find_slot_by_idx(spec: LayoutSpec, idx: int) -> Slot:
    for slot_spec in spec.slots:
        if slot_spec.idx == idx:
            return slot_spec
    raise KeyError(idx)


def style_picture_frame(shape) -> None:
    shape.line.color.rgb = WHITE
    shape.line.width = Pt(1)


def populate_text_placeholders(slide, spec: LayoutSpec) -> None:
    body_slots = [item for item in spec.slots if item.placeholder_type in {"body", "subTitle"}]
    multiple_panels = len([item for item in body_slots if item.role == "Body content"]) > 1
    body_ordinal = 0
    for placeholder in list(slide.placeholders):
        idx = placeholder.placeholder_format.idx
        slot_spec = find_slot_by_idx(spec, idx)
        if slot_spec.placeholder_type in {"pic", "chart", "tbl", "media", "obj"}:
            continue
        if slot_spec.placeholder_type == "title":
            title_text = "Plain fallback layout catalog" if spec.number == 1 else spec.name
            set_shape_text(placeholder, title_text, title=True)
            if spec.number in (25, 41, 97):
                for paragraph in placeholder.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(34)
            if spec.number in (2, 9, 44):
                for paragraph in placeholder.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = WHITE
            continue
        text = body_example(spec, slot_spec, body_ordinal)
        panel = multiple_panels or slot_spec.role in {"Callout", "Quote"} or 73 <= spec.number <= 87
        set_shape_text(placeholder, text, panel=panel)
        if spec.number == 2:
            for paragraph in placeholder.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = WHITE
        if spec.number in (24, 26, 27, 71):
            paragraphs = placeholder.text_frame.paragraphs
            if paragraphs and paragraphs[0].runs:
                paragraphs[0].runs[0].font.size = Pt(36)
                paragraphs[0].runs[0].font.bold = True
            if len(paragraphs) > 1 and paragraphs[1].runs:
                paragraphs[1].runs[0].font.size = Pt(13)
        body_ordinal += 1


def insert_pictures(slide, spec: LayoutSpec, image_path: Path) -> None:
    for placeholder in list(slide.placeholders):
        idx = placeholder.placeholder_format.idx
        slot_spec = find_slot_by_idx(spec, idx)
        if slot_spec.placeholder_type != "pic":
            continue
        picture = placeholder.insert_picture(str(image_path))
        style_picture_frame(picture)


def category_data(series_count: int, *, percentages: bool = False, ranked: bool = False) -> CategoryChartData:
    data = CategoryChartData(number_format="0%" if percentages else "0")
    data.categories = ["North", "Central", "South", "West"]
    values = (
        (42, 36, 31, 25),
        (35, 33, 28, 22),
        (24, 27, 19, 17),
    )
    if ranked:
        data.categories = ["Priority A", "Priority B", "Priority C", "Priority D", "Priority E"]
        values = ((92, 78, 63, 51, 39),)
    if percentages:
        values = tuple(tuple(value / 100 for value in row) for row in values)
    for index in range(series_count):
        data.add_series(("Actual", "Target", "Prior period")[index], values[index])
    return data


def chart_definition(number: int):
    if number == 49:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(1)
    if number == 50:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(2)
    if number == 51:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(3)
    if number == 52:
        return XL_CHART_TYPE.COLUMN_STACKED, category_data(3)
    if number == 53:
        return XL_CHART_TYPE.COLUMN_STACKED_100, category_data(3, percentages=True)
    if number == 54:
        return XL_CHART_TYPE.BAR_CLUSTERED, category_data(1)
    if number == 55:
        return XL_CHART_TYPE.BAR_CLUSTERED, category_data(2)
    if number == 56:
        return XL_CHART_TYPE.BAR_STACKED, category_data(3)
    if number == 57:
        return XL_CHART_TYPE.BAR_STACKED_100, category_data(3, percentages=True)
    if number == 58:
        return XL_CHART_TYPE.BAR_CLUSTERED, category_data(1, ranked=True)
    if number in (59, 60):
        data = CategoryChartData(number_format="0%")
        data.categories = ["Enterprise", "SMB", "Consumer"]
        data.add_series("Share", (0.52, 0.31, 0.17))
        return (XL_CHART_TYPE.PIE if number == 59 else XL_CHART_TYPE.DOUGHNUT), data
    if number == 61:
        return XL_CHART_TYPE.LINE_MARKERS, category_data(2)
    if number == 62:
        return XL_CHART_TYPE.AREA_STACKED, category_data(3)
    if number == 63:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(2)
    if number == 64:
        data = CategoryChartData(number_format="0")
        data.categories = ["Start", "Growth", "Savings", "Investment", "End"]
        data.add_series("Base", (0, 80, 95, 90, 0))
        data.add_series("Movement", (80, 15, 10, 12, 102))
        return XL_CHART_TYPE.COLUMN_STACKED, data
    if number == 65:
        data = XyChartData(number_format="0")
        series = data.add_series("Business units")
        for x, y in ((24, 38), (35, 52), (48, 55), (63, 71), (78, 82)):
            series.add_data_point(x, y)
        return XL_CHART_TYPE.XY_SCATTER, data
    if number == 66:
        data = BubbleChartData(number_format="0")
        series = data.add_series("Markets")
        for x, y, size in ((24, 42, 14), (39, 58, 24), (62, 48, 18), (75, 76, 30)):
            series.add_data_point(x, y, size)
        return XL_CHART_TYPE.BUBBLE, data
    if number == 67:
        data = CategoryChartData(number_format="0")
        data.categories = ["Awareness", "Interest", "Evaluation", "Decision", "Adoption"]
        data.add_series("Count", (100, 78, 55, 36, 24))
        return XL_CHART_TYPE.BAR_CLUSTERED, data
    if number == 72:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, category_data(2)
    raise ValueError(number)


def set_series_colors(chart) -> None:
    for index, series in enumerate(chart.series):
        color = SERIES_COLORS[index % len(SERIES_COLORS)]
        try:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color
        except (AttributeError, ValueError):
            pass
        try:
            series.format.line.color.rgb = color
            series.format.line.width = Pt(2)
        except (AttributeError, ValueError):
            pass


def style_chart(chart, number: int) -> None:
    chart.has_title = False
    chart.chart_style = 10
    set_series_colors(chart)
    multi_series = len(chart.series) > 1
    chart.has_legend = multi_series and number not in (64,)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

    if number in (59, 60):
        chart.has_legend = False
        plot = chart.plots[0]
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.show_category_name = True
        labels.show_percentage = True
        labels.show_value = False
        labels.number_format = "0%"
        labels.position = XL_LABEL_POSITION.BEST_FIT if number == 59 else XL_LABEL_POSITION.CENTER
        if number == 60 and hasattr(plot, "hole_size"):
            plot.hole_size = 62
        return

    axes_required = number in (61, 62, 63, 65, 66)
    try:
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.has_title = False
        chart.value_axis.visible = axes_required
    except ValueError:
        pass
    try:
        chart.category_axis.has_title = False
        chart.category_axis.has_major_gridlines = False
        if number in DESCENDING_BAR_LAYOUTS:
            # PowerPoint renders the first bar category at the bottom by default.
            # The source data is already ranked high-to-low, so reverse the
            # category axis to keep the largest value or stacked total on top.
            chart.category_axis.reverse_order = True
    except (AttributeError, ValueError):
        pass

    labels_preferred = number not in (61, 62, 63, 65, 66)
    if labels_preferred:
        plot = chart.plots[0]
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.show_value = True
        labels.show_category_name = False
        labels.show_series_name = False
        labels.position = (
            XL_LABEL_POSITION.INSIDE_END
            if number in (52, 53, 56, 57, 64)
            else XL_LABEL_POSITION.OUTSIDE_END
        )
        labels.number_format = "0%" if number in (53, 57) else "0"

    if number == 64:
        chart.has_legend = False
        try:
            chart.series[0].format.fill.background()
            chart.series[0].format.line.fill.background()
        except (AttributeError, ValueError):
            pass


def insert_chart(slide, spec: LayoutSpec) -> None:
    if spec.number not in set(range(49, 68)) | {72}:
        return
    chart_placeholder = next(
        placeholder
        for placeholder in slide.placeholders
        if placeholder.placeholder_format.type == PP_PLACEHOLDER.CHART
    )
    chart_type, data = chart_definition(spec.number)
    chart = chart_placeholder.insert_chart(chart_type, data).chart
    style_chart(chart, spec.number)


def style_table(table, number: int) -> None:
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if row_index == 0 else (PALE if row_index % 2 else WHITE)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.name = "Aptos"
                    run.font.size = Pt(11)
                    run.font.color.rgb = WHITE if row_index == 0 else NAVY
                    run.font.bold = row_index == 0


TABLE_SCHEMAS = {
    88: (["Item", "Owner", "Status", "Date"], 4),
    89: (["Criterion", "Option A", "Option B", "Decision"], 4),
    90: (["Line item", "Actual", "Budget", "Variance", "Forecast"], 5),
    91: (["Workstream", "Aug", "Sep", "Oct", "Nov", "Owner"], 5),
    92: (["Date", "Activity", "Owner", "Location", "Status", "Notes"], 4),
    93: (["Task", "Product", "Sales", "Finance", "Operations"], 5),
    94: (["Risk", "Likelihood", "Impact", "Rating", "Mitigation", "Owner", "Review"], 4),
    95: (["Item", "Type", "Description", "Owner", "Due", "Status", "Update"], 4),
}


def insert_table(slide, spec: LayoutSpec) -> None:
    if spec.number not in TABLE_SCHEMAS:
        return
    headers, rows = TABLE_SCHEMAS[spec.number]
    placeholder = next(
        ph for ph in slide.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.TABLE
    )
    table = placeholder.insert_table(rows, len(headers)).table
    for col_index, header in enumerate(headers):
        table.cell(0, col_index).text = header
    for row_index in range(1, rows):
        for col_index, header in enumerate(headers):
            value = f"{header} {row_index}"
            if header in {"Status", "Rating", "Decision"}:
                value = ("Open", "On track", "Review")[row_index % 3]
            table.cell(row_index, col_index).text = value
    style_table(table, spec.number)


def add_shape_text(slide, left, top, width, height, text, *, fill=PALE, color=NAVY, size=14, bold=False):
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = WHITE
    shape.line.width = Pt(1)
    tf = shape.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    run = p.runs[0]
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def fill_object_example(slide, spec: LayoutSpec) -> None:
    object_placeholders = [
        ph for ph in list(slide.placeholders) if ph.placeholder_format.type == PP_PLACEHOLDER.OBJECT
    ]
    if spec.number not in (17, 18, 20, 21, 68, 69, 70):
        for ordinal, placeholder in enumerate(object_placeholders):
            set_shape_text(placeholder, body_example(spec, find_slot_by_idx(spec, placeholder.placeholder_format.idx), ordinal))
        return

    if spec.number in (17, 18, 20, 21):
        for ordinal, placeholder in enumerate(object_placeholders):
            set_shape_text(
                placeholder,
                "Main content\nUse only supplied text, data, tables, images or diagrams",
                panel=True,
            )
        return

    # Leave the underlying object placeholder available for replacement, while
    # the example uses editable PowerPoint primitives rather than a screenshot.
    if spec.number == 68:
        tiles = (
            (0.9, 1.65, 5.1, 3.25, "Product A\n42", BLUE),
            (6.05, 1.65, 3.0, 3.25, "Product B\n24", NAVY),
            (9.1, 1.65, 3.1, 1.55, "Product C\n18", GREEN),
            (9.1, 3.25, 3.1, 1.65, "Product D\n16", AMBER),
            (0.9, 4.95, 5.1, 1.35, "Segment 1", MID),
            (6.05, 4.95, 6.15, 1.35, "Segment 2", MID),
        )
        for left, top, width, height, text, color in tiles:
            add_shape_text(slide, left, top, width, height, text, fill=color, color=WHITE, size=15, bold=True)
    elif spec.number == 69:
        values = ((18, 42, 61, 78), (25, 55, 72, 84), (31, 48, 67, 73), (44, 59, 76, 91))
        colors = (PALE, MID, BLUE, NAVY)
        for row in range(4):
            for col in range(4):
                value = values[row][col]
                fill = colors[min(value // 25, 3)]
                text_color = WHITE if fill in (BLUE, NAVY) else NAVY
                add_shape_text(slide, 2.25 + col * 2.05, 1.65 + row * 1.15, 1.9, 1.0, str(value), fill=fill, color=text_color, bold=True)
        for row, label in enumerate(("North", "Central", "South", "West")):
            add_shape_text(slide, 0.8, 1.65 + row * 1.15, 1.25, 1.0, label, fill=WHITE, size=12)
        for col, label in enumerate(("Q1", "Q2", "Q3", "Q4")):
            add_shape_text(slide, 2.25 + col * 2.05, 6.35, 1.9, 0.4, label, fill=WHITE, size=11)
    elif spec.number == 70:
        regions = (
            (1.0, 2.0, 2.7, 1.5, "North\n42", BLUE),
            (3.8, 2.55, 2.4, 1.7, "West\n31", NAVY),
            (6.35, 1.75, 2.9, 1.65, "Central\n36", GREEN),
            (9.45, 2.65, 2.6, 1.5, "East\n28", AMBER),
            (5.3, 4.55, 3.2, 1.45, "South\n25", MID),
        )
        for left, top, width, height, text, color in regions:
            add_shape_text(slide, left, top, width, height, text, fill=color, color=WHITE, bold=True)


def fill_media_example(slide, spec: LayoutSpec) -> None:
    if spec.number != 48:
        return
    placeholder = next(
        ph for ph in slide.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.MEDIA_CLIP
    )
    left = placeholder.left / EMU_PER_INCH
    top = placeholder.top / EMU_PER_INCH
    width = placeholder.width / EMU_PER_INCH
    height = placeholder.height / EMU_PER_INCH
    placeholder._element.getparent().remove(placeholder._element)
    add_shape_text(
        slide,
        left,
        top,
        width,
        height,
        "▶\nVideo or live demo",
        fill=NAVY,
        color=WHITE,
        size=22,
        bold=True,
    )


def populate_examples(template_path: Path, image_path: Path, output: Path, specs: tuple[LayoutSpec, ...]) -> None:
    prs = Presentation(str(template_path))
    if len(prs.slide_layouts) != len(specs):
        raise RuntimeError(f"Expected {len(specs)} layouts, found {len(prs.slide_layouts)}")
    prs.core_properties.title = "Plain fallback PowerPoint layouts"
    prs.core_properties.subject = "100 opt-in corporate slide layout examples"
    prs.core_properties.author = "brand-faithful-pptx"
    prs.core_properties.comments = "Fallback layouts require explicit user approval before use."

    for index, spec in enumerate(specs):
        layout = prs.slide_layouts[index]
        if layout.name != spec.name:
            raise RuntimeError(f"Layout {index + 1} is {layout.name!r}, expected {spec.name!r}")
        slide = prs.slides.add_slide(layout)
        populate_text_placeholders(slide, spec)
        insert_pictures(slide, spec, image_path)
        insert_chart(slide, spec)
        insert_table(slide, spec)
        fill_object_example(slide, spec)
        fill_media_example(slide, spec)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))


def related_part(source: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def patch_combination_chart(path: Path) -> None:
    """Turn slide 63's second column series into a line series."""

    with zipfile.ZipFile(path) as zf:
        parts = {name: zf.read(name) for name in zf.namelist()}

    slide_rels_name = "ppt/slides/_rels/slide63.xml.rels"
    rels = etree.fromstring(parts[slide_rels_name])
    chart_target = None
    chart_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    for rel in rels:
        if rel.get("Type") == chart_rel_type:
            chart_target = related_part("ppt/slides/slide63.xml", rel.get("Target"))
            break
    if not chart_target:
        raise RuntimeError("slide 63 has no chart relationship")

    chart = etree.fromstring(parts[chart_target])
    plot_area = chart.find(".//c:plotArea", {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"})
    chart_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    if plot_area is None:
        raise RuntimeError("combination chart has no plot area")
    bar = plot_area.find(qname(chart_ns, "barChart"))
    if bar is None:
        raise RuntimeError("combination chart seed is not a bar/column chart")
    series = bar.findall(qname(chart_ns, "ser"))
    if len(series) < 2:
        raise RuntimeError("combination chart seed needs two series")
    line_series = series[-1]
    bar.remove(line_series)

    line_chart = etree.Element(qname(chart_ns, "lineChart"))
    etree.SubElement(line_chart, qname(chart_ns, "grouping"), val="standard")
    etree.SubElement(line_chart, qname(chart_ns, "varyColors"), val="0")
    line_chart.append(line_series)
    category_ref = line_series.find(qname(chart_ns, "cat"))
    marker = etree.Element(qname(chart_ns, "marker"))
    if category_ref is None:
        line_series.append(marker)
    else:
        line_series.insert(line_series.index(category_ref), marker)
    etree.SubElement(line_series, qname(chart_ns, "smooth"), val="0")
    for ax_id in bar.findall(qname(chart_ns, "axId")):
        line_chart.append(copy.deepcopy(ax_id))
    bar.addnext(line_chart)

    parts[chart_target] = etree.tostring(chart, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(parts):
            zf.writestr(name, parts[name])


def build(output: Path) -> None:
    specs = layout_specs()
    with tempfile.TemporaryDirectory(prefix="plain-fallback-pptx-") as tmp:
        tmp_dir = Path(tmp)
        seed = tmp_dir / "seed.pptx"
        custom_template = tmp_dir / "custom-template.pptx"
        populated = tmp_dir / "populated.pptx"
        image_path = tmp_dir / "placeholder.png"

        prs = Presentation()
        prs.slide_width = Inches(SLIDE_WIDTH_IN)
        prs.slide_height = Inches(SLIDE_HEIGHT_IN)
        prs.save(str(seed))
        replace_layout_parts(seed, custom_template, specs)
        create_placeholder_image(image_path)
        populate_examples(custom_template, image_path, populated, specs)
        patch_combination_chart(populated)

        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(populated.read_bytes())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("references/plain-fallback-100-layouts.pptx"),
        help="Output .pptx path",
    )
    args = parser.parse_args()
    build(args.output)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
