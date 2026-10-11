#!/usr/bin/env python3
"""Build the opt-in plain fallback deck: 28 reusable layouts, 100 example slides.

Design contract
---------------
The deck exists so that individual example slides can be pasted into a
corporate presentation with PowerPoint's *Use Destination Theme* option and
come out looking like the destination brand.  That only works when every
visual decision is inherited rather than baked into the slide:

* The nine core layouts carry PowerPoint's standard names and ``type``
  attributes (``Title Slide``/``title``, ``Title and Content``/``obj`` ...), so a
  pasted slide maps onto the destination template's equivalent layout.
* Layouts contain placeholders only: no rules, bars, panels or images.
* Layout and master text styles use theme fonts (``+mj-lt``/``+mn-lt``) and
  theme colours (``schemeClr``).  No RGB values or typeface names appear below
  the theme.
* Slides contain content only.  Placeholder text carries no size, colour or
  font overrides; it inherits from the layout and master.
* Charts are native, editable charts that take their series colours from the
  theme accents.  Diagram examples use native shapes coloured with theme
  colours, so they recolour with the destination theme.
* Picture placeholders are left empty; no stock imagery is embedded.

The generator uses a two-stage workflow because python-pptx cannot create
slide masters or layouts:

1. Create one master and inject the layout OOXML parts.
2. Reopen that package with python-pptx and add one example slide for each of
   the 100 catalogue entries requested in issue #3, each on its best-fit
   layout.
"""

from __future__ import annotations

import argparse
import json
import math
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from lxml import etree
from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P_NS, "a": A_NS, "r": R_NS, "c": C_NS}
NSMAP = {"p": P_NS, "a": A_NS, "r": R_NS}

SLIDE_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
SLIDE_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
SLIDE_LAYOUT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"

DEFAULT_OUTPUT = Path("references/plain-fallback-layouts.pptx")
DEFAULT_ROLES_OUTPUT = Path("references/plain-fallback-placeholder-roles.json")
DEFAULT_CATALOG_OUTPUT = Path("references/plain-fallback-catalog.json")

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
EMU_PER_INCH = 914400

# Content grid (inches).
MARGIN = 0.6
CONTENT_W = SLIDE_WIDTH_IN - 2 * MARGIN  # 12.133
CONTENT_TOP = 1.55
CONTENT_BOTTOM = 6.65
CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP
GUTTER = 0.3

# Theme palette.  These values live only in the theme part; everything else
# refers to them through schemeClr so that a destination theme replaces them.
THEME_COLORS = {
    "dk1": "1B2A3A",
    "lt1": "FFFFFF",
    "dk2": "44546A",
    "lt2": "EEF2F5",
    "accent1": "2F6FD0",
    "accent2": "1B2A3A",
    "accent3": "2E8B57",
    "accent4": "C26A1E",
    "accent5": "8A3FA0",
    "accent6": "5B6B7C",
    "hlink": "2F6FD0",
    "folHlink": "8A3FA0",
}
THEME_MAJOR_FONT = "Aptos Display"
THEME_MINOR_FONT = "Aptos"


# ---------------------------------------------------------------------------
# Layout definitions
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Slot:
    """A genuine placeholder on a slide layout."""

    name: str
    ph_type: str  # title, ctrTitle, subTitle, body, obj, pic
    idx: int | None
    left: float
    top: float
    width: float
    height: float
    role: str = "body_content"
    content_class: str = "core_content"
    note: str | None = None
    lst_style: str = ""  # inner XML for <a:lstStyle>
    anchor: str | None = None
    prompt: str | None = None


@dataclass(frozen=True)
class LayoutSpec:
    name: str
    ooxml_type: str
    slots: tuple[Slot, ...]
    standard: bool = False
    footers: bool = True


def _lvl(level: int, *, size: int | None = None, bold: bool | None = None, align: str | None = None,
         color: str | None = None, bullet_none: bool = False) -> str:
    attrs = []
    if align:
        attrs.append(f'algn="{align}"')
    rpr_attrs = []
    if size:
        rpr_attrs.append(f'sz="{size}"')
    if bold is not None:
        rpr_attrs.append(f'b="{1 if bold else 0}"')
    fill = f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>' if color else ""
    bullet = "<a:buNone/>" if bullet_none else ""
    return (
        f'<a:lvl{level}pPr {" ".join(attrs)}>{bullet}'
        f'<a:defRPr {" ".join(rpr_attrs)}>{fill}</a:defRPr></a:lvl{level}pPr>'
    )


def title(left=MARGIN, top=0.35, width=CONTENT_W, height=1.0, *, lst="", anchor=None) -> Slot:
    return Slot("Title", "title", None, left, top, width, height, "title", "core_content",
                lst_style=lst, anchor=anchor, prompt="Click to add title")


def ctr_title(left, top, width, height, *, lst="", anchor="b") -> Slot:
    return Slot("Title", "ctrTitle", None, left, top, width, height, "title", "core_content",
                lst_style=lst, anchor=anchor, prompt="Click to add title")


def subtitle(idx, left, top, width, height, *, lst="") -> Slot:
    return Slot("Subtitle", "subTitle", idx, left, top, width, height, "subtitle", "optional_supporting",
                "Use only for a supplied subtitle, presenter, date or other functional label.",
                lst_style=lst, prompt="Click to add subtitle")


def content(idx, left, top, width, height, *, name="Content Placeholder", note=None) -> Slot:
    return Slot(name, "obj", idx, left, top, width, height, "body_content", "core_content",
                note or "Primary region for supplied text, chart, table, picture or media.",
                prompt="Click to add text")


def text(idx, left, top, width, height, *, name="Text Placeholder", role="body_content",
         content_class="core_content", note=None, lst="", anchor=None) -> Slot:
    return Slot(name, "body", idx, left, top, width, height, role, content_class, note,
                lst_style=lst, anchor=anchor, prompt="Click to add text")


def picture(idx, left, top, width, height, *, name="Picture Placeholder") -> Slot:
    return Slot(name, "pic", idx, left, top, width, height, "non_text", "non_text",
                "Use only for a supplied image; leave empty otherwise.", prompt="Click icon to add picture")


def columns(count: int, top: float = CONTENT_TOP, height: float = CONTENT_H, gap: float = GUTTER):
    width = (CONTENT_W - gap * (count - 1)) / count
    return [(MARGIN + i * (width + gap), top, width, height) for i in range(count)]


CAPTION_LST = (_lvl(1, size=1400, bullet_none=True)
               + '<a:lvl2pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1200"/></a:lvl2pPr>')
STAT_LST = (_lvl(1, size=5400, bold=True, bullet_none=True)
            + '<a:lvl2pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1600"/></a:lvl2pPr>')
HEADING_LST = _lvl(1, size=2000, bold=True, bullet_none=True)


def _side_caption_layout(name: str, ooxml_type: str, *, picture_left: bool, picture_slot: bool,
                         standard: bool) -> LayoutSpec:
    narrow_w, wide_w = 4.0, CONTENT_W - 4.0 - GUTTER
    narrow_x = MARGIN + wide_w + GUTTER if picture_left else MARGIN
    wide_x = MARGIN if picture_left else MARGIN + narrow_w + GUTTER
    main = (picture(1, wide_x, 0.45, wide_w, 6.2) if picture_slot
            else content(1, wide_x, 0.45, wide_w, 6.2))
    return LayoutSpec(
        name,
        ooxml_type,
        (
            title(narrow_x, 0.45, narrow_w, 1.5, lst=_lvl(1, size=2400), anchor="b"),
            main,
            text(2, narrow_x, 2.15, narrow_w, 4.5, name="Text Placeholder", role="caption",
                 content_class="optional_supporting",
                 note="Use only for a supplied caption, explanation or annotation.",
                 lst=CAPTION_LST + _lvl(1, size=1600)),
        ),
        standard=standard,
    )


FOOTER_IDX = frozenset({10, 11, 12})


def content_idx(n: int) -> int:
    """Return the nth content idx (1-based), skipping the footer idx values 10-12."""

    return n if n < 10 else n + len(FOOTER_IDX)


def _pictures_with_captions(name: str, count: int) -> LayoutSpec:
    slots: list[Slot] = [title()]
    if count in (2, 4):
        for i, (x, _, w, _) in enumerate(columns(count, gap=GUTTER if count == 2 else 0.25)):
            pic_h = 3.8 if count == 2 else 3.3
            slots.append(picture(content_idx(1 + 2 * i), x, CONTENT_TOP, w, pic_h))
            slots.append(text(content_idx(2 + 2 * i), x, CONTENT_TOP + pic_h + 0.15, w, CONTENT_H - pic_h - 0.15,
                              name="Caption", role="caption", content_class="optional_supporting",
                              note="Use only for supplied identification, credit or factual context.",
                              lst=CAPTION_LST))
    elif count == 6:
        cells = columns(3)
        row_h = 2.3
        for row in range(2):
            for col, (x, _, w, _) in enumerate(cells):
                i = row * 3 + col
                y = CONTENT_TOP + row * (row_h + 0.45)
                slots.append(picture(content_idx(1 + 2 * i), x, y, 1.7, row_h))
                slots.append(text(content_idx(2 + 2 * i), x + 1.85, y, w - 1.85, row_h, name="Caption", role="caption",
                                  content_class="optional_supporting",
                                  note="Use only for supplied identification, credit or factual context.",
                                  lst=CAPTION_LST))
    return LayoutSpec(name, "cust", tuple(slots))


def _statistics(name: str, count: int) -> LayoutSpec:
    slots: list[Slot] = [title()]
    lst = STAT_LST if count == 3 else STAT_LST.replace('sz="5400"', 'sz="4400"')
    for i, (x, y, w, h) in enumerate(columns(count)):
        slots.append(text(1 + i, x, y, w, h, name="Statistic", role="body_content",
                          note="First paragraph: supplied figure. Second level: its label or unit.",
                          lst=lst, anchor="ctr"))
    return LayoutSpec(name, "cust", tuple(slots))


def build_layouts() -> tuple[LayoutSpec, ...]:
    full = (0.0, 0.0, SLIDE_WIDTH_IN, SLIDE_HEIGHT_IN)
    two = columns(2)
    narrow_left_w = (CONTENT_W - GUTTER) / 3
    wide_w = CONTENT_W - GUTTER - narrow_left_w
    grid_h = (CONTENT_H - GUTTER) / 2
    return (
        # --- PowerPoint standard layouts (names and types match Office) -----
        LayoutSpec("Title Slide", "title", (
            ctr_title(MARGIN, 1.9, 10.5, 1.9, lst=_lvl(1, size=4400)),
            subtitle(1, MARGIN, 4.0, 10.5, 1.4, lst=_lvl(1, size=2000, bullet_none=True)),
        ), standard=True),
        LayoutSpec("Title and Content", "obj", (
            title(), content(1, MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H),
        ), standard=True),
        LayoutSpec("Section Header", "secHead", (
            title(MARGIN, 2.3, 10.5, 1.7, lst=_lvl(1, size=4000), anchor="b"),
            text(1, MARGIN, 4.15, 10.5, 1.2, name="Text Placeholder", role="subtitle",
                 content_class="optional_supporting",
                 note="Use only for a supplied section number or short description.",
                 lst=_lvl(1, size=2000)),
        ), standard=True),
        LayoutSpec("Two Content", "twoObj", (
            title(), content(1, *two[0]), content(2, *two[1]),
        ), standard=True),
        LayoutSpec("Comparison", "twoTxTwoObj", (
            title(),
            text(1, two[0][0], CONTENT_TOP, two[0][2], 0.6, name="Heading", lst=HEADING_LST, anchor="b"),
            content(2, two[0][0], CONTENT_TOP + 0.75, two[0][2], CONTENT_H - 0.75),
            text(3, two[1][0], CONTENT_TOP, two[1][2], 0.6, name="Heading", lst=HEADING_LST, anchor="b"),
            content(4, two[1][0], CONTENT_TOP + 0.75, two[1][2], CONTENT_H - 0.75),
        ), standard=True),
        LayoutSpec("Title Only", "titleOnly", (title(),), standard=True),
        LayoutSpec("Blank", "blank", (), standard=True),
        _side_caption_layout("Content with Caption", "objTx", picture_left=False, picture_slot=False,
                             standard=True),
        _side_caption_layout("Picture with Caption", "picTx", picture_left=False, picture_slot=True,
                             standard=True),
        # --- Custom layouts: only where the structure differs materially ----
        LayoutSpec("Title Slide with Picture", "cust", (
            picture(1, 0.0, 0.0, SLIDE_WIDTH_IN, 4.6),
            ctr_title(MARGIN, 4.8, 10.5, 1.2, lst=_lvl(1, size=4000)),
            subtitle(2, MARGIN, 6.0, 10.5, 0.7, lst=_lvl(1, size=2000, bullet_none=True)),
        )),
        LayoutSpec("Title Slide with Side Picture", "cust", (
            picture(1, 0.0, 0.0, 6.4, SLIDE_HEIGHT_IN),
            ctr_title(7.0, 1.9, 5.73, 1.9, lst=_lvl(1, size=4000)),
            subtitle(2, 7.0, 4.0, 5.73, 1.4, lst=_lvl(1, size=2000, bullet_none=True)),
        )),
        LayoutSpec("Title and Full Picture", "cust", (
            picture(1, 0.0, 0.0, SLIDE_WIDTH_IN, 5.4),
            title(MARGIN, 5.6, CONTENT_W, 1.1, lst=_lvl(1, size=3600), anchor="t"),
        )),
        LayoutSpec("Statement", "cust", (
            title(1.4, 1.6, SLIDE_WIDTH_IN - 2.8, 3.6, lst=_lvl(1, size=4000, align="ctr"), anchor="ctr"),
        )),
        LayoutSpec("Quote", "cust", (
            text(1, 1.4, 1.4, SLIDE_WIDTH_IN - 2.8, 3.7, name="Quote", role="quote",
                 content_class="optional_supporting", note="Use only for a supplied quotation.",
                 lst=_lvl(1, size=3200, bullet_none=True), anchor="b"),
            text(2, 1.4, 5.3, SLIDE_WIDTH_IN - 2.8, 0.9, name="Attribution", role="attribution",
                 content_class="optional_supporting",
                 note="Use only for the supplied quotation's factual attribution.",
                 lst=_lvl(1, size=1600, bullet_none=True)),
        ), footers=True),
        LayoutSpec("Quote with Picture", "cust", (
            picture(1, MARGIN, 0.6, 4.6, 6.3),
            text(2, 5.6, 1.2, 7.13, 3.9, name="Quote", role="quote", content_class="optional_supporting",
                 note="Use only for a supplied quotation.", lst=_lvl(1, size=2800, bullet_none=True),
                 anchor="b"),
            text(3, 5.6, 5.3, 7.13, 0.9, name="Attribution", role="attribution",
                 content_class="optional_supporting",
                 note="Use only for the supplied quotation's factual attribution.",
                 lst=_lvl(1, size=1600, bullet_none=True)),
        )),
        LayoutSpec("Title, Subtitle and Content", "cust", (
            title(),
            text(1, MARGIN, 1.35, CONTENT_W, 0.6, name="Subtitle", role="subtitle",
                 content_class="optional_supporting",
                 note="Use only for supplied context, such as a measure, unit or period.",
                 lst=_lvl(1, size=1800, bullet_none=True)),
            content(2, MARGIN, 2.1, CONTENT_W, CONTENT_BOTTOM - 2.1),
        )),
        LayoutSpec("Two Content, Narrow Left", "cust", (
            title(),
            content(1, MARGIN, CONTENT_TOP, narrow_left_w, CONTENT_H),
            content(2, MARGIN + narrow_left_w + GUTTER, CONTENT_TOP, wide_w, CONTENT_H),
        )),
        LayoutSpec("Two Content, Narrow Right", "cust", (
            title(),
            content(1, MARGIN, CONTENT_TOP, wide_w, CONTENT_H),
            content(2, MARGIN + wide_w + GUTTER, CONTENT_TOP, narrow_left_w, CONTENT_H),
        )),
        LayoutSpec("Three Content", "cust", (title(),) + tuple(
            content(1 + i, *box) for i, box in enumerate(columns(3)))),
        LayoutSpec("Four Content", "cust", (title(),) + tuple(
            content(1 + i, *box) for i, box in enumerate(columns(4, gap=0.25)))),
        LayoutSpec("Four Content Grid", "cust", (title(),) + tuple(
            content(1 + r * 2 + c, two[c][0], CONTENT_TOP + r * (grid_h + GUTTER), two[c][2], grid_h)
            for r in range(2) for c in range(2))),
        LayoutSpec("Statistic and Description", "cust", (
            title(),
            text(1, MARGIN, CONTENT_TOP, 4.6, CONTENT_H, name="Statistic",
                 note="First paragraph: supplied figure. Second level: its label or unit.",
                 lst=STAT_LST.replace('sz="5400"', 'sz="7200"'), anchor="t"),
            content(2, MARGIN + 4.6 + GUTTER, CONTENT_TOP, CONTENT_W - 4.6 - GUTTER, CONTENT_H),
        )),
        _statistics("Three Statistics", 3),
        _statistics("Four Statistics", 4),
        _pictures_with_captions("Two Pictures with Captions", 2),
        _pictures_with_captions("Four Pictures with Captions", 4),
        _pictures_with_captions("Six Pictures with Captions", 6),
        _side_caption_layout("Picture Left with Caption", "cust", picture_left=True, picture_slot=True,
                             standard=False),
    )


LAYOUTS = build_layouts()
LAYOUT_NAMES = tuple(spec.name for spec in LAYOUTS)
STANDARD_LAYOUT_TYPES = {spec.name: spec.ooxml_type for spec in LAYOUTS if spec.standard}


# ---------------------------------------------------------------------------
# Catalogue: the 100 slide types requested in issue #3
# ---------------------------------------------------------------------------

Body = tuple  # sequence of (text, level, bold) or plain strings


@dataclass(frozen=True)
class Entry:
    number: int
    name: str
    category: str
    layout: str
    title: str | None = None
    bodies: dict[int, tuple] = field(default_factory=dict)
    kind: str = "text"  # text, chart, table, diagram, heatmap, treemap


def E(number, name, category, layout, title=None, bodies=None, kind="text") -> Entry:
    return Entry(number, name, category, layout, title, bodies or {}, kind)


def sub(*lines: str) -> tuple:
    """Level-2 lines (detail under a heading)."""

    return tuple((line, 1, False) for line in lines)


OPEN, PEOPLE, CORE, VISUAL, DATA, PROCESS, TABLE, CLOSE = (
    "Opening and navigation",
    "People and organization",
    "Core content",
    "Visual, product and media",
    "Chart and data",
    "Process and strategy",
    "Table and planning",
    "Closing and administrative",
)

AGENDA = ("1   Context", "2   Evidence", "3   Decision", "4   Next steps")

CATALOG: tuple[Entry, ...] = (
    E(1, "Standard title slide", OPEN, "Title Slide", "Quarterly business review",
      {1: ("Operations team", "October 2026")}),
    E(2, "Hero-image title slide", OPEN, "Title Slide with Picture", "Customer conference 2026",
      {2: ("Opening keynote",)}),
    E(3, "Split-panel title slide", OPEN, "Title Slide with Side Picture", "Market entry proposal",
      {2: ("Strategy office", "October 2026")}),
    E(4, "Minimal title slide", OPEN, "Title Slide", "Annual financial results",
      {1: ("Finance · Fiscal year 2026",)}),
    E(5, "Event or webinar title slide", OPEN, "Title Slide", "Data governance webinar",
      {1: ("Live online · 14 October 2026", "Alex Morgan, Head of Data")}),
    E(6, "Agenda slide", OPEN, "Title and Content", "Agenda", {1: AGENDA}),
    E(7, "Agenda with progress indicator", OPEN, "Title and Content", "Agenda",
      {1: (AGENDA[0], (AGENDA[1], 0, True), AGENDA[2], AGENDA[3])}),
    E(8, "Standard section divider", OPEN, "Section Header", "Market evidence",
      {1: ("Customer research and sales data",)}),
    E(9, "Image-led section divider", OPEN, "Title and Full Picture", "Customer stories"),
    E(10, "Numbered chapter divider", OPEN, "Section Header", "Implementation plan", {1: ("Section 3",)}),
    E(11, "Single-speaker slide", PEOPLE, "Picture with Caption", "Alex Morgan",
      {2: ("Head of Data, Example Organization", "Leads the data governance programme and reporting standards.")}),
    E(12, "Two-speaker slide", PEOPLE, "Two Pictures with Captions", "Speakers",
      {2: ("Alex Morgan",) + sub("Head of Data"), 4: ("Taylor Chen",) + sub("Chief Operating Officer")}),
    E(13, "Panel-speaker slide", PEOPLE, "Four Pictures with Captions", "Panel",
      {2: ("Alex Morgan",) + sub("Moderator"), 4: ("Taylor Chen",) + sub("Operations"),
       6: ("Jordan Singh",) + sub("Finance"), 8: ("Sam Rivera",) + sub("Product")}),
    E(14, "Team grid slide", PEOPLE, "Six Pictures with Captions", "Project team",
      {2: ("Alex Morgan",) + sub("Sponsor"), 4: ("Taylor Chen",) + sub("Programme lead"),
       6: ("Jordan Singh",) + sub("Finance"), 8: ("Sam Rivera",) + sub("Product"),
       13: ("Riley Okafor",) + sub("Engineering"), 15: ("Casey Brown",) + sub("Change management")}),
    E(15, "Executive biography slide", PEOPLE, "Picture with Caption", "Jordan Singh",
      {2: ("Chief Financial Officer", "Joined in 2019 from a regional bank.",
           "Responsible for planning, treasury and investor relations.")}),
    E(16, "Title-only slide", CORE, "Title Only", "Operating model overview"),
    E(17, "Title and content slide", CORE, "Title and Content", "Programme status",
      {1: ("Scope confirmed for all four regions", "Spend within 2% of plan",
           "Two risks escalated to the steering group")}),
    E(18, "Title, subtitle and content slide", CORE, "Title, Subtitle and Content",
      "Support tickets fell after the self-service launch",
      {1: ("Monthly tickets, January to September 2026",),
       2: ("Self-service launched in April", "Tickets fell from 4,200 to 3,450 per month by September")}),
    E(19, "Title and two columns", CORE, "Two Content", "Programme scope",
      {1: ("In scope",) + sub("Customer onboarding", "Billing changes", "Reporting"),
       2: ("Out of scope",) + sub("Pricing review", "Partner contracts")}),
    E(20, "Title with left third and right two-thirds", CORE, "Two Content, Narrow Left", "Service coverage",
      {1: ("Summary",) + sub("Four regions covered", "Two pending"),
       2: ("North: full coverage since March", "Central: full coverage since May",
           "South and West: pilots running until December")}),
    E(21, "Title with left two-thirds and right third", CORE, "Two Content, Narrow Right", "Regional performance",
      {1: ("North exceeded plan in every quarter", "Central recovered after a weak first quarter",
           "South and West are tracking to plan"),
       2: ("Summary",) + sub("Two regions ahead of plan")}),
    E(22, "Title and three columns", CORE, "Three Content", "Programme workstreams",
      {1: ("Process",) + sub("Map and simplify approvals"), 2: ("Technology",) + sub("Replace the legacy portal"),
       3: ("People",) + sub("Train 400 staff by March")}),
    E(23, "Title and four-card grid", CORE, "Four Content", "Capabilities",
      {1: ("Planning",) + sub("Quarterly forecasts"), 2: ("Reporting",) + sub("Monthly dashboards"),
       3: ("Analytics",) + sub("Ad hoc analysis"), 4: ("Governance",) + sub("Data standards")}),
    E(24, "50:50 stat and description", CORE, "Statistic and Description", "Adoption",
      {1: ("42%",) + sub("of eligible staff use the portal weekly"),
       2: ("Up from 28% in January", "Highest in North and Central", "Target: 60% by June 2027")}),
    E(25, "Big-statement slide", CORE, "Statement", "Three decisions are needed today"),
    E(26, "Single callout-stat slide", CORE, "Statistic and Description", "Cycle time",
      {1: ("18 days",) + sub("median order-to-delivery time")}),
    E(27, "Three-stat row", CORE, "Three Statistics", "Year-to-date results",
      {1: ("42%",) + sub("adoption"), 2: ("£3.2m",) + sub("annual value"), 3: ("18 days",) + sub("cycle time")}),
    E(28, "Definition or explainer slide", CORE, "Two Content, Narrow Left", "Definition: active user",
      {1: ("Active user",),
       2: ("A person who signs in at least once in a 28-day period.",
           "Excludes service accounts and test users.")}),
    E(29, "Quotation slide", CORE, "Quote", None,
      {1: ("“The new approval process halved our turnaround time.”",),
       2: ("Name, Role, Organization",)}),
    E(30, "Quotation with portrait", CORE, "Quote with Picture", None,
      {2: ("“We now close the month in four days instead of nine.”",), 3: ("Name, Role, Organization",)}),
    E(31, "Key-takeaways slide", CORE, "Title and Content", "Key takeaways",
      {1: ("1   Adoption is ahead of plan", "2   Costs are within budget", "3   One decision is needed on scope")}),
    E(32, "Executive-summary slide", CORE, "Three Content", "Recommendation: proceed with phase two",
      {1: ("Finding",) + sub("Phase one met all exit criteria"),
       2: ("Impact",) + sub("£3.2m annual value confirmed"),
       3: ("Next step",) + sub("Approve phase two budget")}),
    E(33, "Recommendation and rationale slide", CORE, "Two Content, Narrow Left", "Recommendation",
      {1: ("Adopt option B",),
       2: ("Lowest five-year cost", "Meets all security requirements", "Can start in January")}),
    E(34, "Side-by-side comparison slide", CORE, "Comparison", "Option A versus option B",
      {1: ("Option A",), 2: ("£1.2m five-year cost", "Live in nine months"),
       3: ("Option B",), 4: ("£0.9m five-year cost", "Live in six months")}),
    E(35, "Pros and cons slide", CORE, "Comparison", "Hosted service",
      {1: ("Advantages",), 2: ("No hardware to maintain", "Automatic updates"),
       3: ("Limitations",), 4: ("Annual licence increases", "Limited customisation")}),
    E(36, "Before-and-after slide", CORE, "Two Pictures with Captions", "Checkout page",
      {2: ("Before",) + sub("Five steps"), 4: ("After",) + sub("Two steps")}),
    E(37, "Problem-and-solution slide", CORE, "Comparison", "Late invoices",
      {1: ("Problem",), 2: ("30% of invoices are paid late", "Manual approval queue"),
       3: ("Solution",), 4: ("Automated approval under £5,000", "Weekly exception report")}),
    E(38, "Challenge, approach and result slide", CORE, "Three Content", "Warehouse consolidation",
      {1: ("Challenge",) + sub("Three sites with duplicated stock"),
       2: ("Approach",) + sub("Consolidated to one site over six months"),
       3: ("Result",) + sub("Stock holding down 22%")}),
    E(39, "Customer-proof slide", CORE, "Picture with Caption", "Example Organization",
      {2: ("Reduced onboarding time from 20 days to 12.", "Customer since 2023.")}),
    E(40, "Frequently asked questions slide", CORE, "Two Content", "Frequently asked questions",
      {1: ("When does the pilot start?",) + sub("1 November 2026"),
       2: ("Who can take part?",) + sub("All staff in the North region")}),
    E(41, "Question or discussion-prompt slide", VISUAL, "Statement", "Which option should we fund?"),
    E(42, "Left image and right text", VISUAL, "Picture Left with Caption", "New service centre",
      {2: ("Opened in September 2026", "Handles all customer calls for the North region")}),
    E(43, "Left text and right image", VISUAL, "Picture with Caption", "Field equipment",
      {2: ("Standard kit issued to every engineer from January",)}),
    E(44, "Full-bleed image with text overlay", VISUAL, "Title and Full Picture", "Field operations"),
    E(45, "Image-gallery slide", VISUAL, "Four Pictures with Captions", "Site photographs"),
    E(46, "Product-screenshot slide", VISUAL, "Title and Content", "Dashboard home screen"),
    E(47, "Annotated-screenshot slide", VISUAL, "Content with Caption", "Dashboard filters",
      {2: ("1   Date range", "2   Region", "3   Export")}),
    E(48, "Video or live-demo slide", VISUAL, "Title and Content", "Product demonstration"),
    E(49, "One-series column chart", DATA, "Title and Content", "Adoption by region (%)", kind="chart"),
    E(50, "Two-series column chart", DATA, "Title and Content", "Adoption: actual versus target (%)", kind="chart"),
    E(51, "Multi-series clustered column chart", DATA, "Title and Content", "Units sold by region and channel (thousands)", kind="chart"),
    E(52, "Stacked column chart", DATA, "Title and Content", "Units sold by region and channel (thousands)", kind="chart"),
    E(53, "100% stacked column chart", DATA, "Title and Content", "Sales mix by region", kind="chart"),
    E(54, "One-series horizontal bar chart", DATA, "Title and Content", "Adoption by region (%)", kind="chart"),
    E(55, "Two-series horizontal bar chart", DATA, "Title and Content", "Adoption: actual versus target (%)", kind="chart"),
    E(56, "Stacked horizontal bar chart", DATA, "Title and Content", "Units sold by region and channel (thousands)", kind="chart"),
    E(57, "100% stacked horizontal bar chart", DATA, "Title and Content", "Sales mix by region", kind="chart"),
    E(58, "Ranked bar chart", DATA, "Title and Content", "Priority scores", kind="chart"),
    E(59, "Pie chart", DATA, "Title and Content", "Revenue by segment", kind="chart"),
    E(60, "Donut chart", DATA, "Title and Content", "Revenue by segment", kind="chart"),
    E(61, "Line chart", DATA, "Title and Content", "Monthly active users (thousands)", kind="chart"),
    E(62, "Area or stacked-area chart", DATA, "Title and Content", "Monthly active users by plan (thousands)", kind="chart"),
    E(63, "Combination chart", DATA, "Title and Content", "Monthly active users: actual and target (thousands)", kind="chart"),
    E(64, "Waterfall chart", DATA, "Title and Content", "Operating profit bridge (£m)", kind="chart"),
    E(65, "Scatter-plot slide", DATA, "Title and Content", "Price index versus satisfaction score", kind="chart"),
    E(66, "Bubble-chart slide", DATA, "Title and Content", "Markets: growth, share and size", kind="chart"),
    E(67, "Funnel-chart slide", DATA, "Title and Content", "Sales pipeline (opportunities)", kind="chart"),
    E(68, "Treemap slide", DATA, "Title Only", "Revenue by product (%)", kind="treemap"),
    E(69, "Heatmap slide", DATA, "Title and Content", "Utilisation by region and quarter (%)", kind="heatmap"),
    E(70, "Geographic-map slide", DATA, "Content with Caption", "Customers by region",
      {2: ("North: 42%", "Central: 36%", "South: 25%", "West: 31%")}),
    E(71, "KPI scorecard or dashboard", DATA, "Four Statistics", "Scorecard",
      {1: ("42%",) + sub("adoption"), 2: ("£3.2m",) + sub("annual value"),
       3: ("18 days",) + sub("cycle time"), 4: ("+12%",) + sub("year on year")}),
    E(72, "Chart with takeaway panel", DATA, "Two Content, Narrow Right", "Adoption: actual versus target (%)",
      {2: ("All four regions are above target",) + sub("Largest gap: North, +7 points")}, kind="chart"),
    E(73, "Timeline slide", PROCESS, "Title Only", "Programme timeline", kind="diagram"),
    E(74, "Roadmap slide", PROCESS, "Three Content", "Product roadmap",
      {1: ("Now",) + sub("Single sign-on", "Mobile approvals"),
       2: ("Next",) + sub("Supplier portal", "Spend analytics"),
       3: ("Later",) + sub("Forecasting", "Partner integrations")}),
    E(75, "Horizontal-process slide", PROCESS, "Title Only", "Order process", kind="diagram"),
    E(76, "Vertical-process slide", PROCESS, "Title Only", "Data request process", kind="diagram"),
    E(77, "Chevron-process slide", PROCESS, "Title Only", "Delivery phases", kind="diagram"),
    E(78, "Circular-process slide", PROCESS, "Title Only", "Continuous improvement cycle", kind="diagram"),
    E(79, "Process-funnel slide", PROCESS, "Title Only", "Recruitment funnel", kind="diagram"),
    E(80, "Swimlane-process slide", PROCESS, "Title Only", "Service request handling", kind="diagram"),
    E(81, "Workflow or decision-flow slide", PROCESS, "Title Only", "Content approval", kind="diagram"),
    E(82, "Customer-journey slide", PROCESS, "Title and Content", "Customer journey", kind="table"),
    E(83, "Lifecycle slide", PROCESS, "Title Only", "Asset lifecycle", kind="diagram"),
    E(84, "Organizational-chart slide", PROCESS, "Title Only", "Programme organization", kind="diagram"),
    E(85, "Ecosystem or network slide", PROCESS, "Title Only", "Partner ecosystem", kind="diagram"),
    E(86, "Hierarchy or pyramid slide", PROCESS, "Title Only", "Planning hierarchy", kind="diagram"),
    E(87, "Matrix slide", PROCESS, "Four Content Grid", "Initiatives by impact and effort",
      {1: ("High impact, low effort",) + sub("Automated approvals"),
       2: ("High impact, high effort",) + sub("Platform replacement"),
       3: ("Low impact, low effort",) + sub("Template refresh"),
       4: ("Low impact, high effort",) + sub("Custom integrations")}),
    E(88, "Basic table slide", TABLE, "Title and Content", "Open actions", kind="table"),
    E(89, "Comparison-table slide", TABLE, "Title and Content", "Vendor comparison", kind="table"),
    E(90, "Financial-table slide", TABLE, "Title and Content", "Operating costs (£k)", kind="table"),
    E(91, "Project-plan or Gantt slide", TABLE, "Title and Content", "Project plan", kind="table"),
    E(92, "Calendar or schedule slide", TABLE, "Title and Content", "Event schedule", kind="table"),
    E(93, "RACI slide", TABLE, "Title and Content", "Responsibilities (RACI)", kind="table"),
    E(94, "Risk-register slide", TABLE, "Title and Content", "Risk register", kind="table"),
    E(95, "Issue, action and status tracker", TABLE, "Title and Content", "Issues and actions", kind="table"),
    E(96, "Next-steps slide", CLOSE, "Title and Content", "Next steps", kind="table"),
    E(97, "Q&A slide", CLOSE, "Statement", "Questions"),
    E(98, "Thank-you and contact slide", CLOSE, "Title Slide", "Thank you",
      {1: ("Alex Morgan, Head of Data", "alex.morgan@example.com")}),
    E(99, "Sources, notes and legal slide", CLOSE, "Title and Content", "Sources and notes",
      {1: ("1   Customer survey, September 2026, 1,240 responses",
           "2   Finance ledger, fiscal year 2026 to date",
           "Figures are unaudited.")}),
    E(100, "Blank branded canvas", CLOSE, "Blank"),
)

# Backwards-compatible view used by the verifier: (number, name, category).
LAYOUT_ROWS = tuple((e.number, e.name, e.category) for e in CATALOG)
ENTRY_BY_NUMBER = {e.number: e for e in CATALOG}
NATIVE_CHART_ENTRIES = frozenset(e.number for e in CATALOG if e.kind == "chart")
DESCENDING_BAR_ENTRIES = frozenset({54, 55, 56, 57, 58, 67})


# ---------------------------------------------------------------------------
# OOXML generation: master, layouts and theme
# ---------------------------------------------------------------------------


def emu(inches: float) -> str:
    return str(round(inches * EMU_PER_INCH))


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _xml(fragment: str) -> etree._Element:
    return etree.fromstring(
        f'<root xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}">{fragment}</root>'
    )[0]


def _sp_xml(shape_id: int, name: str, ph_attrs: str, box: tuple[float, float, float, float],
            *, lst: str = "", anchor: str | None = None, prompt: str | None = None) -> str:
    left, top, width, height = box
    anchor_attr = f' anchor="{anchor}"' if anchor else ""
    para = (f'<a:p><a:r><a:rPr lang="en-GB"/><a:t>{prompt}</a:t></a:r></a:p>' if prompt
            else '<a:p><a:endParaRPr lang="en-GB"/></a:p>')
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{emu(left)}" y="{emu(top)}"/><a:ext cx="{emu(width)}" cy="{emu(height)}"/></a:xfrm></p:spPr>'
        f'<p:txBody><a:bodyPr{anchor_attr}/><a:lstStyle>{lst}</a:lstStyle>{para}</p:txBody></p:sp>'
    )


FOOTER_BOXES = {
    "dt": (MARGIN, 6.95, 2.8, 0.35),
    "ftr": (4.4, 6.95, SLIDE_WIDTH_IN - 8.8, 0.35),
    "sldNum": (SLIDE_WIDTH_IN - MARGIN - 2.8, 6.95, 2.8, 0.35),
}


def _footer_shapes(start_id: int, *, layout: bool) -> str:
    idx = {"dt": 10, "ftr": 11, "sldNum": 12} if layout else {"dt": 2, "ftr": 3, "sldNum": 4}
    names = {"dt": "Date Placeholder", "ftr": "Footer Placeholder", "sldNum": "Slide Number Placeholder"}
    out = []
    for offset, kind in enumerate(("dt", "ftr", "sldNum")):
        size = "quarter" if kind == "ftr" else "half"
        out.append(_sp_xml(start_id + offset, names[kind], f'type="{kind}" sz="{size}" idx="{idx[kind]}"',
                           FOOTER_BOXES[kind], prompt=None))
    return "".join(out)


def _group_header() -> str:
    return (
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    )


def _ph_attrs(slot_spec: Slot) -> str:
    if slot_spec.ph_type in ("title", "ctrTitle"):
        # PowerPoint writes title placeholders without idx.  idx="0" on a
        # layout title breaks inheritance in LibreOffice and other consumers.
        return f'type="{slot_spec.ph_type}"'
    if slot_spec.ph_type == "obj":
        return f'idx="{slot_spec.idx}"'
    return f'type="{slot_spec.ph_type}" idx="{slot_spec.idx}"'


def build_layout_xml(spec: LayoutSpec) -> bytes:
    shapes = []
    for shape_id, slot_spec in enumerate(spec.slots, start=2):
        shapes.append(_sp_xml(
            shape_id, f"{slot_spec.name} {shape_id - 1}", _ph_attrs(slot_spec),
            (slot_spec.left, slot_spec.top, slot_spec.width, slot_spec.height),
            lst=slot_spec.lst_style, anchor=slot_spec.anchor, prompt=slot_spec.prompt,
        ))
    if spec.footers:
        shapes.append(_footer_shapes(len(spec.slots) + 2, layout=True))
    preserve = "" if spec.standard else ' preserve="1"'
    xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sldLayout xmlns:a="{A_NS}" xmlns:p="{P_NS}" xmlns:r="{R_NS}" type="{spec.ooxml_type}"{preserve}>'
        f'<p:cSld name="{spec.name}"><p:spTree>{_group_header()}{"".join(shapes)}</p:spTree></p:cSld>'
        f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    )
    return etree.tostring(etree.fromstring(xml.encode()), xml_declaration=True, encoding="UTF-8", standalone=True)


def build_layout_rels() -> bytes:
    root = etree.Element(qname(REL_NS, "Relationships"), nsmap={None: REL_NS})
    etree.SubElement(root, qname(REL_NS, "Relationship"), Id="rId1", Type=SLIDE_MASTER_REL,
                     Target="../slideMasters/slideMaster1.xml")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _font_refs(kind: str) -> str:
    return (f'<a:latin typeface="+{kind}-lt"/><a:ea typeface="+{kind}-ea"/>'
            f'<a:cs typeface="+{kind}-cs"/>')


def _body_level(level: int, size: int, *, bullet: str | None, mar_l: int, indent: int, space_before: int) -> str:
    bullet_xml = (f'<a:buFont typeface="Arial"/><a:buChar char="{bullet}"/>' if bullet else "<a:buNone/>")
    return (
        f'<a:lvl{level}pPr marL="{mar_l}" indent="{indent}" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        f'<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
        f'{bullet_xml}<a:defRPr sz="{size}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        f'{_font_refs("mn")}</a:defRPr></a:lvl{level}pPr>'
    )


def master_tx_styles() -> str:
    title_style = (
        '<p:titleStyle><a:lvl1pPr algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
        '<a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        f'{_font_refs("mj")}</a:defRPr></a:lvl1pPr></p:titleStyle>'
    )
    levels = [
        _body_level(1, 2000, bullet=None, mar_l=0, indent=0, space_before=1000),
        _body_level(2, 1800, bullet="•", mar_l=285750, indent=-285750, space_before=500),
        _body_level(3, 1600, bullet="–", mar_l=571500, indent=-285750, space_before=500),
        _body_level(4, 1400, bullet="•", mar_l=857250, indent=-285750, space_before=500),
        _body_level(5, 1400, bullet="–", mar_l=1143000, indent=-285750, space_before=500),
    ]
    body_style = f'<p:bodyStyle>{"".join(levels)}</p:bodyStyle>'
    other_style = (
        '<p:otherStyle><a:defPPr><a:defRPr lang="en-GB"/></a:defPPr>'
        '<a:lvl1pPr marL="0" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        f'{_font_refs("mn")}</a:defRPr></a:lvl1pPr></p:otherStyle>'
    )
    return f"<p:txStyles>{title_style}{body_style}{other_style}</p:txStyles>"


def master_shapes() -> str:
    footer_lst = ('<a:lvl1pPr><a:defRPr sz="1000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                  '</a:defRPr></a:lvl1pPr>')
    shapes = [
        _sp_xml(2, "Title Placeholder 1", 'type="title"', (MARGIN, 0.35, CONTENT_W, 1.0), anchor="b",
                prompt="Click to edit Master title style"),
        _sp_xml(3, "Text Placeholder 2", 'type="body" idx="1"', (MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H),
                prompt="Click to edit Master text styles"),
    ]
    names = {"dt": "Date Placeholder", "ftr": "Footer Placeholder", "sldNum": "Slide Number Placeholder"}
    for offset, (kind, idx) in enumerate((("dt", 2), ("ftr", 3), ("sldNum", 4)), start=4):
        size = "quarter" if kind == "ftr" else "half"
        align = {"dt": "l", "ftr": "ctr", "sldNum": "r"}[kind]
        lst = footer_lst.replace("<a:lvl1pPr>", f'<a:lvl1pPr algn="{align}">')
        if kind == "sldNum":
            para = ('<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
                    '<a:rPr lang="en-GB"/><a:t>‹#›</a:t></a:fld></a:p>')
        else:
            para = '<a:p><a:endParaRPr lang="en-GB"/></a:p>'
        left, top, width, height = FOOTER_BOXES[kind]
        shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{offset}" name="{names[kind]} {offset - 1}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="{kind}" sz="{size}" idx="{idx}"/></p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(left)}" y="{emu(top)}"/><a:ext cx="{emu(width)}" cy="{emu(height)}"/></a:xfrm></p:spPr>'
            f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle>{lst}</a:lstStyle>{para}</p:txBody></p:sp>'
        )
    return "".join(shapes)


def rebuild_master(master: etree._Element) -> None:
    sp_tree = master.find("p:cSld/p:spTree", NS)
    for child in list(sp_tree):
        sp_tree.remove(child)
    for element in etree.fromstring(
        f'<p:spTree xmlns:a="{A_NS}" xmlns:p="{P_NS}">{_group_header()}{master_shapes()}</p:spTree>'
    ):
        sp_tree.append(element)
    background = master.find("p:cSld/p:bg", NS)
    if background is None:
        c_sld = master.find("p:cSld", NS)
        background = _xml('<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>')
        c_sld.insert(0, background)
    old_styles = master.find("p:txStyles", NS)
    new_styles = etree.fromstring(f'<p:txStyles xmlns:a="{A_NS}" xmlns:p="{P_NS}">'
                                  f'{master_tx_styles()[len("<p:txStyles>"):]}')
    if old_styles is not None:
        old_styles.addprevious(new_styles)
        master.remove(old_styles)
    else:
        master.append(new_styles)


def patch_theme(parts: dict[str, bytes]) -> None:
    theme_name = "ppt/theme/theme1.xml"
    root = etree.fromstring(parts[theme_name])
    root.set("name", "Plain fallback")
    font_scheme = root.find(".//a:themeElements/a:fontScheme", NS)
    font_scheme.set("name", "Plain fallback")
    font_scheme.find("a:majorFont/a:latin", NS).set("typeface", THEME_MAJOR_FONT)
    font_scheme.find("a:minorFont/a:latin", NS).set("typeface", THEME_MINOR_FONT)
    clr_scheme = root.find(".//a:themeElements/a:clrScheme", NS)
    clr_scheme.set("name", "Plain fallback")
    for child in clr_scheme:
        local = etree.QName(child).localname
        if local in THEME_COLORS:
            for existing in list(child):
                child.remove(existing)
            etree.SubElement(child, qname(A_NS, "srgbClr"), val=THEME_COLORS[local])
    parts[theme_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_template(seed: Path, output: Path) -> None:
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
    for child in list(layout_list):
        layout_list.remove(child)
    for rel in list(master_rels):
        if rel.get("Type") == SLIDE_LAYOUT_REL:
            master_rels.remove(rel)
    for override in list(content_types):
        if override.get("PartName", "").startswith("/ppt/slideLayouts/"):
            content_types.remove(override)

    rebuild_master(master)

    for index, spec in enumerate(LAYOUTS, start=1):
        rel_id = f"rIdLayout{index}"
        etree.SubElement(master_rels, qname(REL_NS, "Relationship"), Id=rel_id, Type=SLIDE_LAYOUT_REL,
                         Target=f"../slideLayouts/slideLayout{index}.xml")
        layout_id = etree.SubElement(layout_list, qname(P_NS, "sldLayoutId"), id=str(2147483648 + index))
        layout_id.set(qname(R_NS, "id"), rel_id)
        etree.SubElement(content_types, qname(CT_NS, "Override"),
                         PartName=f"/ppt/slideLayouts/slideLayout{index}.xml",
                         ContentType=SLIDE_LAYOUT_CONTENT_TYPE)
        parts[f"ppt/slideLayouts/slideLayout{index}.xml"] = build_layout_xml(spec)
        parts[f"ppt/slideLayouts/_rels/slideLayout{index}.xml.rels"] = build_layout_rels()

    parts[master_name] = etree.tostring(master, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts[master_rels_name] = etree.tostring(master_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    parts["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8",
                                                  standalone=True)
    patch_theme(parts)

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(parts):
            zf.writestr(name, parts[name])


# ---------------------------------------------------------------------------
# Example slides
# ---------------------------------------------------------------------------


def _normalise_line(line) -> tuple[str, int, bool]:
    if isinstance(line, str):
        return line, 0, False
    text_value, level, *rest = line
    return text_value, level, bool(rest and rest[0])


def fill_text(shape, lines: tuple) -> None:
    """Write text without any size, colour or font overrides."""

    text_frame = shape.text_frame
    text_frame.clear()
    for index, raw in enumerate(lines):
        value, level, bold = _normalise_line(raw)
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.level = level
        run = paragraph.add_run()
        run.text = value
        if bold:
            run.font.bold = True


def placeholder_by_idx(slide, idx: int):
    for placeholder in slide.placeholders:
        if placeholder.placeholder_format.idx == idx:
            return placeholder
    raise KeyError(f"placeholder idx {idx} not on slide")


def adopt_into_placeholder(slide, graphic_frame, placeholder) -> None:
    """Make *graphic_frame* occupy *placeholder* the way PowerPoint does.

    PowerPoint stores a chart or table inserted into a content placeholder as a
    graphicFrame whose ``nvPr`` carries the placeholder's ``<p:ph idx>``.  That
    keeps the object bound to the layout region, so "Use Destination Theme"
    maps it into the destination layout's content placeholder.
    """

    ph = placeholder._element.find(".//p:nvPr/p:ph", NS)
    nv_pr = graphic_frame._element.find("p:nvGraphicFramePr/p:nvPr", NS)
    nv_pr.insert(0, etree.fromstring(etree.tostring(ph)))
    placeholder._element.addprevious(graphic_frame._element)
    placeholder._element.getparent().remove(placeholder._element)


# --- charts -----------------------------------------------------------------

REGIONS = ("North", "Central", "South", "West")
CHANNELS = ("Online", "Retail", "Partner")
REGION_VALUES = ((42, 36, 31, 25), (35, 33, 28, 22), (24, 27, 19, 17))
MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun")


def _category(series_names, values, categories=REGIONS, number_format="0") -> CategoryChartData:
    data = CategoryChartData(number_format=number_format)
    data.categories = list(categories)
    for name, row in zip(series_names, values):
        data.add_series(name, row)
    return data


WATERFALL_START = 80
WATERFALL_STEPS = (("Volume", 15), ("Price", 10), ("Costs", -8))
WATERFALL_CATEGORIES = ("2025 profit",) + tuple(name for name, _ in WATERFALL_STEPS) + ("2026 profit",)
# Theme-linked fills, so the waterfall recolours with the destination theme:
# totals use a neutral grey (50% shade of the background colour), increases accent 1,
# decreases accent 2.
WATERFALL_FILLS = {
    "Total": ("bg1", '<a:lumMod val="50000"/>'),
    "Increase": ("accent1", ""),
    "Decrease": ("accent2", ""),
}


def waterfall_series(start: float, steps) -> tuple[tuple[str, tuple[float, ...]], ...]:
    """Stacked-column waterfall: an invisible base plus Total, Increase and Decrease series.

    Every value is the real magnitude from the source data, so the chart's
    native labels show true figures and the embedded workbook stays editable.
    """

    base, total, increase, decrease = [0.0], [float(start)], [0.0], [0.0]
    running = float(start)
    for _, change in steps:
        after = running + change
        base.append(min(running, after))
        total.append(0.0)
        increase.append(float(change) if change > 0 else 0.0)
        decrease.append(float(-change) if change < 0 else 0.0)
        running = after
    base.append(0.0)
    total.append(running)
    increase.append(0.0)
    decrease.append(0.0)
    return (("Base", tuple(base)), ("Total", tuple(total)), ("Increase", tuple(increase)),
            ("Decrease", tuple(decrease)))


def chart_definition(number: int):
    actual_target = ("Actual", "Target")
    if number in (49, 54):
        return (XL_CHART_TYPE.COLUMN_CLUSTERED if number == 49 else XL_CHART_TYPE.BAR_CLUSTERED,
                _category(("Adoption",), REGION_VALUES[:1]))
    if number in (50, 55, 72):
        chart_type = XL_CHART_TYPE.BAR_CLUSTERED if number == 55 else XL_CHART_TYPE.COLUMN_CLUSTERED
        return chart_type, _category(actual_target, REGION_VALUES[:2])
    if number == 51:
        return XL_CHART_TYPE.COLUMN_CLUSTERED, _category(CHANNELS, REGION_VALUES)
    if number in (52, 56):
        return (XL_CHART_TYPE.COLUMN_STACKED if number == 52 else XL_CHART_TYPE.BAR_STACKED,
                _category(CHANNELS, REGION_VALUES))
    if number in (53, 57):
        totals = [sum(col) for col in zip(*REGION_VALUES)]
        head = [[round(v / t, 3) for v, t in zip(row, totals)] for row in REGION_VALUES[:-1]]
        last = [round(1 - sum(column), 3) for column in zip(*head)]  # each column sums to exactly 100%
        shares = tuple(tuple(row) for row in head) + (tuple(last),)
        return (XL_CHART_TYPE.COLUMN_STACKED_100 if number == 53 else XL_CHART_TYPE.BAR_STACKED_100,
                _category(CHANNELS, shares, number_format="0%"))
    if number == 58:
        return XL_CHART_TYPE.BAR_CLUSTERED, _category(
            ("Score",), ((92, 78, 63, 51, 39),),
            categories=("Supplier portal", "Mobile approvals", "Spend analytics", "Forecasting", "Partner API"))
    if number in (59, 60):
        data = CategoryChartData(number_format="0%")
        data.categories = ["Enterprise", "Mid-market", "Small business"]
        data.add_series("Share of revenue", (0.52, 0.31, 0.17))
        return (XL_CHART_TYPE.PIE if number == 59 else XL_CHART_TYPE.DOUGHNUT), data
    if number in (61, 63):
        values = ((118, 124, 131, 140, 146, 155), (120, 125, 130, 135, 140, 145))
        chart_type = XL_CHART_TYPE.LINE_MARKERS if number == 61 else XL_CHART_TYPE.COLUMN_CLUSTERED
        return chart_type, _category(actual_target, values, categories=MONTHS)
    if number == 62:
        values = ((70, 72, 75, 79, 81, 85), (33, 35, 37, 40, 42, 45), (15, 17, 19, 21, 23, 25))
        return XL_CHART_TYPE.AREA_STACKED, _category(("Standard", "Professional", "Enterprise"), values,
                                                     categories=MONTHS)
    if number == 64:
        data = CategoryChartData(number_format="0")
        data.categories = list(WATERFALL_CATEGORIES)
        for name, values in waterfall_series(WATERFALL_START, WATERFALL_STEPS):
            data.add_series(name, values)
        return XL_CHART_TYPE.COLUMN_STACKED, data
    if number == 65:
        data = XyChartData(number_format="0")
        series = data.add_series("Products")
        for x, y in ((82, 61), (88, 66), (94, 70), (97, 64), (103, 74), (108, 72), (115, 79)):
            series.add_data_point(x, y)
        return XL_CHART_TYPE.XY_SCATTER, data
    if number == 66:
        data = BubbleChartData(number_format="0")
        series = data.add_series("Markets")
        for x, y, size in ((4, 12, 14), (7, 22, 24), (11, 15, 18), (15, 30, 30)):
            series.add_data_point(x, y, size)
        return XL_CHART_TYPE.BUBBLE, data
    if number == 67:
        return XL_CHART_TYPE.BAR_CLUSTERED, _category(
            ("Opportunities",), ((1200, 640, 310, 140, 85),),
            categories=("Leads", "Qualified", "Proposal", "Negotiation", "Won"))
    raise ValueError(number)


def _dlbls_xml(*, position: str | None, light_text: bool, number_format: str | None = None,
               per_point_formats: dict[int, str] | None = None, show_cat: bool = False,
               show_percent: bool = False, show_val: bool = True) -> str:
    """Return a <c:dLbls> element that uses theme colours only."""

    tx_pr = ""
    if light_text:
        tx_pr = ('<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>'
                 '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr>'
                 '<a:endParaRPr lang="en-GB"/></a:p></c:txPr>')
    flags = (f'<c:showLegendKey val="0"/><c:showVal val="{int(show_val)}"/>'
             f'<c:showCatName val="{int(show_cat)}"/><c:showSerName val="0"/>'
             f'<c:showPercent val="{int(show_percent)}"/><c:showBubbleSize val="0"/>')
    pos = f'<c:dLblPos val="{position}"/>' if position else ""
    points = ""
    for idx, fmt in (per_point_formats or {}).items():
        points += (f'<c:dLbl><c:idx val="{idx}"/><c:numFmt formatCode="{fmt}" sourceLinked="0"/>'
                   f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{tx_pr}{pos}{flags}</c:dLbl>')
    num_fmt = f'<c:numFmt formatCode="{number_format}" sourceLinked="0"/>' if number_format else ""
    separator = "<c:separator>\n</c:separator>" if show_cat and (show_percent or show_val) else ""
    return (f'<c:dLbls xmlns:c="{C_NS}" xmlns:a="{A_NS}">{points}{num_fmt}'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{tx_pr}{pos}{flags}{separator}</c:dLbls>')


def _set_series_dlbls(ser: etree._Element, dlbls_xml: str) -> None:
    existing = ser.find("c:dLbls", NS)
    if existing is not None:
        ser.remove(existing)
    new = etree.fromstring(dlbls_xml)
    # dLbls follows tx/spPr/invertIfNegative/pictureOptions/dPt in CT_*Ser.
    anchor = None
    for tag in ("dPt", "pictureOptions", "invertIfNegative", "explosion", "marker", "spPr", "tx", "order"):
        found = ser.findall(f"c:{tag}", NS)
        if found:
            anchor = found[-1]
            break
    if anchor is None:
        ser.insert(0, new)
    else:
        anchor.addnext(new)


def _set_series_no_fill(ser: etree._Element) -> None:
    sp_pr = ser.find("c:spPr", NS)
    if sp_pr is not None:
        ser.remove(sp_pr)
    sp_pr = etree.fromstring(f'<c:spPr xmlns:c="{C_NS}" xmlns:a="{A_NS}"><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')
    tx = ser.find("c:tx", NS)
    (tx if tx is not None else ser.find("c:order", NS)).addnext(sp_pr)


def _set_series_fill(ser: etree._Element, scheme: str, modifiers: str = "") -> None:
    sp_pr = ser.find("c:spPr", NS)
    if sp_pr is not None:
        ser.remove(sp_pr)
    sp_pr = etree.fromstring(
        f'<c:spPr xmlns:c="{C_NS}" xmlns:a="{A_NS}"><a:solidFill><a:schemeClr val="{scheme}">{modifiers}'
        '</a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>')
    tx = ser.find("c:tx", NS)
    (tx if tx is not None else ser.find("c:order", NS)).addnext(sp_pr)


def _remove_plot_dlbls(chart_space: etree._Element) -> None:
    for plot in chart_space.iterfind(".//c:plotArea/*", NS):
        dlbls = plot.find("c:dLbls", NS)
        if dlbls is not None:
            plot.remove(dlbls)


def style_chart(chart, number: int) -> None:
    """Restrained defaults: no title, no gridlines, theme colours throughout."""

    chart.has_title = False
    chart.chart_style = 2  # default Office style: series colours come from theme accents
    chart.font.size = Pt(14)
    chart_space = chart._chartSpace
    series_elements = chart_space.findall(".//c:plotArea/*/c:ser", NS)
    series_count = len(series_elements)

    has_axes = number not in (59, 60)
    if has_axes:
        value_axis = chart.value_axis
        value_axis.has_major_gridlines = False
        value_axis.has_minor_gridlines = False
        value_axis.has_title = False
        category_axis = chart.category_axis
        category_axis.has_major_gridlines = False
        category_axis.has_title = False
        # Keep the continuous scale only where the audience needs it.
        value_axis.visible = number in (61, 62, 63, 65, 66)
        if number in DESCENDING_BAR_ENTRIES:
            category_axis.reverse_order = True

    chart.has_legend = series_count > 1 and number not in (64,)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

    _remove_plot_dlbls(chart_space)

    if number in (59, 60):
        chart.has_legend = False
        dlbls = _dlbls_xml(position="ctr" if number == 60 else "inEnd", light_text=True, number_format="0%",
                           show_cat=True, show_percent=True, show_val=False)
        _set_series_dlbls(series_elements[0], dlbls)
        if number == 60:
            hole = chart_space.find(".//c:doughnutChart/c:holeSize", NS)
            if hole is not None:
                hole.set("val", "55")
        return

    if number in (61, 62, 63, 65, 66):
        if number == 65:
            for series in chart.plots[0].series:
                series.format.line.fill.background()
                series.marker.style = XL_MARKER_STYLE.CIRCLE
                series.marker.size = 9
                series.marker.format.fill.solid()
                series.marker.format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
                series.marker.format.line.fill.background()
        return

    if number == 64:
        base, total, increase, decrease = series_elements
        _set_series_no_fill(base)
        for ser, (scheme, mods) in zip((total, increase, decrease), WATERFALL_FILLS.values()):
            _set_series_fill(ser, scheme, mods)
        # Zero-height points get no label (empty format sections); signs are
        # added by the number format, so the labels stay linked to the data.
        _set_series_dlbls(total, _dlbls_xml(position="ctr", light_text=True, number_format="#,##0;;;"))
        _set_series_dlbls(increase, _dlbls_xml(position="ctr", light_text=True, number_format="+#,##0;;;"))
        _set_series_dlbls(decrease, _dlbls_xml(position="ctr", light_text=True, number_format="-#,##0;;;"))
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        legend_pos = chart_space.find(".//c:legend/c:legendPos", NS)
        legend_pos.addnext(etree.fromstring(
            f'<c:legendEntry xmlns:c="{C_NS}"><c:idx val="0"/><c:delete val="1"/></c:legendEntry>'))
        gap = chart_space.find(".//c:barChart/c:gapWidth", NS)
        if gap is not None:
            gap.set("val", "60")
        return

    stacked = number in (52, 53, 56, 57)
    fmt = "0%" if number in (53, 57) else "#,##0"
    for ser in series_elements:
        _set_series_dlbls(ser, _dlbls_xml(position="ctr" if stacked else "outEnd", light_text=stacked,
                                          number_format=fmt))
    if stacked:
        gap = chart_space.find(".//c:barChart/c:gapWidth", NS)
        if gap is not None:
            gap.set("val", "60")


def patch_combination_chart(chart) -> None:
    """Move the second series of a clustered column chart onto a line plot."""

    plot_area = chart._chartSpace.find(".//c:plotArea", NS)
    bar = plot_area.find("c:barChart", NS)
    series = bar.findall("c:ser", NS)
    line_series = series[-1]
    bar.remove(line_series)
    for tag in ("invertIfNegative",):
        element = line_series.find(f"c:{tag}", NS)
        if element is not None:
            line_series.remove(element)
    line_chart = etree.SubElement(plot_area, qname(C_NS, "lineChart"))
    etree.SubElement(line_chart, qname(C_NS, "grouping"), val="standard")
    etree.SubElement(line_chart, qname(C_NS, "varyColors"), val="0")
    line_chart.append(line_series)
    marker = etree.fromstring(f'<c:marker xmlns:c="{C_NS}"><c:symbol val="none"/></c:marker>')
    anchor = line_series.find("c:spPr", NS)
    if anchor is None:
        anchor = line_series.find("c:tx", NS)
    anchor.addnext(marker)
    etree.SubElement(line_series, qname(C_NS, "smooth"), val="0")
    etree.SubElement(line_chart, qname(C_NS, "marker"), val="1")
    for ax_id in bar.findall("c:axId", NS):
        line_chart.append(etree.fromstring(etree.tostring(ax_id)))
    bar.addnext(line_chart)


def insert_chart(slide, entry: Entry) -> None:
    placeholder = placeholder_by_idx(slide, 1)
    chart_type, data = chart_definition(entry.number)
    frame = slide.shapes.add_chart(chart_type, placeholder.left, placeholder.top, placeholder.width,
                                   placeholder.height, data)
    chart = frame.chart
    if entry.number == 63:
        patch_combination_chart(chart)
    style_chart(chart, entry.number)
    adopt_into_placeholder(slide, frame, placeholder)


# --- tables -----------------------------------------------------------------

TABLES: dict[int, tuple[tuple[str, ...], tuple[tuple[str, ...], ...]]] = {
    82: (("Stage", "Discover", "Evaluate", "Buy", "Onboard", "Renew"),
         (("Customer goal", "Understand options", "Compare suppliers", "Agree terms", "Go live", "Review value"),
          ("Touchpoint", "Website", "Sales call", "Contract", "Training", "Account review"),
          ("Owner", "Marketing", "Sales", "Legal", "Services", "Account team"))),
    88: (("Action", "Owner", "Status", "Due"),
         (("Confirm regional scope", "Alex Morgan", "Open", "31 Oct"),
          ("Approve budget", "Jordan Singh", "In review", "7 Nov"),
          ("Publish training plan", "Casey Brown", "Complete", "14 Oct"))),
    89: (("Criterion", "Vendor A", "Vendor B", "Vendor C"),
         (("Five-year cost", "£1.2m", "£0.9m", "£1.1m"),
          ("Time to deploy", "9 months", "6 months", "8 months"),
          ("Security certification", "Yes", "Yes", "Pending"))),
    90: (("Line item", "Actual", "Budget", "Variance", "Forecast"),
         (("Staff", "1,240", "1,200", "40", "1,260"),
          ("Technology", "610", "650", "-40", "640"),
          ("Facilities", "320", "300", "20", "330"),
          ("Total", "2,170", "2,150", "20", "2,230"))),
    91: (("Workstream", "Oct", "Nov", "Dec", "Jan", "Owner"),
         (("Design", "■", "■", "", "", "Taylor Chen"),
          ("Build", "", "■", "■", "", "Riley Okafor"),
          ("Test", "", "", "■", "■", "Sam Rivera"),
          ("Launch", "", "", "", "■", "Alex Morgan"))),
    92: (("Time", "Session", "Speaker", "Room"),
         (("09:00", "Welcome", "Alex Morgan", "Hall A"),
          ("09:30", "Market update", "Taylor Chen", "Hall A"),
          ("11:00", "Workshops", "Various", "Rooms 1 to 4"),
          ("13:00", "Lunch", "", "Atrium"))),
    93: (("Task", "Product", "Sales", "Finance", "Operations"),
         (("Set pricing", "R", "C", "A", "I"),
          ("Forecast demand", "C", "R", "A", "I"),
          ("Plan capacity", "I", "C", "C", "A/R"),
          ("Report results", "I", "I", "A/R", "C"))),
    94: (("Risk", "Likelihood", "Impact", "Mitigation", "Owner"),
         (("Supplier delay", "Medium", "High", "Second supplier on standby", "Sam Rivera"),
          ("Data migration errors", "Low", "High", "Two trial migrations", "Riley Okafor"),
          ("Low adoption", "Medium", "Medium", "Training and champions", "Casey Brown"))),
    95: (("Item", "Type", "Owner", "Due", "Status"),
         (("Login failures", "Issue", "Riley Okafor", "18 Oct", "Open"),
          ("Update user guide", "Action", "Casey Brown", "25 Oct", "In progress"),
          ("Licence renewal", "Action", "Jordan Singh", "1 Nov", "Complete"))),
    96: (("Action", "Owner", "Due"),
         (("Approve phase two budget", "Steering group", "31 Oct"),
          ("Confirm pilot sites", "Taylor Chen", "7 Nov"),
          ("Start staff training", "Casey Brown", "14 Nov"),
          ("Report progress", "Alex Morgan", "5 Dec"))),
}

HEATMAP = (("Region", "Q1", "Q2", "Q3", "Q4"),
           (("North", 18, 42, 61, 78), ("Central", 25, 55, 72, 84),
            ("South", 31, 48, 67, 73), ("West", 44, 59, 76, 91)))


def _table_in_placeholder(slide, headers, rows, *, row_height=0.5, font_size=14):
    placeholder = placeholder_by_idx(slide, 1)
    n_rows = len(rows) + 1
    frame = slide.shapes.add_table(n_rows, len(headers), placeholder.left, placeholder.top, placeholder.width,
                                   Inches(row_height * n_rows))
    table = frame.table
    for row_index, row in enumerate((headers,) + tuple(rows)):
        table.rows[row_index].height = Inches(row_height)
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = str(value)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    # Size only; colour and font come from the theme table style.
                    run.font.size = Pt(font_size)
    adopt_into_placeholder(slide, frame, placeholder)
    return table


def insert_table(slide, entry: Entry) -> None:
    headers, rows = TABLES[entry.number]
    font_size = 12 if len(headers) >= 6 else 14
    _table_in_placeholder(slide, headers, rows, font_size=font_size)


def insert_heatmap(slide, entry: Entry) -> None:
    headers, rows = HEATMAP
    table = _table_in_placeholder(slide, headers, rows, row_height=0.85, font_size=14)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row[1:], start=1):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
            # Four sequential bins of the theme accent: light to dark.
            brightness = 0.8 if value < 30 else 0.6 if value < 50 else 0.0 if value < 70 else -0.5
            cell.fill.fore_color.brightness = brightness
            run = cell.text_frame.paragraphs[0].runs[0]
            run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1 if brightness > 0.5 else MSO_THEME_COLOR.BACKGROUND_1
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# --- native-shape diagrams --------------------------------------------------


@dataclass
class Box:
    shape: object

    @property
    def cx(self) -> float:
        return (self.shape.left + self.shape.width / 2) / EMU_PER_INCH

    @property
    def cy(self) -> float:
        return (self.shape.top + self.shape.height / 2) / EMU_PER_INCH

    @property
    def hw(self) -> float:
        return self.shape.width / 2 / EMU_PER_INCH

    @property
    def hh(self) -> float:
        return self.shape.height / 2 / EMU_PER_INCH


def _style_text(shape, lines, *, size=14, color=MSO_THEME_COLOR.TEXT_1, align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE, bold_first=False):
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.vertical_anchor = anchor
    for margin in ("margin_left", "margin_right"):
        setattr(text_frame, margin, Inches(0.08))
    for index, value in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = value
        run.font.size = Pt(size if index == 0 else size - 2)
        run.font.color.theme_color = color
        if bold_first and index == 0:
            run.font.bold = True


def node(slide, left, top, width, height, lines, *, shape=MSO_AUTO_SHAPE_TYPE.RECTANGLE, filled=False,
         size=14, brightness=0.0, bold_first=True) -> Box:
    item = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    if filled:
        item.fill.solid()
        item.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        item.fill.fore_color.brightness = brightness
        item.line.fill.background()
        text_color = MSO_THEME_COLOR.TEXT_1 if brightness > 0.5 else MSO_THEME_COLOR.BACKGROUND_1
    else:
        item.fill.background()
        item.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
        item.line.width = Pt(1.25)
        text_color = MSO_THEME_COLOR.TEXT_1
    item.shadow.inherit = False
    _style_text(item, lines, size=size, color=text_color, bold_first=bold_first)
    return Box(item)


def label(slide, left, top, width, height, lines, *, size=14, align=PP_ALIGN.CENTER,
          anchor=MSO_ANCHOR.MIDDLE, bold_first=False):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    _style_text(box, lines, size=size, align=align, anchor=anchor, bold_first=bold_first)
    return box


def _edge_point(box: Box, toward_x: float, toward_y: float) -> tuple[float, float]:
    dx, dy = toward_x - box.cx, toward_y - box.cy
    if dx == 0 and dy == 0:
        return box.cx, box.cy
    scale = min(box.hw / abs(dx) if dx else math.inf, box.hh / abs(dy) if dy else math.inf)
    return box.cx + dx * scale, box.cy + dy * scale


def line(slide, x1, y1, x2, y2, *, arrow=True, elbow=False):
    connector = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT,
                                           Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    connector.line.color.theme_color = MSO_THEME_COLOR.TEXT_1
    connector.line.width = Pt(1.25)
    if arrow:
        ln = connector.line._get_or_add_ln()
        ln.append(etree.fromstring(f'<a:tailEnd xmlns:a="{A_NS}" type="triangle" w="med" len="med"/>'))
    return connector


def link(slide, a: Box, b: Box, *, arrow=True, gap=0.08):
    x1, y1 = _edge_point(a, b.cx, b.cy)
    x2, y2 = _edge_point(b, a.cx, a.cy)
    length = math.hypot(x2 - x1, y2 - y1) or 1
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    return line(slide, x1 + ux * gap, y1 + uy * gap, x2 - ux * gap, y2 - uy * gap, arrow=arrow)


def diagram_timeline(slide):
    milestones = (("Oct 2026", "Kick-off"), ("Dec 2026", "Design signed off"), ("Mar 2027", "Pilot live"),
                  ("Jun 2027", "Full rollout"), ("Sep 2027", "Review"))
    y = 3.9
    start, end = 1.2, SLIDE_WIDTH_IN - 1.2
    line(slide, start - 0.4, y, end + 0.4, y, arrow=False)
    step = (end - start) / (len(milestones) - 1)
    for i, (date, event) in enumerate(milestones):
        x = start + i * step
        dot = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x - 0.12), Inches(y - 0.12), Inches(0.24),
                                     Inches(0.24))
        dot.fill.solid()
        dot.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        dot.line.fill.background()
        label(slide, x - 1.1, y - 1.05, 2.2, 0.7, (date,), size=16, anchor=MSO_ANCHOR.BOTTOM, bold_first=True)
        label(slide, x - 1.1, y + 0.3, 2.2, 0.8, (event,), size=14, anchor=MSO_ANCHOR.TOP)


def diagram_horizontal(slide):
    steps = (("1  Order", "Customer submits"), ("2  Check", "Credit and stock"), ("3  Pick", "Warehouse"),
             ("4  Ship", "Carrier collects"), ("5  Invoice", "Finance bills"))
    width, gap = 1.95, (CONTENT_W - 5 * 1.95) / 4
    boxes = [node(slide, MARGIN + i * (width + gap), 3.0, width, 1.5, lines) for i, lines in enumerate(steps)]
    for a, b in zip(boxes, boxes[1:]):
        link(slide, a, b)


def diagram_vertical(slide):
    steps = (("1  Request", "Submitted through the portal"), ("2  Review", "Data owner checks purpose"),
             ("3  Approve", "Governance board signs off"), ("4  Deliver", "Extract shared within five days"))
    height, gap = 0.85, (CONTENT_H - 4 * 0.85) / 3
    boxes = [node(slide, 3.4, CONTENT_TOP + i * (height + gap), 6.5, height, lines) for i, lines in enumerate(steps)]
    for a, b in zip(boxes, boxes[1:]):
        link(slide, a, b)


def diagram_chevron(slide):
    phases = ("Discover", "Design", "Build", "Test", "Launch")
    details = ("Interviews and data review", "Target process agreed", "Configuration complete",
               "User acceptance", "Go live and support")
    width, overlap = 2.55, 0.12
    total = 5 * width - 4 * overlap
    left = MARGIN + (CONTENT_W - total) / 2
    for i, (phase, detail) in enumerate(zip(phases, details)):
        x = left + i * (width - overlap)
        kind = MSO_AUTO_SHAPE_TYPE.PENTAGON if i == 0 else MSO_AUTO_SHAPE_TYPE.CHEVRON
        node(slide, x, 2.6, width, 1.0, (phase,), shape=kind, filled=True, size=16)
        label(slide, x + 0.15, 3.8, width - 0.45, 1.2, (detail,), size=14, anchor=MSO_ANCHOR.TOP)


def _ring(slide, items, *, rx, ry, cx=SLIDE_WIDTH_IN / 2, cy=4.1, width=2.4, height=0.95):
    boxes = []
    for i, lines in enumerate(items):
        angle = -math.pi / 2 + 2 * math.pi * i / len(items)
        x, y = cx + rx * math.cos(angle), cy + ry * math.sin(angle)
        boxes.append(node(slide, x - width / 2, y - height / 2, width, height, lines,
                          shape=MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE))
    for a, b in zip(boxes, boxes[1:] + boxes[:1]):
        link(slide, a, b)
    return boxes


def diagram_circular(slide):
    _ring(slide, (("Plan",), ("Do",), ("Check",), ("Act",), ("Standardise",)), rx=3.6, ry=1.95)


def diagram_lifecycle(slide):
    _ring(slide, (("Acquire",), ("Operate",), ("Maintain",), ("Retire",)), rx=3.3, ry=1.85)


def diagram_funnel(slide):
    stages = (("Applications", "1,200"), ("Screened", "480"), ("Interviewed", "150"), ("Offers", "40"))
    height, gap = 1.05, 0.2
    max_w, min_w = 10.0, 4.0
    for i, (stage, value) in enumerate(stages):
        width = max_w - (max_w - min_w) * i / (len(stages) - 1)
        node(slide, (SLIDE_WIDTH_IN - width) / 2, CONTENT_TOP + 0.2 + i * (height + gap), width, height,
             (f"{stage}: {value}",), filled=True, brightness=(0.0, 0.2, 0.4, 0.6)[i], size=16)


def diagram_swimlane(slide):
    lanes = ("Customer", "Service desk", "Engineering")
    lane_h = CONTENT_H / 3
    for i, lane in enumerate(lanes):
        top = CONTENT_TOP + i * lane_h
        line(slide, MARGIN, top, MARGIN + CONTENT_W, top, arrow=False)
        label(slide, MARGIN, top, 1.7, lane_h, (lane,), size=14, align=PP_ALIGN.LEFT, bold_first=True)
    line(slide, MARGIN, CONTENT_BOTTOM, MARGIN + CONTENT_W, CONTENT_BOTTOM, arrow=False)
    steps = ((0, "Raise request"), (1, "Triage"), (2, "Fix"), (1, "Confirm fix"), (0, "Close request"))
    width = 1.75
    gap = (CONTENT_W - 1.9 - 5 * width) / 4
    boxes = []
    for i, (lane, text_value) in enumerate(steps):
        top = CONTENT_TOP + lane * lane_h + (lane_h - 0.8) / 2
        boxes.append(node(slide, MARGIN + 1.9 + i * (width + gap), top, width, 0.8, (text_value,), bold_first=False))
    for a, b in zip(boxes, boxes[1:]):
        link(slide, a, b)


def diagram_workflow(slide):
    draft = node(slide, MARGIN, 3.1, 2.1, 1.0, ("Draft",), shape=MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE)
    review = node(slide, 3.4, 3.1, 2.1, 1.0, ("Review",))
    decision = node(slide, 6.0, 2.7, 2.8, 1.8, ("Approved?",), shape=MSO_AUTO_SHAPE_TYPE.DIAMOND)
    publish = node(slide, 10.3, 3.1, 2.1, 1.0, ("Publish",), shape=MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE)
    revise = node(slide, 6.35, 5.35, 2.1, 1.0, ("Revise",))
    link(slide, draft, review)
    link(slide, review, decision)
    link(slide, decision, publish)
    link(slide, decision, revise)
    # Revise loops back to review.
    line(slide, revise.cx - revise.hw - 0.08, revise.cy, review.cx, revise.cy, arrow=False)
    line(slide, review.cx, revise.cy, review.cx, review.cy + review.hh + 0.08)
    label(slide, 8.9, 3.05, 1.2, 0.4, ("Yes",), size=14)
    label(slide, 7.45, 4.6, 0.8, 0.4, ("No",), size=14)


def diagram_org(slide):
    top = node(slide, (SLIDE_WIDTH_IN - 3.0) / 2, CONTENT_TOP, 3.0, 1.0, ("Alex Morgan", "Programme sponsor"))
    children = (("Taylor Chen", "Operations"), ("Jordan Singh", "Finance"), ("Riley Okafor", "Technology"))
    width = 3.0
    gap = (CONTENT_W - 3 * width) / 2
    boxes = [node(slide, MARGIN + i * (width + gap), 3.6, width, 1.0, lines) for i, lines in enumerate(children)]
    bus_y = 3.05
    line(slide, top.cx, top.cy + top.hh, top.cx, bus_y, arrow=False)
    line(slide, boxes[0].cx, bus_y, boxes[-1].cx, bus_y, arrow=False)
    for child in boxes:
        line(slide, child.cx, bus_y, child.cx, child.cy - child.hh, arrow=False)
    grandchildren = (("Casey Brown", "Change"), ("Sam Rivera", "Product"))
    for i, lines in enumerate(grandchildren):
        parent = boxes[2]
        box = node(slide, parent.cx - parent.hw + 0.4, 5.05 + i * 0.85, 2.6, 0.7, lines, size=12)
        line(slide, parent.cx - parent.hw + 0.2, parent.cy + parent.hh, parent.cx - parent.hw + 0.2, box.cy,
             arrow=False)
        line(slide, parent.cx - parent.hw + 0.2, box.cy, box.cx - box.hw, box.cy, arrow=False)


def diagram_ecosystem(slide):
    hub = node(slide, (SLIDE_WIDTH_IN - 2.6) / 2, 3.55, 2.6, 1.1, ("Our platform",), shape=MSO_AUTO_SHAPE_TYPE.OVAL,
               filled=True)
    partners = (("Payment providers",), ("Logistics partners",), ("Resellers",), ("Data suppliers",),
                ("Systems integrators",))
    rx, ry = 4.4, 2.0
    for i, lines in enumerate(partners):
        angle = -math.pi / 2 + 2 * math.pi * i / len(partners)
        x, y = hub.cx + rx * math.cos(angle), hub.cy + ry * math.sin(angle)
        box = node(slide, x - 1.25, y - 0.42, 2.5, 0.84, lines, shape=MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                   bold_first=False)
        link(slide, hub, box, arrow=False)


def diagram_pyramid(slide):
    levels = ("Vision", "Strategy", "Programmes", "Operations")
    total_h, base_w = CONTENT_H - 0.1, 9.0
    level_h = total_h / len(levels)
    apex_x = SLIDE_WIDTH_IN / 2
    for i, level_name in enumerate(levels):
        top_w = base_w * i / len(levels)
        bottom_w = base_w * (i + 1) / len(levels)
        top = CONTENT_TOP + i * level_h
        if i == 0:
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, Inches(apex_x - bottom_w / 2),
                                           Inches(top), Inches(bottom_w), Inches(level_h))
        else:
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.TRAPEZOID, Inches(apex_x - bottom_w / 2),
                                           Inches(top), Inches(bottom_w), Inches(level_h - 0.06))
            inset = (bottom_w - top_w) / 2
            shape.adjustments[0] = inset / min(bottom_w, level_h - 0.06)
        brightness = (0.0, 0.2, 0.4, 0.6)[i]
        shape.fill.solid()
        shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        shape.fill.fore_color.brightness = brightness
        shape.line.fill.background()
        shape.shadow.inherit = False
        color = MSO_THEME_COLOR.TEXT_1 if brightness > 0.5 else MSO_THEME_COLOR.BACKGROUND_1
        _style_text(shape, (level_name,), size=16, color=color, anchor=MSO_ANCHOR.BOTTOM if i == 0 else MSO_ANCHOR.MIDDLE,
                    bold_first=True)


def diagram_treemap(slide):
    """Treemap with areas proportional to value (slice-and-dice layout)."""

    items = (("Product A", 42), ("Product B", 24), ("Product C", 18), ("Product D", 16))
    left, top, width, height = MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H
    total = sum(v for _, v in items)
    a_w = width * items[0][1] / total
    rest_w = width - a_w
    rest_total = total - items[0][1]
    b_h = height * items[1][1] / rest_total
    c_w = rest_w * items[2][1] / (items[2][1] + items[3][1])
    rects = (
        (left, top, a_w, height),
        (left + a_w, top, rest_w, b_h),
        (left + a_w, top + b_h, c_w, height - b_h),
        (left + a_w + c_w, top + b_h, rest_w - c_w, height - b_h),
    )
    accents = (MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3, MSO_THEME_COLOR.ACCENT_4)
    for (name, value), (x, y, w, h), accent in zip(items, rects, accents):
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.theme_color = accent
        shape.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
        shape.line.width = Pt(2)
        shape.shadow.inherit = False
        _style_text(shape, (name, f"{value}%"), size=18, color=MSO_THEME_COLOR.BACKGROUND_1,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bold_first=True)


DIAGRAMS: dict[int, Callable] = {
    68: diagram_treemap,
    73: diagram_timeline,
    75: diagram_horizontal,
    76: diagram_vertical,
    77: diagram_chevron,
    78: diagram_circular,
    79: diagram_funnel,
    80: diagram_swimlane,
    81: diagram_workflow,
    83: diagram_lifecycle,
    84: diagram_org,
    85: diagram_ecosystem,
    86: diagram_pyramid,
}


# --- slide assembly ---------------------------------------------------------


def add_example(prs, entry: Entry) -> None:
    layout = next(item for item in prs.slide_layouts if item.name == entry.layout)
    slide = prs.slides.add_slide(layout)
    if entry.title is not None and slide.shapes.title is not None:
        slide.shapes.title.text_frame.text = entry.title
    for idx, lines in entry.bodies.items():
        fill_text(placeholder_by_idx(slide, idx), lines)
    if entry.kind == "chart":
        insert_chart(slide, entry)
    elif entry.kind == "table":
        insert_table(slide, entry)
    elif entry.kind == "heatmap":
        insert_heatmap(slide, entry)
    elif entry.kind in ("diagram", "treemap"):
        DIAGRAMS[entry.number](slide)


def populate_examples(template_path: Path, output: Path) -> None:
    prs = Presentation(str(template_path))
    names = tuple(layout.name for layout in prs.slide_layouts)
    if names != LAYOUT_NAMES:
        raise RuntimeError(f"Unexpected layouts: {names!r}")
    prs.core_properties.title = "Plain fallback PowerPoint layouts"
    prs.core_properties.subject = f"{len(LAYOUTS)} opt-in layouts with {len(CATALOG)} example slides"
    prs.core_properties.author = "brand-faithful-pptx"
    prs.core_properties.comments = "Fallback layouts require explicit user approval before use."
    for entry in CATALOG:
        add_example(prs, entry)
    prs.save(str(output))


# --- generated reference data ----------------------------------------------


def role_overrides() -> dict:
    """Reviewed placeholder roles for scripts/inspect_template.py --role-overrides."""

    overrides: dict[str, dict] = {}
    for layout_index, spec in enumerate(LAYOUTS):
        for slot_spec in spec.slots:
            if slot_spec.ph_type in ("title", "ctrTitle", "pic", "subTitle"):
                continue  # placeholder type already gives a high-confidence role
            overrides[f"{layout_index}:{slot_spec.idx}"] = {
                "layout_name": spec.name,
                "placeholder_name": None,  # filled in from the built deck
                "content_role": slot_spec.role,
                "content_class": slot_spec.content_class,
                "note": slot_spec.note or "Use only for supplied content.",
            }
    return {"version": 1, "overrides": overrides}


def write_role_overrides(deck: Path, output: Path) -> None:
    data = role_overrides()
    prs = Presentation(str(deck))
    for key, item in data["overrides"].items():
        layout_index, idx = (int(part) for part in key.split(":"))
        layout = prs.slide_layouts[layout_index]
        item["placeholder_name"] = next(ph.name for ph in layout.placeholders if ph.placeholder_format.idx == idx)
    output.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def write_catalog(output: Path) -> None:
    data = {
        "description": "Issue #3 slide types mapped to the fallback deck's example slides and layouts.",
        "layouts": [
            {"name": spec.name, "type": spec.ooxml_type, "standard_office_layout": spec.standard}
            for spec in LAYOUTS
        ],
        "entries": [
            {"slide": entry.number, "name": entry.name, "category": entry.category, "layout": entry.layout,
             "example": entry.kind}
            for entry in CATALOG
        ],
    }
    output.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def build(output: Path, roles_output: Path | None = None, catalog_output: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="plain-fallback-pptx-") as tmp:
        tmp_dir = Path(tmp)
        seed = tmp_dir / "seed.pptx"
        template = tmp_dir / "template.pptx"
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_WIDTH_IN)
        prs.slide_height = Inches(SLIDE_HEIGHT_IN)
        prs.save(str(seed))
        build_template(seed, template)
        output.parent.mkdir(parents=True, exist_ok=True)
        populate_examples(template, output)
    if roles_output:
        write_role_overrides(output, roles_output)
    if catalog_output:
        write_catalog(catalog_output)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Output .pptx path")
    parser.add_argument("--roles-output", type=Path, default=DEFAULT_ROLES_OUTPUT,
                        help="Where to write the reviewed placeholder-role overrides JSON")
    parser.add_argument("--catalog-output", type=Path, default=DEFAULT_CATALOG_OUTPUT,
                        help="Where to write the slide-type to layout catalogue JSON")
    args = parser.parse_args()
    build(args.output, args.roles_output, args.catalog_output)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
