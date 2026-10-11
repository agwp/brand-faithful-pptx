from __future__ import annotations

import copy
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Callable

from lxml import etree


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from scripts import build_plain_fallback_deck as builder  # noqa: E402
from scripts import verify_plain_fallback_deck as verifier  # noqa: E402


REFERENCE_DECK = REPO_ROOT / "references" / "plain-fallback-layouts.pptx"
ROLES_FILE = REPO_ROOT / "references" / "plain-fallback-placeholder-roles.json"
CATALOG_FILE = REPO_ROOT / "references" / "plain-fallback-catalog.json"
P_NS = verifier.P_NS
A_NS = verifier.A_NS
C_NS = verifier.C_NS
REL_NS = verifier.REL_NS
NS = verifier.NS
REL = verifier.REL


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def to_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def layout_part(name: str) -> str:
    index = verifier.expected_layout_names().index(name) + 1
    return f"ppt/slideLayouts/slideLayout{index}.xml"


def chart_part_for_slide(parts: dict[str, bytes], slide: int) -> str:
    rels = etree.fromstring(parts[f"ppt/slides/_rels/slide{slide}.xml.rels"])
    target = next(r.get("Target") for r in rels.findall("rel:Relationship", REL) if r.get("Type") == verifier.CHART_REL)
    return verifier._resolve_target(f"ppt/slides/slide{slide}.xml", target)


class CatalogueTests(unittest.TestCase):
    def test_catalogue_covers_all_issue_3_slide_types(self) -> None:
        self.assertEqual(len(builder.CATALOG), 100)
        self.assertEqual([e.number for e in builder.CATALOG], list(range(1, 101)))
        self.assertEqual(builder.CATALOG[0].name, "Standard title slide")
        self.assertEqual(builder.CATALOG[48].name, "One-series column chart")
        self.assertEqual(builder.CATALOG[-1].name, "Blank branded canvas")

    def test_every_entry_uses_a_defined_layout(self) -> None:
        names = set(builder.LAYOUT_NAMES)
        for entry in builder.CATALOG:
            self.assertIn(entry.layout, names, entry.name)

    def test_every_layout_is_used_by_an_example(self) -> None:
        used = {entry.layout for entry in builder.CATALOG}
        self.assertEqual(used, set(builder.LAYOUT_NAMES))

    def test_layouts_are_consolidated_and_core_layouts_are_standard(self) -> None:
        self.assertLessEqual(len(builder.LAYOUTS), 30)
        self.assertEqual(builder.STANDARD_LAYOUT_TYPES, verifier.REQUIRED_STANDARD_LAYOUTS)

    def test_generated_reference_files_match_the_deck(self) -> None:
        catalog = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))
        self.assertEqual(len(catalog["entries"]), 100)
        self.assertEqual([item["name"] for item in catalog["layouts"]], list(builder.LAYOUT_NAMES))
        roles = json.loads(ROLES_FILE.read_text(encoding="utf-8"))
        self.assertEqual(roles["version"], 1)
        for key, item in roles["overrides"].items():
            layout_index = int(key.split(":")[0])
            self.assertEqual(item["layout_name"], builder.LAYOUT_NAMES[layout_index], key)


class PlainFallbackVerifierTests(unittest.TestCase):
    maxDiff = None

    def inspect_mutation(self, callback: Callable[[dict[str, bytes]], None]) -> dict:
        with zipfile.ZipFile(REFERENCE_DECK) as source:
            parts = {name: source.read(name) for name in source.namelist()}
        callback(parts)
        with tempfile.TemporaryDirectory() as temp_directory:
            output = Path(temp_directory) / "mutated.pptx"
            with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as target:
                for name in sorted(parts):
                    target.writestr(name, parts[name])
            return verifier.inspect_deck(output)

    def mutate_xml(self, part: str, change: Callable[[etree._Element], None]) -> dict:
        def mutate(parts: dict[str, bytes]) -> None:
            root = etree.fromstring(parts[part])
            change(root)
            parts[part] = to_bytes(root)

        return self.inspect_mutation(mutate)

    def assert_has_error(self, report: dict, text: str) -> None:
        self.assertFalse(report["ok"])
        self.assertTrue(any(text in error for error in report["errors"]),
                        f"Expected an error containing {text!r}; errors were {report['errors']!r}")

    # --- acceptance -----------------------------------------------------

    def test_reference_deck_passes_all_acceptance_checks(self) -> None:
        report = verifier.inspect_deck(REFERENCE_DECK)
        self.assertTrue(report["ok"], report["errors"])
        self.assertEqual(report["master_count"], 1)
        self.assertEqual(report["layout_count"], len(builder.LAYOUTS))
        self.assertEqual(report["standard_layout_count"], 9)
        self.assertEqual(report["slide_count"], 100)
        self.assertEqual(report["native_chart_part_count"], 20)
        self.assertEqual(report["embedded_workbook_count"], 20)

    # --- layouts and master ---------------------------------------------

    def test_wrong_layout_name_is_rejected(self) -> None:
        def change(root):
            root.find("p:cSld", NS).set("name", "Invented cover")

        report = self.mutate_xml(layout_part("Title Slide"), change)
        self.assert_has_error(report, "Layout 1 is named 'Invented cover'")

    def test_custom_type_on_standard_layout_is_rejected(self) -> None:
        report = self.mutate_xml(layout_part("Title and Content"), lambda root: root.set("type", "cust"))
        self.assert_has_error(report, "Standard layout 'Title and Content' with type 'obj' is missing")

    def test_decorative_shape_on_layout_is_rejected(self) -> None:
        def change(root):
            sp_tree = root.find("p:cSld/p:spTree", NS)
            rule = etree.SubElement(sp_tree, qname(P_NS, "sp"))
            nv = etree.SubElement(rule, qname(P_NS, "nvSpPr"))
            etree.SubElement(nv, qname(P_NS, "cNvPr"), id="99", name="Accent rule")
            etree.SubElement(nv, qname(P_NS, "cNvSpPr"))
            etree.SubElement(nv, qname(P_NS, "nvPr"))
            etree.SubElement(rule, qname(P_NS, "spPr"))

        report = self.mutate_xml(layout_part("Title and Content"), change)
        self.assert_has_error(report, "contains a non-placeholder sp 'Accent rule'")

    def test_title_placeholder_with_idx_zero_is_rejected(self) -> None:
        def change(root):
            ph = next(p for p in root.findall(".//p:ph", NS) if p.get("type") == "title")
            ph.set("idx", "0")

        report = self.mutate_xml(layout_part("Title Only"), change)
        self.assert_has_error(report, "title placeholder carries idx='0'")

    def test_duplicate_placeholder_idx_is_rejected(self) -> None:
        def change(root):
            placeholders = [p for p in root.findall(".//p:ph", NS) if p.get("idx") in ("1", "2")]
            placeholders[1].set("idx", placeholders[0].get("idx"))

        report = self.mutate_xml(layout_part("Two Content"), change)
        self.assert_has_error(report, "duplicate placeholder indices")

    def test_hard_coded_colour_on_layout_is_rejected(self) -> None:
        def change(root):
            def_rpr = root.find(".//a:defRPr", NS)
            fill = etree.SubElement(def_rpr, qname(A_NS, "solidFill"))
            etree.SubElement(fill, qname(A_NS, "srgbClr"), val="2F80ED")

        report = self.mutate_xml(layout_part("Title Slide"), change)
        self.assert_has_error(report, "hard-coded colour(s) (a:srgbClr)")

    def test_master_without_footer_placeholders_is_rejected(self) -> None:
        def change(root):
            for shape in root.findall("p:cSld/p:spTree/p:sp", NS):
                ph = shape.find(".//p:ph", NS)
                if ph is not None and ph.get("type") in ("dt", "ftr", "sldNum"):
                    shape.getparent().remove(shape)

        report = self.mutate_xml("ppt/slideMasters/slideMaster1.xml", change)
        self.assert_has_error(report, "Slide master lacks a sldNum placeholder")

    def test_missing_layout_back_relationship_is_rejected(self) -> None:
        def change(root):
            for rel in list(root):
                if rel.get("Type") == verifier.SLIDE_MASTER_REL:
                    root.remove(rel)

        report = self.mutate_xml("ppt/slideLayouts/_rels/slideLayout1.xml.rels", change)
        self.assert_has_error(report, "Layout 1 must have exactly one back-relationship")

    # --- slides -----------------------------------------------------------

    def test_slide_using_wrong_layout_is_rejected(self) -> None:
        def change(root):
            rel = next(r for r in root.findall("rel:Relationship", REL) if r.get("Type") == verifier.SLIDE_LAYOUT_REL)
            rel.set("Target", "../" + layout_part("Blank").split("/", 1)[1])

        report = self.mutate_xml("ppt/slides/_rels/slide1.xml.rels", change)
        self.assert_has_error(report, "Slide 1 uses layout 'Blank'")

    def test_placeholder_font_size_override_is_rejected(self) -> None:
        report = self.mutate_xml("ppt/slides/slide1.xml",
                                 lambda root: root.find(".//p:sp//a:rPr", NS).set("sz", "2800")
                                 if root.find(".//p:sp//a:rPr", NS) is not None
                                 else root.find(".//p:sp//a:r", NS).insert(0, etree.Element(qname(A_NS, "rPr"), sz="2800")))
        self.assert_has_error(report, "placeholder text overrides font size")

    def test_named_typeface_on_slide_is_rejected(self) -> None:
        def change(root):
            run = root.find(".//p:sp//a:r", NS)
            rpr = run.find("a:rPr", NS)
            if rpr is None:
                rpr = etree.Element(qname(A_NS, "rPr"))
                run.insert(0, rpr)
            etree.SubElement(rpr, qname(A_NS, "latin"), typeface="Aptos")

        report = self.mutate_xml("ppt/slides/slide17.xml", change)
        self.assert_has_error(report, "names typeface 'Aptos'")

    def test_ornamental_microcopy_is_rejected(self) -> None:
        def change(root):
            root.find(".//a:t", NS).text = "Where insight meets impact"

        report = self.mutate_xml("ppt/slides/slide1.xml", change)
        self.assert_has_error(report, "contains prohibited ornamental copy")

    def test_image_relationship_on_slide_is_rejected(self) -> None:
        def change(root):
            etree.SubElement(root, qname(REL_NS, "Relationship"), Id="rIdImageTest", Type=verifier.IMAGE_REL,
                             Target="../slideLayouts/slideLayout1.xml")

        report = self.mutate_xml("ppt/slides/_rels/slide49.xml.rels", change)
        self.assert_has_error(report, "Slide 49 contains a picture")

    # --- charts -----------------------------------------------------------

    def test_chart_title_and_gridlines_are_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 49)
            root = etree.fromstring(parts[name])
            root.find("c:chart", NS).insert(0, etree.Element(qname(C_NS, "title")))
            root.find(".//c:valAx", NS).insert(0, etree.Element(qname(C_NS, "majorGridlines")))
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "contains a chart or axis title")
        self.assert_has_error(report, "contains major gridlines")

    def test_hard_coded_series_colour_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 50)
            root = etree.fromstring(parts[name])
            ser = root.find(".//c:ser", NS)
            sp_pr = etree.fromstring(f'<c:spPr xmlns:c="{C_NS}" xmlns:a="{A_NS}"><a:solidFill>'
                                     '<a:srgbClr val="17324D"/></a:solidFill></c:spPr>')
            ser.find("c:tx", NS).addnext(sp_pr)
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Slide 50 chart")
        self.assert_has_error(report, "hard-coded colour(s) (a:srgbClr)")

    def test_horizontal_bar_axis_must_put_highest_value_at_top(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 54)
            root = etree.fromstring(parts[name])
            root.find(".//c:catAx/c:scaling/c:orientation", NS).set("val", "minMax")
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "horizontal bar chart must reverse its category axis")

    def test_horizontal_bar_values_must_be_sorted_descending(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 54)
            root = etree.fromstring(parts[name])
            root.find(".//c:barChart/c:ser[1]/c:val/c:numRef/c:numCache/c:pt[@idx='3']/c:v", NS).text = "99"
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "horizontal bar categories are not sorted highest-to-lowest")

    def _waterfall_series(self, root):
        return {verifier._series_name(s): s for s in root.findall(".//c:barChart/c:ser", NS)}

    def test_waterfall_that_does_not_add_up_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 64)
            root = etree.fromstring(parts[name])
            total = self._waterfall_series(root)["Total"]
            total.find("c:val/c:numRef/c:numCache/c:pt[@idx='4']/c:v", NS).text = "117"
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "does not add up")

    def test_waterfall_without_a_decrease_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 64)
            root = etree.fromstring(parts[name])
            decrease = self._waterfall_series(root)["Decrease"]
            for value in decrease.findall("c:val/c:numRef/c:numCache/c:pt/c:v", NS):
                value.text = "0"
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "must show at least one increase and one decrease")

    def test_waterfall_base_labels_are_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 64)
            root = etree.fromstring(parts[name])
            series = self._waterfall_series(root)
            series["Base"].find("c:spPr", NS).addnext(copy.deepcopy(series["Increase"].find("c:dLbls", NS)))
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "labels its invisible base series")

    def test_waterfall_without_theme_colours_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 64)
            root = etree.fromstring(parts[name])
            decrease = self._waterfall_series(root)["Decrease"]
            scheme = decrease.find("c:spPr/a:solidFill/a:schemeClr", NS)
            scheme.getparent().replace(scheme, etree.Element(qname(A_NS, "srgbClr"), val="C00000"))
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Decrease series must use a theme colour fill")

    def test_scatter_with_connecting_line_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 65)
            root = etree.fromstring(parts[name])
            ln = root.find(".//c:scatterChart/c:ser/c:spPr/a:ln", NS)
            for child in list(ln):
                ln.remove(child)
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "scatter series must be markers only")

    def test_missing_chart_workbook_is_rejected(self) -> None:
        def mutate(parts):
            for name in [n for n in parts if n.startswith("ppt/embeddings/")][:1]:
                parts.pop(name)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "resolves to missing part")

    def test_typed_in_data_label_is_rejected(self) -> None:
        def mutate(parts):
            name = chart_part_for_slide(parts, 64)
            root = etree.fromstring(parts[name])
            dlbls = root.find(".//c:ser/c:dLbls", NS)
            dlbl = etree.fromstring(
                f'<c:dLbl xmlns:c="{C_NS}" xmlns:a="{A_NS}"><c:idx val="1"/><c:tx><c:rich><a:bodyPr/><a:p><a:r>'
                '<a:t>+15</a:t></a:r></a:p></c:rich></c:tx><c:showLegendKey val="0"/><c:showVal val="1"/>'
                '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
                '<c:showBubbleSize val="0"/></c:dLbl>')
            dlbls.insert(0, dlbl)
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "data label(s) with typed-in text")

    def test_text_box_over_chart_is_rejected(self) -> None:
        def change(root):
            sp_tree = root.find("p:cSld/p:spTree", NS)
            box = etree.fromstring(
                f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:nvSpPr><p:cNvPr id="90" name="Label"/>'
                '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
                '<a:p><a:r><a:t>42</a:t></a:r></a:p></p:txBody></p:sp>')
            sp_tree.append(box)

        report = self.mutate_xml("ppt/slides/slide49.xml", change)
        self.assert_has_error(report, "Native chart slide 49 has 1 free-floating shape(s) or text box(es)")

    def test_chart_not_bound_to_placeholder_is_rejected(self) -> None:
        def change(root):
            frame = root.find(".//p:graphicFrame", NS)
            ph = frame.find("p:nvGraphicFramePr/p:nvPr/p:ph", NS)
            ph.getparent().remove(ph)

        report = self.mutate_xml("ppt/slides/slide49.xml", change)
        self.assert_has_error(report, "chart is not bound to a content placeholder")

    def test_chart_relationship_on_diagram_slide_is_rejected(self) -> None:
        def mutate(parts):
            name = "ppt/slides/_rels/slide68.xml.rels"
            root = etree.fromstring(parts[name])
            etree.SubElement(root, qname(REL_NS, "Relationship"), Id="rIdChartTest", Type=verifier.CHART_REL,
                             Target="../" + chart_part_for_slide(parts, 49).split("/", 1)[1])
            parts[name] = to_bytes(root)

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Slide 68 is not a supported native-chart example but contains chart content")


if __name__ == "__main__":
    unittest.main()
