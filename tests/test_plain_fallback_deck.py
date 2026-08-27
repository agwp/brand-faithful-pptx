from __future__ import annotations

import copy
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Callable

from lxml import etree


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from scripts import verify_plain_fallback_deck as verifier  # noqa: E402


REFERENCE_DECK = REPO_ROOT / "references" / "plain-fallback-100-layouts.pptx"
P_NS = verifier.P_NS
A_NS = verifier.A_NS
C_NS = verifier.C_NS
REL_NS = verifier.REL_NS
CT_NS = verifier.CT_NS
NS = verifier.NS
REL = verifier.REL


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


class PlainFallbackVerifierTests(unittest.TestCase):
    maxDiff = None

    def mutate_deck(
        self,
        callback: Callable[[dict[str, bytes]], None],
    ) -> tuple[tempfile.TemporaryDirectory[str], Path]:
        temp_directory = tempfile.TemporaryDirectory()
        output = Path(temp_directory.name) / "mutated.pptx"
        with zipfile.ZipFile(REFERENCE_DECK) as source:
            parts = {name: source.read(name) for name in source.namelist()}
        callback(parts)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for name in sorted(parts):
                target.writestr(name, parts[name])
        return temp_directory, output

    def inspect_mutation(
        self,
        callback: Callable[[dict[str, bytes]], None],
    ) -> dict:
        temp_directory, output = self.mutate_deck(callback)
        try:
            return verifier.inspect_deck(output)
        finally:
            temp_directory.cleanup()

    def assert_has_error(self, report: dict, text: str) -> None:
        self.assertFalse(report["ok"])
        self.assertTrue(
            any(text in error for error in report["errors"]),
            f"Expected an error containing {text!r}; errors were {report['errors']!r}",
        )

    def test_expected_layout_catalog_is_complete_and_exact(self) -> None:
        names = verifier.expected_layout_names()
        self.assertEqual(len(names), 100)
        self.assertEqual(names[0], "Standard title slide")
        self.assertEqual(names[48], "One-series column chart")
        self.assertEqual(names[71], "Chart with takeaway panel")
        self.assertEqual(names[-1], "Blank branded canvas")

    def test_reference_deck_passes_all_acceptance_checks(self) -> None:
        report = verifier.inspect_deck(REFERENCE_DECK)
        self.assertTrue(report["ok"], report["errors"])
        self.assertEqual(report["master_count"], 1)
        self.assertEqual(report["layout_count"], 100)
        self.assertEqual(report["slide_count"], 100)
        self.assertEqual(report["native_chart_part_count"], 20)
        self.assertEqual(report["embedded_workbook_count"], 20)
        self.assertEqual(report["editable_shape_data_slide_count"], 4)

    def test_wrong_layout_name_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slideLayouts/slideLayout1.xml"
            root = etree.fromstring(parts[name])
            root.set("matchingName", "Invented cover")
            root.find("p:cSld", NS).set("name", "Invented cover")
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Layout 1 is named 'Invented cover'")

    def test_slide_using_wrong_layout_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slides/_rels/slide1.xml.rels"
            root = etree.fromstring(parts[name])
            layout_rel = next(
                rel
                for rel in root.findall("rel:Relationship", REL)
                if rel.get("Type") == verifier.SLIDE_LAYOUT_REL
            )
            layout_rel.set("Target", "../slideLayouts/slideLayout2.xml")
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Slide 1 uses ppt/slideLayouts/slideLayout2.xml")

    def test_layout_100_with_placeholder_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            source = etree.fromstring(parts["ppt/slideLayouts/slideLayout99.xml"])
            source_shape = next(
                shape
                for shape in source.findall(".//p:sp", NS)
                if shape.find("p:nvSpPr/p:nvPr/p:ph", NS) is not None
            )
            name = "ppt/slideLayouts/slideLayout100.xml"
            root = etree.fromstring(parts[name])
            root.find("p:cSld/p:spTree", NS).append(copy.deepcopy(source_shape))
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Layout 100 must be a blank branded canvas")

    def test_missing_layout_back_relationship_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slideLayouts/_rels/slideLayout1.xml.rels"
            root = etree.fromstring(parts[name])
            for rel in list(root):
                if rel.get("Type") == verifier.SLIDE_MASTER_REL:
                    root.remove(rel)
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Layout 1 must have exactly one back-relationship")

    def test_missing_layout_content_type_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "[Content_Types].xml"
            root = etree.fromstring(parts[name])
            for override in list(root):
                if override.get("PartName") == "/ppt/slideLayouts/slideLayout1.xml":
                    root.remove(override)
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "Layout 1 part ppt/slideLayouts/slideLayout1.xml lacks")

    def test_chart_title_and_gridlines_are_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/charts/chart1.xml"
            root = etree.fromstring(parts[name])
            chart = root.find("c:chart", NS)
            chart.insert(0, etree.Element(qname(C_NS, "title")))
            axis = root.find(".//c:valAx", NS)
            axis.insert(0, etree.Element(qname(C_NS, "majorGridlines")))
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "contains a chart or axis title")
        self.assert_has_error(report, "contains major gridlines")

    def test_horizontal_bar_axis_must_put_highest_value_at_top(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/charts/chart6.xml"
            root = etree.fromstring(parts[name])
            orientation = root.find(".//c:catAx/c:scaling/c:orientation", NS)
            orientation.set("val", "minMax")
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(
            report,
            "horizontal bar chart must reverse its category axis",
        )

    def test_horizontal_bar_values_must_be_sorted_descending(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/charts/chart6.xml"
            root = etree.fromstring(parts[name])
            last_value = root.find(
                ".//c:barChart/c:ser[1]/c:val/c:numRef/c:numCache/c:pt[@idx='3']/c:v",
                NS,
            )
            last_value.text = "99"
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(
            report,
            "horizontal bar categories are not sorted highest-to-lowest",
        )

    def test_missing_chart_workbook_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            parts.pop("ppt/embeddings/Microsoft_Excel_Sheet1.xlsx")

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "resolves to missing part")

    def test_image_relationship_on_chart_slide_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slides/_rels/slide49.xml.rels"
            root = etree.fromstring(parts[name])
            etree.SubElement(
                root,
                qname(REL_NS, "Relationship"),
                Id="rIdImageVerifierTest",
                Type=verifier.IMAGE_REL,
                Target="../media/image1.png",
            )
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "may be using an image as a chart")

    def test_ornamental_microcopy_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slides/slide1.xml"
            root = etree.fromstring(parts[name])
            first_text = root.find(".//a:t", NS)
            first_text.text = "Where insight meets impact"
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(report, "contains prohibited ornamental copy")

    def test_chart_relationship_on_editable_shape_slide_is_rejected(self) -> None:
        def mutate(parts: dict[str, bytes]) -> None:
            name = "ppt/slides/_rels/slide68.xml.rels"
            root = etree.fromstring(parts[name])
            etree.SubElement(
                root,
                qname(REL_NS, "Relationship"),
                Id="rIdChartVerifierTest",
                Type=verifier.CHART_REL,
                Target="../charts/chart1.xml",
            )
            parts[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        report = self.inspect_mutation(mutate)
        self.assert_has_error(
            report,
            "Slide 68 is not a supported native-chart example but contains chart content",
        )


if __name__ == "__main__":
    unittest.main()
