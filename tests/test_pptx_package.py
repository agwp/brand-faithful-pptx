from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from pptx_package import pack_pptx


class PackPptxTests(unittest.TestCase):
    @staticmethod
    def _create_package(package_dir: Path) -> None:
        (package_dir / "ppt").mkdir(parents=True)
        (package_dir / "[Content_Types].xml").write_text("<Types/>")
        (package_dir / "ppt" / "presentation.xml").write_text("<p:presentation/>")

    def test_rejects_output_inside_input_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            package_dir = Path(temp_dir) / "package"
            self._create_package(package_dir)
            output_path = package_dir / "rebuilt.pptx"

            with self.assertRaisesRegex(ValueError, "inside the input directory"):
                pack_pptx(package_dir, output_path)

            self.assertFalse(output_path.exists())
            self.assertFalse(output_path.with_suffix(".pptx.tmp").exists())

    def test_packs_output_outside_input_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root_dir = Path(temp_dir)
            package_dir = root_dir / "package"
            self._create_package(package_dir)
            output_path = root_dir / "rebuilt.pptx"

            pack_pptx(package_dir, output_path)

            with zipfile.ZipFile(output_path) as archive:
                self.assertEqual(
                    archive.namelist(),
                    ["[Content_Types].xml", "ppt/presentation.xml"],
                )
                self.assertIsNone(archive.testzip())

    def test_rejects_in_tree_output_symlinked_outside_input_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root_dir = Path(temp_dir)
            package_dir = root_dir / "package"
            self._create_package(package_dir)
            output_path = package_dir / "rebuilt.pptx"
            output_path.symlink_to(root_dir / "external.pptx")

            with self.assertRaisesRegex(ValueError, "inside the input directory"):
                pack_pptx(package_dir, output_path)

            self.assertTrue(output_path.is_symlink())
            self.assertFalse(output_path.with_suffix(".pptx.tmp").exists())

    def test_rejects_temporary_symlink_target_inside_input_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root_dir = Path(temp_dir)
            package_dir = root_dir / "package"
            self._create_package(package_dir)
            output_path = root_dir / "rebuilt.pptx"
            tmp_path = output_path.with_suffix(".pptx.tmp")
            tmp_path.symlink_to(package_dir / "injected.tmp")

            with self.assertRaisesRegex(ValueError, "temporary output file.*inside the input directory"):
                pack_pptx(package_dir, output_path)

            self.assertTrue(tmp_path.is_symlink())
            self.assertFalse(output_path.exists())


if __name__ == "__main__":
    unittest.main()
