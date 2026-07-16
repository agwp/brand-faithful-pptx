#!/usr/bin/env python3
"""Pack and unpack PowerPoint .pptx files as ordinary directories.

A .pptx file is an Open Packaging Convention ZIP archive. Some coding agents
and chat clients mishandle binary .pptx files, but they can safely inspect and
edit the expanded XML/media directory. This helper provides a deterministic
round trip between the binary .pptx and an unpacked folder.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path


DEFAULT_EXCLUDES = {".DS_Store", "Thumbs.db"}


def _safe_members(zf: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    members: list[zipfile.ZipInfo] = []
    for info in zf.infolist():
        name = info.filename
        path = Path(name)
        if name.startswith("/") or ".." in path.parts:
            raise ValueError(f"unsafe archive member path: {name!r}")
        members.append(info)
    return members


def unpack_pptx(pptx_path: Path, output_dir: Path, force: bool = False) -> None:
    if output_dir.exists():
        if not force:
            raise FileExistsError(f"output directory already exists: {output_dir}")
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.testzip()
        for info in _safe_members(zf):
            zf.extract(info, output_dir)


def _iter_files(root: Path):
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        if path.name in DEFAULT_EXCLUDES:
            continue
        yield path


def _require_path_outside_directory(path: Path, directory: Path, label: str) -> None:
    resolved_directory = directory.resolve()
    locations = (path.resolve(), path.parent.resolve() / path.name)
    for location in locations:
        try:
            location.relative_to(resolved_directory)
        except ValueError:
            continue
        raise ValueError(f"{label} must not be inside the input directory: {path}")


def pack_pptx(input_dir: Path, pptx_path: Path, force: bool = False) -> None:
    required = [input_dir / "[Content_Types].xml", input_dir / "ppt" / "presentation.xml"]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise FileNotFoundError("input directory is not an unpacked .pptx package; missing " + ", ".join(missing))

    tmp_path = pptx_path.with_suffix(pptx_path.suffix + ".tmp")
    _require_path_outside_directory(pptx_path, input_dir, "output file")
    _require_path_outside_directory(tmp_path, input_dir, "temporary output file")

    if pptx_path.exists() and not force:
        raise FileExistsError(f"output file already exists: {pptx_path}")
    pptx_path.parent.mkdir(parents=True, exist_ok=True)

    if tmp_path.exists() or tmp_path.is_symlink():
        tmp_path.unlink()

    try:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in _iter_files(input_dir):
                arcname = path.relative_to(input_dir).as_posix()
                zf.write(path, arcname)
        tmp_path.replace(pptx_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    unpack = subparsers.add_parser("unpack", help="expand a .pptx into a directory")
    unpack.add_argument("pptx", type=Path, help="source .pptx file")
    unpack.add_argument("output_dir", type=Path, help="destination directory")
    unpack.add_argument("--force", action="store_true", help="replace the destination if it exists")

    pack = subparsers.add_parser("pack", help="zip an unpacked .pptx directory into a .pptx")
    pack.add_argument("input_dir", type=Path, help="source unpacked package directory")
    pack.add_argument("pptx", type=Path, help="destination .pptx file")
    pack.add_argument("--force", action="store_true", help="replace the destination if it exists")

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "unpack":
            unpack_pptx(args.pptx, args.output_dir, args.force)
        elif args.command == "pack":
            pack_pptx(args.input_dir, args.pptx, args.force)
        else:  # pragma: no cover
            raise AssertionError(args.command)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
