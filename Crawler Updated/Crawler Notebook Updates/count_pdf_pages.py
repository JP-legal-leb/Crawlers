#!/usr/bin/env python3
from __future__ import annotations


import argparse
import sys
from pathlib import Path

try:
    from PyPDF2 import PdfReader
except ImportError as exc:
    sys.stderr.write(
        "PyPDF2 is required. Install it with 'pip install PyPDF2'.\n"
    )
    raise SystemExit(1) from exc


def count_pages(pdf_path: Path) -> int:
    """Return the number of pages in the given PDF file."""
    reader = PdfReader(str(pdf_path))
    return len(reader.pages)


def iter_pdfs(root: Path, recursive: bool):
    """Yield PDF files under root."""
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from root.glob(pattern)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Report page counts for PDF files in a folder."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=".",
        help="folder to scan (defaults to current directory)",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="include PDFs in subdirectories",
    )
    args = parser.parse_args()

    root = Path("NCAR_PDFs").expanduser().resolve()
    if not root.exists():
        parser.error(f"Folder '{root}' does not exist.")
    if not root.is_dir():
        parser.error(f"'{root}' is not a folder.")

    found_any = False
    total_pages = 0

    for pdf_path in sorted(iter_pdfs(root, args.recursive)):
        found_any = True
        try:
            page_count = count_pages(pdf_path)
        except Exception as exc:  # noqa: BLE001
            rel_path = pdf_path if pdf_path == root else pdf_path.relative_to(root)
            sys.stderr.write(f"{rel_path}: ERROR ({exc})\n")
            continue

        rel_path = pdf_path if pdf_path == root else pdf_path.relative_to(root)
        print(f"{rel_path}: {page_count}")
        total_pages += page_count

    if not found_any:
        sys.stderr.write("No PDF files found.\n")
        return

    print(f"Total pages: {total_pages}")


if __name__ == "__main__":
    main()
