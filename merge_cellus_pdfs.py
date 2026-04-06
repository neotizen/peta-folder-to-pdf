#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Merge all PDFs in each immediate subfolder of --root into "
            "{subfolder_name}_yyyymmdd.pdf and save to --output-dir."
        )
    )
    parser.add_argument(
        "--root",
        type=Path,
        required=True,
        help="Root folder that contains target subfolders.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Output folder for merged PDFs. Defaults to --root.",
    )
    parser.add_argument(
        "--date",
        default=datetime.now().strftime("%Y%m%d"),
        help="Date string in yyyymmdd format. Defaults to today.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Scan and print what would be merged without creating files.",
    )
    return parser.parse_args()


def is_auto_merged_file(pdf_path: Path, subfolder_name: str) -> bool:
    if pdf_path.suffix.lower() != ".pdf":
        return False
    stem = pdf_path.stem
    prefix = f"{subfolder_name}_"
    if not stem.startswith(prefix):
        return False
    tail = stem[len(prefix) :]
    return len(tail) == 8 and tail.isdigit()


def collect_source_pdfs(subfolder: Path) -> list[Path]:
    pdfs: list[Path] = []
    for path in subfolder.rglob("*.pdf"):
        if not path.is_file():
            continue
        if is_auto_merged_file(path, subfolder.name):
            continue
        pdfs.append(path)
    return sorted(pdfs, key=lambda p: str(p).lower())


def merge_pdfs(source_pdfs: list[Path], output_pdf: Path) -> None:
    writer = PdfWriter()
    try:
        for src_pdf in source_pdfs:
            reader = PdfReader(str(src_pdf))
            for page in reader.pages:
                writer.add_page(page)
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        with output_pdf.open("wb") as fp:
            writer.write(fp)
    finally:
        writer.close()


def main() -> int:
    args = parse_args()
    root: Path = args.root.expanduser().resolve()
    output_dir: Path = (
        args.output_dir.expanduser().resolve() if args.output_dir else root
    )
    date_str: str = args.date

    if len(date_str) != 8 or not date_str.isdigit():
        print(f"ERROR\tinvalid --date value: {date_str} (expected yyyymmdd)")
        return 2
    if not root.exists() or not root.is_dir():
        print(f"ERROR\troot folder not found: {root}")
        return 2

    subfolders = sorted([p for p in root.iterdir() if p.is_dir()], key=lambda p: p.name)

    processed = 0
    skipped = 0
    failed = 0

    for sub in subfolders:
        source_pdfs = collect_source_pdfs(sub)
        if not source_pdfs:
            skipped += 1
            print(f"SKIP\t{sub.name}\t(no source pdf)")
            continue

        output_pdf = output_dir / f"{sub.name}_{date_str}.pdf"

        if args.dry_run:
            print(f"DRY\t{sub.name}\t{len(source_pdfs)} files\t-> {output_pdf.name}")
            processed += 1
            continue

        try:
            if output_pdf.exists():
                output_pdf.unlink()
            merge_pdfs(source_pdfs, output_pdf)
            processed += 1
            print(f"OK\t{sub.name}\t{len(source_pdfs)} files\t-> {output_pdf.name}")
        except Exception as exc:  # pragma: no cover
            failed += 1
            print(f"FAIL\t{sub.name}\t{type(exc).__name__}: {exc}")

    print(
        f"DONE\tprocessed={processed}\tskipped={skipped}\tfailed={failed}\tdate={date_str}"
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
