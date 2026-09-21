"""Command-line interface for the KPLC PDF renamer."""

from __future__ import annotations

import argparse

from pathlib import Path

from .renamer import (
    BillData,
    discover_pdfs,
    extract_bill_data,
    load_station_lookup,
    merge_bill_data,
    process_files,
    ocr_pdf,
)

DEFAULT_INPUT_DIRECTORY = "raw_files"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "paths",
        nargs="*",
        default=[DEFAULT_INPUT_DIRECTORY],
        help=f"PDF files, directories, or glob patterns (default: {DEFAULT_INPUT_DIRECTORY})",
    )
    parser.add_argument("--apply", action="store_true", help="perform the renames")
    parser.add_argument(
        "--workbook",
        default="excel/Power Bills Template.xlsx",
        help="Excel template containing station names and account numbers",
    )
    parser.add_argument(
        "--lookup-sheet",
        default="August",
        help="worksheet containing station and account lookup data",
    )
    parser.add_argument(
        "--output-sheet",
        default="August",
        help="worksheet where extracted bill values are written",
    )
    parser.add_argument(
        "--output",
        default="excel/August KPLC Power Bills.xlsx",
        help="output workbook for extracted bill fields",
    )
    parser.add_argument(
        "--on-conflict",
        choices=("skip", "error"),
        default="skip",
        help="what to do when the destination already exists (default: skip)",
    )
    args = parser.parse_args()

    files = discover_pdfs(args.paths)
    if not files:
        parser.error("no PDF files found")

    workbook = Path(args.workbook)
    try:
        station_lookup = (
            load_station_lookup(workbook, args.lookup_sheet) if workbook.exists() else {}
        )
    except (RuntimeError, ValueError) as error:
        parser.error(str(error))
    ocr_cache: dict[Path, str] = {}

    def cached_ocr(path: Path) -> str:
        if path not in ocr_cache:
            ocr_cache[path] = ocr_pdf(path)
        return ocr_cache[path]

    results = process_files(
        files,
        apply=args.apply,
        on_conflict=args.on_conflict,
        station_lookup=station_lookup,
        ocr=cached_ocr,
    )
    for result in results:
        destination = f" -> {result.destination}" if result.destination else ""
        print(f"{result.status}: {result.source}{destination} ({result.message})")
    bills: list[BillData] = []
    for path, text in ocr_cache.items():
        bill = extract_bill_data(text)
        if bill is not None:
            bills.append(bill)
    if args.apply and bills:
        merge_bill_data(workbook, Path(args.output), bills, args.output_sheet)
        print(f"merged: {len(bills)} bill(s) -> {args.output}")
    return 1 if any(result.status == "error" for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())