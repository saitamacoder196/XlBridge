"""CLI entry point for XlBridge."""

import argparse
import logging
import sys

from xlbridge.extractor import extract
from xlbridge.injector import inject


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="xlbridge",
        description="Extract and inject Excel cell content for translation.",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # extract
    p_extract = subparsers.add_parser("extract", help="Extract Excel content to TXT")
    p_extract.add_argument("--input", "-i", required=True, help="Input Excel file")
    p_extract.add_argument("--output", "-o", required=True, help="Output TXT file")
    p_extract.add_argument("--sheet", "-s", action="append", help="Sheet name to extract (repeatable, default: all)")

    # inject
    p_inject = subparsers.add_parser("inject", help="Inject translated TXT back into Excel")
    p_inject.add_argument("--input", "-i", required=True, help="Original Excel file")
    p_inject.add_argument("--translation", "-t", required=True, help="Translated TXT file")
    p_inject.add_argument("--output", "-o", default=None, help="Output Excel file (default: *_translated.xlsx)")

    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if args.command == "extract":
        extract(args.input, args.output, sheet_names=args.sheet)
    elif args.command == "inject":
        output = inject(args.input, args.translation, output_path=args.output)
        print(f"Output written to: {output}")


if __name__ == "__main__":
    main()
