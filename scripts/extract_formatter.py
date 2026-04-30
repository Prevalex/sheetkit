from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from sheetkit import extract_formatter_range_to_file, extract_formatter_to_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="extract_formatter",
        description="Extract formatter from a styled Excel worksheet.",
    )
    parser.add_argument("file_name", help="Source .xlsx file with a styled sample block.")
    parser.add_argument("json_file", help="Output formatter JSON file.")
    parser.add_argument("-sn", "--sheet-name", dest="sheet_name", default=None, help="Worksheet name.")
    parser.add_argument(
        "-m",
        "--mode",
        dest="mode",
        choices=("sample", "range"),
        default="sample",
        help="Extraction mode: sample (header/zebra template) or range (copy full rows/columns). Default: sample.",
    )
    parser.add_argument(
        "-cs",
        "--columns",
        dest="columns",
        type=int,
        default=None,
        help="Number of columns to scan.",
    )
    parser.add_argument(
        "-rs",
        "--rows",
        dest="rows",
        type=int,
        default=None,
        help="Number of rows to scan.",
    )
    parser.add_argument(
        "-p",
        "--priority",
        dest="priority",
        choices=("row", "col"),
        default="row",
        help="Formatter priority: row or col. Default: row.",
    )
    parser.add_argument(
        "-sc",
        "--start-cell",
        dest="start_cell",
        default="A1",
        help="Top-left cell of the sample block. Default: A1.",
    )
    parser.add_argument(
        "-hr",
        "--header-rows",
        "--header",
        dest="header",
        type=int,
        default=2,
        help="Number of header rows/columns in sample mode. Default: 2.",
    )
    parser.add_argument(
        "-z",
        "--zebra",
        dest="zebra",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Sample mode: import zebra styles from two sample data rows. Use --no-zebra to import only base style.",
    )
    parser.add_argument(
        "-name",
        "--name",
        dest="name",
        default=None,
        help="Optional formatter name to store in JSON.",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Suppress console output.",
    )
    parser.add_argument(
        "-log",
        "--log",
        dest="log_file",
        default=None,
        help="Append run protocol to a log file.",
    )
    return parser


def _emit(lines: list[str], *, quiet: bool, log_file: str | None) -> None:
    text = "\n".join(lines)
    if not quiet and text:
        print(text)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as f:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{ts}] extract_formatter\n")
            if text:
                f.write(text)
                f.write("\n")
            f.write("\n")


def main() -> int:
    parser = build_parser()
    raw_argv = ["--log" if token == "-log" else token for token in sys.argv[1:]]
    args = parser.parse_args(raw_argv)

    try:
        if args.mode == "range":
            if args.columns is None or args.columns < 1:
                parser.error("range mode requires -cs/--columns >= 1")
            if args.rows is None or args.rows < 1:
                parser.error("range mode requires -rs/--rows >= 1")

            result = extract_formatter_range_to_file(
                file_name=Path(args.file_name),
                json_file=Path(args.json_file),
                sheet_name=args.sheet_name,
                columns=args.columns,
                rows=args.rows,
                start_cell=args.start_cell,
                priority=args.priority,
                name=args.name,
            )
        else:
            result = extract_formatter_to_file(
                file_name=Path(args.file_name),
                json_file=Path(args.json_file),
                sheet_name=args.sheet_name,
                columns=args.columns,
                rows=args.rows,
                start_cell=args.start_cell,
                header=args.header,
                zebra=args.zebra,
                priority=args.priority,
                name=args.name,
            )
        _emit([str(result.resolve())], quiet=args.quiet, log_file=args.log_file)
        return 0

    except Exception as e:
        _emit([str(e)], quiet=args.quiet, log_file=args.log_file)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

