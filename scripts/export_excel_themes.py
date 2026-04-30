from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from sheetkit.themes import DEFAULT_OFFICE_THEME_DIR, DEFAULT_USER_THEME_DIR, build_formatter_from_theme, load_preset_file
from sheetkit.tools import export_excel_themes as export_excel_themes_from_dir
from sheetkit import save_formatter


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="export_excel_themes",
        description="Export Office and/or user .thmx themes into JSON presets.",
    )
    parser.add_argument(
        "-of",
        "--office-folder",
        dest="office_folder",
        nargs="?",
        const=str(DEFAULT_OFFICE_THEME_DIR),
        default=None,
        help="Office themes folder. If the flag is present without a value, the default Office folder is used.",
    )
    parser.add_argument(
        "-uf",
        "--user-folder",
        dest="user_folder",
        nargs="?",
        const=str(DEFAULT_USER_THEME_DIR),
        default=None,
        help="User themes folder. If the flag is present without a value, the default user folder is used.",
    )
    parser.add_argument(
        "-sf",
        "--save-folder",
        dest="save_folder",
        nargs="?",
        const=".",
        default=".",
        help="Folder where JSON presets will be saved. If omitted or used without a value, the current folder is used.",
    )
    parser.add_argument(
        "-cm",
        "--color-mode",
        dest="color_mode",
        choices=("hex", "ref"),
        default=None,
        help=(
            "Also export formatter JSON files from exported themes using this color mode. "
            "Creates both <theme>_row.json and <theme>_col.json in <save-folder>/formatters."
        ),
    )
    parser.add_argument(
        "-s",
        "--strict",
        action="store_true",
        help="Return non-zero exit code when no themes were exported.",
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
            f.write(f"[{ts}] export_excel_themes\n")
            if text:
                f.write(text)
                f.write("\n")
            f.write("\n")


def main() -> int:
    parser = build_parser()
    raw_argv = ["--log" if token == "-log" else token for token in sys.argv[1:]]
    args = parser.parse_args(raw_argv)

    if args.office_folder is None and args.user_folder is None:
        parser.error("Nothing to export. Specify -of and/or -uf.")

    output_dir = Path(args.save_folder)
    output_dir.mkdir(parents=True, exist_ok=True)

    used_names = {file.stem for file in output_dir.glob("*.json")}
    exported: list[Path] = []

    if args.office_folder is not None:
        exported.extend(
            export_excel_themes_from_dir(
                Path(args.office_folder),
                output_dir=output_dir,
                is_user_theme=False,
                used_names=used_names,
            )
        )

    if args.user_folder is not None:
        exported.extend(
            export_excel_themes_from_dir(
                Path(args.user_folder),
                output_dir=output_dir,
                is_user_theme=True,
                used_names=used_names,
            )
        )

    lines: list[str] = []
    for json_file in exported:
        lines.append(str(json_file.resolve()))

    if args.color_mode is not None:
        formatter_dir = output_dir / "formatters"
        formatter_dir.mkdir(parents=True, exist_ok=True)

        for theme_file in exported:
            theme = load_preset_file(theme_file, expected_kind="themes")
            for priority in ("row", "col"):
                formatter = build_formatter_from_theme(
                    theme,
                    color_mode=args.color_mode,
                    priority=priority,
                )
                formatter_file = formatter_dir / f"{theme_file.stem}_{priority}.json"
                save_formatter(
                    formatter,
                    formatter_file,
                    name=f"{theme_file.stem}_{priority}",
                    theme=theme,
                )
                lines.append(str(formatter_file.resolve()))

    if not exported:
        msg = "No themes exported."
        if args.strict:
            _emit([msg], quiet=args.quiet, log_file=args.log_file)
            return 1
        lines.append(msg)

    _emit(lines, quiet=args.quiet, log_file=args.log_file)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

