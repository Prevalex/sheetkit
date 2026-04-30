from __future__ import annotations

import json
from os import PathLike
from pathlib import Path

from .themes import (
    DEFAULT_OFFICE_THEME_DIR,
    DEFAULT_USER_THEME_DIR,
    THEMES_DIR,
    USER_THEME_NAME_PREFIX,
    iter_thmx_files,
    load_thmx_theme,
    make_theme_preset_name,

)


def export_excel_theme(
        thmx_file: str | PathLike,
        json_file: str | PathLike,
) -> Path:
    """
    Converts one .thmx theme file to a JSON preset file.
    """
    theme_info = load_thmx_theme(thmx_file)
    json_path = Path(json_file)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(theme_info, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return json_path


def export_excel_themes(
        theme_dir: Path,
        *,
        output_dir: Path,
        is_user_theme: bool,
        used_names: set[str],
) -> list[Path]:
    """
    Exports all .thmx files from a directory into JSON theme presets.
    """
    exported: list[Path] = []

    for thmx_file in iter_thmx_files(theme_dir):
        preset_name = make_theme_preset_name(thmx_file.stem, is_user_theme=is_user_theme)
        if preset_name in used_names and is_user_theme:
            preset_name = f"{USER_THEME_NAME_PREFIX + preset_name.removeprefix(USER_THEME_NAME_PREFIX)}"

        json_file = output_dir / f"{preset_name}.json"
        export_excel_theme(thmx_file, json_file)
        exported.append(json_file)
        used_names.add(preset_name)

    return exported


def main() -> None:
    output_dir = THEMES_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    used_names = {file.stem for file in output_dir.glob("*.json")}
    exported: list[Path] = []

    exported.extend(
        export_excel_themes(
            DEFAULT_OFFICE_THEME_DIR,
            output_dir=output_dir,
            is_user_theme=False,
            used_names=used_names,
        )
    )

    exported.extend(
        export_excel_themes(
            DEFAULT_USER_THEME_DIR,
            output_dir=output_dir,
            is_user_theme=True,
            used_names=used_names,
        )
    )

    for json_file in exported:
        print(json_file)


if __name__ == "__main__":
    main()

