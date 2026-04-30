from copy import deepcopy
from pathlib import Path
import json
import shutil
import uuid
from zipfile import ZipFile

import pytest

from sheetkit import themes as themes_module
from sheetkit.themes import (
    build_formatter_from_theme,
    get_theme,
    import_theme,
    import_themes,
    load_formatter,
    save_formatter,
    resolve_formatter_colors,
    load_thmx_theme,
)


def _write_test_thmx(
    path: Path,
    *,
    scheme_name: str = "Codex Test Theme",
    accent2: str = "ED7D31",
) -> None:
    theme_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{scheme_name}">
  <a:themeElements>
    <a:clrScheme name="Codex Colors">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FEFEFE"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="EFEFEF"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="{accent2}"><a:tint val="60000"/></a:srgbClr></a:accent2>
      <a:accent3><a:srgbClr val="70AD47"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="A5A5A5"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Codex Fonts">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
"""
    variant_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Codex Variant">
  <a:themeElements>
    <a:clrScheme name="Variant Colors">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="C00000"><a:shade val="75000"/></a:srgbClr></a:accent1>
      <a:accent2><a:srgbClr val="00B050"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Variant Fonts">
      <a:majorFont><a:latin typeface="Variant Major"/></a:majorFont>
      <a:minorFont><a:latin typeface="Variant Minor"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>
"""
    with ZipFile(path, "w") as archive:
        archive.writestr("theme/theme/theme1.xml", theme_xml)
        archive.writestr("themeVariants/variant1/theme/theme/theme1.xml", variant_xml)


@pytest.fixture
def restore_themes() -> None:
    snapshot = deepcopy(themes_module.THEMES)
    try:
        yield
    finally:
        themes_module.THEMES.clear()
        themes_module.THEMES.update(snapshot)


@pytest.fixture
def workspace_tmp_dir() -> None:
    base_dir = Path(__file__).resolve().parent / "_tmp"
    temp_dir = base_dir / f"themes-{uuid.uuid4().hex}"
    temp_dir.mkdir(parents=True, exist_ok=True)
    try:
        yield temp_dir
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_build_formatter_from_theme_auto_uses_header_zebra_and_types() -> None:
    spec = build_formatter_from_theme(
        get_theme("office_theme"),
        types=["0.00", None, "@"],
        header=2,
        zebra=True,
    )

    assert 0 in spec["row"]
    assert 1 in spec["row"]
    assert -1 in spec["row"]
    assert -2 in spec["row"]
    assert spec["row"][0]["bold"] is True
    assert spec["col"][0]["num_format"] == "0.00"
    assert spec["col"][0]["align"] == "right"
    assert spec["col"][2]["num_format"] == "@"
    assert spec["col"][2]["align"] == "left"


def test_build_formatter_from_theme_uses_theme_colors_and_fonts() -> None:
    theme = {
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "44546A",
            "light2": "E7E6E6",
            "accent1": "4472C4",
            "accent2": "ED7D31",
        },
        "fonts": {"major": "Aptos Display", "minor": "Aptos"},
    }

    spec = build_formatter_from_theme(
        theme,
        header=2,
        zebra=True,
        types=["@", "0.00"],
        font_name="Aptos",
        font_size=12,
    )

    assert spec["row"][0]["fg_color"] == "4472C4"
    assert spec["row"][0]["font_color"] == "FFFFFF"
    assert spec["row"][0]["font_name"] == "Aptos"
    assert spec["row"][0]["font_size"] == 12
    assert spec["row"][1]["fg_color"] == "8FAADC"
    assert spec["row"][1]["font_color"] == "000000"
    assert "fg_color" not in spec["row"][-1]
    assert spec["row"][-2]["fg_color"] == "DAE3F3"
    assert spec["col"][0]["num_format"] == "@"
    assert spec["col"][1]["num_format"] == "0.00"


def test_build_formatter_from_theme_uses_variant_when_valid() -> None:
    theme = {
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "44546A",
            "light2": "E7E6E6",
            "accent1": "4472C4",
            "accent2": "ED7D31",
        },
        "variants": [
            {
                "colors": {
                    "accent1": "C00000",
                    "accent2": "00B050",
                    "light2": "F9F9F9",
                }
            }
        ],
        "fonts": {"minor": "Aptos"},
    }

    spec = build_formatter_from_theme(theme, header=2, zebra=True, variant=1)

    assert spec["row"][0]["fg_color"] == "C00000"
    assert spec["row"][1]["fg_color"] == "FF4040"
    assert spec["row"][1]["font_color"] == "000000"
    assert spec["row"][-2]["fg_color"] == "FFBFBF"


def test_build_formatter_from_theme_ignores_invalid_variant_and_falls_back_to_base() -> None:
    theme = {
        "colors": {
            "accent1": "4472C4",
            "accent2": "ED7D31",
            "light1": "FFFFFF",
            "light2": "E7E6E6",
            "dark2": "44546A",
        },
        "fonts": {"minor": "Aptos"},
    }

    spec = build_formatter_from_theme(theme, header=2, zebra=True, variant=99)

    assert spec["row"][0]["fg_color"] == "4472C4"
    assert spec["row"][1]["fg_color"] == "8FAADC"


def test_build_formatter_from_theme_uses_requested_accent_when_present() -> None:
    theme = {
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "44546A",
            "accent1": "4472C4",
            "accent2": "ED7D31",
        },
        "fonts": {"minor": "Aptos"},
    }

    spec = build_formatter_from_theme(theme, header=2, zebra=True, accent=2)

    assert spec["row"][0]["fg_color"] == "ED7D31"
    assert spec["row"][1]["fg_color"] == "F4B183"
    assert spec["row"][-2]["fg_color"] == "FBE5D6"


def test_build_formatter_from_theme_uses_accent1_when_requested_accent_is_invalid() -> None:
    theme = {
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "44546A",
            "accent1": "4472C4",
            "accent2": "ED7D31",
        },
        "fonts": {"minor": "Aptos"},
    }

    spec = build_formatter_from_theme(theme, header=2, zebra=True, accent=99)

    assert spec["row"][0]["fg_color"] == "4472C4"
    assert spec["row"][1]["fg_color"] == "8FAADC"


def test_load_thmx_theme_parses_colors_fonts_and_variants(workspace_tmp_dir: Path) -> None:
    thmx_path = workspace_tmp_dir / "codex_theme.thmx"
    _write_test_thmx(thmx_path)

    theme_info = load_thmx_theme(thmx_path)

    assert theme_info["kind"] == "themes"
    assert theme_info["source"] == thmx_path.name
    keys = list(theme_info.keys())
    assert keys[0] == "kind"
    assert keys[1] == "source"
    assert theme_info["scheme_name"] == "Codex Test Theme"
    assert theme_info["colors"]["accent1"] == "4472C4"
    assert theme_info["colors"]["light2"] == "EFEFEF"
    assert theme_info["color_transforms"]["accent2"]["transforms"][0]["op"] == "tint"
    assert theme_info["color_transforms"]["accent2"]["transforms"][0]["val"] == "60000"
    assert theme_info["fonts"]["major"] == "Aptos Display"
    assert theme_info["fonts"]["minor"] == "Aptos"
    assert theme_info["variants"][0]["scheme_name"] == "Codex Variant"
    assert theme_info["variants"][0]["colors"]["accent1"] == "C00000"
    assert theme_info["variants"][0]["color_transforms"]["accent1"]["transforms"][0]["op"] == "shade"
    assert theme_info["variants"][0]["color_transforms"]["accent1"]["transforms"][0]["val"] == "75000"


def test_load_thmx_theme_extracts_application_and_version(workspace_tmp_dir: Path) -> None:
    thmx_path = workspace_tmp_dir / "meta_theme.thmx"
    _write_test_thmx(thmx_path)
    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>
"""
    with ZipFile(thmx_path, "a") as archive:
        archive.writestr("docProps/app.xml", app_xml)

    theme_info = load_thmx_theme(thmx_path)

    assert theme_info["application"] == "Microsoft Office PowerPoint"
    assert theme_info["version"] == "16.0000"


def test_import_theme_registers_normalized_key(workspace_tmp_dir: Path, restore_themes: None) -> None:
    thmx_path = workspace_tmp_dir / "codex_theme.thmx"
    _write_test_thmx(thmx_path)

    assert themes_module.THEMES == {}

    imported = import_theme(thmx_path)

    assert "codex_test_theme" in themes_module.THEMES
    assert imported is themes_module.THEMES["codex_test_theme"]
    assert imported["kind"] == "themes"
    assert imported["scheme_name"] == "Codex Test Theme"


def test_get_theme_can_auto_import_from_file_path(workspace_tmp_dir: Path, restore_themes: None) -> None:
    thmx_path = workspace_tmp_dir / "codex_theme.thmx"
    _write_test_thmx(thmx_path)

    theme_info = get_theme(thmx_path)

    assert theme_info["scheme_name"] == "Codex Test Theme"
    assert themes_module.THEMES["codex_test_theme"]["colors"]["accent2"] == "ED7D31"


def test_get_theme_loads_json_when_explicit_json_path(workspace_tmp_dir: Path, restore_themes: None) -> None:
    json_file = workspace_tmp_dir / "mytheme.json"
    json_file.write_text(
        json.dumps(
            {
                "name": "mytheme",
                "kind": "themes",
                "scheme_name": "My Json Theme",
                "colors": {"accent1": "123456"},
                "fonts": {"minor": "Calibri"},
            }
        ),
        encoding="utf-8",
    )

    theme_info = get_theme(json_file)

    assert theme_info["scheme_name"] == "My Json Theme"
    assert theme_info["colors"]["accent1"] == "123456"


def test_get_theme_requires_kind_in_json_file(workspace_tmp_dir: Path, restore_themes: None) -> None:
    json_file = workspace_tmp_dir / "bad_theme.json"
    json_file.write_text(
        json.dumps(
            {
                "name": "bad_theme",
                "scheme_name": "Bad Theme",
                "colors": {"accent1": "123456"},
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="must contain key 'kind'"):
        get_theme(json_file)


def test_get_theme_name_prefers_json_preset_over_thmx_dirs(
    workspace_tmp_dir: Path,
    restore_themes: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    THEMES_DIR = workspace_tmp_dir / "themes"
    user_dir = workspace_tmp_dir / "user_themes"
    office_dir = workspace_tmp_dir / "office_themes"
    THEMES_DIR.mkdir(parents=True, exist_ok=True)
    user_dir.mkdir(parents=True, exist_ok=True)
    office_dir.mkdir(parents=True, exist_ok=True)

    (THEMES_DIR / "brand.json").write_text(
        json.dumps(
            {
                "name": "brand",
                "kind": "themes",
                "scheme_name": "Brand Json",
                "colors": {"accent1": "ABCDEF"},
                "fonts": {"minor": "Calibri"},
            }
        ),
        encoding="utf-8",
    )
    _write_test_thmx(user_dir / "Brand.thmx", scheme_name="Brand User", accent2="AA0000")
    _write_test_thmx(office_dir / "Brand.thmx", scheme_name="Brand Office", accent2="00AA00")

    monkeypatch.setattr(themes_module, "THEMES_DIR", THEMES_DIR)
    monkeypatch.setattr(themes_module, "DEFAULT_USER_THEME_DIR", user_dir)
    monkeypatch.setattr(themes_module, "DEFAULT_OFFICE_THEME_DIR", office_dir)

    theme_info = get_theme("brand")

    assert theme_info["scheme_name"] == "Brand Json"
    assert theme_info["colors"]["accent1"] == "ABCDEF"


def test_get_theme_prefers_user_theme_dir_over_office(
    workspace_tmp_dir: Path,
    restore_themes: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    user_dir = workspace_tmp_dir / "user_themes"
    office_dir = workspace_tmp_dir / "office_themes"
    user_dir.mkdir(parents=True, exist_ok=True)
    office_dir.mkdir(parents=True, exist_ok=True)

    _write_test_thmx(user_dir / "Brand.thmx", scheme_name="Brand User", accent2="AA0000")
    _write_test_thmx(office_dir / "Brand.thmx", scheme_name="Brand Office", accent2="00AA00")

    monkeypatch.setattr(themes_module, "DEFAULT_USER_THEME_DIR", user_dir)
    monkeypatch.setattr(themes_module, "DEFAULT_OFFICE_THEME_DIR", office_dir)

    theme_info = get_theme("Brand")

    assert theme_info["scheme_name"] == "Brand User"
    assert theme_info["colors"]["accent2"] == "AA0000"


def test_import_theme_can_resolve_by_name_in_default_dirs(
    workspace_tmp_dir: Path,
    restore_themes: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    user_dir = workspace_tmp_dir / "user_themes"
    user_dir.mkdir(parents=True, exist_ok=True)
    _write_test_thmx(user_dir / "Theme1.thmx", scheme_name="Theme One")

    monkeypatch.setattr(themes_module, "DEFAULT_USER_THEME_DIR", user_dir)
    monkeypatch.setattr(themes_module, "DEFAULT_OFFICE_THEME_DIR", workspace_tmp_dir / "missing_office")

    imported = import_theme("Theme1")

    assert imported["scheme_name"] == "Theme One"
    assert "theme_one" in themes_module.THEMES


def test_import_themes_uses_default_dirs_and_prefers_user(
    workspace_tmp_dir: Path,
    restore_themes: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    user_dir = workspace_tmp_dir / "user_themes"
    office_dir = workspace_tmp_dir / "office_themes"
    user_dir.mkdir(parents=True, exist_ok=True)
    office_dir.mkdir(parents=True, exist_ok=True)

    _write_test_thmx(user_dir / "Brand.thmx", scheme_name="Brand", accent2="AA0000")
    _write_test_thmx(office_dir / "Brand.thmx", scheme_name="Brand", accent2="00AA00")
    _write_test_thmx(office_dir / "Facet.thmx", scheme_name="Facet Office", accent2="123456")

    monkeypatch.setattr(themes_module, "DEFAULT_USER_THEME_DIR", user_dir)
    monkeypatch.setattr(themes_module, "DEFAULT_OFFICE_THEME_DIR", office_dir)

    imported = import_themes(theme_dir=None)

    assert "brand" in imported
    assert imported["brand"]["colors"]["accent2"] == "AA0000"
    assert "facet_office" in imported
    assert themes_module.THEMES["brand"]["colors"]["accent2"] == "AA0000"


def test_build_formatter_from_theme_uses_preloaded_fonts_by_default() -> None:
    spec = build_formatter_from_theme(
        get_theme("office_theme"),
        types=["0"],
        header=1,
        zebra=False,
    )

    assert spec["row"][0]["font_name"] == "Calibri"
    assert spec["row"][0]["font_size"] == 11


def test_theme_api_is_consistent(workspace_tmp_dir: Path, restore_themes: None) -> None:
    thmx_path = workspace_tmp_dir / "codex_theme.thmx"
    _write_test_thmx(thmx_path, scheme_name="Canonical Wrapper Theme")

    theme_a = get_theme(thmx_path)
    imported_a = import_theme(thmx_path)
    imported_many = import_themes(theme_dir=workspace_tmp_dir)
    fmt = build_formatter_from_theme(theme_a, header=1, zebra=True)
    assert theme_a["scheme_name"] == imported_a["scheme_name"]
    assert set(imported_many.keys())
    assert 0 in fmt["row"]


def test_formatter_wrappers_delegate_load_save(workspace_tmp_dir: Path) -> None:
    formatter = {"row": {-1: {"bold": True}}, "col": {}}
    json_file = workspace_tmp_dir / "formatter.json"

    written = save_formatter(formatter, json_file)
    loaded = load_formatter(written)

    assert written == json_file
    assert loaded["row"][-1]["bold"] is True


def test_load_formatter_registers_embedded_theme(workspace_tmp_dir: Path, restore_themes: None) -> None:
    theme = {
        "name": "unit_embedded",
        "kind": "themes",
        "scheme_name": "Unit Embedded",
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "222222",
            "accent1": "123456",
        },
        "fonts": {"minor": "Aptos"},
    }
    formatter = build_formatter_from_theme(theme, header=1, zebra=True, color_mode="ref")
    json_file = workspace_tmp_dir / "embedded_formatter.json"

    save_formatter(formatter, json_file, name="embedded_formatter", theme=theme)
    loaded = load_formatter(json_file)
    resolved = resolve_formatter_colors(loaded)

    assert loaded["row"][0]["fg_color"] == "unit_embedded:Accent1"
    assert resolved["row"][0]["fg_color"] == "123456"


def test_build_formatter_from_theme_with_color_refs_and_resolve() -> None:
    theme = get_theme("office_theme")
    formatter_ref = build_formatter_from_theme(theme, header=2, zebra=True, color_mode="ref")
    assert formatter_ref["row"][0]["fg_color"] == "office_theme:Accent1"
    assert formatter_ref["row"][1]["fg_color"] == "office_theme:Accent1+40"
    assert formatter_ref["row"][-2]["fg_color"] == "office_theme:Accent1+80"

    formatter_hex = resolve_formatter_colors(formatter_ref, theme=theme)
    assert formatter_hex["row"][0]["fg_color"] == "4472C4"


def test_build_formatter_from_theme_with_color_refs_uses_theme_name() -> None:
    theme = {
        "name": "facet",
        "colors": {
            "dark1": "000000",
            "light1": "FFFFFF",
            "dark2": "2C3C43",
            "accent1": "90C226",
        },
        "fonts": {"minor": "Trebuchet MS"},
    }

    formatter_ref = build_formatter_from_theme(theme, header=1, zebra=True, color_mode="ref")

    assert formatter_ref["row"][0]["fg_color"] == "facet:Accent1"
    assert formatter_ref["row"][0]["border_bottom_color"] == "facet:Dark2"


def test_build_formatter_from_theme_supports_column_priority() -> None:
    formatter = build_formatter_from_theme(
        get_theme("office_theme"),
        header=2,
        zebra=True,
        priority="col",
        types=["@", "0.00"],
    )

    assert formatter["priority"] == "col"
    assert formatter["col"][0]["fg_color"] == "4472C4"
    assert formatter["col"][1]["fg_color"] == "8FAADC"
    assert formatter["col"][-2]["fg_color"] == "DAE3F3"
    assert formatter["row"][0]["num_format"] == "@"
    assert formatter["row"][1]["num_format"] == "0.00"


