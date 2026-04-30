import pytest

from sheetkit.formatter import (
    build_row_formats_with_columns,
    translate_formatter,
    validate_format_dict,
)


class DummyWorkbook:
    def __init__(self) -> None:
        self.calls: list[dict[str, object]] = []

    def add_format(self, props: dict[str, object]) -> dict[str, object]:
        self.calls.append(props)
        return props


def test_validate_format_dict_accepts_supported_color_formats() -> None:
    fmt = {
        "fg_color": "steelblue",
        "font_color": (255, 255, 255),
        "border_bottom_color": "#D9E1F2",
        "bold": True,
    }

    validate_format_dict(fmt)


def test_validate_format_dict_rejects_unknown_key() -> None:
    with pytest.raises(ValueError, match="Unknown format key"):
        validate_format_dict({"no_such_key": 1})


def test_build_row_formats_with_columns_creates_default_row_for_column_only_spec() -> None:
    format_spec = {
        "col": {
            0: {"num_format": "0"},
            1: {"num_format": "0.00"},
        },
        "row": {},
    }

    formatter = build_row_formats_with_columns(format_spec, max_cols=2)

    assert -1 in formatter
    assert formatter[-1][0]["num_format"] == "0"
    assert formatter[-1][1]["num_format"] == "0.00"


def test_build_row_formats_with_columns_merges_defaults_and_row_override() -> None:
    format_spec = {
        "col": {
            -1: {"align": "right"},
            0: {"num_format": "0.00"},
        },
        "row": {
            -1: {"font_name": "Calibri"},
            2: {"align": "left", "bold": True},
        },
    }

    formatter = build_row_formats_with_columns(format_spec, max_cols=2)

    assert formatter[2][0] == {
        "align": "left",
        "num_format": "0.00",
        "font_name": "Calibri",
        "bold": True,
    }
    assert formatter[2][1] == {
        "align": "left",
        "font_name": "Calibri",
        "bold": True,
    }


def test_build_row_formats_with_columns_returns_independent_dicts() -> None:
    format_spec = {
        "col": {0: {"num_format": "0"}},
        "row": {-1: {"font_name": "Calibri"}},
    }

    formatter = build_row_formats_with_columns(format_spec, max_cols=2)
    formatter[-1][0]["font_name"] = "Arial"

    assert formatter[-1][1]["font_name"] == "Calibri"


def test_build_row_formats_with_columns_supports_column_priority() -> None:
    format_spec = {
        "priority": "col",
        "col": {
            -1: {"fg_color": "FFFFFF"},
            -2: {"fg_color": "EAF3F8"},
            0: {"bold": True, "fg_color": "4472C4"},
        },
        "row": {
            0: {"num_format": "@"},
            1: {"num_format": "0.00"},
        },
    }

    formatter = build_row_formats_with_columns(format_spec, max_cols=4, max_rows=2)

    assert formatter[0][0]["bold"] is True
    assert formatter[0][0]["fg_color"] == "4472C4"
    assert formatter[0][0]["num_format"] == "@"
    assert formatter[1][1]["fg_color"] == "FFFFFF"
    assert formatter[1][2]["fg_color"] == "EAF3F8"
    assert formatter[1][2]["num_format"] == "0.00"


def test_translate_formatter_requires_workbook_for_xlsxwriter() -> None:
    with pytest.raises(ValueError, match="must provide 'workbook'"):
        translate_formatter({-1: [{"bold": True}]}, engine="xlsxwriter")


def test_translate_formatter_reuses_cached_xlsxwriter_formats() -> None:
    workbook = DummyWorkbook()
    formatter = {
        -1: [{"bold": True}, {"bold": True}],
        0: [{"bold": True}, {"italic": True}],
    }

    result = translate_formatter(formatter, engine="xlsxwriter", workbook=workbook)

    assert len(workbook.calls) == 2
    assert result[-1][0] is result[-1][1]
    assert result[-1][0] is result[0][0]
    assert result[0][1]["italic"] is True


def test_translate_formatter_builds_openpyxl_styles() -> None:
    formatter = {
        -1: [
            {
                "bold": True,
                "align": "center",
                "fg_color": "4472C4",
                "font_color": "FFFFFF",
                "num_format": "0.00",
            }
        ]
    }

    result = translate_formatter(formatter, engine="openpyxl")
    style = result[-1][0]

    assert style.alignment is not None
    assert style.alignment.horizontal == "center"
    assert style.fill is not None
    assert style.fill.fgColor.rgb == "FF4472C4"
    assert style.font is not None
    assert style.font.bold is True
    assert style.font.color is not None
    assert style.font.color.rgb == "FFFFFFFF"
    assert style.number_format == "0.00"

