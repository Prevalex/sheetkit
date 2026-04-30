from sheetkit.colors import color_to_hex, normalize_color_value
from sheetkit.formatter import validate_format_dict


def test_normalize_color_value_supports_explicit_theme_tuple_reference() -> None:
    assert normalize_color_value(("wisp", "accent3")) == "9F8351"


def test_normalize_color_value_supports_theme_name_prefix() -> None:
    assert normalize_color_value("office_theme:accent1") == "4472C4"


def test_normalize_color_value_supports_semantic_theme_slots() -> None:
    assert normalize_color_value("Text1") == "000000"
    assert normalize_color_value("Background2") == "E7E6E6"
    assert normalize_color_value("Hyperlink") == "0563C1"


def test_normalize_color_value_supports_theme_shifts() -> None:
    assert normalize_color_value("Accent1+40") == "8FAADC"
    assert normalize_color_value("Background2-20") == "B9B8B8"


def test_color_to_hex_supports_theme_reference_with_alpha() -> None:
    assert color_to_hex(("office_theme", "accent2"), with_alpha=True) == "FFED7D31"


def test_validate_format_dict_accepts_theme_color_references() -> None:
    validate_format_dict(
        {
            "fg_color": ("office_theme", "accent1"),
            "font_color": "Accent2+25",
            "border_bottom_color": "Hyperlink",
        }
    )

