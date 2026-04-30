import sheetkit


def test_root_package_exports_preferred_api() -> None:
    assert callable(sheetkit.write_sheet)
    assert callable(sheetkit.save_formatted_xlsx)
    assert callable(sheetkit.build_formatter_from_theme)
    assert callable(sheetkit.resolve_formatter_colors)
    assert callable(sheetkit.load_thmx_theme)
    assert callable(sheetkit.get_theme)
    assert callable(sheetkit.import_theme)
    assert callable(sheetkit.import_themes)
    assert callable(sheetkit.save_formatter)
    assert callable(sheetkit.load_formatter)
    assert callable(sheetkit.load_format_preset)
    assert callable(sheetkit.list_preset_themes)
    assert callable(sheetkit.list_preset_formatters)
    assert callable(sheetkit.save_preset_file)
    assert callable(sheetkit.load_preset_file)
    assert callable(sheetkit.extract_formatter_from_sheet)
    assert callable(sheetkit.extract_formatter_to_file)
    assert callable(sheetkit.extract_formatter_range_from_sheet)
    assert callable(sheetkit.extract_formatter_range_to_file)


def test_root_package_exports_core_types() -> None:
    assert sheetkit.EngineLiteral is not None
    assert sheetkit.FormatPriorityLiteral is not None
    assert sheetkit.ModeLiteral is not None

