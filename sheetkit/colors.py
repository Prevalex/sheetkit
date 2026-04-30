import re
import colorsys
import math
from collections.abc import Mapping
from typing import Any

from .types import ColorInput

#
################# [ STANDARD COLORS SECTION ] #####################################
#
# The complete dictionary of standard CSS colors (148 colors) is the official set from the W3C (SVG/CSS Level 4).
# See the W3C CSS Color Module specification: https://www.w3.org/TR/css-color-4/#named-colors
CSS_COLORS = {
    "aliceblue": "F0F8FF",
    "antiquewhite": "FAEBD7",
    "aqua": "00FFFF",
    "aquamarine": "7FFFD4",
    "azure": "F0FFFF",
    "beige": "F5F5DC",
    "bisque": "FFE4C4",
    "black": "000000",
    "blanchedalmond": "FFEBCD",
    "blue": "0000FF",
    "blueviolet": "8A2BE2",
    "brown": "A52A2A",
    "burlywood": "DEB887",
    "cadetblue": "5F9EA0",
    "chartreuse": "7FFF00",
    "chocolate": "D2691E",
    "coral": "FF7F50",
    "cornflowerblue": "6495ED",
    "cornsilk": "FFF8DC",
    "crimson": "DC143C",
    "cyan": "00FFFF",
    "darkblue": "00008B",
    "darkcyan": "008B8B",
    "darkgoldenrod": "B8860B",
    "darkgray": "A9A9A9",
    "darkgreen": "006400",
    "darkgrey": "A9A9A9",
    "darkkhaki": "BDB76B",
    "darkmagenta": "8B008B",
    "darkolivegreen": "556B2F",
    "darkorange": "FF8C00",
    "darkorchid": "9932CC",
    "darkred": "8B0000",
    "darksalmon": "E9967A",
    "darkseagreen": "8FBC8F",
    "darkslateblue": "483D8B",
    "darkslategray": "2F4F4F",
    "darkslategrey": "2F4F4F",
    "darkturquoise": "00CED1",
    "darkviolet": "9400D3",
    "deeppink": "FF1493",
    "deepskyblue": "00BFFF",
    "dimgray": "696969",
    "dimgrey": "696969",
    "dodgerblue": "1E90FF",
    "firebrick": "B22222",
    "floralwhite": "FFFAF0",
    "forestgreen": "228B22",
    "fuchsia": "FF00FF",
    "gainsboro": "DCDCDC",
    "ghostwhite": "F8F8FF",
    "gold": "FFD700",
    "goldenrod": "DAA520",
    "gray": "808080",
    "green": "008000",
    "greenyellow": "ADFF2F",
    "grey": "808080",
    "honeydew": "F0FFF0",
    "hotpink": "FF69B4",
    "indianred": "CD5C5C",
    "indigo": "4B0082",
    "ivory": "FFFFF0",
    "khaki": "F0E68C",
    "lavender": "E6E6FA",
    "lavenderblush": "FFF0F5",
    "lawngreen": "7CFC00",
    "lemonchiffon": "FFFACD",
    "lightblue": "ADD8E6",
    "lightcoral": "F08080",
    "lightcyan": "E0FFFF",
    "lightgoldenrodyellow": "FAFAD2",
    "lightgray": "D3D3D3",
    "lightgreen": "90EE90",
    "lightgrey": "D3D3D3",
    "lightpink": "FFB6C1",
    "lightsalmon": "FFA07A",
    "lightseagreen": "20B2AA",
    "lightskyblue": "87CEFA",
    "lightslategray": "778899",
    "lightslategrey": "778899",
    "lightsteelblue": "B0C4DE",
    "lightyellow": "FFFFE0",
    "lime": "00FF00",
    "limegreen": "32CD32",
    "linen": "FAF0E6",
    "magenta": "FF00FF",
    "maroon": "800000",
    "mediumaquamarine": "66CDAA",
    "mediumblue": "0000CD",
    "mediumorchid": "BA55D3",
    "mediumpurple": "9370DB",
    "mediumseagreen": "3CB371",
    "mediumslateblue": "7B68EE",
    "mediumspringgreen": "00FA9A",
    "mediumturquoise": "48D1CC",
    "mediumvioletred": "C71585",
    "midnightblue": "191970",
    "mintcream": "F5FFFA",
    "mistyrose": "FFE4E1",
    "moccasin": "FFE4B5",
    "navajowhite": "FFDEAD",
    "navy": "000080",
    "oldlace": "FDF5E6",
    "olive": "808000",
    "olivedrab": "6B8E23",
    "orange": "FFA500",
    "orangered": "FF4500",
    "orchid": "DA70D6",
    "palegoldenrod": "EEE8AA",
    "palegreen": "98FB98",
    "paleturquoise": "AFEEEE",
    "palevioletred": "DB7093",
    "papayawhip": "FFEFD5",
    "peachpuff": "FFDAB9",
    "peru": "CD853F",
    "pink": "FFC0CB",
    "plum": "DDA0DD",
    "powderblue": "B0E0E6",
    "purple": "800080",
    "rebeccapurple": "663399",
    "red": "FF0000",
    "rosybrown": "BC8F8F",
    "royalblue": "4169E1",
    "saddlebrown": "8B4513",
    "salmon": "FA8072",
    "sandybrown": "F4A460",
    "seagreen": "2E8B57",
    "seashell": "FFF5EE",
    "sienna": "A0522D",
    "silver": "C0C0C0",
    "skyblue": "87CEEB",
    "slateblue": "6A5ACD",
    "slategray": "708090",
    "slategrey": "708090",
    "snow": "FFFAFA",
    "springgreen": "00FF7F",
    "steelblue": "4682B4",
    "tan": "D2B48C",
    "teal": "008080",
    "thistle": "D8BFD8",
    "tomato": "FF6347",
    "turquoise": "40E0D0",
    "violet": "EE82EE",
    "wheat": "F5DEB3",
    "white": "FFFFFF",
    "whitesmoke": "F5F5F5",
    "yellow": "FFFF00",
    "yellowgreen": "9ACD32",
}


# Хелпер: CSS → HEX
def css_color_to_hex(name: str) -> str:
    """
    Converts CSS color name ('red', 'steelblue') -> HEX 'RRGGBB'.
    """
    if not isinstance(name, str):
        raise TypeError(f"Color name must be string, got {type(name)}")

    key = name.lower().replace(" ", "")
    if key not in CSS_COLORS:
        raise ValueError(f"Unknown CSS color name: '{name}'")

    return CSS_COLORS[key]


#
################ [ СЕКЦИЯ HEX-RGB ] ##################################
#

# =====================================================
# Converting HEX <-> ARGB/RGB Colors
# --------------------------------------------
# xlsxwriter:
# Accepts colors in RGB HEX format:
# "FF0000" or "#FF0000"
# (both are acceptable).
# DOES NOT understand ARGB - Alpha transparency is not supported.
#
# openpyxl:
# Requires ARGB: "FFFF0000"
# (the first byte is alpha, almost always "FF").
# Accepts any variants without #. #
# That is:
# Library Format
# xlsxwriter "#RRGGBB" or "RRGGBB"
# openpyxl "AARRGGBB"
# ------------------------------------------------

# HEX → RGB
def hex_to_rgb(color: str) -> tuple[int, int, int]:
    """Converts '#RRGGBB' or 'RRGGBB' → (r, g, b)."""
    r, g, b, _ = hex_to_rgba(color)
    return r, g, b


# HEX → RGBA
def hex_to_rgba(color: str) -> tuple[int, int, int, int]:
    """
    Converts a color string from HEX format to (r, g, b, a).

    Supported formats:
    - 'RRGGBB'
    - '#RRGGBB'
    - '0xRRGGBB'
    - 'AARRGGBB'
    - '#AARRGGBB'

    Returns:
    (r, g, b, a) — components 0..255
    """

    if not isinstance(color, str):
        raise TypeError(f"color must be str, got: {type(color)}")

    s = color.strip()

    # Убираем возможные префиксы
    if s.startswith("#"):
        s = s[1:]
    if s.lower().startswith("0x"):
        s = s[2:]

    if len(s) == 6:
        # RRGGBB
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        a = 255
    elif len(s) == 8:
        # AARRGGBB
        a = int(s[0:2], 16)
        r = int(s[2:4], 16)
        g = int(s[4:6], 16)
        b = int(s[6:8], 16)
    else:
        raise ValueError(
            f"Invalid HEX color '{color}'. Must be 6 or 8 hex digits."
        )

    return r, g, b, a


# RGB → HEX
def rgb_to_hex(r: int, g: int, b: int, prefix: str = "") -> str:
    """
    Converts (r, g, b) → 'RRGGBB'
    prefix='#' converts to '#RRGGBB' for convenience in xlsxwriter.
    """
    return f"{prefix}{r:02X}{g:02X}{b:02X}"


# RGBA → HEX
def rgba_to_hex(r: int, g: int, b: int, a: int = 255, prefix: str = "") -> str:
    """
    Converts (r, g, b, a) → 'AARRGGBB'
    prefix='' can be left blank for openpyxl.
    """
    return f"{prefix}{a:02X}{r:02X}{g:02X}{b:02X}"


#
################ [ UNIVERSAL CONVERSION OF ANY COLOR REPRESENTATION TO HEX ] ###############################
#

# -----------------------------------------------------------------
# Universal converter: any format → HEX (RGB or ARGB)
#
# To be able to write:
#
# color_to_hex("tomato")
# color_to_hex("#aabbcc")
# color_to_hex("80FFAACC")
# color_to_hex((255,100,20))
# color_to_hex((200,120,50,120))
# -----------------------------------------------------------------
EXCEL_THEME_COLOR_ALIASES = {
    "text1": "dark1",
    "text2": "dark2",
    "background1": "light1",
    "background2": "light2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hyperlink": "hlink",
    "followedhyperlink": "folHlink",
    "followed_hyperlink": "folHlink",
}

THEME_COLOR_REF_RE = re.compile(
    r"^(?P<slot>[a-z][a-z0-9_ ]*)(?P<shift>[+-]\d{1,3})?$",
    re.IGNORECASE,
)


def _normalize_excel_theme_slot_name(name: str) -> str:
    """
    Casts the custom name theme-color to the internal key of colors-scheme.
    """
    key = re.sub(r"[\s_]+", "", name).lower()
    key = EXCEL_THEME_COLOR_ALIASES.get(key, key)
    if key not in {
        "dark1",
        "dark2",
        "light1",
        "light2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    }:
        raise ValueError(f"Unknown theme color slot: {name!r}")
    return key


def _resolve_exel_theme_info(theme: str | Mapping[str, Any] | None) -> Mapping[str, Any]:
    """
    Returns an Excel theme dictionary. If theme=None, use the Office Theme.
    """
    from .themes import get_theme

    if theme is None:
        return get_theme("office_theme")
    if isinstance(theme, str):
        return get_theme(theme)
    return theme


def _apply_theme_color_shift(base_color: str, shift: int) -> str:
    """
    Applies a simple Excel-like hue correction: +N -> mix with white, -N -> mix with black.
    """
    if shift == 0:
        return base_color

    amount = min(abs(shift), 100) / 100.0
    target = "FFFFFF" if shift > 0 else "000000"
    return _mix_hex_colors(base_color, target, amount)


def _resolve_theme_color_reference(
        value: str | tuple[str, str],
        *,
        theme: str | Mapping[str, Any] | None = None,
) -> str:
    """
    Converts the theme-color reference to a real RRGGBB HEX color.

    Supports:
    - ('wisp', 'accent3')
    - 'wisp:accent3'
    - 'Accent1+40'
    - 'Text1'
    - 'Background2-20'
    - 'Hyperlink'
    """
    theme_name_or_info: str | Mapping[str, Any] | None = theme
    ref_value: str

    if isinstance(value, tuple):
        theme_name_or_info, ref_value = value
    else:
        if ":" in value:
            theme_name, ref_value = value.split(":", 1)
            theme_name_or_info = theme_name.strip()
        else:
            ref_value = value

    match = THEME_COLOR_REF_RE.match(ref_value.strip())
    if not match:
        raise ValueError(f"Invalid theme color reference: {value!r}")

    slot_name = _normalize_excel_theme_slot_name(match.group("slot"))
    shift_group = match.group("shift")
    shift = int(shift_group) if shift_group else 0

    theme_info = _resolve_exel_theme_info(theme_name_or_info)
    colors = theme_info.get("colors", theme_info)
    if not isinstance(colors, Mapping):
        raise ValueError("Theme info must contain a 'colors' mapping")

    try:
        base_color = str(colors[slot_name]).upper()
    except KeyError as e:
        raise ValueError(f"Theme does not contain color slot {slot_name!r}") from e

    return _apply_theme_color_shift(base_color, shift)


def color_to_hex(
        color: ColorInput | tuple[str, str],
        *,
        with_alpha: bool = False,
        theme: str | Mapping[str, Any] | None = None,
) -> str:
    """
       Universal function:
    - CSS colors ('tomato')
    - HEX colors ('#FFAACC', 'FFAACC', '80FFAACC')
    - RGB (r,g,b)
    - RGBA (r,g,b,a)

    Returns RRGGBB or AARRGGBB depending on with_alpha.
    """

    # --- explicit theme reference ---
    if isinstance(color, tuple) and len(color) == 2 and all(isinstance(item, str) for item in color):
        base = _resolve_theme_color_reference(color, theme=theme)
        return f"FF{base}" if with_alpha else base

    # --- RGB/RGBA tuple ---
    if isinstance(color, tuple):
        if len(color) == 3:
            r, g, b = color
            a = 255
        elif len(color) == 4:
            r, g, b, a = color
        else:
            raise ValueError("Tuple color must be (r,g,b) or (r,g,b,a)")

        return (
            rgba_to_hex(r, g, b, a)
            if with_alpha
            else rgb_to_hex(r, g, b)
        )

    # --- string ---
    if isinstance(color, str):
        name = color.lower().strip()

        # CSS name?
        if name in CSS_COLORS:
            base = CSS_COLORS[name]
            return (
                f"FF{base}" if with_alpha else base
            )

        # theme color reference?
        try:
            base = _resolve_theme_color_reference(color, theme=theme)
        except ValueError:
            pass
        else:
            return f"FF{base}" if with_alpha else base

        # HEX?
        r, g, b, a = hex_to_rgba(name)
        return (
            rgba_to_hex(r, g, b, a)
            if with_alpha
            else rgb_to_hex(r, g, b)
        )

    raise TypeError("Unsupported color format")


def normalize_color_value(
        value: Any,
        *,
        with_alpha: bool = False,
        theme: str | Mapping[str, Any] | None = None,
) -> str | None:
    """
        Normalizes a color value to a canonical HEX format.
    We assume that a logical "normalized" color is simply a HEX string:
    without #, uppercase #, # RRGGBB, or AARRGGBB depending on with_alpha.

    Accepts:
    - HEX string: 'FFAACC', '#FFAACC', '80FFAACC', '0xFFAACC'
    - CSS name: 'red', 'steelblue', ...
    - tuple (r,g,b) or (r,g,b,a), where 0..255
    - theme name

    Returns:
    - 'RRGGBB' (with_alpha=False)
    - 'AARRGGBB' (with_alpha=True)
    - or None if value is None.

    Raises ValueError if the value is invalid.
    """
    if value is None:
        return None

    try:
        # color_to_hex уже умеет и строки, и tuple
        hex_str = color_to_hex(value, with_alpha=with_alpha, theme=theme)
    except (TypeError, ValueError) as e:
        raise ValueError(f"Invalid color value {value!r}: {e}") from e

    return hex_str


# ✅
# ✅=============== [👉 EXCEL COLOR CONVERSION SECTION] ================
# ✅

def validate_key_color_value(
        value: Any,
        key: str,
        *,
        theme: str | Mapping[str, Any] | None = None,
) -> None:
    """
    Checks that value is a valid color for the Boolean formatter.

    Allowed values:
    - HEX strings (6 or 8 digits, with or without #, 0x)
    - CSS names ('red', 'steelblue', ...)
    - (r,g,b) / (r,g,b,a), where 0..255

    Returns nothing, but raises ValueError on error.
    """
    if value is None:
        return None
    try:
        # for validation you can always require with_alpha=True - this is the strictest case
        _ = normalize_color_value(value, with_alpha=True, theme=theme)
    except ValueError as e:
        raise ValueError(f"Invalid color for key {key!r}: {value!r} ({e})") from e


def _mix_hex_colors(c1: str, c2: str, t: float) -> str:
    """
    Linearly blends two HEX colors (RRGGBB or prefixed with '#'):
    result = (1 - t) * c1 + t * c2, t in [0,1].
    Returns 'RRGGBB'.
    """
    r1, g1, b1, _ = hex_to_rgba(c1)
    r2, g2, b2, _ = hex_to_rgba(c2)
    r = int(round(r1 * (1.0 - t) + r2 * t))
    g = int(round(g1 * (1.0 - t) + g2 * t))
    b = int(round(b1 * (1.0 - t) + b2 * t))
    return rgb_to_hex(r, g, b)


def _relative_luminance(hex_color: str) -> float:
    """
    Calculates WCAG relative luminance for an RGB HEX color.
    """
    r, g, b, _ = hex_to_rgba(hex_color)

    def _to_linear(channel: int) -> float:
        c = channel / 255.0
        if c <= 0.03928:
            return c / 12.92
        return math.pow((c + 0.055) / 1.055, 2.4)

    r_lin = _to_linear(r)
    g_lin = _to_linear(g)
    b_lin = _to_linear(b)
    return (0.2126 * r_lin) + (0.7152 * g_lin) + (0.0722 * b_lin)


def _contrast_ratio(c1: str, c2: str) -> float:
    """
    Calculates WCAG contrast ratio between two RGB HEX colors.
    """
    l1 = _relative_luminance(c1)
    l2 = _relative_luminance(c2)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _pick_contrast_color(background: str, *, light_color: str, dark_color: str) -> str:
    """
    Picks the better-contrast color (light or dark) for a given background.
    """
    light_ratio = _contrast_ratio(background, light_color)
    dark_ratio = _contrast_ratio(background, dark_color)
    return light_color if light_ratio >= dark_ratio else dark_color


def _apply_excel_tint(base_color: str, tint: float) -> str:
    """
    Applies Excel-like tint using HLS luminance transform.

    tint range:
    -1.0 .. 0.0 -> darker
     0.0 .. 1.0 -> lighter
    """
    clamped_tint = max(-1.0, min(1.0, float(tint)))

    # RGB (0..1) -> HLS (0..1)
    r, g, b, _ = hex_to_rgba(base_color)
    r_f = r / 255.0
    g_f = g / 255.0
    b_f = b / 255.0
    h, l, s = colorsys.rgb_to_hls(r_f, g_f, b_f)

    # Excel logic on luminance.
    if clamped_tint < 0.0:
        l_new = l * (1.0 + clamped_tint)
    else:
        l_new = l * (1.0 - clamped_tint) + clamped_tint

    r_n, g_n, b_n = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l_new)), s)

    # Keep stable channel quantization for RGB conversion.
    r_out = int(round(r_n * 255.0))
    g_out = int(round(g_n * 255.0))
    b_out = int(round(b_n * 255.0))
    return rgb_to_hex(r_out, g_out, b_out)


def _excel_accent_percent(base_accent: str, percent: int) -> str:
    """
    Converts Excel UI style 'Accent N, P%' to a resulting RGB HEX color.

    Example:
    - 60% -> tint +0.4 (40% to white)
    - 20% -> tint +0.8 (80% to white)
    """
    clamped_percent = max(0, min(100, int(percent)))
    tint = 1.0 - (clamped_percent / 100.0)
    return _apply_excel_tint(base_accent, tint)
