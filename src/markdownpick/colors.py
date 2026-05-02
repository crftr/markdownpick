"""Comprehensive color name resolution with nearest-color fallback.

Supports both docx (6-char hex RGB) and xlsx (8-char ARGB) formats.
Uses Euclidean distance in RGB space for nearest-color matching.
"""

from __future__ import annotations

import math

# ── 150+ named CSS/X11 colors ────────────────────────────────────────────────

CSS_COLORS: dict[str, tuple[int, int, int]] = {
    # Blacks
    "black": (0, 0, 0),
    "very dark": (10, 10, 10),
    "near-black": (20, 20, 20),
    "charcoal": (51, 51, 51),

    # Grays
    "dimgray": (105, 105, 105),
    "gray": (128, 128, 128),
    "dark gray": (130, 130, 130),
    "darkslategray": (47, 79, 79),
    "lightslategray": (119, 136, 153),
    "slategray": (112, 128, 144),
    "darkslateblue": (72, 61, 139),
    "silver": (192, 192, 192),
    "lightgray": (211, 211, 211),
    "light gray": (211, 211, 211),
    "gainsboro": (220, 220, 220),
    "whitesmoke": (245, 245, 245),
    "snow": (255, 250, 250),
    "ivory": (255, 255, 240),
    "offwhite": (245, 245, 245),
    "white": (255, 255, 255),
    "seashell": (255, 245, 238),
    "floralwhite": (255, 250, 240),

    # Reds
    "darkred": (139, 0, 0),
    "dark red": (139, 0, 0),
    "maroon": (128, 0, 0),
    "firebrick": (178, 34, 34),
    "crimson": (220, 20, 60),
    "red": (255, 0, 0),
    "indianred": (205, 92, 92),
    "lightcoral": (240, 128, 128),
    "salmon": (250, 128, 114),
    "darksalmon": (233, 150, 122),
    "lightsalmon": (255, 160, 122),
    "orangered": (255, 69, 0),
    "tomato": (255, 99, 71),
    "coral": (255, 127, 80),
    "mistyrose": (255, 228, 225),

    # Pinks / Roses
    "deeppink": (255, 20, 147),
    "hotpink": (255, 105, 180),
    "palevioletred": (219, 112, 147),
    "mediumvioletred": (199, 21, 133),
    "pink": (255, 192, 203),
    "lightpink": (255, 182, 193),
    "palepink": (248, 188, 178),
    "rosybrown": (188, 143, 143),
    "lavenderblush": (255, 240, 245),

    # Oranges
    "darkorange": (255, 140, 0),
    "orange": (255, 165, 0),
    "sandybrown": (244, 164, 96),
    "chocolate": (210, 105, 30),
    "peru": (205, 133, 63),
    "papayawhip": (255, 239, 213),
    "moccasin": (255, 228, 181),
    "blanchedalmond": (255, 235, 205),
    "navajowhite": (255, 222, 173),
    "peachpuff": (255, 218, 185),

    # Yellows / Golds
    "darkgoldenrod": (184, 134, 11),
    "goldenrod": (218, 165, 32),
    "darkkhaki": (189, 183, 107),
    "palegoldenrod": (238, 232, 170),
    "khaki": (240, 230, 140),
    "lightgoldenrodyellow": (250, 250, 210),
    "lemonchiffon": (255, 250, 205),
    "lightyellow": (255, 255, 224),
    "yellow": (255, 255, 0),
    "gold": (255, 215, 0),
    "yellowgreen": (154, 205, 50),

    # Greens
    "darkgreen": (0, 100, 0),
    "green": (0, 128, 0),
    "olivedrab": (107, 142, 35),
    "olive": (128, 128, 0),
    "darkolivegreen": (85, 107, 47),
    "forestgreen": (34, 139, 34),
    "seagreen": (46, 139, 87),
    "lightseagreen": (32, 178, 170),
    "mediumseagreen": (60, 179, 113),
    "limegreen": (50, 205, 50),
    "lime": (0, 255, 0),
    "springgreen": (0, 255, 127),
    "mediumspringgreen": (0, 250, 154),
    "lawngreen": (124, 252, 0),
    "chartreuse": (127, 255, 0),
    "greenyellow": (173, 255, 47),
    "palegreen": (152, 251, 152),
    "lightgreen": (144, 238, 144),
    "darkseagreen": (143, 205, 169),
    "honeydew": (240, 255, 240),
    "mintcream": (245, 255, 250),

    # Cyans / Teals / Aquas
    "darkcyan": (0, 139, 139),
    "teal": (0, 128, 128),
    "mediumturquoise": (72, 209, 204),
    "turquoise": (64, 224, 208),
    "aquamarine": (127, 255, 212),
    "mediumaquamarine": (102, 205, 170),
    "paleturquoise": (175, 238, 238),
    "lightcyan": (224, 255, 255),
    "aqua": (0, 255, 255),
    "cyan": (0, 255, 255),

    # Blues
    "midnightblue": (25, 25, 112),
    "navy": (0, 0, 128),
    "darkblue": (0, 0, 139),
    "mediumblue": (0, 0, 205),
    "blue": (0, 0, 255),
    "royalblue": (65, 105, 225),
    "cornflowerblue": (100, 149, 237),
    "steelblue": (70, 130, 180),
    "lightsteelblue": (176, 196, 222),
    "dodgerblue": (30, 144, 255),
    "deepskyblue": (0, 191, 255),
    "skyblue": (135, 206, 235),
    "lightskyblue": (135, 206, 250),
    "lightblue": (173, 216, 230),
    "powderblue": (176, 224, 230),
    "aliceblue": (240, 248, 255),
    "cadetblue": (95, 158, 160),

    # Purples / Violets / Magentas / Indigos
    "indigo": (75, 0, 130),
    "darkviolet": (148, 0, 211),
    "darkmagenta": (139, 0, 139),
    "purple": (128, 0, 128),
    "darkorchid": (153, 50, 204),
    "mediumorchid": (186, 85, 211),
    "blueviolet": (138, 43, 226),
    "fuchsia": (255, 0, 255),
    "magenta": (255, 0, 255),
    "violet": (238, 130, 238),
    "orchid": (218, 112, 214),
    "plum": (221, 160, 221),
    "thistle": (216, 191, 216),
    "lavender": (230, 230, 250),
    "mediumpurple": (147, 112, 219),
    "slateblue": (106, 90, 205),
    "mediumslateblue": (123, 104, 238),
    "heliotrope": (223, 115, 255),
    "wisteria": (201, 160, 220),

    # Browns / Tans
    "saddlebrown": (139, 69, 19),
    "brown": (165, 42, 42),
    "sienna": (160, 82, 45),
    "burlywood": (222, 184, 135),
    "tan": (210, 180, 140),
    "wheat": (245, 222, 179),
    "bisque": (255, 228, 196),
    "antiquewhite": (250, 235, 215),
    "linen": (250, 240, 230),
    "beige": (245, 245, 220),
    "oldlace": (253, 245, 230),
    "cornsilk": (255, 248, 220),
    "cream": (255, 253, 208),

    # Brand / theme colors
    "dark navy": (15, 52, 96),
    "hyperlink": (0, 112, 192),
    "followed hyperlink": (128, 0, 128),
    "accent 1": (68, 114, 196),
    "accent 2": (112, 173, 71),
    "accent 3": (255, 192, 0),
    "accent 4": (255, 153, 0),
    "accent 5": (192, 80, 77),
    "accent 6": (155, 86, 173),
}

# ── Color resolver ────────────────────────────────────────────────────────────


def _rgb_distance(
    r1: int, g1: int, b1: int,
    r2: int, g2: int, b2: int,
) -> float:
    """Euclidean distance between two RGB colors."""
    return math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2)


class ColorResolver:
    """Precomputed color lookup with nearest-color fallback."""

    def __init__(self, color_dict: dict[str, tuple[int, int, int]]) -> None:
        self._colors = {name.lower().strip(): rgb for name, rgb in color_dict.items()}
        self._rgb_list = list(self._colors.items())

    def exact_match(self, hex_str: str) -> str | None:
        """Look up exact 6-char hex RGB. Returns name or None."""
        key = hex_str.upper().strip()
        if not key or key in ("000000",):
            return None
        for name, (r, g, b) in self._rgb_list:
            candidate = f"{r:02X}{g:02X}{b:02X}"
            if candidate == key:
                return name
        return None

    def nearest_name(self, r: int, g: int, b: int) -> str:
        """Find the named color closest to the given RGB values."""
        best_name = f"#{r:02X}{g:02X}{b:02X}"
        best_dist = float("inf")
        for name, (cr, cg, cb) in self._rgb_list:
            dist = _rgb_distance(r, g, b, cr, cg, cb)
            if dist < best_dist:
                best_dist = dist
                best_name = name
        return best_name

    def resolve(self, hex_str: str) -> str:
        """Resolve a hex color string to a human-readable name.

        Tries exact match first, falls back to nearest color.
        """
        key = hex_str.upper().strip()
        if not key or key in ("000000",):
            return ""

        exact = self.exact_match(key)
        if exact:
            return exact

        if len(key) == 6:
            try:
                r, g, b = int(key[0:2], 16), int(key[2:4], 16), int(key[4:6], 16)
            except ValueError:
                return f"#{key}"
        else:
            return f"#{key}"

        nearest = self.nearest_name(r, g, b)
        return nearest


# ── Module-level resolvers ────────────────────────────────────────────────────

_docx_resolver = ColorResolver(CSS_COLORS)
_xlsx_resolver = ColorResolver(CSS_COLORS)


def docx_color_name(hex_6: str | None) -> str | None:
    """Resolve a 6-char hex RGB color (docx format) to a name.

    Returns empty string for black/None.
    """
    if not hex_6:
        return None
    cleaned = hex_6.strip()
    if not cleaned or cleaned.upper() == "000000":
        return None
    return _docx_resolver.resolve(cleaned)


def xlsx_color_name(argb_8: str | None) -> str | None:
    """Resolve an 8-char ARGB color (xlsx format) to a name.

    Strips alpha prefix, then resolves the RGB portion.
    Returns empty string for black/None.
    """
    if not argb_8:
        return None
    cleaned = argb_8.strip().upper()
    if not cleaned:
        return None
    # Strip alpha channel
    if len(cleaned) == 8:
        cleaned = cleaned[2:]
    if cleaned == "000000":
        return None
    return _xlsx_resolver.resolve(cleaned)
