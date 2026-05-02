"""Tests for the color module."""

from markdownpick.colors import (
    CSS_COLORS,
    ColorResolver,
    _rgb_distance,
    docx_color_name,
    xlsx_color_name,
)


def test_rgb_distance_identical():
    assert _rgb_distance(0, 0, 0, 0, 0, 0) == 0.0
    assert _rgb_distance(255, 255, 255, 255, 255, 255) == 0.0


def test_rgb_distance_different():
    d = _rgb_distance(255, 0, 0, 0, 0, 255)
    assert d > 0


def test_resolver_exact_match():
    resolver = ColorResolver(CSS_COLORS)
    # Red = FF0000
    assert resolver.exact_match("FF0000") == "red"


def test_resolver_nearest_name_red_adjacent():
    resolver = ColorResolver(CSS_COLORS)
    name = resolver.nearest_name(254, 1, 0)
    assert name == "red"


def test_resolver_nearest_name_blue_adjacent():
    resolver = ColorResolver(CSS_COLORS)
    name = resolver.nearest_name(1, 0, 254)
    assert "blue" in name.lower() or name == "blue"


def test_resolver_fallback_to_hex():
    resolver = ColorResolver(CSS_COLORS)
    result = resolver.resolve("ZZZZZZ")
    assert result == "#ZZZZZZ"


def test_docx_color_name_exact():
    assert docx_color_name("FF0000") == "red"


def test_docx_color_name_none():
    assert docx_color_name(None) is None
    assert docx_color_name("") is None
    assert docx_color_name("000000") is None


def test_docx_color_name_near_black():
    name = docx_color_name("0A0F1E")
    assert name is not None


def test_xlsx_color_name_exact():
    assert xlsx_color_name("FFFF0000") == "red"


def test_xlsx_color_name_strips_alpha():
    assert xlsx_color_name("FFFF0000") == "red"


def test_xlsx_color_name_none():
    assert xlsx_color_name(None) is None
    assert xlsx_color_name("FF000000") is None


def test_colors_has_expected_families():
    names = {n.lower() for n in CSS_COLORS}
    assert "red" in names
    assert "blue" in names
    assert "green" in names
    assert "yellow" in names
    assert "orange" in names
    assert "purple" in names
    assert "black" in names
    assert "white" in names
    assert "gray" in names
    assert "brown" in names


def test_nearest_magenta_to_purple_family():
    resolver = ColorResolver(CSS_COLORS)
    name = resolver.nearest_name(128, 0, 128)
    assert name == "purple"


def test_nearest_gold():
    resolver = ColorResolver(CSS_COLORS)
    name = resolver.nearest_name(255, 215, 0)
    assert name == "gold"
