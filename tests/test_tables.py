"""Tests for the shared table renderer."""

from markdownpick.tables import markdown_table


def test_basic_table():
    result = markdown_table(["A", "B"], [["1", "2"], ["3", "4"]])
    lines = result.split("\n")
    assert len(lines) == 4
    assert "A" in lines[0]
    assert "B" in lines[0]
    assert "---" in lines[1]
    assert "1" in lines[2]
    assert "3" in lines[3]


def test_empty_rows():
    result = markdown_table(["A"], [])
    lines = result.split("\n")
    assert len(lines) == 2


def test_single_cell():
    result = markdown_table(["X"], [["y"]])
    assert "| X |" in result
    assert "| y |" in result


def test_pipe_escaping():
    result = markdown_table(["A|B"], [["1|2"]])
    assert "\\|" in result


def test_short_row_padded():
    result = markdown_table(["A", "B", "C"], [["1"]])
    lines = result.split("\n")
    assert lines[2].count("|") >= 4  # 3 cols + leading = 4 pipes minimum
