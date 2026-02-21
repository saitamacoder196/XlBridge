"""Tests for the TXT parser."""

import os
import tempfile

from xlbridge.parser import parse_txt


def _write_tmp(content: str) -> str:
    fd, path = tempfile.mkstemp(suffix=".txt")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(content)
    return path


# ─── 2-column format (backward-compatible) ───────────────────────────────────

def test_parse_basic():
    path = _write_tmp(
        "# comment\n"
        "\n"
        "[Sheet1]!A1|Hello\n"
        "[Sheet1]!B2|World\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 2
    assert entries[0].sheet == "Sheet1"
    assert entries[0].coord == "A1"
    assert entries[0].value == "Hello"
    assert entries[0].en is None
    assert entries[0].vi is None
    assert entries[1].coord == "B2"
    assert entries[1].value == "World"


def test_parse_multiline_escaped():
    path = _write_tmp("[Sheet1]!A1|Line1\\nLine2\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 1
    assert entries[0].value == "Line1\nLine2"


def test_parse_skips_invalid_lines():
    path = _write_tmp(
        "# header\n"
        "bad line\n"
        "[Sheet1]!A1|OK\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 1
    assert entries[0].value == "OK"


def test_parse_japanese():
    path = _write_tmp("[Sheet1]!A1|プロジェクト名\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].value == "プロジェクト名"


# ─── 4-column format ──────────────────────────────────────────────────────────

def test_parse_4col_basic():
    path = _write_tmp(
        "[変更履歴]!A1|変更履歴|Change history|Lịch sử thay đổi\n"
        "[変更履歴]!B2|共通関数ID|Common function ID|ID hàm chung\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert len(entries) == 2

    e0 = entries[0]
    assert e0.sheet == "変更履歴"
    assert e0.coord == "A1"
    assert e0.value == "変更履歴"
    assert e0.en == "Change history"
    assert e0.vi == "Lịch sử thay đổi"

    e1 = entries[1]
    assert e1.value == "共通関数ID"
    assert e1.en == "Common function ID"
    assert e1.vi == "ID hàm chung"


def test_parse_4col_escaped_newline_in_translation():
    path = _write_tmp("[Sheet1]!A1|原文|Line1\\nLine2|Dòng1\\nDòng2\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].value == "原文"
    assert entries[0].en == "Line1\nLine2"
    assert entries[0].vi == "Dòng1\nDòng2"


def test_parse_4col_missing_vi():
    """Only 3 columns present (no VI column) → vi is None."""
    path = _write_tmp("[Sheet1]!A1|Original|English only\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].value == "Original"
    assert entries[0].en == "English only"
    assert entries[0].vi is None


def test_parse_mixed_2col_and_4col():
    """A file may contain both formats (e.g. partially translated)."""
    path = _write_tmp(
        "[Sheet1]!A1|Original only\n"
        "[Sheet1]!A2|原文|English|Tiếng Việt\n"
    )
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].en is None
    assert entries[0].vi is None
    assert entries[1].en == "English"
    assert entries[1].vi == "Tiếng Việt"


def test_parse_4col_pipe_separator_is_column_boundary():
    """|  in the 4-column format is a column separator, not part of the value."""
    path = _write_tmp("[Sheet1]!A1|col1|col2|col3\n")
    entries = parse_txt(path)
    os.unlink(path)

    assert entries[0].value == "col1"
    assert entries[0].en == "col2"
    assert entries[0].vi == "col3"
