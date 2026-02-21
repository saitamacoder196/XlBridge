"""Parse XlBridge TXT format.

Supported line formats
----------------------
2-column (original / backward-compatible):
    [SheetName]!CellRef|OriginalValue

4-column (translated):
    [SheetName]!CellRef|OriginalValue|EnglishValue|VietnameseValue

The pipe character ``|`` is used as a column separator; individual
column values must not contain a literal pipe.
"""

import logging
import re
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

# Matches: [SheetName]!CellRef|<rest>
# <rest> is split on '|' afterwards to extract up to 3 columns.
LINE_PATTERN = re.compile(r"^\[(.+?)\]!([A-Za-z]+\d+)\|(.+)$")


def _unescape(s: str) -> str:
    """Replace literal '\\n' sequences with real newlines."""
    return s.replace("\\n", "\n")


@dataclass
class CellEntry:
    sheet: str
    coord: str
    value: str              # column 1 — original text
    en: str | None = field(default=None)  # column 2 — English translation
    vi: str | None = field(default=None)  # column 3 — Vietnamese translation


def parse_txt(file_path: str) -> list[CellEntry]:
    """Parse a XlBridge TXT file into a list of CellEntry objects.

    Lines starting with '#' are treated as comments and skipped.
    Empty lines are skipped.
    Both 2-column and 4-column formats are accepted in the same file.
    """
    entries: list[CellEntry] = []

    with open(file_path, "r", encoding="utf-8") as f:
        for line_num, raw_line in enumerate(f, start=1):
            line = raw_line.rstrip("\n").rstrip("\r")
            if not line or line.startswith("#"):
                continue

            match = LINE_PATTERN.match(line)
            if not match:
                logger.warning("Line %d: invalid format, skipping: %s", line_num, line)
                continue

            sheet, coord, rest = match.groups()
            parts = rest.split("|")

            value = _unescape(parts[0])
            en    = _unescape(parts[1]) if len(parts) > 1 and parts[1] else None
            vi    = _unescape(parts[2]) if len(parts) > 2 and parts[2] else None

            entries.append(CellEntry(sheet=sheet, coord=coord, value=value, en=en, vi=vi))

    has_en = sum(1 for e in entries if e.en is not None)
    has_vi = sum(1 for e in entries if e.vi is not None)
    logger.info(
        "Parsed %d entries from %s  (en=%d, vi=%d)",
        len(entries), file_path, has_en, has_vi,
    )
    return entries
