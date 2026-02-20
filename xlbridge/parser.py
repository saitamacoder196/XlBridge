"""Parse XlBridge TXT format."""

import logging
import re
from dataclasses import dataclass

logger = logging.getLogger(__name__)

LINE_PATTERN = re.compile(r"^\[(.+?)\]!([A-Z]+\d+)\|(.*)$")


@dataclass
class CellEntry:
    sheet: str
    coord: str
    value: str


def parse_txt(file_path: str) -> list[CellEntry]:
    """Parse a XlBridge TXT file into a list of CellEntry objects.

    Lines starting with '#' are treated as comments and skipped.
    Empty lines are skipped.
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
            sheet, coord, value = match.groups()
            value = value.replace("\\n", "\n")
            entries.append(CellEntry(sheet=sheet, coord=coord, value=value))

    logger.info("Parsed %d entries from %s", len(entries), file_path)
    return entries
