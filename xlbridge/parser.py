"""Parse XlBridge TXT format.

Supported address types
-----------------------
Cell  :  [SheetName]!A1|value
Shape :  [SheetName]!shape:TextBox 1|text
Note  :  [SheetName]!note:A1|comment

Supported column counts
-----------------------
2-column (original / backward-compatible):
    [SheetName]!addr|OriginalValue

4-column (translated):
    [SheetName]!addr|OriginalValue|EnglishValue|VietnameseValue

The pipe character ``|`` is used as a column separator; individual
column values must not contain a literal pipe.
"""

import logging
import re
from dataclasses import dataclass, field
from typing import Literal

logger = logging.getLogger(__name__)

# Address patterns
_CELL_ADDR  = r'[A-Za-z]+\d+'               # e.g. A1, BZ100
_SHAPE_ADDR = r'shape:[^|\r\n]+'            # e.g. shape:TextBox 1
_NOTE_ADDR  = r'note:[A-Za-z]+\d+'          # e.g. note:A1

LINE_PATTERN = re.compile(
    rf'^\[(.+?)\]!({_CELL_ADDR}|{_SHAPE_ADDR}|{_NOTE_ADDR})\|(.+)$'
)


def _unescape(s: str) -> str:
    """Replace literal '\\n' sequences with real newlines."""
    return s.replace('\\n', '\n')


@dataclass
class CellEntry:
    sheet: str
    coord: str   # "A1"  |  "shape:TextBox 1"  |  "note:A1"
    value: str   # original text (column 1)
    en: str | None = field(default=None)   # English translation (column 2)
    vi: str | None = field(default=None)   # Vietnamese translation (column 3)

    # ── type helpers ─────────────────────────────────────────────────────────

    @property
    def entry_type(self) -> Literal['cell', 'shape', 'note']:
        if self.coord.startswith('shape:'):
            return 'shape'
        if self.coord.startswith('note:'):
            return 'note'
        return 'cell'

    @property
    def shape_name(self) -> str:
        """Shape name — only valid when ``entry_type == 'shape'``."""
        return self.coord[len('shape:'):]

    @property
    def note_coord(self) -> str:
        """Cell coordinate — only valid when ``entry_type == 'note'``."""
        return self.coord[len('note:'):]


def parse_txt(file_path: str) -> list[CellEntry]:
    """Parse a XlBridge TXT file into a list of CellEntry objects.

    Lines starting with '#' are treated as comments and skipped.
    Empty lines are skipped.
    Both 2-column and 4-column formats are accepted in the same file.
    All three address types (cell, shape, note) are accepted.
    """
    entries: list[CellEntry] = []

    with open(file_path, 'r', encoding='utf-8') as f:
        for line_num, raw_line in enumerate(f, start=1):
            line = raw_line.rstrip('\n').rstrip('\r')
            if not line or line.startswith('#'):
                continue

            match = LINE_PATTERN.match(line)
            if not match:
                logger.warning('Line %d: invalid format, skipping: %s', line_num, line)
                continue

            sheet, coord, rest = match.groups()
            parts = rest.split('|')

            value = _unescape(parts[0])
            en    = _unescape(parts[1]) if len(parts) > 1 and parts[1] else None
            vi    = _unescape(parts[2]) if len(parts) > 2 and parts[2] else None

            entries.append(CellEntry(sheet=sheet, coord=coord, value=value, en=en, vi=vi))

    has_en = sum(1 for e in entries if e.en is not None)
    has_vi = sum(1 for e in entries if e.vi is not None)
    cells  = sum(1 for e in entries if e.entry_type == 'cell')
    shapes = sum(1 for e in entries if e.entry_type == 'shape')
    notes  = sum(1 for e in entries if e.entry_type == 'note')
    logger.info(
        'Parsed %d entries from %s  (cells=%d, shapes=%d, notes=%d, en=%d, vi=%d)',
        len(entries), file_path, cells, shapes, notes, has_en, has_vi,
    )
    return entries
