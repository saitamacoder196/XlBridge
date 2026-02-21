"""Feature 2: Inject translated content back into Excel.

Supported languages for the ``lang`` parameter
-----------------------------------------------
None  — use the original value (column 1); backward-compatible default.
"en"  — use the English translation (column 2).
        Falls back to the original value if the EN column is absent.
"vi"  — use the Vietnamese translation (column 3).
        Falls back to the original value if the VI column is absent.

Any other language string raises ``ValueError`` (not supported).
"""

import logging
import shutil
from pathlib import Path
from typing import Literal

import openpyxl

from xlbridge.parser import CellEntry, parse_txt

logger = logging.getLogger(__name__)

Lang = Literal["en", "vi"] | None

SUPPORTED_LANGS: set[str] = {"en", "vi"}


def _pick_value(entry: CellEntry, lang: Lang) -> tuple[str, bool]:
    """Select the cell value for the requested language.

    Returns:
        (value, used_fallback) — used_fallback is True when the requested
        translation column was absent and the original text was used instead.
    """
    if lang == "en":
        if entry.en is not None:
            return entry.en, False
        logger.debug("EN translation missing for %s!%s — using original", entry.sheet, entry.coord)
        return entry.value, True

    if lang == "vi":
        if entry.vi is not None:
            return entry.vi, False
        logger.debug("VI translation missing for %s!%s — using original", entry.sheet, entry.coord)
        return entry.value, True

    # lang is None → original value (backward-compatible)
    return entry.value, False


def inject(
    input_path: str,
    translation_path: str,
    output_path: str | None = None,
    lang: Lang = None,
) -> str:
    """Inject text from a TXT file into an Excel file.

    Creates a new file; never overwrites the original.

    Args:
        input_path:       Path to the original .xlsx file.
        translation_path: Path to the TXT file (2-column or 4-column format).
        output_path:      Path for the output file.
                          Defaults to ``*_<lang>.xlsx`` when lang is set,
                          otherwise ``*_translated.xlsx``.
        lang:             Target language — ``"en"``, ``"vi"``, or ``None``.
                          ``None`` injects the original value (column 1).

    Returns:
        The path of the output file.

    Raises:
        ValueError: If ``lang`` is not None, ``"en"``, or ``"vi"``.
    """
    if lang is not None and lang not in SUPPORTED_LANGS:
        raise ValueError(
            f"Language '{lang}' is not supported. "
            f"Supported values: {sorted(SUPPORTED_LANGS)} or None."
        )

    if output_path is None:
        p = Path(input_path)
        suffix = f"_{lang}" if lang else "_translated"
        output_path = str(p.with_name(f"{p.stem}{suffix}{p.suffix}"))

    # Copy original so all formatting / formulas are preserved
    shutil.copy2(input_path, output_path)

    entries = parse_txt(translation_path)
    wb = openpyxl.load_workbook(output_path)

    applied = 0
    skipped_sheet = 0
    fallback = 0

    for entry in entries:
        if entry.sheet not in wb.sheetnames:
            logger.warning("Sheet '%s' not found, skipping %s!%s", entry.sheet, entry.sheet, entry.coord)
            skipped_sheet += 1
            continue

        value, used_fallback = _pick_value(entry, lang)
        if used_fallback:
            fallback += 1

        wb[entry.sheet][entry.coord] = value
        applied += 1

    wb.save(output_path)
    wb.close()

    logger.info(
        "Injected %d/%d entries into %s  (lang=%s, fallback=%d, sheet_missing=%d)",
        applied, len(entries), output_path,
        lang or "original", fallback, skipped_sheet,
    )
    return output_path
