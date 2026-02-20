"""Feature 2: Inject translated content back into Excel."""

import logging
import shutil
from pathlib import Path

import openpyxl

from xlbridge.parser import parse_txt

logger = logging.getLogger(__name__)


def inject(input_path: str, translation_path: str, output_path: str | None = None) -> str:
    """Inject translated text from a TXT file into an Excel file.

    Creates a new file; never overwrites the original.

    Args:
        input_path: Path to the original .xlsx file.
        translation_path: Path to the translated .txt file.
        output_path: Path for the output file. Defaults to *_translated.xlsx.

    Returns:
        The path of the output file.
    """
    if output_path is None:
        p = Path(input_path)
        output_path = str(p.with_name(f"{p.stem}_translated{p.suffix}"))

    # Copy original to output first so all formatting is preserved
    shutil.copy2(input_path, output_path)

    entries = parse_txt(translation_path)
    wb = openpyxl.load_workbook(output_path)

    applied = 0
    for entry in entries:
        if entry.sheet not in wb.sheetnames:
            logger.warning("Sheet '%s' not found, skipping entry %s", entry.sheet, entry.coord)
            continue
        ws = wb[entry.sheet]
        ws[entry.coord] = entry.value
        applied += 1

    wb.save(output_path)
    wb.close()
    logger.info("Injected %d/%d entries into %s", applied, len(entries), output_path)
    return output_path
