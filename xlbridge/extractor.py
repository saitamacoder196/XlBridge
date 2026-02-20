"""Feature 1: Extract Excel cell content to TXT."""

import logging
from datetime import date
from pathlib import Path

import openpyxl

from xlbridge.utils import cell_coord, is_merged_slave

logger = logging.getLogger(__name__)


def extract(input_path: str, output_path: str, sheet_names: list[str] | None = None) -> None:
    """Extract non-empty cells from an Excel file to a TXT file.

    Args:
        input_path: Path to the source .xlsx file.
        output_path: Path for the output .txt file.
        sheet_names: Optional list of sheet names to extract. None means all sheets.
    """
    wb = openpyxl.load_workbook(input_path, data_only=True)
    source_name = Path(input_path).name
    lines: list[str] = []

    sheets = sheet_names if sheet_names else wb.sheetnames
    for name in sheets:
        if name not in wb.sheetnames:
            logger.warning("Sheet '%s' not found, skipping.", name)
            continue
        ws = wb[name]
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                if is_merged_slave(ws, col, row):
                    continue
                cell = ws.cell(row=row, column=col)
                if cell.value is None or str(cell.value).strip() == "":
                    continue
                coord = cell_coord(col, row)
                value = str(cell.value).replace("\n", "\\n")
                lines.append(f"[{name}]!{coord}|{value}")

    wb.close()

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"# XlBridge Export\n")
        f.write(f"# Source: {source_name}\n")
        f.write(f"# Date: {date.today().isoformat()}\n")
        f.write(f"# Encoding: UTF-8\n")
        f.write(f"\n")
        for line in lines:
            f.write(line + "\n")

    logger.info("Extracted %d cells to %s", len(lines), output_path)
