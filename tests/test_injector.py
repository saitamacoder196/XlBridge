"""Tests for the injector."""

import openpyxl

from xlbridge.injector import inject


def _create_test_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Original"
    ws["B1"] = "Keep"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)
    wb.save(str(path))
    wb.close()


def _write_txt(path, content):
    path.write_text(content, encoding="utf-8")


def test_inject_basic(tmp_path):
    xlsx = tmp_path / "test.xlsx"
    txt = tmp_path / "trans.txt"
    _create_test_xlsx(xlsx)
    _write_txt(txt, "[Sheet1]!A1|Translated\n")

    output = inject(str(xlsx), str(txt))
    wb = openpyxl.load_workbook(output)

    assert wb["Sheet1"]["A1"].value == "Translated"
    assert wb["Sheet1"]["B1"].value == "Keep"
    wb.close()


def test_inject_preserves_format(tmp_path):
    xlsx = tmp_path / "test.xlsx"
    txt = tmp_path / "trans.txt"
    _create_test_xlsx(xlsx)
    _write_txt(txt, "[Sheet1]!A1|New\n")

    output = inject(str(xlsx), str(txt))
    wb = openpyxl.load_workbook(output)

    assert wb["Sheet1"]["A1"].value == "New"
    assert wb["Sheet1"]["A1"].font.bold is True
    assert wb["Sheet1"]["A1"].font.size == 14
    wb.close()


def test_inject_default_output_name(tmp_path):
    xlsx = tmp_path / "test.xlsx"
    txt = tmp_path / "trans.txt"
    _create_test_xlsx(xlsx)
    _write_txt(txt, "[Sheet1]!A1|Test\n")

    output = inject(str(xlsx), str(txt))
    assert "_translated" in output


def test_inject_missing_sheet_skipped(tmp_path):
    xlsx = tmp_path / "test.xlsx"
    txt = tmp_path / "trans.txt"
    _create_test_xlsx(xlsx)
    _write_txt(txt, "[NoSheet]!A1|Test\n")

    output = inject(str(xlsx), str(txt))
    wb = openpyxl.load_workbook(output)
    assert wb["Sheet1"]["A1"].value == "Original"
    wb.close()
