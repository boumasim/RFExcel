from pathlib import Path

from openpyxl import Workbook

from rfexcel.RFExcelLibrary import RFExcelLibrary


def _make_offset_xlsx(path: str, start_col: int = 2) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.cell(row=1, column=start_col,     value="Name")
    ws.cell(row=1, column=start_col + 1, value="Score")
    ws.cell(row=2, column=start_col,     value="Alice")
    ws.cell(row=2, column=start_col + 1, value=90)
    ws.cell(row=3, column=start_col,     value="Bob")
    ws.cell(row=3, column=start_col + 1, value=75)
    wb.save(path)
    wb.close()


def _make_gap_xlsx(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.cell(row=1, column=1, value="Name")
    ws.cell(row=1, column=3, value="Score")
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=2, column=3, value=90)
    ws.cell(row=3, column=1, value="Bob")
    ws.cell(row=3, column=3, value=75)
    wb.save(path)
    wb.close()


# ---------------------------------------------------------------------------
# xlsx edit
# ---------------------------------------------------------------------------

class TestOffsetTableXlsxEdit:

    def test_correct_row_count(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset.xlsx")
        _make_offset_xlsx(path)
        lib.load_workbook(path)
        assert len(lib.get_rows()) == 2

    def test_header_keys_are_correct(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset.xlsx")
        _make_offset_xlsx(path)
        lib.load_workbook(path)
        rows = lib.get_rows()
        assert list(rows[0]) == ["Name", "Score"]

    def test_values_are_mapped_to_correct_columns(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset.xlsx")
        _make_offset_xlsx(path)
        lib.load_workbook(path)
        rows = lib.get_rows()
        assert rows[0]["Name"] == "Alice"
        assert rows[0]["Score"] == 90
        assert rows[1]["Name"] == "Bob"
        assert rows[1]["Score"] == 75

    def test_column_c_start(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset_c.xlsx")
        _make_offset_xlsx(path, start_col=3)
        lib.load_workbook(path)
        rows = lib.get_rows()
        assert rows[0]["Name"] == "Alice"
        assert rows[0]["Score"] == 90


# ---------------------------------------------------------------------------
# xlsx stream
# ---------------------------------------------------------------------------

class TestOffsetTableXlsxStream:

    def test_correct_row_count(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset.xlsx")
        _make_offset_xlsx(path)
        lib.load_workbook(path, read_only=True)
        assert len(lib.get_rows()) == 2

    def test_values_are_mapped_to_correct_columns(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "offset.xlsx")
        _make_offset_xlsx(path)
        lib.load_workbook(path, read_only=True)
        rows = lib.get_rows()
        assert rows[0]["Name"] == "Alice"
        assert rows[0]["Score"] == 90
        assert rows[1]["Name"] == "Bob"
        assert rows[1]["Score"] == 75


# ---------------------------------------------------------------------------
# xlsx edit
# ---------------------------------------------------------------------------

class TestGapColumnXlsxEdit:

    def test_correct_row_count(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "gap.xlsx")
        _make_gap_xlsx(path)
        lib.load_workbook(path)
        assert len(lib.get_rows()) == 2

    def test_header_keys_exclude_empty_column(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "gap.xlsx")
        _make_gap_xlsx(path)
        lib.load_workbook(path)
        rows = lib.get_rows()
        assert list(rows[0]) == ["Name", "Score"]

    def test_values_skip_gap_column_correctly(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "gap.xlsx")
        _make_gap_xlsx(path)
        lib.load_workbook(path)
        rows = lib.get_rows()
        assert rows[0]["Name"] == "Alice"
        assert rows[0]["Score"] == 90
        assert rows[1]["Name"] == "Bob"
        assert rows[1]["Score"] == 75


# ---------------------------------------------------------------------------
# xlsx stream
# ---------------------------------------------------------------------------

class TestGapColumnXlsxStream:

    def test_correct_row_count(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "gap.xlsx")
        _make_gap_xlsx(path)
        lib.load_workbook(path, read_only=True)
        assert len(lib.get_rows()) == 2

    def test_values_skip_gap_column_correctly(self, lib: RFExcelLibrary, tmp_path: Path):
        path = str(tmp_path / "gap.xlsx")
        _make_gap_xlsx(path)
        lib.load_workbook(path, read_only=True)
        rows = lib.get_rows()
        assert rows[0]["Name"] == "Alice"
        assert rows[0]["Score"] == 90
        assert rows[1]["Name"] == "Bob"
        assert rows[1]["Score"] == 75
