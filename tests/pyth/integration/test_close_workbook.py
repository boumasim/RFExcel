"""Integration tests for the Close Workbook keyword.

Covers:
  - Positive: close after load, close after create; active_workbook becomes None.
  - Negative / edge: close when nothing is loaded (must be silent); close twice
    in a row (must not raise); workbook is not accessible after close.
"""
import pytest

from tests.pyth.conftest import CSV_FILE, XLS_FILE, XLSX_FILE

# ─── positive ─────────────────────────────────────────────────────────────────

class TestCloseWorkbookPositive:

    def test_close_after_load_xlsx_sets_active_workbook_to_none(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        assert lib._active_workbook is None

    def test_close_after_load_xlsx_stream_sets_active_workbook_to_none(self, lib):
        lib.load_workbook(XLSX_FILE, read_only=True)
        lib.close()
        assert lib._active_workbook is None

    def test_close_after_load_xls_sets_active_workbook_to_none(self, lib):
        lib.load_workbook(XLS_FILE)
        lib.close()
        assert lib._active_workbook is None

    def test_close_after_load_csv_sets_active_workbook_to_none(self, lib):
        lib.load_workbook(CSV_FILE)
        lib.close()
        assert lib._active_workbook is None

    def test_close_after_create_xlsx_sets_active_workbook_to_none(self, lib, tmp_path):
        lib.create_workbook(str(tmp_path / "new.xlsx"))
        lib.close()
        assert lib._active_workbook is None

    def test_close_after_create_csv_sets_active_workbook_to_none(self, lib, tmp_path):
        lib.create_workbook(str(tmp_path / "new.csv"))
        lib.close()
        assert lib._active_workbook is None


# ─── negative / edge ──────────────────────────────────────────────────────────

class TestCloseWorkbookEdge:

    def test_close_without_open_does_not_raise(self, lib):
        """Calling Close Workbook when nothing is loaded must be a no-op."""
        lib.close()  # must not raise
        assert lib._active_workbook is None

    def test_close_twice_does_not_raise(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.close()  # second call — must be silent

    def test_get_rows_after_close_returns_empty_list(self, lib):
        lib.load_workbook(XLSX_FILE)
        lib.close()
        assert lib.get_rows() == []

    def test_reload_after_close_works(self, lib):
        """The library must be fully reusable after a close."""
        lib.load_workbook(XLSX_FILE)
        lib.close()
        lib.load_workbook(XLSX_FILE)
        assert len(lib.get_rows()) == 4

    def test_listener_closes_workbook_automatically(self, lib):
        """The end_test listener must call close; simulate it directly."""
        lib.load_workbook(XLSX_FILE)
        lib.end_test("some test", {})
        assert lib._active_workbook is None

    def test_close_csv_stream_closes_file_handle(self, lib):
        """CSV stream resource holds an open file handle; close must release it."""
        lib.load_workbook(CSV_FILE, read_only=True)
        resource = lib._active_workbook._resource  # type: ignore[union-attr]
        lib.close()
        # After close the handle must be closed so the file is not locked
        assert resource._handle.closed

    def test_close_then_reload_then_close_again(self, lib):
        """Full lifecycle repeated twice must work without errors."""
        for _ in range(2):
            lib.load_workbook(XLSX_FILE)
            assert lib._active_workbook is not None
            lib.close()
            assert lib._active_workbook is None
