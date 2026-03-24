from typing import Any, Callable, TypeAlias

import pytest
from openpyxl import Workbook

from rfexcel.model.raw_data.csv_raw_row_data import CsvRawRowData
from rfexcel.model.raw_data.i_raw_row_data import IRawRowData
from rfexcel.model.raw_data.null_raw_row_data import NullRawRowData
from rfexcel.model.raw_data.xls_raw_row_data import XlsRawRowData
from rfexcel.model.raw_data.xlsx_raw_row_data import XlsxRawRowData

RawFactory: TypeAlias = Callable[[list[Any]], IRawRowData]

# ---------------------------------------------------------------------------
# Factories — normalise a conceptual value list to each format's storage type
# ---------------------------------------------------------------------------

def _make_csv(values: list[Any]) -> IRawRowData:
    """CSV reader always yields strings; None entries become empty strings."""
    return CsvRawRowData([str(v) if v is not None else "" for v in values])


def _make_xls(values: list[Any]) -> IRawRowData:
    return XlsRawRowData(values)


def _make_xlsx_value_only(values: list[Any]) -> IRawRowData:
    return XlsxRawRowData(tuple(values), value_only=True)


def _make_xlsx_cell_mode(values: list[Any]) -> IRawRowData:
    wb = Workbook()
    ws = wb.active
    for col, value in enumerate(values, start=1):
        ws.cell(row=1, column=col, value=value)
    row_data = XlsxRawRowData(tuple(ws[1]), value_only=False)
    wb.close()
    return row_data


_FACTORIES: list[RawFactory] = [
    _make_csv,
    _make_xls,
    _make_xlsx_value_only,
    _make_xlsx_cell_mode,
]
_IDS = ["csv", "xls", "xlsx_value_only", "xlsx_cell_mode"]


# ---------------------------------------------------------------------------
# other edge cases
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("factory", _FACTORIES, ids=_IDS)
def test_col_zero_returns_empty_string(factory: RawFactory) -> None:
    """Column index 0 must never wrap to the last element via negative indexing."""
    row = factory(["first", "second"])
    assert row.get_dict_row_data({"invalid_col": 0, "valid_col": 2}) == {
        "invalid_col": "",
        "valid_col": "second",
    }

@pytest.mark.parametrize("factory", _FACTORIES, ids=_IDS)
def test_get_header_map_skips_none_or_empty_column(factory: RawFactory) -> None:
    """A blank/None cell in a header row must not produce a phantom key."""
    row = factory(["Name", None, "Age"])
    assert row.get_header_map() == {"Name": 1, "Age": 3}


@pytest.mark.parametrize("factory", _FACTORIES, ids=_IDS)
def test_get_header_map_skips_whitespace_only_column(factory: RawFactory) -> None:
    """A whitespace-only cell must be excluded from the header map."""
    row = factory(["Name", "   ", "Age"])
    assert row.get_header_map() == {"Name": 1, "Age": 3}

def test_csv_header_keys_are_stripped() -> None:
    row = CsvRawRowData(["  Product ID  ", " Description", "Location "])
    assert row.get_header_map() == {
        "Product ID": 1,
        "Description": 2,
        "Location": 3,
    }

def test_null_get_list_row_data_warns_about_row_data(monkeypatch: pytest.MonkeyPatch) -> None:
    messages: list[str] = []
    monkeypatch.setattr("rfexcel.model.raw_data.null_raw_row_data.logger.warn", messages.append)

    result = NullRawRowData().get_list_row_data()

    assert result == []
    assert messages == ["No row data values were returned"]