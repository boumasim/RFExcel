from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class XlsRawRowData(IRawRowData):
    def __init__(self, data: list[str | int | float | bool]):
        self._data = data