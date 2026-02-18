from rfexcel.model.raw_data.i_raw_row_data import IRawRowData


class CsvRawRowData(IRawRowData):
    def __init__(self, data: list[str]):
        self._data = data