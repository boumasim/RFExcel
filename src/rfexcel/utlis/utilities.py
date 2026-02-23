from robot.utils import DotDict  # type: ignore

from rfexcel.utlis.types import DictRowData


def search_in_row(source_row: DictRowData, search_criteria: DictRowData, partial_match: bool) -> bool:
    """Returns True if ALL rules in search_criteria match source_row (AND logic).

    Each key-value pair in search_criteria is one rule. A rule matches when:
    - partial_match=False: the value in source_row equals the criteria value exactly.
    - partial_match=True:  the criteria value is a substring of the row value.

    A key in search_criteria that does not exist in source_row causes an
    immediate False return — the criterion cannot be satisfied.
    Returns True only when every rule in search_criteria produces a match.
    An empty search_criteria always returns True.
    """
    for key, criteria_value in search_criteria.items():  # type: ignore[misc]
        key_str: str = str(key)  # type: ignore[arg-type]
        criteria_str: str = str(criteria_value)  # type: ignore[arg-type]
        if key_str not in source_row:
            return False
        row_value: str = str(source_row[key_str])  # type: ignore[arg-type]
        if partial_match:
            if criteria_str not in row_value:
                return False
        else:
            if criteria_str != row_value:
                return False
    return True


def convert_string_to_dict_row_data(data: str | dict[str, str], delimiter: str = ';') -> DictRowData:
    """Converts a string like ``animal=cat;person=Ted`` into a DictRowData.

    Each segment separated by ``delimiter`` must contain ``=``. Everything
    before the first ``=`` is the key; everything after is the value. This
    means values that themselves contain ``=`` (e.g. URLs) are handled
    correctly. Whitespace around keys and values is stripped. Segments
    without ``=`` are silently ignored.
    """
    if(isinstance(data, dict)):
        return DotDict(data)
    result: DictRowData = DotDict()
    for segment in data.split(delimiter):
        segment = segment.strip()
        if '=' not in segment:
            continue
        key, _, value = segment.partition('=')
        result[key.strip()] = value.strip()
    return result
