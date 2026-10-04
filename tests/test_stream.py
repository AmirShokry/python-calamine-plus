from io import BytesIO
from pathlib import Path

import pytest
from python_calamine_plus import (
    CalamineWorkbook,
    StreamingNotSupported,
    StreamInvalidated,
    WorkbookClosed,
    WorksheetNotFound,
)

PATH = Path(__file__).parent / "data"


def _normalize(rows):
    """Strip trailing empty cells and trailing empty rows so layouts can be compared."""
    out = []
    for row in rows:
        row = list(row)
        while row and row[-1] == "":
            row.pop()
        out.append(row)
    while out and not out[-1]:
        out.pop()
    return out


@pytest.mark.parametrize("filename", ["base.xlsx", "any_sheets.xlsx", "issue139.xlsx"])
def test_stream_matches_full_load(filename):
    reader = CalamineWorkbook.from_path(PATH / filename)
    for name in reader.sheet_names:
        try:
            expected = reader.get_sheet_by_name(name).to_python(skip_empty_area=False)
        except Exception:
            continue  # e.g. chart sheets, not readable as a range either
        streamed = list(reader.stream_sheet_by_name(name))
        # Same number of rows (no extra rows for formatted-but-empty cells)...
        assert len(streamed) == len(expected), name
        # ...and same values (widths differ: streamed rows always start at column A).
        assert _normalize(streamed) == _normalize(expected), name


def test_stream_by_index():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    by_index = list(reader.stream_sheet_by_index(0))
    by_name = list(reader.stream_sheet_by_name(reader.sheet_names[0]))
    assert by_index == by_name
    assert reader.stream_sheet_by_index(0).name == reader.sheet_names[0]


def test_stream_is_lazy_iterator():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    stream = reader.stream_sheet_by_name(reader.sheet_names[0])
    assert iter(stream) is stream
    first = next(stream)
    assert isinstance(first, list)


def test_stream_from_filelike():
    data = (PATH / "base.xlsx").read_bytes()
    expected = list(
        CalamineWorkbook.from_path(PATH / "base.xlsx").stream_sheet_by_index(0)
    )
    assert (
        list(CalamineWorkbook.from_filelike(BytesIO(data)).stream_sheet_by_index(0))
        == expected
    )


def test_stream_closed_workbook():
    with CalamineWorkbook.from_path(PATH / "base.xlsx") as reader:
        stream = reader.stream_sheet_by_index(0)
        next(stream)
    with pytest.raises(WorkbookClosed):
        next(stream)


def test_stream_invalidated_by_other_access():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    expected = list(reader.stream_sheet_by_index(0))

    first = reader.stream_sheet_by_index(0)
    next(first)
    second = reader.stream_sheet_by_index(0)
    with pytest.raises(StreamInvalidated):
        next(first)
    assert list(second) == expected

    third = reader.stream_sheet_by_index(0)
    next(third)
    reader.get_sheet_by_index(0)
    with pytest.raises(StreamInvalidated):
        next(third)


def test_stream_exhausted_and_restarted():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    stream = reader.stream_sheet_by_index(0)
    rows = list(stream)
    assert list(stream) == []  # exhausted streams stay exhausted, no error
    for _ in range(3):  # abandoned (dropped) streams don't break new ones
        next(reader.stream_sheet_by_index(0))
    assert list(reader.stream_sheet_by_index(0)) == rows


def test_stream_errors():
    reader = CalamineWorkbook.from_path(PATH / "base.xlsx")
    with pytest.raises(WorksheetNotFound):
        reader.stream_sheet_by_name("nope")
    with pytest.raises(WorksheetNotFound):
        reader.stream_sheet_by_index(100)

    with pytest.raises(StreamingNotSupported):
        CalamineWorkbook.from_path(PATH / "base.xls").stream_sheet_by_index(0)

    reader.close()
    with pytest.raises(WorkbookClosed):
        reader.stream_sheet_by_index(0)
