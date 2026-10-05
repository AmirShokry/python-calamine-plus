"""Fixes to the native API (also visible through compatibility="openpyxl")."""

import pytest
from python_calamine_plus import CalamineWorkbook, load_workbook

openpyxl = pytest.importorskip("openpyxl")

URL = "https://example.com/a?x=1&y=2"
LITERAL = "https://example.com/?q=&amp;"  # the text "&amp;" itself, not an entity


@pytest.fixture(scope="module")
def links_xlsx(tmp_path_factory):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Links"
    ws["A1"] = "link"
    ws["A1"].hyperlink = URL  # stored as "...&amp;y=2" in the .rels part
    ws["A2"] = "literal"
    ws["A2"].hyperlink = LITERAL  # stored as "...&amp;amp;"
    wb.create_sheet("Empty")
    path = tmp_path_factory.mktemp("native") / "links.xlsx"
    wb.save(path)
    return path


def test_hyperlink_targets_are_entity_decoded(links_xlsx):
    stream = CalamineWorkbook.from_path(links_xlsx).stream_sheet_by_name(
        "Links", rich=True
    )
    rows = list(stream)
    assert rows[0][0].hyperlink["target"] == URL
    assert rows[1][0].hyperlink["target"] == LITERAL
    assert [h["target"] for h in stream.hyperlinks] == [URL, LITERAL]
    expected = openpyxl.load_workbook(links_xlsx)["Links"]
    assert [expected["A1"].hyperlink.target, expected["A2"].hyperlink.target] == [
        URL,
        LITERAL,
    ]


def test_hyperlink_targets_in_compatibility_mode(links_xlsx):
    ws = load_workbook(links_xlsx, compatibility="openpyxl")["Links"]
    assert ws["A1"].hyperlink.target == URL
    assert ws["A2"].hyperlink.target == LITERAL  # decoded once, not twice


def test_iter_rows_of_empty_sheet(links_xlsx):
    sheet = CalamineWorkbook.from_path(links_xlsx).get_sheet_by_name("Empty")
    assert list(sheet.iter_rows()) == []
    assert sheet.to_python() == []
    ws = load_workbook(links_xlsx, compatibility="openpyxl")["Empty"]
    assert list(ws.iter_rows()) == [] and list(ws.values) == []
