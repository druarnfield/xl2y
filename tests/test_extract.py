import polars as pl
import pytest

from tests import fixtures
from xl2y import extract, reader
from xl2y.errors import EmptySheetError


def _extracted(tmp_path, builder, sheet, **opts):
    p = builder(tmp_path / "f.xlsx")
    return extract.extract_table(reader.read_sheet(p, sheet), **opts)


def test_simple_table(tmp_path):
    ex = _extracted(tmp_path, fixtures.simple_book, "Data")
    assert ex.df.columns == ["Store", "Date", "Revenue"]
    assert ex.df.height == 3
    assert ex.excel_rows == [2, 3, 4]


def test_banner_becomes_comment_not_header(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report")
    assert ex.df.columns == ["Store", "Date", "Units", "Revenue"]
    assert ex.comments[0]["kind"] == "merged"
    assert "CONFIDENTIAL" in ex.comments[0]["text"]


def test_sparse_rows_default_comment(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report")
    texts = [c["text"] for c in ex.comments if c["kind"] == "sparse"]
    assert "Northern Region" in texts
    assert ex.df.height == 3  # section rows stripped


def test_sparse_rows_section_mode(tmp_path):
    ex = _extracted(
        tmp_path, fixtures.nasty_book, "Report", sparse_rows="section"
    )
    assert ex.df.columns[0] == "section"
    assert ex.df["section"].to_list() == [
        "Northern Region",
        "Northern Region",
        "Southern Region",
    ]


def test_multi_row_header(tmp_path):
    ex = _extracted(
        tmp_path, fixtures.multi_header_book, "Wide", header_rows=2
    )
    assert ex.df.columns == ["Store", "Revenue Q1", "Revenue Q2"]


def test_hidden_skipped_when_asked(tmp_path):
    ex = _extracted(
        tmp_path, fixtures.hidden_book, "Hidden", skip_hidden=True
    )
    assert ex.df.columns == ["A", "B"]
    assert ex.df.height == 2
    assert ex.excel_rows == [2, 4]


def test_empty_sheet_raises(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    with pytest.raises(EmptySheetError):
        extract.extract_table(reader.read_sheet(p, "Empty"))


def test_mixed_type_column_becomes_text(tmp_path):
    ex = _extracted(tmp_path, fixtures.simple_book, "Data")
    assert ex.df["Revenue"].dtype == pl.Utf8


def test_units_column_is_integer(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report")
    assert ex.df["Units"].dtype == pl.Int64
    assert ex.df["Units"].to_list() == [10, 4, 7]
