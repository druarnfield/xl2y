import pytest

import xl2y
from tests import fixtures


def test_load_single_sheet(tmp_path):
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"))
    assert t.df.height == 3
    assert t.source["sheet"] == "Data"


def test_load_picks_best_sheet(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx"))
    assert t.source["sheet"] == "Report"  # Empty sheet loses


def test_load_explicit_sheet_missing(tmp_path):
    with pytest.raises(xl2y.SheetNotFoundError):
        xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"), sheet="X")


def test_load_hints_forwarded(tmp_path):
    t = xl2y.load(
        fixtures.nasty_book(tmp_path / "n.xlsx"),
        sheet="Report",
        sparse_rows="section",
    )
    assert "section" in t.df.columns


def test_load_all(tmp_path):
    ts = xl2y.load_all(fixtures.nasty_book(tmp_path / "n.xlsx"))
    assert list(ts) == ["Report"]  # empty sheet skipped, logged
    assert ts["Report"].df.height == 3
