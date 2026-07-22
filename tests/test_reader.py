import time

import pytest

from tests import fixtures
from xl2y import reader
from xl2y.errors import SheetNotFoundError, UnsupportedFormatError


def test_rejects_xls(tmp_path):
    p = tmp_path / "old.xls"
    p.touch()
    with pytest.raises(UnsupportedFormatError, match="Convert"):
        reader.check_format(p)


def test_rejects_unknown_suffix(tmp_path):
    p = tmp_path / "f.csv"
    p.touch()
    with pytest.raises(UnsupportedFormatError):
        reader.check_format(p)


def test_sheet_meta_lists_sheets_in_order(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    meta = reader.sheet_meta(p)
    assert list(meta) == ["Report", "Empty"]
    member, is_chart = meta["Report"]
    assert member.startswith("xl/") and not is_chart


def test_read_sheet_returns_rawsheet(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    raw = reader.read_sheet(p, "Report")
    assert raw.name == "Report"
    assert raw.grid[0][0] == "2025 Results — CONFIDENTIAL"
    assert (1, 1, 1, 4) in raw.merged        # A1:D1 as (min_r, min_c, max_r, max_c)
    assert raw.grid[1] == []                  # blank row preserved as placeholder


def test_hidden_rows_and_cols_detected(tmp_path):
    p = fixtures.hidden_book(tmp_path / "h.xlsx")
    raw = reader.read_sheet(p, "Hidden")
    assert 2 in raw.hidden_rows               # Excel row 3, 0-indexed
    assert (2, 2) in raw.hidden_col_intervals # column C, 0-indexed interval


def test_formula_cells_detected(tmp_path):
    p = fixtures.formula_book(tmp_path / "f.xlsx")
    raw = reader.read_sheet(p, "Calc")
    assert (2, 2) in raw.formulas             # B2, 1-indexed


def test_bloated_file_loads_fast(tmp_path):
    p = fixtures.bloated_book(tmp_path / "b.xlsx")
    t0 = time.monotonic()
    raw = reader.read_sheet(p, "Bloat")
    assert time.monotonic() - t0 < 5.0
    assert len(raw.grid) == 3                 # styled ghost cell ignored


def test_missing_sheet_raises_with_available(tmp_path):
    p = fixtures.simple_book(tmp_path / "s.xlsx")
    with pytest.raises(SheetNotFoundError, match="Data"):
        reader.read_sheet(p, "Nope")
