import pytest

from tests import fixtures
from xl2y import reader
from xl2y.errors import UnsupportedFormatError


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
