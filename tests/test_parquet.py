import json

import polars as pl
import pytest

import xl2y
from tests import fixtures


def test_to_parquet_roundtrip_with_lineage(tmp_path):
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx")).clean()
    out = tmp_path / "out.parquet"
    t.to_parquet(out)
    assert pl.read_parquet(out).height == 3
    meta = pl.read_parquet_metadata(out)
    payload = json.loads(meta["xl2y"])
    assert payload["source"]["sheet"] == "Data"
    assert payload["lineage"][0]["verb"] == "load"


def test_to_parquet_refuses_overwrite(tmp_path):
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"))
    out = tmp_path / "o.parquet"
    t.to_parquet(out)
    with pytest.raises(FileExistsError):
        t.to_parquet(out)
    t.to_parquet(out, overwrite=True)  # fine


def test_tableset_fans_out(tmp_path):
    ts = xl2y.load_all(fixtures.nasty_book(tmp_path / "n.xlsx"))
    manifest = ts.clean().to_parquet(tmp_path / "out")
    assert (tmp_path / "out" / "report.parquet").exists()
    assert manifest["report"]["rows"] == 3
