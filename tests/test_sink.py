import polars as pl

import xl2y
from tests import fixtures


def test_kitchen_sink_on_nasty_book(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx")).kitchen_sink()
    # winner keeps the section labels and yields typed columns
    assert "section" in t.df.columns
    assert t.df.height == 3
    assert t.df["revenue"].dtype.is_numeric()


def test_kitchen_sink_on_multi_header(tmp_path):
    t = xl2y.load(
        fixtures.multi_header_book(tmp_path / "m.xlsx")
    ).kitchen_sink()
    assert t.df.columns == ["store", "revenue_q1", "revenue_q2"]


def test_kitchen_sink_records_candidates(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx")).kitchen_sink()
    entry = t.lineage[-1]
    assert entry["verb"] == "kitchen_sink"
    assert len(entry["candidates"]) >= 4
    assert entry["winner"]["score"] == max(
        c["score"] for c in entry["candidates"]
    )


def test_kitchen_sink_without_raw_falls_back_to_clean():
    from xl2y.table import Table

    t = Table(df=pl.DataFrame({"A": ["1", "2"]}), source={})
    assert t.kitchen_sink().df["a"].dtype.is_numeric()  # degraded = clean()


def test_kitchen_sink_unpivots_wide_years(tmp_path):
    t = xl2y.load(fixtures.wide_years_book(tmp_path / "w.xlsx")).kitchen_sink()
    assert t.df.columns == ["store", "period", "value"]
    assert t.df.height == 6
    assert set(t.df["period"].to_list()) == {"2021", "2022", "2023"}
    assert t.lineage[-1]["winner"]["opts"] == {"unpivot": True}
