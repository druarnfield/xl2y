import polars as pl
import pytest

from xl2y.table import Table


def _table():
    df = pl.DataFrame({"a": ["1", "2"], "b": ["x", "y"]})
    return Table(
        df=df,
        source={"path": "t.xlsx", "sheet": "S"},
        excel_rows=[2, 3],
        lineage=[],
    )


def test_collect_returns_polars():
    assert isinstance(_table().collect(), pl.DataFrame)


def test_apply_is_immutable_and_logged():
    t = _table()
    t2 = t.apply(lambda df: df.head(1))
    assert t.df.height == 2 and t2.df.height == 1
    assert t2.lineage[-1]["verb"] == "apply"
    assert t2.excel_rows is None  # row identity lost


def test_apply_same_height_keeps_excel_rows():
    t2 = _table().apply(
        lambda df: df.with_columns(pl.col("a").alias("c"))
    )
    assert t2.excel_rows == [2, 3]


def test_cast():
    t2 = _table().cast(a="int_")
    assert t2.df["a"].dtype == pl.Int64
    assert t2.lineage[-1]["verb"] == "cast"


def test_cast_unknown_column_raises():
    with pytest.raises(ValueError, match="nope"):
        _table().cast(nope="int_")


def test_clean_snake_cases_and_coerces(tmp_path):
    import xl2y
    from tests import fixtures

    t = xl2y.load(fixtures.typed_book(tmp_path / "t.xlsx")).clean()
    assert t.df.columns == ["money", "pct", "when", "count", "note"]
    assert t.df["money"].to_list() == [1234.5, -500.0, None]
    assert t.df["pct"].to_list() == [0.15, 0.075, None]
    assert str(t.df["when"].dtype) == "Date"
    assert t.lineage[-1]["verb"] == "clean"


def test_clean_snake_case_collision_dedupes():
    df = pl.DataFrame({"Total Rev": [1], "total_rev": [2]})
    t = Table(df=df, source={}).clean()
    assert t.df.columns == ["total_rev", "total_rev_2"]


def test_clean_drops_fully_null_columns():
    df = pl.DataFrame(
        {"a": [1, 2], "empty": [None, None]},
        schema={"a": pl.Int64, "empty": pl.Utf8},
    )
    t = Table(df=df, source={}).clean()
    assert t.df.columns == ["a"]
    assert t.lineage[-1]["dropped_columns"] == ["empty"]
