import polars as pl
import pytest

import xl2y
from xl2y import patterns
from xl2y.errors import SchemaError
from xl2y.table import Table


def _table():
    df = pl.DataFrame(
        {
            "store": ["Syd", "Mel", None],
            "revenue": ["$1,000", "(50)", "abc"],
            "email": ["a@b.co", "bad", "c@d.co"],
        }
    )
    return Table(
        df=df,
        source={"path": "t.xlsx", "sheet": "S"},
        excel_rows=[2, 3, 4],
    )


SCHEMA = xl2y.Schema(
    store=xl2y.str_(nullable=False),
    revenue=xl2y.float_(min=0),
    email=xl2y.str_(pattern=patterns.EMAIL),
)


def test_conform_raises_with_all_problems():
    with pytest.raises(SchemaError) as ei:
        _table().conform(SCHEMA)
    msg = str(ei.value)
    assert "store" in msg and "revenue" in msg and "email" in msg
    assert len(ei.value.problems) == 4
    cast_p = [p for p in ei.value.problems if p.rule == "cast"][0]
    assert cast_p.rows == [4]


def test_conform_happy_path_casts():
    df = pl.DataFrame(
        {"store": ["Syd"], "revenue": ["$1,000"], "email": ["a@b.co"]}
    )
    t = Table(df=df, source={}).conform(SCHEMA)
    assert t.df["revenue"].to_list() == [1000.0]
    assert t.lineage[-1]["verb"] == "conform"


def test_missing_column_is_error():
    df = pl.DataFrame({"store": ["Syd"]})
    with pytest.raises(SchemaError, match="revenue"):
        Table(df=df, source={}).conform(SCHEMA)


def test_extra_columns_modes():
    schema = xl2y.Schema(a=xl2y.int_(), extra_columns="error")
    df = pl.DataFrame({"a": [1], "b": [2]})
    with pytest.raises(SchemaError, match="b"):
        Table(df=df, source={}).conform(schema)
    schema2 = xl2y.Schema(a=xl2y.int_(), extra_columns="drop")
    assert Table(df=df, source={}).conform(schema2).df.columns == ["a"]


def test_cat_and_bounds_and_check():
    schema = xl2y.Schema(
        state=xl2y.cat_("NSW", "VIC"),
        n=xl2y.int_(min=0, max=10, check=pl.col("n") != 7),
    )
    df = pl.DataFrame({"state": ["NSW", "QLD"], "n": [7, 11]})
    with pytest.raises(SchemaError) as ei:
        Table(df=df, source={}).conform(schema)
    rules = {p.rule for p in ei.value.problems}
    assert {"allowed_values", "max", "check"} <= rules


def test_abn_check_via_schema():
    schema = xl2y.Schema(abn=xl2y.str_(check=lambda df: df["abn"].map_elements(
        patterns.abn_valid, return_dtype=pl.Boolean
    )))
    df = pl.DataFrame({"abn": ["51 824 753 556", "00 000 000 000"]})
    with pytest.raises(SchemaError) as ei:
        Table(df=df, source={}).conform(schema)
    assert ei.value.problems[0].rule == "check"
    assert ei.value.problems[0].count == 1


def test_quarantine_splits_rows():
    t = _table().conform(SCHEMA, on_error="quarantine")
    assert t.df.height == 1  # only the fully-clean row survives
    assert t.rejects.height == 2
    assert t.errors  # problems still reported


def test_report_keeps_everything():
    t = _table().conform(SCHEMA, on_error="report")
    assert t.df.height == 3
    assert len(t.errors) == 4


def test_bad_on_error_value():
    with pytest.raises(ValueError, match="on_error"):
        _table().conform(SCHEMA, on_error="explode")
