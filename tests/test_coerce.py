import polars as pl

from xl2y import coerce


def _col(vals):
    return pl.DataFrame({"x": pl.Series("x", vals, dtype=pl.Utf8)})


def test_numeric_forms():
    df = _col(["$1,234.50", "(500)", "15%", "A$ 2,000", "-7"])
    out = coerce.coerce_columns(df, dayfirst=True)[0]
    assert out["x"].to_list() == [1234.5, -500.0, 0.15, 2000.0, -7.0]


def test_whole_numbers_become_int():
    out, _ = coerce.coerce_columns(_col(["1,000", "2", "3"]), dayfirst=True)
    assert out["x"].dtype == pl.Int64


def test_na_tokens_nulled_even_in_text_columns():
    out, _ = coerce.coerce_columns(_col(["hello", "N/A", "--"]), dayfirst=True)
    assert out["x"].to_list() == ["hello", None, None]


def test_below_threshold_stays_text():
    vals = ["1", "2", "x", "y", "z"]  # 40% numeric < 80%
    out, _ = coerce.coerce_columns(_col(vals), dayfirst=True)
    assert out["x"].dtype == pl.Utf8


def test_dayfirst_dates():
    out, _ = coerce.coerce_columns(
        _col(["03/04/2025", "04/04/2025"]), dayfirst=True
    )
    assert out["x"].dtype == pl.Date
    assert out["x"][0].month == 4


def test_monthfirst_dates():
    out, _ = coerce.coerce_columns(
        _col(["03/04/2025", "04/04/2025"]), dayfirst=False
    )
    assert out["x"].dtype == pl.Date
    assert out["x"][0].month == 3


def test_bare_numbers_are_not_dates():
    out, _ = coerce.coerce_columns(
        _col(["2021", "2022", "2023x"]), dayfirst=True
    )
    assert out["x"].dtype == pl.Utf8


def test_events_reported():
    _, events = coerce.coerce_columns(
        _col(["1", "2", "3", "4", "oops"]), dayfirst=True
    )
    assert events[0]["event"] == "coerced_numeric"
    assert events[0]["failed"] == 1


def test_nested_parens_not_fabricated():
    # "((500))" is malformed; must NOT become a number (matches prototype).
    out, _ = coerce.coerce_columns(
        _col(["((500))", "(500))", "x", "y", "z"]), dayfirst=True
    )
    assert out["x"].dtype == pl.Utf8  # too few parse -> stays text


def test_single_paren_still_negative():
    out, _ = coerce.coerce_columns(_col(["(500)", "(1,000)", "5"]), dayfirst=True)
    assert out["x"].to_list() == [-500, -1000, 5]


def test_iso_slash_and_month_word_dates():
    out, _ = coerce.coerce_columns(
        _col(["2025/03/04", "2025/03/05", "2025/03/06"]), dayfirst=True
    )
    assert out["x"].dtype == pl.Date
    out2, _ = coerce.coerce_columns(
        _col(["3-Apr-2025", "4-Apr-2025", "5-Apr-2025"]), dayfirst=True
    )
    assert out2["x"].dtype == pl.Date
    assert out2["x"][0].month == 4
