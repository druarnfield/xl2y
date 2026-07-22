"""Vectorized coercion of text columns.

Mirrors ``reference/excel_loader.py``'s ``_coerce_columns`` /
``_parse_numberish`` but as polars expressions rather than per-value Python
callbacks. Per string column, in order:

1. Common NA tokens ("N/A", "-", "#REF!", ...) become null (applied even if
   the column stays text).
2. If at least ``COERCE_MIN_FRACTION`` of the non-null values parse as
   tolerant numbers ("1,234", "$1,000.50", "(500)", "15%"), the column
   becomes Float64 (or Int64 when every value is whole).
3. Otherwise, if that fraction parse as dates *and* at least half look
   date-like, it becomes Date.
4. Otherwise it stays text (with NA tokens still nulled).
"""

from __future__ import annotations

import logging

import polars as pl

logger = logging.getLogger(__name__)

COERCE_MIN_FRACTION = 0.8

NA_TOKENS = {
    "n/a",
    "na",
    "#n/a",
    "#value!",
    "#ref!",
    "#div/0!",
    "-",
    "--",
    "–",
    "—",
    "none",
    "null",
    "nil",
}

_CURRENCY = r"(?i)^(?:a\$|au\$|aud|nz\$|us\$|usd|[$£€¥])\s*"

_DATE_FORMATS_DAYFIRST = [
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d.%m.%Y",
    "%d/%m/%y",
    "%Y-%m-%d",
    "%d %b %Y",
    "%d %B %Y",
]
_DATE_FORMATS_MONTHFIRST = [
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%m/%d/%y",
    "%Y-%m-%d",
    "%b %d %Y",
    "%B %d %Y",
]
_DATEY = (
    r"(?i)[-/:.]|\d\s+\w|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
)


def numberish(col: str) -> pl.Expr:
    """Expression parsing '1,234', '$1,000.50', '(500)', '15%' -> Float64
    (null when it does not look like a number). Accounting parentheses and a
    leading minus both mean negative; a trailing percent divides by 100."""
    s = pl.col(col).str.strip_chars()
    paren = s.str.contains(r"^\(.*\)$")
    core = (
        pl.when(paren)
        .then(s.str.strip_chars("()"))
        .otherwise(s)
        .str.strip_chars()
    )
    minus = core.str.starts_with("-")
    core = core.str.replace(r"^-", "").str.strip_chars()
    core = core.str.replace(_CURRENCY, "")
    pct = core.str.ends_with("%")
    core = core.str.replace(r"%$", "")
    core = core.str.replace_all(",", "").str.replace_all(r"\s", "")
    n = core.cast(pl.Float64, strict=False)
    n = pl.when(pct).then(n / 100).otherwise(n)
    return pl.when(paren | minus).then(-n).otherwise(n)


def datish(col: str, dayfirst: bool) -> pl.Expr:
    """Expression parsing common date strings -> Date (null on no match)."""
    fmts = _DATE_FORMATS_DAYFIRST if dayfirst else _DATE_FORMATS_MONTHFIRST
    tries = [
        pl.col(col).str.strip_chars().str.to_date(f, strict=False)
        for f in fmts
    ]
    return pl.coalesce(tries)


def coerce_columns(
    df: pl.DataFrame, dayfirst: bool, sheet: str = "?"
) -> tuple[pl.DataFrame, list[dict]]:
    """Return ``(coerced_df, events)``. Only Utf8 columns are touched."""
    events: list[dict] = []
    for col in list(df.columns):
        if df[col].dtype != pl.Utf8:
            continue

        # NA tokens -> null (kept even if the column stays text).
        base = df.with_columns(
            pl.when(
                pl.col(col)
                .str.strip_chars()
                .str.to_lowercase()
                .is_in(list(NA_TOKENS))
            )
            .then(None)
            .otherwise(pl.col(col))
            .alias(col)
        )
        nn = base[col].drop_nulls()
        if nn.is_empty():
            df = base
            continue

        # --- numeric ---------------------------------------------------- #
        num = base.select(numberish(col).alias("v"))["v"]
        parsed = num.drop_nulls().len()
        if parsed / nn.len() >= COERCE_MIN_FRACTION:
            failed = nn.len() - parsed
            out = base.with_columns(numberish(col).alias(col))
            v = out[col].drop_nulls()
            if v.len() and bool(((v % 1) == 0).all()):
                out = out.with_columns(pl.col(col).cast(pl.Int64))
            events.append(
                {"event": "coerced_numeric", "column": col, "failed": failed}
            )
            if failed:
                logger.warning(
                    "Sheet %r: column %r coerced to numeric; %d value(s) "
                    "became null.",
                    sheet,
                    col,
                    failed,
                )
            df = out
            continue

        # --- datetime --------------------------------------------------- #
        dt = base.select(datish(col, dayfirst).alias("v"))["v"]
        datey = base.select(pl.col(col).str.contains(_DATEY).alias("d"))["d"]
        datey_nn = datey.drop_nulls()
        datey_ratio = datey_nn.sum() / datey_nn.len() if datey_nn.len() else 0
        if (
            dt.drop_nulls().len() / nn.len() >= COERCE_MIN_FRACTION
            and datey_ratio >= 0.5
        ):
            failed = nn.len() - dt.drop_nulls().len()
            events.append(
                {
                    "event": "coerced_date",
                    "column": col,
                    "failed": failed,
                    "dayfirst": dayfirst,
                }
            )
            logger.info(
                "Sheet %r: column %r parsed as dates (dayfirst=%s).",
                sheet,
                col,
                dayfirst,
            )
            df = base.with_columns(datish(col, dayfirst).alias(col))
            continue

        df = base  # text, NA tokens applied
    return df, events
