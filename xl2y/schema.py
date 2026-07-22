"""A deliberately small schema: column names, types, nullability, bounds,
patterns, allowed values, and a per-column escape-hatch ``check=``.

``Schema.conform`` runs two phases, *collecting* every failure rather than
stopping at the first:

1. **cast** each declared column to its type with the same tolerant parsers
   as :mod:`xl2y.coerce`;
2. **validate** nullability, bounds, patterns, allowed values, and checks.

It returns ``(cast_df, problems)``. :meth:`xl2y.table.Table.conform` decides
what to do with the problems (raise / quarantine / report).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

import polars as pl

from xl2y.coerce import datish, numberish

_TRUE = {"true", "yes", "y", "1", "t"}
_FALSE = {"false", "no", "n", "0", "f"}


# --------------------------------------------------------------------------- #
# Column types
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class ColumnType:
    kind: str  # str | int | float | bool | date | datetime | cat
    nullable: bool = True
    min: Any = None
    max: Any = None
    pattern: str | None = None
    values: tuple = ()
    check: Any = None  # pl.Expr or Callable[[pl.DataFrame], boolean Series]


def str_(
    nullable: bool = True, pattern: str | None = None, check: Any = None
) -> ColumnType:
    return ColumnType("str", nullable=nullable, pattern=pattern, check=check)


def int_(
    nullable: bool = True,
    min: Any = None,
    max: Any = None,
    check: Any = None,
) -> ColumnType:
    return ColumnType("int", nullable=nullable, min=min, max=max, check=check)


def float_(
    nullable: bool = True,
    min: Any = None,
    max: Any = None,
    check: Any = None,
) -> ColumnType:
    return ColumnType("float", nullable=nullable, min=min, max=max, check=check)


def bool_(nullable: bool = True, check: Any = None) -> ColumnType:
    return ColumnType("bool", nullable=nullable, check=check)


def date_(
    nullable: bool = True,
    min: Any = None,
    max: Any = None,
    check: Any = None,
) -> ColumnType:
    return ColumnType("date", nullable=nullable, min=min, max=max, check=check)


def datetime_(
    nullable: bool = True,
    min: Any = None,
    max: Any = None,
    check: Any = None,
) -> ColumnType:
    return ColumnType(
        "datetime", nullable=nullable, min=min, max=max, check=check
    )


def cat_(*values: Any, nullable: bool = True, check: Any = None) -> ColumnType:
    return ColumnType("cat", nullable=nullable, values=tuple(values), check=check)


# --------------------------------------------------------------------------- #
# Problems
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class Problem:
    column: str
    rule: str
    count: int
    rows: list  # sample of offending Excel rows (or df positions), capped
    message: str


def _fmt_rows(rows: list) -> str:
    return ", ".join(str(r) for r in rows)


def _record(
    problems: list[Problem],
    bad: set[int],
    column: str,
    rule: str,
    mask: pl.Series,
    excel_rows: list[int] | None,
    make_msg: Callable[[int, list], str],
) -> None:
    """If ``mask`` has any True rows, append a Problem and add those row
    positions to the running ``bad`` set (used for quarantine)."""
    idx = [i for i, v in enumerate(mask.to_list()) if v]
    if not idx:
        return
    count = len(idx)
    if excel_rows is not None:
        rows = [excel_rows[i] for i in idx[:10] if i < len(excel_rows)]
    else:
        rows = idx[:10]
    bad.update(idx)
    problems.append(Problem(column, rule, count, rows, make_msg(count, rows)))


# --------------------------------------------------------------------------- #
# Casting
# --------------------------------------------------------------------------- #


def _cast_expr(col: str, kind: str, dtype: pl.DataType, dayfirst: bool) -> pl.Expr:
    c = pl.col(col)
    if kind in ("str", "cat"):
        return c.cast(pl.Utf8, strict=False)
    if kind in ("int", "float"):
        n = numberish(col) if dtype == pl.Utf8 else c.cast(pl.Float64, strict=False)
        if kind == "int":
            return n.round(0).cast(pl.Int64, strict=False)
        return n
    if kind == "bool":
        if dtype == pl.Boolean:
            return c
        lc = (
            c.cast(pl.Utf8, strict=False)
            .str.strip_chars()
            .str.to_lowercase()
        )
        return (
            pl.when(lc.is_in(list(_TRUE)))
            .then(True)
            .when(lc.is_in(list(_FALSE)))
            .then(False)
            .otherwise(None)
        )
    if kind == "date":
        return datish(col, dayfirst) if dtype == pl.Utf8 else c.cast(
            pl.Date, strict=False
        )
    if kind == "datetime":
        if dtype == pl.Utf8:
            return pl.coalesce(
                [
                    c.str.strip_chars().str.to_datetime(strict=False),
                    datish(col, dayfirst).cast(pl.Datetime("us")),
                ]
            )
        return c.cast(pl.Datetime("us"), strict=False)
    return c


# --------------------------------------------------------------------------- #
# Schema
# --------------------------------------------------------------------------- #


class Schema:
    def __init__(self, *, extra_columns: str = "keep", **columns: ColumnType):
        if extra_columns not in ("keep", "drop", "error"):
            raise ValueError(
                "extra_columns must be 'keep', 'drop', or 'error'"
            )
        self.columns: dict[str, ColumnType] = columns
        self.extra_columns = extra_columns

    def conform(
        self,
        df: pl.DataFrame,
        excel_rows: list[int] | None,
        dayfirst: bool = True,
    ) -> tuple[pl.DataFrame, list[Problem], list[int]]:
        """Return ``(cast_df, problems, bad_row_positions)``.

        ``bad_row_positions`` are df-row indices with at least one row-level
        failure (used for quarantine); column-level problems (missing/extra)
        do not contribute rows.
        """
        problems: list[Problem] = []
        bad: set[int] = set()
        present = set(df.columns)

        # -- presence --------------------------------------------------- #
        for name in self.columns:
            if name not in present:
                problems.append(
                    Problem(name, "missing", 0, [], "declared column is missing")
                )
        extras = [c for c in df.columns if c not in self.columns]
        if extras and self.extra_columns == "error":
            for c in extras:
                problems.append(
                    Problem(c, "extra", 0, [], "unexpected column")
                )
        elif extras and self.extra_columns == "drop":
            df = df.drop(extras)

        # -- cast ------------------------------------------------------- #
        out = df
        for name, ct in self.columns.items():
            if name not in out.columns:
                continue
            before_null = df[name].is_null()
            expr = _cast_expr(name, ct.kind, df[name].dtype, dayfirst)
            out = out.with_columns(expr.alias(name))
            fail = (~before_null) & out[name].is_null()
            _record(
                problems,
                bad,
                name,
                "cast",
                fail,
                excel_rows,
                lambda c, r, k=ct.kind: f"{c} value(s) could not be cast to "
                f"{k}_ (rows {_fmt_rows(r)})",
            )

        # -- validate --------------------------------------------------- #
        for name, ct in self.columns.items():
            if name not in out.columns:
                continue
            s = out[name]

            if not ct.nullable:
                _record(
                    problems,
                    bad,
                    name,
                    "not_null",
                    s.is_null(),
                    excel_rows,
                    lambda c, r: f"{c} null value(s) (rows {_fmt_rows(r)})",
                )

            if ct.min is not None:
                _record(
                    problems,
                    bad,
                    name,
                    "min",
                    s.is_not_null() & (s < ct.min),
                    excel_rows,
                    lambda c, r, m=ct.min: f"{c} value(s) < {m} "
                    f"(rows {_fmt_rows(r)})",
                )

            if ct.max is not None:
                _record(
                    problems,
                    bad,
                    name,
                    "max",
                    s.is_not_null() & (s > ct.max),
                    excel_rows,
                    lambda c, r, m=ct.max: f"{c} value(s) > {m} "
                    f"(rows {_fmt_rows(r)})",
                )

            if ct.pattern is not None:
                matches = s.str.contains(f"^(?:{ct.pattern})$")
                _record(
                    problems,
                    bad,
                    name,
                    "pattern",
                    s.is_not_null() & (~matches.fill_null(False)),
                    excel_rows,
                    lambda c, r: f"{c} value(s) fail pattern "
                    f"(rows {_fmt_rows(r)})",
                )

            if ct.kind == "cat" and ct.values:
                in_set = s.is_in(list(ct.values))
                _record(
                    problems,
                    bad,
                    name,
                    "allowed_values",
                    s.is_not_null() & (~in_set.fill_null(False)),
                    excel_rows,
                    lambda c, r, v=list(ct.values): f"{c} value(s) not in {v} "
                    f"(rows {_fmt_rows(r)})",
                )

            if ct.check is not None:
                res = _eval_check(ct.check, out)
                _record(
                    problems,
                    bad,
                    name,
                    "check",
                    (~res).fill_null(False),
                    excel_rows,
                    lambda c, r: f"{c} value(s) fail check "
                    f"(rows {_fmt_rows(r)})",
                )

        return out, problems, sorted(bad)


def _eval_check(check: Any, df: pl.DataFrame) -> pl.Series:
    if callable(check) and not isinstance(check, pl.Expr):
        check = check(df)
    if isinstance(check, pl.Expr):
        return df.select(check.alias("_c"))["_c"]
    return pl.Series(check)
