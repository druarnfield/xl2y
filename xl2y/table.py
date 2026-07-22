"""The immutable pipeline object.

Every verb returns a *new* :class:`Table` carrying one more lineage entry, so
a pipeline is just a reusable function ``lambda t: t.clean().cast(...)`` and
``dry_run()`` / parquet metadata can always explain what happened.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any, Callable

import polars as pl

# String aliases accepted by cast() and mapped to polars dtypes. The schema
# module's ColumnType objects reuse these same names via their .kind.
_CAST_TARGETS: dict[str, pl.DataType] = {
    "str_": pl.Utf8,
    "int_": pl.Int64,
    "float_": pl.Float64,
    "bool_": pl.Boolean,
    "date_": pl.Date,
    "datetime_": pl.Datetime("us"),
}


@dataclass(frozen=True)
class Table:
    """Immutable pipeline state."""

    df: pl.DataFrame
    source: dict
    excel_rows: list[int] | None = None  # None once row identity is lost
    lineage: list[dict] = field(default_factory=list)
    comments: list[dict] = field(default_factory=list)
    rejects: pl.DataFrame | None = None
    errors: list = field(default_factory=list)
    _raw: Any = None  # RawSheet, kept for kitchen_sink re-extraction

    # -- internal --------------------------------------------------------- #

    def _step(
        self,
        verb: str,
        df: pl.DataFrame,
        excel_rows: list[int] | None,
        **detail: Any,
    ) -> "Table":
        entry = {"verb": verb, **detail}
        return replace(
            self, df=df, excel_rows=excel_rows, lineage=[*self.lineage, entry]
        )

    # -- terminal / escape hatches --------------------------------------- #

    def collect(self) -> pl.DataFrame:
        return self.df

    def apply(self, fn: Callable[[pl.DataFrame], pl.DataFrame]) -> "Table":
        """Run an arbitrary ``df -> df`` callable as a pipeline step.

        Excel-row identity is preserved only if the row count is unchanged;
        otherwise it is dropped to None rather than guessed.
        """
        out = fn(self.df)
        rows = (
            self.excel_rows
            if (self.excel_rows is not None and out.height == self.df.height)
            else None
        )
        name = getattr(fn, "__name__", "<callable>")
        return self._step("apply", out, rows, fn=name)

    def cast(self, **types: Any) -> "Table":
        """Cast named columns to explicit types.

        Values may be the string alias ("int_") or a schema ColumnType.
        """
        unknown = set(types) - set(self.df.columns)
        if unknown:
            raise ValueError(f"cast: unknown column(s) {sorted(unknown)}")
        exprs = []
        recorded: dict[str, str] = {}
        for col, t in types.items():
            kind = getattr(t, "kind", None)
            alias = f"{kind}_" if kind else t
            if alias not in _CAST_TARGETS:
                raise ValueError(f"cast: unknown target type {t!r} for {col!r}")
            exprs.append(pl.col(col).cast(_CAST_TARGETS[alias]))
            recorded[col] = alias
        return self._step(
            "cast", self.df.with_columns(exprs), self.excel_rows, types=recorded
        )
