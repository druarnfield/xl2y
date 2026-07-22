"""The immutable pipeline object.

Every verb returns a *new* :class:`Table` carrying one more lineage entry, so
a pipeline is just a reusable function ``lambda t: t.clean().cast(...)`` and
``dry_run()`` / parquet metadata can always explain what happened.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any, Callable

import polars as pl

from xl2y.coerce import coerce_columns
from xl2y.util import snake_case, unique_name

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


def _describe_event(ev: dict) -> str:
    kind = ev.get("event")
    if kind == "sparse_row_stripped":
        return f"stripped sparse row {ev['excel_row']}: {ev['text']!r}"
    if kind == "section_label":
        return f"section label row {ev['excel_row']}: {ev['text']!r}"
    if kind == "merged_comment":
        return f"merged title row {ev['excel_row']}: {ev['text']!r}"
    if kind == "hidden_skipped":
        return f"skipped {ev['rows']} hidden row(s), {ev['cols']} hidden col(s)"
    if kind == "columns_ignored":
        return f"used columns {ev['used']}, ignored other block(s)"
    if kind == "uncached_formulas":
        return f"{ev['count']} uncached formula cell(s) read as null"
    return str(ev)


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

    def dry_run(self, n: int = 8) -> "Table":
        """Print a summary of the current table and everything the pipeline
        has done, then return ``self`` UNCHANGED. The only verb that prints;
        drop it from the chain when you are done developing."""
        src = self.source
        print(
            f"{src.get('path', '?')} [{src.get('sheet', '?')}] — "
            f"{self.df.height} rows × {self.df.width} cols"
        )
        if self.df.width:
            print(
                "  dtypes: "
                + ", ".join(
                    f"{c}: {self.df[c].dtype}" for c in self.df.columns
                )
            )
        print(self.df.head(n))
        if self.lineage:
            print("Pipeline:")
            for entry in self.lineage:
                print(f"  • {entry.get('verb', '?')}")
                for ev in entry.get("events", []):
                    print(f"      - {_describe_event(ev)}")
                if entry.get("renames"):
                    print(f"      - renamed {len(entry['renames'])} column(s)")
                if entry.get("dropped_columns"):
                    print(
                        "      - dropped empty column(s): "
                        f"{entry['dropped_columns']}"
                    )
                for co in entry.get("coercions", []):
                    kind = co["event"].replace("coerced_", "")
                    extra = f" ({co['failed']} failed)" if co.get("failed") else ""
                    print(f"      - {co['column']} → {kind}{extra}")
        if self.comments:
            print("Comments/notes stripped:")
            for c in self.comments:
                print(
                    f"  [{c['kind']} @ Excel row {c['excel_row']}] {c['text']}"
                )
        return self

    def clean(self) -> "Table":
        """The safe, deterministic tidy-up: snake_case headers, strip
        whitespace, drop fully-empty rows/columns, then coerce text columns
        (NA tokens, tolerant numerics, dates)."""
        df = self.df
        detail: dict[str, Any] = {}

        # 1. snake_case + dedupe column names.
        taken: set[str] = set()
        mapping: dict[str, str] = {}
        renames: dict[str, str] = {}
        for c in df.columns:
            new = unique_name(snake_case(str(c)), taken)
            taken.add(new)
            mapping[c] = new
            if new != c:
                renames[c] = new
        df = df.rename(mapping)
        if renames:
            detail["renames"] = renames

        # 2. Strip whitespace on text columns; "" -> null.
        str_cols = [c for c in df.columns if df[c].dtype == pl.Utf8]
        if str_cols:
            df = df.with_columns(
                [
                    pl.when(pl.col(c).str.strip_chars().str.len_chars() == 0)
                    .then(None)
                    .otherwise(pl.col(c).str.strip_chars())
                    .alias(c)
                    for c in str_cols
                ]
            )

        # 3a. Drop fully-null rows, keeping excel_rows in sync.
        excel_rows = self.excel_rows
        if df.height and df.width:
            keep = df.select(
                pl.any_horizontal(pl.all().is_not_null()).alias("_k")
            )["_k"]
            if not bool(keep.all()):
                if excel_rows is not None:
                    excel_rows = [
                        er for er, k in zip(excel_rows, keep.to_list()) if k
                    ]
                df = df.filter(keep)

        # 3b. Drop fully-null columns.
        if df.height:
            null_cols = [
                c for c in df.columns if df[c].null_count() == df.height
            ]
            if null_cols:
                df = df.drop(null_cols)
                detail["dropped_columns"] = null_cols

        # 4. Coerce text columns.
        df, events = coerce_columns(
            df,
            dayfirst=self.source.get("dayfirst", True),
            sheet=self.source.get("sheet", "?"),
        )
        if events:
            detail["coercions"] = events

        return self._step("clean", df, excel_rows, **detail)

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


class TableSet:
    """An ordered collection of :class:`Table`, keyed by original sheet name.

    Verbs map over every member and return a new ``TableSet`` so a whole
    workbook flows through one pipeline: ``load_all(p).clean().to_parquet(d)``.
    """

    def __init__(self, tables: dict[str, Table]):
        self._tables: dict[str, Table] = dict(tables)

    def __iter__(self):
        return iter(self._tables)

    def __getitem__(self, key: str) -> Table:
        return self._tables[key]

    def __len__(self) -> int:
        return len(self._tables)

    def __repr__(self) -> str:
        return f"TableSet({list(self._tables)!r})"

    def items(self):
        return self._tables.items()

    def keys(self):
        return self._tables.keys()

    def values(self):
        return self._tables.values()

    def _map(self, verb: str, *a: Any, **k: Any) -> "TableSet":
        return TableSet(
            {key: getattr(t, verb)(*a, **k) for key, t in self._tables.items()}
        )

    def clean(self, *a: Any, **k: Any) -> "TableSet":
        return self._map("clean", *a, **k)

    def kitchen_sink(self, *a: Any, **k: Any) -> "TableSet":
        return self._map("kitchen_sink", *a, **k)

    def cast(self, *a: Any, **k: Any) -> "TableSet":
        return self._map("cast", *a, **k)

    def apply(self, *a: Any, **k: Any) -> "TableSet":
        return self._map("apply", *a, **k)

    def conform(self, *a: Any, **k: Any) -> "TableSet":
        return self._map("conform", *a, **k)

    def dry_run(self, *a: Any, **k: Any) -> "TableSet":
        for t in self._tables.values():
            t.dry_run(*a, **k)
        return self
