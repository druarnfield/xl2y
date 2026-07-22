"""xl2y — Excel in, cleaning pipeline, Parquet out."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from xl2y import extract, patterns, reader
from xl2y.errors import (
    EmptySheetError,
    SchemaError,
    SheetNotFoundError,
    UnsupportedFormatError,
    Xl2yError,
)
from xl2y.extract import Extracted
from xl2y.reader import RawSheet
from xl2y.schema import (
    Schema,
    bool_,
    cat_,
    date_,
    datetime_,
    float_,
    int_,
    str_,
)
from xl2y.table import Table, TableSet

logger = logging.getLogger(__name__)

__all__ = [
    "load",
    "load_all",
    "Table",
    "TableSet",
    "Schema",
    "str_",
    "int_",
    "float_",
    "bool_",
    "date_",
    "datetime_",
    "cat_",
    "patterns",
    "Xl2yError",
    "UnsupportedFormatError",
    "EmptySheetError",
    "SheetNotFoundError",
    "SchemaError",
]


def _to_table(
    path: Path, ex: Extracted, raw: RawSheet, dayfirst: bool
) -> Table:
    return Table(
        df=ex.df,
        source={
            "path": str(path),
            "sheet": ex.sheet_name,
            "dayfirst": dayfirst,
        },
        excel_rows=ex.excel_rows,
        lineage=[{"verb": "load", "events": ex.events}],
        comments=ex.comments,
        _raw=raw,
    )


def load(path: str | Path, sheet: str | None = None, **hints: Any) -> Table:
    """Load one table.

    With ``sheet`` given, load exactly that worksheet. Otherwise load every
    worksheet and keep the one that extracts the biggest table (rows ×
    columns), logging the choice when there is more than one candidate.

    ``hints`` forward to :func:`xl2y.extract.extract_table` (``header_rows``,
    ``sparse_rows``, ``skip_hidden``, ``check_formula_cache``,
    ``header_min_fill``); ``dayfirst`` is stored for later coercion.
    """
    dayfirst = hints.pop("dayfirst", True)
    path = Path(path)

    if sheet is not None:
        raw = reader.read_sheet(path, sheet)
        ex = extract.extract_table(raw, **hints)
        return _to_table(path, ex, raw, dayfirst)

    best: tuple[int, Extracted, RawSheet] | None = None
    candidates = 0
    for raw in reader.read_all(path):
        try:
            ex = extract.extract_table(raw, **hints)
        except EmptySheetError:
            continue
        candidates += 1
        score = ex.df.height * ex.df.width
        if best is None or score > best[0]:
            best = (score, ex, raw)

    if best is None:
        raise EmptySheetError(f"{path.name}: no sheet contains a table.")
    if candidates > 1:
        logger.info(
            "Selected sheet %r from %d candidate sheet(s).",
            best[1].sheet_name,
            candidates,
        )
    return _to_table(path, best[1], best[2], dayfirst)


def load_all(path: str | Path, **hints: Any) -> TableSet:
    """Load every worksheet into a :class:`TableSet`, keyed by sheet name.

    Empty sheets are skipped (with an INFO log). The same ``hints`` as
    :func:`load` apply to each sheet.
    """
    dayfirst = hints.pop("dayfirst", True)
    path = Path(path)

    tables: dict[str, Table] = {}
    for raw in reader.read_all(path):
        try:
            ex = extract.extract_table(raw, **hints)
        except EmptySheetError:
            logger.info("Skipping empty sheet: %r", raw.name)
            continue
        tables[raw.name] = _to_table(path, ex, raw, dayfirst)
    return TableSet(tables)
