"""``kitchen_sink``: try several interpretations of a messy sheet and keep the
one that looks most like a competent table.

The tournament varies the number of header rows and how sparse rows are
handled. It deliberately only considers information-*preserving* sparse modes
(``section`` forward-fills them into a column; ``keep`` leaves them as data) —
never ``comment``, which would discard rows — so "just get it into something"
never silently loses data. Each candidate is extracted, cleaned, and scored;
the winner and all also-rans are recorded in lineage for ``dry_run``.
"""

from __future__ import annotations

import logging
import re

import polars as pl

from xl2y.errors import EmptySheetError
from xl2y.extract import extract_table
from xl2y.reader import RawSheet

logger = logging.getLogger(__name__)

CANDIDATE_GRID = [
    {"header_rows": h, "sparse_rows": s}
    for h in (1, 2, 3)
    for s in ("section", "keep")
]

_MONTHS = {
    "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct",
    "nov", "dec", "january", "february", "march", "april", "june", "july",
    "august", "september", "october", "november", "december",
}
_YEAR = re.compile(r"(?:19|20)\d{2}")
_ISO = re.compile(r"\d{4}_\d{2}_\d{2}$")

# A long-format (tidy) table is preferred over a wide period-per-column one.
UNPIVOT_BONUS = 0.25


def _looks_period(name: str) -> bool:
    n = name.lower()
    if _YEAR.search(n):
        return True
    if n.strip("_") in _MONTHS:
        return True
    return bool(_ISO.search(n))


def _try_unpivot(df: pl.DataFrame) -> pl.DataFrame | None:
    """If the table ends in a run of >=3 numeric, period-named columns
    (2021, jan, 2025-01-01, ...), reshape it to long form; else None."""
    period_cols: list[str] = []
    for c in reversed(df.columns):
        if df[c].dtype.is_numeric() and _looks_period(c):
            period_cols.append(c)
        else:
            break
    period_cols.reverse()
    if len(period_cols) < 3:
        return None
    index = [c for c in df.columns if c not in period_cols]
    if not index:  # need at least one identifier column to pivot around
        return None
    out = df.unpivot(
        on=period_cols,
        index=index,
        variable_name="period",
        value_name="value",
    )
    return out.with_columns(pl.col("period").str.replace(r"^col_", ""))


def score(df: pl.DataFrame) -> float:
    """A 'competent table' score. Rewards typed and dense columns; penalises
    ``col_N`` fallback names, forced dedup suffixes, and degenerate shapes."""
    if df.height == 0 or df.width == 0:
        return float("-inf")
    cells = df.height * df.width
    nulls = sum(df.null_count().row(0))
    density = 1 - nulls / cells
    typed = sum(1 for d in df.dtypes if d != pl.Utf8) / df.width
    fallback = sum(1 for c in df.columns if c.startswith("col_")) / df.width
    dupey = (
        sum(1 for c in df.columns if c.rsplit("_", 1)[-1].isdigit()) / df.width
    )
    shape_penalty = 0.5 if df.height < 2 else 0.0
    return typed + 1.5 * density - 2.0 * fallback - 1.0 * dupey - shape_penalty


def run_tournament(raw: RawSheet, source: dict):
    """Return ``(winning_extracted, winning_cleaned_table, lineage_entry)``."""
    # Imported here to avoid a table <-> sink import cycle.
    from xl2y.table import Table

    results = []  # (score, opts, extracted, cleaned_table)
    for opts in CANDIDATE_GRID:
        try:
            ex = extract_table(raw, **opts)
        except (ValueError, EmptySheetError):
            continue
        tbl = Table(
            df=ex.df,
            source=source,
            excel_rows=ex.excel_rows,
            comments=ex.comments,
        )
        cleaned = tbl.clean()
        results.append((score(cleaned.df), opts, ex, cleaned))

    if not results:
        raise EmptySheetError(
            f"Sheet {raw.name!r}: no candidate interpretation produced a table."
        )

    results.sort(key=lambda r: r[0], reverse=True)
    best_score, best_opts, best_ex, best_clean = results[0]
    candidates = [
        {
            "opts": o,
            "score": round(s, 4),
            "shape": (c.df.height, c.df.width),
        }
        for s, o, _, c in results
    ]

    winner = {
        "df": best_clean.df,
        "excel_rows": best_ex.excel_rows,
        "comments": best_ex.comments,
        "opts": best_opts,
        "score": best_score,
    }

    # Extra candidate: a wide period-per-column table reshaped to long form.
    unpiv = _try_unpivot(best_clean.df)
    if unpiv is not None:
        us = score(unpiv) + UNPIVOT_BONUS
        candidates.append(
            {
                "opts": {"unpivot": True},
                "score": round(us, 4),
                "shape": (unpiv.height, unpiv.width),
            }
        )
        if us > winner["score"]:
            # unpivot multiplies rows, so Excel-row identity no longer maps.
            winner = {
                "df": unpiv,
                "excel_rows": None,
                "comments": [],
                "opts": {"unpivot": True},
                "score": us,
            }

    entry = {
        "verb": "kitchen_sink",
        "candidates": candidates,
        "winner": {
            "opts": winner["opts"],
            "score": round(winner["score"], 4),
            "shape": (winner["df"].height, winner["df"].width),
        },
    }
    logger.info(
        "kitchen_sink: chose %s (score %.3f) from %d candidate(s).",
        winner["opts"],
        winner["score"],
        len(candidates),
    )
    return winner["df"], winner["excel_rows"], winner["comments"], entry
