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
    entry = {
        "verb": "kitchen_sink",
        "candidates": candidates,
        "winner": {
            "opts": best_opts,
            "score": round(best_score, 4),
            "shape": (best_clean.df.height, best_clean.df.width),
        },
    }
    logger.info(
        "kitchen_sink: chose %s (score %.3f) from %d candidate(s).",
        best_opts,
        best_score,
        len(results),
    )
    return best_ex, best_clean, entry
