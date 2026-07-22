"""Turn a :class:`~xl2y.reader.RawSheet` grid into a clean polars table.

Ported from ``reference/excel_loader.py`` (``_extract_table`` and friends).
Two deliberate departures from the prototype:

* **Polars, not pandas.** The body is assembled column by column: a column
  whose non-null values share one primitive type becomes a typed ``Series``;
  anything mixed is stringified and left for :mod:`xl2y.coerce` to re-type.
* **Headers keep their original text.** snake_casing is ``clean()``'s job,
  and value coercion happens later too. Extraction only *shapes* the table
  (strips titles/comments, applies merges, finds the header and bounding box).

Every heuristic that removes or reinterprets a row also records an ``events``
entry, which seeds the pipeline's lineage.
"""

from __future__ import annotations

import datetime as _dt
import logging
from bisect import bisect_right
from dataclasses import dataclass, field
from typing import Any, Literal

import polars as pl

from xl2y.errors import EmptySheetError
from xl2y.reader import RawSheet

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Tunable heuristics (identical to the prototype)
# --------------------------------------------------------------------------- #

COMMENT_MERGE_WIDTH_FRACTION = 0.5
COMMENT_MERGE_MIN_COLS = 3
SPARSE_ROW_FRACTION = 0.25
SPARSE_ROW_MIN_COLS = 4


@dataclass
class Extracted:
    """A shaped (but not yet cleaned/coerced) table plus provenance."""

    sheet_name: str
    df: pl.DataFrame
    excel_rows: list[int]  # 1-indexed Excel row per df row
    comments: list[dict] = field(default_factory=list)
    events: list[dict] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #


def _is_empty(v: Any) -> bool:
    return v is None or str(v).strip() == ""


def _is_non_numeric_text(v: Any) -> bool:
    """True for strings that don't look like numbers ("1,234" is numeric)."""
    if not isinstance(v, str):
        return False
    try:
        float(v.strip().replace(",", ""))
        return False
    except ValueError:
        return True


def _unique_name(base: str, taken: set[str]) -> str:
    name, n = base, 2
    while name in taken:
        name, n = f"{base}_{n}", n + 1
    return name


def build_columns(header: list[list[Any]], width: int) -> list[str]:
    """Join multi-row header parts with spaces; keep original text.

    Deduplicates collisions ("Revenue", "Revenue_2") and falls back to
    ``col_{i}`` for empty headers. No snake_casing here.
    """
    if not header:
        return [f"col_{i}" for i in range(width)]
    parts_per_col: list[list[str]] = []
    for c in range(width):
        parts, last = [], None
        for row in header:
            v = row[c] if c < len(row) else None
            v = str(v).strip() if v is not None and str(v).strip() else None
            if v is not None and v != last:
                parts.append(v)
                last = v
        parts_per_col.append(parts)
    names: list[str] = []
    taken: set[str] = set()
    for i, parts in enumerate(parts_per_col):
        base = " ".join(parts) if parts else f"col_{i}"
        name = _unique_name(base, taken)
        taken.add(name)
        names.append(name)
    return names


def _build_series(name: str, values: list[Any]) -> pl.Series:
    """Type a single column defensively: never let polars raise on mixed
    input. A single-primitive-type column becomes that type; a numeric
    int/float mix becomes Float64; anything else is stringified."""
    types = {type(v) for v in values if v is not None}
    if not types:
        return pl.Series(name, values, dtype=pl.Utf8)
    try:
        if types == {bool}:
            return pl.Series(name, values, dtype=pl.Boolean)
        if types <= {int}:
            return pl.Series(name, values, dtype=pl.Int64)
        if types <= {int, float}:
            # An integer column with a stray decimal-formatted cell arrives as
            # an int/float mix; match pandas convert_dtypes and keep it Int64
            # when every value is whole.
            s = pl.Series(
                name,
                [float(v) if v is not None else None for v in values],
                dtype=pl.Float64,
            )
            nn = s.drop_nulls()
            if nn.len() and bool(((nn % 1) == 0).all()):
                return s.cast(pl.Int64)
            return s
        if types == {str}:
            return pl.Series(name, values, dtype=pl.Utf8)
        if types <= {_dt.datetime, _dt.date, _dt.time}:
            return pl.Series(name, values)  # polars infers Date/Datetime/Time
    except (TypeError, OverflowError, pl.exceptions.PolarsError):
        pass
    return pl.Series(
        name, [None if v is None else str(v) for v in values], dtype=pl.Utf8
    )


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #


def extract_table(
    raw: RawSheet,
    header_rows: int = 1,
    sparse_rows: Literal["comment", "keep", "section"] = "comment",
    skip_hidden: bool = False,
    check_formula_cache: bool = True,
) -> Extracted:
    title = raw.name
    # Work on a copy: merge propagation below fills and pads cells, and
    # kitchen_sink re-extracts the same RawSheet under different options.
    grid = [list(row) for row in raw.grid]
    merged_ranges = raw.merged
    if not grid:
        raise EmptySheetError(f"Sheet {title!r} is empty")

    events: list[dict] = []
    formula_cells = raw.formulas if check_formula_cache else ()

    # ------------------------------------------------------------------ #
    # 1. Propagate merged values across the full merged rectangle.
    # ------------------------------------------------------------------ #
    uncached_coords = {
        (fr - 1, fc - 1)
        for fr, fc in formula_cells
        if fr - 1 >= len(grid)
        or fc - 1 >= len(grid[fr - 1])
        or grid[fr - 1][fc - 1] is None
    }

    propagated: set[tuple[int, int]] = set()  # 0-indexed copies (not anchors)
    for min_r, min_c, max_r, max_c in merged_ranges:
        if min_r - 1 >= len(grid) or min_c - 1 >= len(grid[min_r - 1]):
            continue
        anchor = grid[min_r - 1][min_c - 1]
        if anchor is None:
            continue
        for r in range(min_r - 1, min(max_r, len(grid))):
            if len(grid[r]) < max_c:  # pad trimmed rows inside the merge
                grid[r].extend([None] * (max_c - len(grid[r])))
            for c in range(min_c - 1, max_c):
                if (r, c) != (min_r - 1, min_c - 1):
                    if grid[r][c] is None:
                        grid[r][c] = anchor
                    propagated.add((r, c))

    # ------------------------------------------------------------------ #
    # 2. Bounding box: trim empty rows/cols, optionally drop hidden.
    # ------------------------------------------------------------------ #
    hidden_rows = raw.hidden_rows if skip_hidden else set()
    grid_cols = max((len(r) for r in grid), default=0)
    hidden_cols = (
        {
            c
            for lo, hi in raw.hidden_col_intervals
            for c in range(max(lo, 0), min(hi, grid_cols - 1) + 1)
        }
        if skip_hidden
        else set()
    )
    if hidden_rows or hidden_cols:
        logger.info(
            "Sheet %r: skipping %d hidden row(s), %d hidden column(s).",
            title,
            len(hidden_rows),
            len(hidden_cols),
        )
        events.append(
            {
                "event": "hidden_skipped",
                "rows": len(hidden_rows),
                "cols": len(hidden_cols),
            }
        )

    def visible_rows() -> list[int]:
        return [r for r in range(len(grid)) if r not in hidden_rows]

    row_has = {
        r: any(not _is_empty(v) for v in grid[r]) for r in visible_rows()
    }
    data_rows = [r for r, has in row_has.items() if has]
    if not data_rows:
        raise EmptySheetError(f"Sheet {title!r} has no data")
    top, bottom = data_rows[0], data_rows[-1]

    n_cols = max(len(grid[r]) for r in data_rows)
    col_counts = [
        sum(
            1
            for r in range(top, bottom + 1)
            if r not in hidden_rows
            and len(grid[r]) > c
            and not _is_empty(grid[r][c])
        )
        for c in range(n_cols)
    ]
    runs: list[tuple[int, int]] = []  # (start, end) inclusive, non-empty cols
    start = None
    for c, count in enumerate(col_counts + [0]):
        if count and start is None:
            start = c
        elif not count and start is not None:
            runs.append((start, c - 1))
            start = None
    left, right = max(
        runs, key=lambda run: sum(col_counts[run[0] : run[1] + 1])
    )
    table_cols = [c for c in range(left, right + 1) if c not in hidden_cols]
    width = len(table_cols)
    if len(runs) > 1:
        logger.info(
            "Sheet %r: multiple column blocks found; using columns %d-%d as "
            "the table and ignoring the rest.",
            title,
            left + 1,
            right + 1,
        )
        events.append(
            {"event": "columns_ignored", "used": (left + 1, right + 1)}
        )

    # ------------------------------------------------------------------ #
    # 3. Merged-comment detection.
    # ------------------------------------------------------------------ #
    merged_comment_rows: dict[int, tuple[int, int]] = {}  # row -> anchor (r,c)
    if width >= COMMENT_MERGE_MIN_COLS:
        for min_r, min_c, max_r, max_c in merged_ranges:
            if min_r - 1 >= len(grid) or min_c - 1 >= len(grid[min_r - 1]):
                continue
            if _is_empty(grid[min_r - 1][min_c - 1]):  # formatting-only merge
                continue
            overlap = sum(1 for c in table_cols if min_c - 1 <= c <= max_c - 1)
            if overlap < width * COMMENT_MERGE_WIDTH_FRACTION:
                continue
            span_rows = [
                r
                for r in range(min_r - 1, min(max_r, len(grid)))
                if r not in hidden_rows
            ]
            clean = all(
                (_is_empty(grid[r][c]) if c < len(grid[r]) else True)
                for r in span_rows
                for c in table_cols
                if not (min_c - 1 <= c <= max_c - 1)
            )
            if clean:
                for r in span_rows:
                    merged_comment_rows.setdefault(r, (min_r - 1, min_c - 1))

    # ------------------------------------------------------------------ #
    # 4. Classify rows: comments, then headers, then body.
    # ------------------------------------------------------------------ #
    header: list[list[Any]] = []
    body_rows: list[tuple[int, list[Any], str | None]] = []
    raw_comments: list[tuple[int, str, str]] = []  # (excel_row, text, kind)
    seen_merge_anchors: set[tuple[int, int]] = set()
    current_section: str | None = None

    def row_text(r: int, values: list[Any]) -> str:
        parts = []
        for c, v in zip(table_cols, values):
            if _is_empty(v):
                continue
            if (r, c) in propagated:
                continue
            parts.append(str(v).strip())
        return " ".join(parts)

    for r in range(top, bottom + 1):
        if r in hidden_rows:
            continue
        values = [grid[r][c] if c < len(grid[r]) else None for c in table_cols]

        # A wide merged banner normally reads as a title to strip. But when the
        # caller asked for a multi-row header, a banner among the first
        # `header_rows` rows is part of that header, not a comment sitting above
        # it — pull banners out only once the header region is filled. (For the
        # default single-row header, a leading banner still strips, so the real
        # header below it is found.)
        strip_merged = r in merged_comment_rows and not (
            header_rows > 1 and len(header) < header_rows
        )
        if strip_merged:
            anchor = merged_comment_rows[r]
            if anchor not in seen_merge_anchors:
                seen_merge_anchors.add(anchor)
                text = row_text(r, values)
                if text:
                    raw_comments.append((anchor[0] + 1, text, "merged"))
                    events.append(
                        {
                            "event": "merged_comment",
                            "excel_row": anchor[0] + 1,
                            "text": text,
                        }
                    )
            continue

        own_non_null = [
            v
            for c, v in zip(table_cols, values)
            if not _is_empty(v) and (r, c) not in propagated
        ]
        non_null = [v for v in values if not _is_empty(v)]

        if not non_null:
            continue

        if (
            sparse_rows in ("comment", "section")
            and len(header) >= header_rows
            and width >= SPARSE_ROW_MIN_COLS
            and 0 < len(own_non_null) <= max(1, int(width * SPARSE_ROW_FRACTION))
            and len(own_non_null) == len(non_null)
            and all(_is_non_numeric_text(v) for v in own_non_null)
        ):
            text = row_text(r, values)
            if sparse_rows == "section":
                current_section = text
                logger.info(
                    "Sheet %r: row %d used as section label %r.",
                    title,
                    r + 1,
                    text,
                )
                events.append(
                    {"event": "section_label", "excel_row": r + 1, "text": text}
                )
            else:
                logger.warning(
                    "Sheet %r: treating sparse row %d as a comment "
                    "(text: %.80r). Pass sparse_rows='keep' to keep such rows "
                    "as data, or sparse_rows='section' to forward-fill them "
                    "into a section column.",
                    title,
                    r + 1,
                    text,
                )
                raw_comments.append((r + 1, text, "sparse"))
                events.append(
                    {
                        "event": "sparse_row_stripped",
                        "excel_row": r + 1,
                        "text": text,
                    }
                )
            continue

        if len(header) < header_rows:
            header.append(values)
            continue

        if own_non_null:
            body_rows.append((r + 1, values, current_section))

    if len(header) < header_rows:
        raise ValueError(
            f"Sheet {title!r}: requested header_rows={header_rows} but only "
            f"{len(header)} non-comment row(s) available."
        )

    uncached = sum(
        1
        for (r, c) in uncached_coords
        if top <= r <= bottom and c in table_cols and r not in hidden_rows
    )
    if uncached:
        logger.warning(
            "Sheet %r: %d formula cell(s) have no cached value and will read "
            "as missing. Open and save the file in Excel to populate them.",
            title,
            uncached,
        )
        events.append({"event": "uncached_formulas", "count": uncached})

    # ------------------------------------------------------------------ #
    # 5. Build the DataFrame column by column; link comments to rows.
    # ------------------------------------------------------------------ #
    columns = build_columns(header, width)
    series: list[pl.Series] = []
    if sparse_rows == "section" and any(s is not None for _, _, s in body_rows):
        section_name = _unique_name("section", set(columns))
        series.append(
            _build_series(section_name, [s for _, _, s in body_rows])
        )
    for i, name in enumerate(columns):
        series.append(
            _build_series(name, [vals[i] for _, vals, _ in body_rows])
        )
    df = pl.DataFrame(series)

    body_excel_rows = [er for er, _, _ in body_rows]
    comments = []
    for excel_row, text, kind in sorted(raw_comments):
        pos = bisect_right(body_excel_rows, excel_row - 1)
        comments.append(
            {
                "text": text,
                "excel_row": excel_row,
                "kind": kind,
                "before_df_position": pos
                if pos < len(body_excel_rows)
                else None,
            }
        )

    return Extracted(
        sheet_name=title,
        df=df,
        excel_rows=body_excel_rows,
        comments=comments,
        events=events,
    )
