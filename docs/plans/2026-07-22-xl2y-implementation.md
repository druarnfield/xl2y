# xl2y Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build the xl2y library (Excel → cleaning pipeline → Parquet) per `docs/plans/2026-07-22-xl2y-design.md`.

**Architecture:** Streaming two-pass Excel reader (XML structure pass + openpyxl value pass) feeding a polars-backed immutable `Table` with chainable verbs and lineage tracking. Extraction heuristics ported from the existing prototype `reference/excel_loader.py` (pandas → polars). Schema validation and a `kitchen_sink()` candidate tournament on top.

**Tech Stack:** Python ≥3.13, polars, openpyxl, pytest, uv.

**Working directory:** `/Users/dru/Documents/Development/python/xl2y` (NO worktree — work directly on `main`).

**Conventions for every task:**
- TDD: write the failing test first, see it fail, implement, see it pass, commit.
- Run tests with `uv run pytest tests/ -v` (or a specific file/test).
- Commit messages: conventional style (`feat:`, `test:`, `refactor:`, `chore:`), **never mention Claude**.
- The prototype `reference/excel_loader.py` is the porting source. Line references below point at it. Its logic is trusted (it works); the port changes pandas → polars and reshapes outputs — do not "improve" heuristics while porting.

---

## Task 1: Project setup

**Files:**
- Move: `excel_loader.py` → `reference/excel_loader.py`
- Delete: `main.py`, `test_excel_load.py` (prototype scratch)
- Modify: `pyproject.toml`
- Create: `xl2y/__init__.py`, `xl2y/errors.py`, `tests/__init__.py` (empty), `tests/test_package.py`

**Step 1: Move prototype and commit it (preserve history)**

```bash
mkdir -p reference
git mv 2>/dev/null; mv excel_loader.py reference/excel_loader.py
rm main.py test_excel_load.py
git add reference/excel_loader.py .gitignore .python-version README.md uv.lock pyproject.toml
git commit -m "chore: import prototype loader as reference, remove scratch files"
```

**Step 2: Update `pyproject.toml`** (full contents):

```toml
[project]
name = "xl2y"
version = "0.1.0"
description = "Excel in, cleaning pipeline, Parquet out"
readme = "README.md"
requires-python = ">=3.13"
dependencies = [
    "openpyxl>=3.1.5",
    "polars>=1.17",
]

[dependency-groups]
dev = ["pytest>=8"]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["xl2y"]
```

Then: `uv sync` (installs polars + pytest, makes `xl2y` importable).

**Step 3: Write the failing test** — `tests/test_package.py`:

```python
def test_package_imports():
    import xl2y
    assert xl2y.__all__  # public API declared


def test_error_hierarchy():
    from xl2y.errors import (
        Xl2yError, UnsupportedFormatError, EmptySheetError,
        SheetNotFoundError, SchemaError,
    )
    for exc in (UnsupportedFormatError, EmptySheetError,
                SheetNotFoundError, SchemaError):
        assert issubclass(exc, Xl2yError)
```

Run: `uv run pytest tests/test_package.py -v` → FAIL (no package).

**Step 4: Implement**

`xl2y/errors.py`:

```python
class Xl2yError(Exception):
    """Base for all xl2y errors."""


class UnsupportedFormatError(Xl2yError, ValueError):
    """Spreadsheet format openpyxl cannot open (.xls, .xlsb, ...)."""


class EmptySheetError(Xl2yError, ValueError):
    """Sheet contains no data."""


class SheetNotFoundError(Xl2yError, KeyError):
    """Requested sheet does not exist; message lists available sheets."""


class SchemaError(Xl2yError, ValueError):
    """Validation failed. Carries every problem, not just the first."""

    def __init__(self, message: str, problems: list):
        super().__init__(message)
        self.problems = problems
```

`xl2y/__init__.py` (grows over the project; start minimal):

```python
from xl2y.errors import (
    Xl2yError, UnsupportedFormatError, EmptySheetError,
    SheetNotFoundError, SchemaError,
)

__all__ = [
    "Xl2yError", "UnsupportedFormatError", "EmptySheetError",
    "SheetNotFoundError", "SchemaError",
]
```

**Step 5:** `uv run pytest tests/test_package.py -v` → PASS. Commit: `feat: package skeleton with error hierarchy`.

---

## Task 2: Fixture workbook builders

All tests use in-code generated workbooks. No binary fixtures in git.

**Files:**
- Create: `tests/fixtures.py`, `tests/test_fixtures.py`

**Step 1: Failing test** — `tests/test_fixtures.py`:

```python
from openpyxl import load_workbook
from tests import fixtures


def test_simple_book(tmp_path):
    p = fixtures.simple_book(tmp_path / "s.xlsx")
    ws = load_workbook(p)["Data"]
    assert ws["A1"].value == "Store"
    assert ws.max_row == 4


def test_nasty_book_has_banner_merge(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    ws = load_workbook(p)["Report"]
    assert any(str(r) == "A1:D1" for r in ws.merged_cells.ranges)
```

Run → FAIL.

**Step 2: Implement** — `tests/fixtures.py`:

```python
"""Generated fixture workbooks. Each builder returns the saved path."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill

ROWS = [
    ["Sydney", "01/02/2025", "$1,234"],
    ["Melbourne", "02/02/2025", "(500)"],
    ["Brisbane", "03/02/2025", "N/A"],
]


def simple_book(path: Path) -> Path:
    """One sheet 'Data', clean 1-row header, 3 data rows."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Store", "Date", "Revenue"])
    for r in ROWS:
        ws.append(r)
    wb.save(path)
    return path


def nasty_book(path: Path) -> Path:
    """Sheet 'Report': merged title banner (A1:D1), blank row, header,
    data with a sparse section row, plus a second empty sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "2025 Results — CONFIDENTIAL"
    ws.merge_cells("A1:D1")
    # row 2 blank
    ws.append([])
    ws.append(["Store", "Date", "Units", "Revenue"])
    ws.append(["Northern Region"])                      # sparse section row
    ws.append(["Sydney", "01/02/2025", 10, "$1,234"])
    ws.append(["Newcastle", "01/02/2025", 4, "$400"])
    ws.append(["Southern Region"])
    ws.append(["Melbourne", "02/02/2025", 7, "(500)"])
    wb.create_sheet("Empty")
    wb.save(path)
    return path


def multi_header_book(path: Path) -> Path:
    """Two header rows: 'Revenue' merged over Q1/Q2."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Wide"
    ws.append(["Store", "Revenue", None])
    ws.merge_cells("B1:C1")
    ws.append([None, "Q1", "Q2"])
    ws.append(["Sydney", 1, 2])
    ws.append(["Melbourne", 3, 4])
    wb.save(path)
    return path


def bloated_book(path: Path) -> Path:
    """Real table is 3x2 but a styled empty cell inflates the dimension
    to row 1,048,000. Must load fast with bounded memory."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bloat"
    ws.append(["A", "B"])
    ws.append([1, 2])
    ws.append([3, 4])
    ws.cell(row=1_048_000, column=1).fill = PatternFill("solid", fgColor="FFFF00")
    wb.save(path)
    return path


def hidden_book(path: Path) -> Path:
    """Header + 3 data rows; row 3 (Excel) hidden, column C hidden."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Hidden"
    ws.append(["A", "B", "Scratch"])
    ws.append([1, 2, "x"])
    ws.append([9, 9, "stale"])
    ws.append([3, 4, "y"])
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["C"].hidden = True
    wb.save(path)
    return path


def formula_book(path: Path) -> Path:
    """Formula cell with no cached value (never opened in Excel)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws.append(["A", "Total"])
    ws.append([1, "=SUM(A2:A2)"])
    wb.save(path)
    return path


def typed_book(path: Path) -> Path:
    """Columns exercising coercion: currency, percent, NA tokens,
    day-first dates, plain ints, free text."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Types"
    ws.append(["money", "pct", "when", "count", "note"])
    ws.append(["$1,234.50", "15%", "03/04/2025", 1, "hello"])
    ws.append(["(500)", "7.5%", "04/04/2025", 2, "world"])
    ws.append(["N/A", "-", "05/04/2025", 3, "-"])
    wb.save(path)
    return path
```

**Step 3:** Run → PASS. Commit: `test: fixture workbook builders`.

---

## Task 3: reader — format check and sheet metadata

**Files:**
- Create: `xl2y/reader.py`, `tests/test_reader.py`

**Port sources:** `reference/excel_loader.py` lines 365–378 (`_check_format`, verbatim), 405–433 (`_sheet_meta`, verbatim), constant `SUPPORTED_SUFFIXES` line 110.

**Step 1: Failing test** — `tests/test_reader.py`:

```python
import pytest

from tests import fixtures
from xl2y import reader
from xl2y.errors import UnsupportedFormatError


def test_rejects_xls(tmp_path):
    p = tmp_path / "old.xls"
    p.touch()
    with pytest.raises(UnsupportedFormatError, match="Convert"):
        reader.check_format(p)


def test_rejects_unknown_suffix(tmp_path):
    p = tmp_path / "f.csv"
    p.touch()
    with pytest.raises(UnsupportedFormatError):
        reader.check_format(p)


def test_sheet_meta_lists_sheets_in_order(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    meta = reader.sheet_meta(p)
    assert list(meta) == ["Report", "Empty"]
    member, is_chart = meta["Report"]
    assert member.startswith("xl/") and not is_chart
```

Run → FAIL.

**Step 2: Implement** `xl2y/reader.py`: module docstring, then port `_check_format` → `check_format(path)` (raise `xl2y.errors.UnsupportedFormatError`) and `_sheet_meta` → `sheet_meta(path)` exactly as in the prototype (they have no pandas dependency). Also port `_localname` (line 405).

**Step 3:** Run → PASS. Commit: `feat: reader format check and sheet metadata`.

---

## Task 4: reader — structure pass, value pass, RawSheet

**Files:**
- Modify: `xl2y/reader.py`
- Test: `tests/test_reader.py` (append)

**Port sources:** lines 436–510 (`_sheet_structure`, verbatim), 513–539 (`_read_grid`, verbatim), 390–391 (`_is_empty`).

**Step 1: Failing tests** (append to `tests/test_reader.py`):

```python
import time


def test_read_sheet_returns_rawsheet(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    raw = reader.read_sheet(p, "Report")
    assert raw.name == "Report"
    assert raw.grid[0][0] == "2025 Results — CONFIDENTIAL"
    assert (1, 1, 1, 4) in raw.merged        # A1:D1 as (min_r, min_c, max_r, max_c)
    assert raw.grid[1] == []                  # blank row preserved as placeholder


def test_hidden_rows_and_cols_detected(tmp_path):
    p = fixtures.hidden_book(tmp_path / "h.xlsx")
    raw = reader.read_sheet(p, "Hidden")
    assert 2 in raw.hidden_rows               # Excel row 3, 0-indexed
    assert (2, 2) in raw.hidden_col_intervals # column C, 0-indexed interval


def test_formula_cells_detected(tmp_path):
    p = fixtures.formula_book(tmp_path / "f.xlsx")
    raw = reader.read_sheet(p, "Calc")
    assert (2, 2) in raw.formulas             # B2, 1-indexed


def test_bloated_file_loads_fast(tmp_path):
    p = fixtures.bloated_book(tmp_path / "b.xlsx")
    t0 = time.monotonic()
    raw = reader.read_sheet(p, "Bloat")
    assert time.monotonic() - t0 < 5.0
    assert len(raw.grid) == 3                 # styled ghost cell ignored


def test_missing_sheet_raises_with_available(tmp_path):
    from xl2y.errors import SheetNotFoundError
    p = fixtures.simple_book(tmp_path / "s.xlsx")
    with pytest.raises(SheetNotFoundError, match="Data"):
        reader.read_sheet(p, "Nope")
```

Run → FAIL.

**Step 2: Implement.** Add to `xl2y/reader.py`:

```python
@dataclass
class RawSheet:
    """Everything the extractor needs about one worksheet, fully in memory
    but bounded by the true data extent (never the declared dimension)."""
    name: str
    grid: list[list]                       # trimmed rows; [] = empty row
    merged: list[tuple[int, int, int, int]]  # (min_r, min_c, max_r, max_c) 1-idx
    hidden_rows: set[int]                  # 0-indexed
    hidden_col_intervals: list[tuple[int, int]]  # 0-indexed inclusive
    formulas: list[tuple[int, int]]        # (row, col) 1-indexed
```

Then `read_sheet(path, sheet_name) -> RawSheet` and `read_all(path) -> Iterator[RawSheet]`:
- `check_format`, `sheet_meta`; unknown name → `SheetNotFoundError(f"Sheet {name!r} not found. Available: {list(meta)}")`; chartsheet requested by name → `ValueError` (per prototype lines 204–207). `read_all` skips chartsheets with a `logger.info`.
- Open workbook `load_workbook(path, data_only=True, read_only=True)` **once**, `zipfile.ZipFile(path)` once; for each wanted sheet: `ws.reset_dimensions()`, run ported `_sheet_structure` (rename `sheet_structure`), then ported `_read_grid` (rename `read_grid`) bounded by the extent. Close the workbook in `finally` (prototype lines 215–258 shows the pattern).
- `read_all` must be a generator that yields one `RawSheet` at a time (memory: peak = largest sheet).

**Step 3:** Run → PASS. Commit: `feat: streaming structure and value passes into RawSheet`.

---

## Task 5: extract — table extraction to polars

The big port. Input `RawSheet` + options; output an `Extracted` result with a polars DataFrame, original-text headers, comments, and Excel row mapping.

**Files:**
- Create: `xl2y/extract.py`, `tests/test_extract.py`

**Port sources:** lines 542–837 (`_extract_table`), 847–869 (`_build_columns`), 840–844 (`_unique_name`), 394–402 (`_is_non_numeric_text`), heuristic constants lines 59–104.

**Key changes from the prototype (everything else ports verbatim):**
1. **No pandas.** Build the DataFrame per-column: for each column, collect values; if all non-None values share one primitive type (`str`/`int`/`float`/`bool`/`datetime`) build a typed `pl.Series`, else cast every non-None value to `str` (mixed columns become text; coercion re-types them later). Never let polars raise on mixed input.
2. **Headers keep original text.** `build_columns` joins multi-row header parts with a single space, strips, deduplicates with `_unique_name` ("Revenue", "Revenue_2"), and falls back to `col_{i}` for empty headers. NO snake_casing here — that is `clean()`'s job.
3. **No value coercion here** (`coerce_values` parameter and `_coerce_columns` call are dropped; coercion moves to Task 8).
4. **Return type:**

```python
@dataclass
class Extracted:
    sheet_name: str
    df: pl.DataFrame
    excel_rows: list[int]        # 1-indexed Excel row per df row
    comments: list[dict]         # {"text", "excel_row", "kind", "before_df_position"}
    events: list[dict]           # lineage seeds: comments stripped, sparse rows,
                                 # hidden skipped, column runs ignored, uncached formulas
```

Every `logger.warning`/`logger.info` in the prototype ALSO appends a dict to `events` (e.g. `{"event": "sparse_row_stripped", "excel_row": 5, "text": "Northern Region"}`). Keep the log calls.

**Step 1: Failing tests** — `tests/test_extract.py`:

```python
import pytest

from tests import fixtures
from xl2y import extract, reader
from xl2y.errors import EmptySheetError


def _extracted(tmp_path, builder, sheet, **opts):
    p = builder(tmp_path / "f.xlsx")
    return extract.extract_table(reader.read_sheet(p, sheet), **opts)


def test_simple_table(tmp_path):
    ex = _extracted(tmp_path, fixtures.simple_book, "Data")
    assert ex.df.columns == ["Store", "Date", "Revenue"]
    assert ex.df.height == 3
    assert ex.excel_rows == [2, 3, 4]


def test_banner_becomes_comment_not_header(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report")
    assert ex.df.columns == ["Store", "Date", "Units", "Revenue"]
    assert ex.comments[0]["kind"] == "merged"
    assert "CONFIDENTIAL" in ex.comments[0]["text"]


def test_sparse_rows_default_comment(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report")
    texts = [c["text"] for c in ex.comments if c["kind"] == "sparse"]
    assert "Northern Region" in texts
    assert ex.df.height == 3            # section rows stripped


def test_sparse_rows_section_mode(tmp_path):
    ex = _extracted(tmp_path, fixtures.nasty_book, "Report",
                    sparse_rows="section")
    assert ex.df.columns[0] == "section"
    assert ex.df["section"].to_list() == [
        "Northern Region", "Northern Region", "Southern Region"]


def test_multi_row_header(tmp_path):
    ex = _extracted(tmp_path, fixtures.multi_header_book, "Wide",
                    header_rows=2)
    assert ex.df.columns == ["Store", "Revenue Q1", "Revenue Q2"]


def test_hidden_skipped_when_asked(tmp_path):
    ex = _extracted(tmp_path, fixtures.hidden_book, "Hidden",
                    skip_hidden=True)
    assert ex.df.columns == ["A", "B"]
    assert ex.df.height == 2
    assert ex.excel_rows == [2, 4]


def test_empty_sheet_raises(tmp_path):
    p = fixtures.nasty_book(tmp_path / "n.xlsx")
    with pytest.raises(EmptySheetError):
        extract.extract_table(reader.read_sheet(p, "Empty"))


def test_mixed_type_column_becomes_text(tmp_path):
    ex = _extracted(tmp_path, fixtures.simple_book, "Data")
    assert ex.df["Revenue"].dtype.is_(__import__("polars").Utf8)
```

Run → FAIL.

**Step 2: Implement** `xl2y/extract.py` with signature:

```python
def extract_table(
    raw: RawSheet,
    header_rows: int = 1,
    sparse_rows: Literal["comment", "keep", "section"] = "comment",
    skip_hidden: bool = False,
    check_formula_cache: bool = True,
) -> Extracted: ...
```

Port `_extract_table` section by section (the prototype's numbered comment blocks 1–5 map directly). At block 5, replace the pandas DataFrame build with the per-column typed/str build described above; `section` column inserted first when applicable (dedupe its name against headers with `_unique_name`). `excel_rows` = the `body_excel_rows` list the prototype already computes (line 822). Uncached-formula count (lines 794–807) becomes an event + warning.

**Step 3:** Run → PASS (iterate; this is the largest single task — if it exceeds ~an hour, split the polars-build change into its own commit first using pandas-free stub assertions).

**Step 4:** Commit: `feat: table extraction heuristics ported to polars`.

---

## Task 6: Table object, apply, cast, collect

**Files:**
- Create: `xl2y/table.py`, `tests/test_table.py`

**Step 1: Failing tests** — `tests/test_table.py`:

```python
import polars as pl
import pytest

from xl2y.table import Table


def _table():
    df = pl.DataFrame({"a": ["1", "2"], "b": ["x", "y"]})
    return Table(df=df, source={"path": "t.xlsx", "sheet": "S"},
                 excel_rows=[2, 3], lineage=[])


def test_collect_returns_polars():
    assert isinstance(_table().collect(), pl.DataFrame)


def test_apply_is_immutable_and_logged():
    t = _table()
    t2 = t.apply(lambda df: df.head(1))
    assert t.df.height == 2 and t2.df.height == 1
    assert t2.lineage[-1]["verb"] == "apply"
    assert t2.excel_rows is None          # row identity lost


def test_apply_same_height_keeps_excel_rows():
    t2 = _table().apply(lambda df: df.with_columns(pl.col("a").alias("c")))
    assert t2.excel_rows == [2, 3]


def test_cast():
    t2 = _table().cast(a="int_")
    assert t2.df["a"].dtype == pl.Int64
    assert t2.lineage[-1]["verb"] == "cast"


def test_cast_unknown_column_raises():
    with pytest.raises(ValueError, match="nope"):
        _table().cast(nope="int_")
```

Run → FAIL.

**Step 2: Implement** `xl2y/table.py`:

```python
from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any, Callable

import polars as pl

_CAST_TARGETS = {
    "str_": pl.Utf8, "int_": pl.Int64, "float_": pl.Float64,
    "bool_": pl.Boolean, "date_": pl.Date, "datetime_": pl.Datetime("us"),
}


@dataclass(frozen=True)
class Table:
    """Immutable pipeline state. Every verb returns a new Table with one
    more lineage entry."""
    df: pl.DataFrame
    source: dict
    excel_rows: list[int] | None = None   # None once row identity is lost
    lineage: list[dict] = field(default_factory=list)
    comments: list[dict] = field(default_factory=list)
    rejects: pl.DataFrame | None = None
    errors: list = field(default_factory=list)
    _raw: Any = None                      # RawSheet, kept for kitchen_sink

    def _step(self, verb: str, df: pl.DataFrame,
              excel_rows: list[int] | None, **detail) -> "Table":
        entry = {"verb": verb, **detail}
        return replace(self, df=df, excel_rows=excel_rows,
                       lineage=[*self.lineage, entry])

    def collect(self) -> pl.DataFrame:
        return self.df

    def apply(self, fn: Callable[[pl.DataFrame], pl.DataFrame]) -> "Table":
        out = fn(self.df)
        rows = self.excel_rows if (self.excel_rows is not None
                                   and out.height == self.df.height) else None
        name = getattr(fn, "__name__", "<callable>")
        return self._step("apply", out, rows, fn=name)

    def cast(self, **types: str) -> "Table":
        unknown = set(types) - set(self.df.columns)
        if unknown:
            raise ValueError(f"cast: unknown column(s) {sorted(unknown)}")
        exprs = [pl.col(c).cast(_CAST_TARGETS[t]) for c, t in types.items()]
        return self._step("cast", self.df.with_columns(exprs),
                          self.excel_rows, types=types)
```

(`cast` accepts either the string name `"int_"` or the schema type objects once Task 11 exists — revisit there.)

**Step 3:** Run → PASS. Commit: `feat: immutable Table with apply, cast, collect, lineage`.

---

## Task 7: load() and load_all() entry points

**Files:**
- Create: `tests/test_load.py`
- Modify: `xl2y/__init__.py`, `xl2y/table.py` (add `TableSet`)

**Step 1: Failing tests** — `tests/test_load.py`:

```python
import pytest

import xl2y
from tests import fixtures


def test_load_single_sheet(tmp_path):
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"))
    assert t.df.height == 3
    assert t.source["sheet"] == "Data"


def test_load_picks_best_sheet_and_logs(tmp_path, caplog):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx"))
    assert t.source["sheet"] == "Report"       # Empty sheet loses


def test_load_explicit_sheet_missing(tmp_path):
    with pytest.raises(xl2y.SheetNotFoundError):
        xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"), sheet="X")


def test_load_hints_forwarded(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx"),
                  sheet="Report", sparse_rows="section")
    assert "section" in t.df.columns


def test_load_all(tmp_path):
    ts = xl2y.load_all(fixtures.nasty_book(tmp_path / "n.xlsx"))
    assert list(ts) == ["Report"]              # empty sheet skipped, logged
    assert ts["Report"].df.height == 3
```

Run → FAIL.

**Step 2: Implement.** In `xl2y/__init__.py`:

- `load(path, sheet=None, **hints) -> Table`: if `sheet` given, `read_sheet` + `extract_table`; else iterate `read_all`, extract each non-empty sheet, and pick the winner by `df.height * df.width` (log the choice at INFO when >1 candidate). Wrap into `Table` with `source={"path": str(path), "sheet": ...}`, `comments=ex.comments`, `excel_rows=ex.excel_rows`, `lineage=[{"verb": "load", "events": ex.events}]`, `_raw=raw`.
- `load_all(path, **hints) -> TableSet`: extract every non-empty worksheet (skip `EmptySheetError` with INFO log), key by snake_case sheet name — port `_snake_case` (prototype lines 381–387) into `xl2y/util.py` along with the key-dedup loop (lines 251–255).
- `TableSet` in `table.py`: wraps `dict[str, Table]`; `__iter__`, `__getitem__`, `__len__`; verb methods `clean/kitchen_sink/cast/apply/conform/dry_run` map over members returning a new `TableSet` (implement generically: `def _map(self, verb, *a, **k)`); `to_parquet(dir)` arrives in Task 10.
- Export `load`, `load_all`, `Table`, `TableSet` in `__all__`.

**Step 3:** Run → PASS. Commit: `feat: load and load_all entry points`.

---

## Task 8: coerce — polars-native value coercion + clean()

**Files:**
- Create: `xl2y/coerce.py`, `tests/test_coerce.py`
- Modify: `xl2y/table.py` (add `clean()`)

**Port sources:** constants `NA_TOKENS`, `COERCE_MIN_FRACTION`, `_CURRENCY_RE` (lines 85–108); logic of `_parse_numberish` (877–904) and `_coerce_columns` (907–984) — reimplemented as polars expressions, NOT `map_elements`.

**Step 1: Failing tests** — `tests/test_coerce.py`:

```python
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
    vals = ["1", "2", "x", "y", "z"]        # 40% numeric < 80%
    out, _ = coerce.coerce_columns(_col(vals), dayfirst=True)
    assert out["x"].dtype == pl.Utf8


def test_dayfirst_dates():
    out, _ = coerce.coerce_columns(_col(["03/04/2025", "04/04/2025"]),
                                   dayfirst=True)
    assert out["x"].dtype == pl.Date
    assert out["x"][0].month == 4


def test_bare_numbers_are_not_dates():
    out, _ = coerce.coerce_columns(_col(["2021", "2022", "2023x"]),
                                   dayfirst=True)
    assert out["x"].dtype == pl.Utf8


def test_events_reported():
    _, events = coerce.coerce_columns(_col(["1", "2", "3", "4", "oops"]),
                                      dayfirst=True)
    assert events[0]["event"] == "coerced_numeric"
    assert events[0]["failed"] == 1
```

Run → FAIL.

**Step 2: Implement** `xl2y/coerce.py`:

```python
"""Vectorized coercion of text columns: NA tokens, tolerant numerics
('$1,234', '(500)', '15%'), and dates. Mirrors reference/excel_loader.py
_coerce_columns but as polars expressions."""
from __future__ import annotations

import logging

import polars as pl

logger = logging.getLogger(__name__)

COERCE_MIN_FRACTION = 0.8
NA_TOKENS = {"n/a", "na", "#n/a", "#value!", "#ref!", "#div/0!", "-", "--",
             "–", "—", "none", "null", "nil"}
_CURRENCY = r"(?i)^(?:a\$|au\$|aud|nz\$|us\$|usd|[$£€¥])\s*"

_DATE_FORMATS_DAYFIRST = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y",
                          "%Y-%m-%d", "%d %b %Y", "%d %B %Y"]
_DATE_FORMATS_MONTHFIRST = ["%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y",
                            "%Y-%m-%d", "%b %d %Y", "%B %d %Y"]
_DATEY = r"(?i)[-/:.]|\d\s+\w|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"


def numberish(col: str) -> pl.Expr:
    """Parse '1,234', '$1,000.50', '(500)', '15%' -> Float64 (null on fail)."""
    s = (pl.col(col).str.strip_chars()
         .str.replace_all(" ", ""))
    paren = s.str.contains(r"^\(.*\)$")
    core = pl.when(paren).then(s.str.strip_chars("()")).otherwise(s).str.strip_chars()
    minus = core.str.starts_with("-")
    core = pl.when(minus).then(core.str.strip_prefix("-")).otherwise(core).str.strip_chars()
    core = core.str.replace(_CURRENCY, "")
    pct = core.str.ends_with("%")
    core = pl.when(pct).then(core.str.strip_suffix("%")).otherwise(core)
    core = core.str.replace_all(",", "").str.replace_all(" ", "")
    n = pl.when(core.str.len_chars() > 0).then(core).otherwise(None) \
          .cast(pl.Float64, strict=False)
    n = pl.when(pct).then(n / 100).otherwise(n)
    return pl.when(paren | minus).then(-n).otherwise(n)


def datish(col: str, dayfirst: bool) -> pl.Expr:
    fmts = _DATE_FORMATS_DAYFIRST if dayfirst else _DATE_FORMATS_MONTHFIRST
    tries = [pl.col(col).str.strip_chars().str.to_date(f, strict=False)
             for f in fmts]
    return pl.coalesce(tries)


def coerce_columns(df: pl.DataFrame, dayfirst: bool,
                   sheet: str = "?") -> tuple[pl.DataFrame, list[dict]]:
    """Returns (coerced df, lineage events). Per string column:
    NA tokens -> null; then numeric if >=80% parse; then date if >=80%
    parse AND >=50% look date-like; else stays text."""
    events: list[dict] = []
    for col in df.columns:
        if df[col].dtype != pl.Utf8:
            continue
        base = df.with_columns(
            pl.when(pl.col(col).str.strip_chars().str.to_lowercase()
                    .is_in(list(NA_TOKENS)))
            .then(None).otherwise(pl.col(col)).alias(col))
        nn = base[col].drop_nulls()
        if nn.is_empty():
            df = base
            continue

        num = base.select(numberish(col).alias("v"))["v"]
        ok = num.drop_nulls().len() / nn.len() if nn.len() else 0
        if ok >= COERCE_MIN_FRACTION:
            failed = nn.len() - num.drop_nulls().len()
            out = base.with_columns(numberish(col).alias(col))
            v = out[col].drop_nulls()
            if v.len() and (v % 1 == 0).all():
                out = out.with_columns(pl.col(col).cast(pl.Int64))
            events.append({"event": "coerced_numeric", "column": col,
                           "failed": int(failed)})
            if failed:
                logger.warning("Sheet %r: column %r coerced to numeric; "
                               "%d value(s) became null.", sheet, col, failed)
            df = out
            continue

        dt = base.select(datish(col, dayfirst).alias("v"))["v"]
        datey = base.select(pl.col(col).str.contains(_DATEY).alias("d"))["d"]
        if (nn.len() and dt.drop_nulls().len() / nn.len() >= COERCE_MIN_FRACTION
                and (datey.drop_nulls().sum() / max(datey.drop_nulls().len(), 1)) >= 0.5):
            failed = nn.len() - dt.drop_nulls().len()
            events.append({"event": "coerced_date", "column": col,
                           "failed": int(failed), "dayfirst": dayfirst})
            df = base.with_columns(datish(col, dayfirst).alias(col))
            continue

        df = base   # text, NA tokens applied
    return df, events
```

(Exact polars method names to verify while implementing: `str.strip_prefix` / `str.strip_suffix` exist in polars ≥0.20; if the installed version lacks them, use `str.replace(r"^-", "")` / `str.replace(r"%$", "")`.)

**Step 3:** Run → PASS. Commit: `feat: polars-native value coercion`.

**Step 4: clean() verb — failing test** (append to `tests/test_table.py`):

```python
def test_clean_snake_cases_and_coerces(tmp_path):
    import xl2y
    from tests import fixtures
    t = xl2y.load(fixtures.typed_book(tmp_path / "t.xlsx")).clean()
    assert t.df.columns == ["money", "pct", "when", "count", "note"]
    assert t.df["money"].to_list() == [1234.5, -500.0, None]
    assert t.df["pct"].to_list() == [0.15, 0.075, None]
    assert str(t.df["when"].dtype) == "Date"
    assert t.lineage[-1]["verb"] == "clean"


def test_clean_drops_fully_empty_rows_and_cols(tmp_path):
    # build inline: a book with an all-empty column captured in the run
    ...  # covered implicitly by extraction; assert no all-null columns remain


def test_clean_snake_case_collision_dedupes(tmp_path):
    import polars as pl
    from xl2y.table import Table
    df = pl.DataFrame({"Total Rev": [1], "total_rev": [2]})
    t = Table(df=df, source={}).clean()
    assert t.df.columns == ["total_rev", "total_rev_2"]
```

**Step 5: Implement `clean()`** on `Table`:
1. Rename columns via `snake_case` from `xl2y/util.py`, deduped with `unique_name` (record `renames={old: new}` in lineage only for changed names).
2. Strip whitespace on all Utf8 columns (`.str.strip_chars()`), empty string → null.
3. Drop fully-null rows (filter; keep `excel_rows` in sync via a boolean mask computed BEFORE dropping) and fully-null columns (record which).
4. `coerce.coerce_columns(df, dayfirst=self.source.get("dayfirst", True))`; merge its events into the lineage entry.

**Step 6:** Run full suite `uv run pytest -v` → PASS. Commit: `feat: clean() verb`.

---

## Task 9: dry_run()

**Files:**
- Modify: `xl2y/table.py`
- Test: `tests/test_table.py` (append)

**Step 1: Failing test:**

```python
def test_dry_run_prints_and_returns_self(tmp_path, capsys):
    import xl2y
    from tests import fixtures
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx")).clean()
    t2 = t.dry_run()
    out = capsys.readouterr().out
    assert t2 is t
    assert "3 rows" in out and "store" in out
    assert "Northern Region" in out       # stripped rows are surfaced
```

**Step 2: Implement.** `dry_run(n=8)` prints: `path [sheet] — {rows} rows × {cols} cols`, dtypes line, `df.head(n)` (polars' repr is already good), then a "what happened" section walking `lineage`: one line per event (comments stripped with text, sparse rows, coercions with column + failure count, renames count). Returns `self` unchanged — the ONLY verb that prints and the only one returning the same object.

**Step 3:** PASS → commit: `feat: dry_run() pipeline peek`.

---

## Task 10: to_parquet with embedded lineage

**Files:**
- Modify: `xl2y/table.py` (Table.to_parquet, TableSet.to_parquet)
- Test: `tests/test_parquet.py` (create)

**Step 1: Failing tests:**

```python
import json

import polars as pl

import xl2y
from tests import fixtures


def test_to_parquet_roundtrip_with_lineage(tmp_path):
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx")).clean()
    out = tmp_path / "out.parquet"
    t.to_parquet(out)
    assert pl.read_parquet(out).height == 3
    meta = pl.read_parquet_metadata(out)
    payload = json.loads(meta["xl2y"])
    assert payload["source"]["sheet"] == "Data"
    assert payload["lineage"][0]["verb"] == "load"


def test_to_parquet_refuses_overwrite(tmp_path):
    import pytest
    t = xl2y.load(fixtures.simple_book(tmp_path / "s.xlsx"))
    out = tmp_path / "o.parquet"
    t.to_parquet(out)
    with pytest.raises(FileExistsError):
        t.to_parquet(out)
    t.to_parquet(out, overwrite=True)     # fine


def test_tableset_fans_out(tmp_path):
    ts = xl2y.load_all(fixtures.nasty_book(tmp_path / "n.xlsx"))
    manifest = ts.clean().to_parquet(tmp_path / "out")
    assert (tmp_path / "out" / "report.parquet").exists()
    assert manifest["report"]["rows"] == 3
```

**Step 2: Implement.**
- `Table.to_parquet(path, overwrite=False)`: `FileExistsError` unless overwrite (message per prototype line 320); payload = `{"source", "lineage", "comments"}` JSON with `default=str`; `self.df.write_parquet(path, metadata={"xl2y": json.dumps(...)})`. Returns the path.
- `TableSet.to_parquet(dir, overwrite=False)`: mkdir parents, one file per key, return manifest dict per design (`parquet_path`, `sheet_name`, `rows`, `columns`, `comments`) — port shape from prototype lines 350–356.

**Step 3:** PASS → commit: `feat: parquet export with embedded lineage metadata`.

---

## Task 11: patterns module

**Files:**
- Create: `xl2y/patterns.py`, `tests/test_patterns.py`

**Step 1: Failing tests** (representative; write all):

```python
import re

from xl2y import patterns


def _full(p, s):
    return re.fullmatch(p, s) is not None


def test_email():
    assert _full(patterns.EMAIL, "a.b+c@example.com.au")
    assert not _full(patterns.EMAIL, "not an email")


def test_phone_au():
    for ok in ["0412 345 678", "0412345678", "+61 412 345 678", "(02) 9123 4567"]:
        assert _full(patterns.PHONE_AU, ok), ok
    assert not _full(patterns.PHONE_AU, "12345")


def test_abn_checksum():
    assert patterns.abn_valid("51 824 753 556")      # ATO's example ABN
    assert not patterns.abn_valid("51 824 753 557")


def test_helpers():
    assert _full(patterns.digits(4), "1234") and not _full(patterns.digits(4), "12345")
    assert _full(patterns.one_of("Y", "N"), "Y")
    assert _full(patterns.any_of(patterns.EMAIL, patterns.UUID),
                 "a@b.co")
```

**Step 2: Implement** `xl2y/patterns.py` — all regex strings (no compiled objects, so they drop straight into polars `str.contains`):

- `EMAIL = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"`
- `URL`, `UUID`, `IPV4`, `PHONE_E164 = r"\+[1-9]\d{7,14}"`, `PHONE_AU` (mobile/landline, optional +61, tolerant of spaces/parens), `POSTCODE_AU = r"\d{4}"`, `BSB = r"\d{3}-?\d{3}"`, `DATE_ISO = r"\d{4}-\d{2}-\d{2}"`, `CURRENCY = r"-?\(?\$?\s?[\d,]+(?:\.\d+)?\)?"`
- Helpers returning regex strings: `any_of(*pats)` → `"(?:p1|p2)"`, `exact(literal)` → `re.escape`, `digits(n)` → `rf"\d{{{n}}}"`, `one_of(*values)` → alternation of escaped literals.
- Checksums as plain predicates: `abn_valid(s)`, `acn_valid(s)`, `tfn_valid(s)` (strip spaces, standard weighted-digit algorithms — ABN: subtract 1 from first digit, weights [10,1,3,5,7,9,11,13,15,17,19], mod 89; ACN: weights [8,7,6,5,4,3,2,1] complement mod 10; TFN: weights [1,4,3,7,5,8,6,9,10] mod 11).

**Step 3:** PASS → commit: `feat: patterns module with AU-flavoured validators`.

---

## Task 12: Schema, type constructors, conform (raise mode)

**Files:**
- Create: `xl2y/schema.py`, `tests/test_schema.py`
- Modify: `xl2y/table.py` (`conform` verb), `xl2y/__init__.py` (export `Schema`, constructors, `patterns`)

**Step 1: Failing tests:**

```python
import polars as pl
import pytest

import xl2y
from xl2y import patterns
from xl2y.errors import SchemaError
from xl2y.table import Table


def _table():
    df = pl.DataFrame({
        "store": ["Syd", "Mel", None],
        "revenue": ["$1,000", "(50)", "abc"],
        "email": ["a@b.co", "bad", "c@d.co"],
    })
    return Table(df=df, source={"path": "t.xlsx", "sheet": "S"},
                 excel_rows=[2, 3, 4])


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
    # cast failure cites the original Excel row
    cast_p = [p for p in ei.value.problems if p.rule == "cast"][0]
    assert cast_p.rows == [4]


def test_conform_happy_path_casts():
    df = pl.DataFrame({"store": ["Syd"], "revenue": ["$1,000"],
                       "email": ["a@b.co"]})
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
        n=xl2y.int_(min=0, max=10,
                    check=pl.col("n") != 7),
    )
    df = pl.DataFrame({"state": ["NSW", "QLD"], "n": [7, 11]})
    with pytest.raises(SchemaError) as ei:
        Table(df=df, source={}).conform(schema)
    rules = {p.rule for p in ei.value.problems}
    assert {"allowed_values", "max", "check"} <= rules
```

**Step 2: Implement** `xl2y/schema.py`:

```python
@dataclass(frozen=True)
class ColumnType:
    kind: str                       # "str" | "int" | "float" | "bool" | "date" | "datetime" | "cat"
    nullable: bool = True
    min: Any = None
    max: Any = None
    pattern: str | None = None
    values: tuple = ()              # cat_ only
    check: Any = None               # pl.Expr or Callable[[pl.DataFrame], pl.Series]


def str_(nullable=True, pattern=None, check=None) -> ColumnType: ...
def int_(nullable=True, min=None, max=None, check=None) -> ColumnType: ...
def float_(nullable=True, min=None, max=None, check=None) -> ColumnType: ...
def bool_(nullable=True, check=None) -> ColumnType: ...
def date_(nullable=True, min=None, max=None, check=None) -> ColumnType: ...
def datetime_(nullable=True, min=None, max=None, check=None) -> ColumnType: ...
def cat_(*values, nullable=True, check=None) -> ColumnType: ...


@dataclass(frozen=True)
class Problem:
    column: str
    rule: str            # "missing" | "cast" | "not_null" | "min" | "max"
                         # | "pattern" | "allowed_values" | "check" | "extra"
    count: int
    rows: list[int]      # original Excel rows if known, else df positions
    message: str


class Schema:
    def __init__(self, *, extra_columns: str = "keep", **columns: ColumnType): ...
    def conform(self, df, excel_rows) -> tuple[pl.DataFrame, list[Problem]]: ...
```

`Schema.conform` phases:
1. **Presence**: missing declared columns → `Problem("missing")`; extras per `extra_columns` mode.
2. **Cast**: per column by `kind` — numeric kinds through `coerce.numberish`, date kinds through `coerce.datish` (with `str.to_datetime` fallbacks for `datetime`), `bool` maps `{"true","yes","y","1"} / {"false","no","n","0"}` case-insensitively, `str`/`cat` cast to Utf8. Values non-null before and null after = cast failures → `Problem("cast", rows=...)`. Row numbers: take `excel_rows[i]` where available.
3. **Validate** on the cast frame: `not_null`, `min`, `max`, `pattern` (wrap as `^(?:...)$` full match), `allowed_values` for `cat_`, `check` (expr → `df.select(expr)`; callable → `fn(df)` must return boolean Series; False rows are failures).

Cap `rows` samples at 10 per problem, but `count` is the true total.

`Table.conform(schema, on_error="raise")` (Task 13 adds the other modes): call `schema.conform`, build lineage entry; if problems and `on_error == "raise"`, format per design:

```
3 problems in 'report.xlsx' [Q3]:
  revenue: 4 values < 0 (rows 12, 40, 88, 101)
  ...
```

and `raise SchemaError(msg, problems)`.

Update `Table.cast` (Task 6) to also accept `ColumnType` values, mapping `kind` → polars dtype.

**Step 3:** PASS → commit: `feat: Schema, type constructors, conform with full-report raise`.

---

## Task 13: conform lenient modes

**Files:**
- Modify: `xl2y/table.py`, `tests/test_schema.py` (append)

**Step 1: Failing tests:**

```python
def test_quarantine_splits_rows():
    t = _table().conform(SCHEMA, on_error="quarantine")
    assert t.df.height == 1                 # only fully-clean row survives
    assert t.rejects.height == 2
    assert t.errors                          # problems still reported


def test_report_keeps_everything():
    t = _table().conform(SCHEMA, on_error="report")
    assert t.df.height == 3
    assert len(t.errors) == 4
```

**Step 2: Implement.** In `conform`: compute a per-row bad mask (union of all problem rows). `"quarantine"` → `rejects` = bad rows (pre-cast originals where cast failed — keep it simple: rows from the cast frame), df = clean rows, `excel_rows` filtered, WARNING logged with counts. `"report"` → df unchanged (cast still applied where it succeeded), `errors` = problems. Both record the mode + problem count in lineage.

**Step 3:** PASS → commit: `feat: conform quarantine and report modes`.

---

## Task 14: kitchen_sink candidate tournament

**Files:**
- Create: `xl2y/sink.py`, `tests/test_sink.py`
- Modify: `xl2y/table.py` (verb), `xl2y/extract.py` (nothing — reuse), `xl2y/__init__.py`

**Step 1: Failing tests:**

```python
import xl2y
from tests import fixtures


def test_kitchen_sink_on_nasty_book(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx")).kitchen_sink()
    # winner should use section mode (denser, no col_ fallbacks, typed cols)
    assert "section" in t.df.columns
    assert t.df.height == 3
    assert t.df["revenue"].dtype.is_numeric()


def test_kitchen_sink_on_multi_header(tmp_path):
    t = xl2y.load(fixtures.multi_header_book(tmp_path / "m.xlsx")).kitchen_sink()
    assert t.df.columns == ["store", "revenue_q1", "revenue_q2"]


def test_kitchen_sink_records_candidates(tmp_path):
    t = xl2y.load(fixtures.nasty_book(tmp_path / "n.xlsx")).kitchen_sink()
    entry = t.lineage[-1]
    assert entry["verb"] == "kitchen_sink"
    assert len(entry["candidates"]) >= 4
    assert entry["winner"]["score"] == max(c["score"] for c in entry["candidates"])


def test_kitchen_sink_without_raw_falls_back_to_clean():
    import polars as pl
    from xl2y.table import Table
    t = Table(df=pl.DataFrame({"A": ["1", "2"]}), source={})
    assert t.kitchen_sink().df["a"].dtype.is_numeric()   # degraded = clean()
```

**Step 2: Implement** `xl2y/sink.py`:

```python
CANDIDATE_GRID = [
    {"header_rows": h, "sparse_rows": s}
    for h in (1, 2, 3)
    for s in ("comment", "section")
]


def score(df: pl.DataFrame) -> float:
    """'Competent table' score: typed columns good, nulls bad,
    fallback/duplicate names bad, degenerate shapes bad."""
    if df.height == 0 or df.width == 0:
        return float("-inf")
    typed = sum(1 for d in df.dtypes if d != pl.Utf8) / df.width
    cells = df.height * df.width
    density = 1 - (sum(df.null_count().row(0)) / cells)
    fallback = sum(1 for c in df.columns if c.startswith("col_")) / df.width
    dupey = sum(1 for c in df.columns if c.rsplit("_", 1)[-1].isdigit()) / df.width
    shape_penalty = 0.5 if df.height < 2 else 0.0
    return 2.0 * typed + density - 1.5 * fallback - 0.5 * dupey - shape_penalty
```

`run_tournament(raw: RawSheet, source) -> Table`: for each candidate, `extract_table(raw, **opts)` (catch `ValueError`/`EmptySheetError` → skip), wrap in a Table, `.clean()`, score. Keep the winner's Table; append a `kitchen_sink` lineage entry listing every candidate `{"opts", "score", "shape"}` and the winner. `Table.kitchen_sink()`: if `self._raw` is None (constructed manually / post-apply), log a warning and return `self.clean()`; else delegate to `run_tournament`.

**Step 3:** PASS → commit: `feat: kitchen_sink candidate tournament`.

---

## Task 15: kitchen_sink unpivot candidate (stretch)

**Files:**
- Modify: `xl2y/sink.py`, `tests/test_sink.py`

**Step 1: Failing test:** fixture inline — header `["Store", 2021, 2022, 2023]`, two data rows of ints. `kitchen_sink()` should produce columns `["store", "period", "value"]` with height 6.

**Step 2: Implement.** After the base tournament, take the winner; if ≥3 trailing columns (a) share a numeric dtype and (b) have names matching years (`^(19|20)\d{2}$`), month names, or ISO dates — generate one more candidate via `df.unpivot(index=<other cols>, variable_name="period", value_name="value")`, `.clean()` it, score with a +0.25 tidy bonus, and let it compete. Record in candidates list as `{"opts": {"unpivot": True}, ...}`.

**Step 3:** PASS → commit: `feat: unpivot candidate in kitchen_sink`.

---

## Task 16: docs, project CLAUDE.md, cleanup

**Files:**
- Modify: `README.md`
- Create: `CLAUDE.md`, `.claude/` docs if useful
- Verify: full suite, `__all__` complete

**Step 1:** `README.md`: one-paragraph pitch, install, the two-ends-of-spectrum example from the design doc, verb table, schema example with patterns, link to design doc.

**Step 2:** `CLAUDE.md` (project root — written for future agent sessions):

```markdown
# xl2y

Excel → cleaning pipeline → Parquet. Design: docs/plans/2026-07-22-xl2y-design.md

## Commands
- Test: `uv run pytest tests/ -v`
- Add dep: `uv add <pkg>` (runtime) / `uv add --dev <pkg>`

## Architecture (read in this order)
- reader.py: two-pass streaming Excel read → RawSheet. Never trust declared
  dimensions; extent comes from actual values.
- extract.py: heuristics turning a RawSheet grid into a polars df +
  comments + excel_rows. Ported from reference/excel_loader.py — that file
  is the frozen prototype, do not edit it.
- table.py: immutable Table; every verb returns a new Table + lineage entry.
- coerce.py / schema.py / patterns.py / sink.py: see design doc.

## Conventions
- Polars only — no pandas anywhere.
- Type constructors ALL have trailing underscores (str_, date_, ...), even
  where not shadowing builtins. Consistency is deliberate.
- excel_rows tracks original Excel row numbers through the pipeline; any
  verb that can't maintain it must set it to None, never guess.
- Errors: everything under Xl2yError; SchemaError carries .problems.
- Library never prints except dry_run(). Warnings via logging.
- TDD: failing test first. Fixtures are generated in tests/fixtures.py,
  never binary files in git.
- Commits: conventional style, no AI/tool references.
```

**Step 3:** Full suite: `uv run pytest -v` → all PASS. Also run the two design-doc examples end-to-end in a scratch script against `fixtures.nasty_book` output.

**Step 4:** Commit: `docs: README, project CLAUDE.md`.

---

## Task order & dependencies

```
1 setup → 2 fixtures → 3 reader-meta → 4 reader-passes → 5 extract
  → 6 table → 7 load → 8 coerce+clean → 9 dry_run → 10 parquet
  → 11 patterns → 12 schema → 13 lenient modes → 14 sink → 15 unpivot → 16 docs
```

11 (patterns) can run any time after 1. Everything else is sequential.
