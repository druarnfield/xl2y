# xl2y — working notes

Excel → cleaning pipeline → Parquet. One purpose, tiny verb set.
Design: `docs/plans/2026-07-22-xl2y-design.md`.
Implementation plan: `docs/plans/2026-07-22-xl2y-implementation.md`.

## Commands
- Test: `uv run pytest -q`  (single: `uv run pytest tests/test_x.py::test_y -v`)
- Add dep: `uv add <pkg>` (runtime) / `uv add --dev <pkg>`
- Deps are intentionally just `polars` + `openpyxl`. No pandas, no pyarrow.

## Architecture (read in this order)
- `reader.py` — two-pass streaming Excel read → `RawSheet`. NEVER trust the
  declared dimension; the real extent comes from actual values. Pass 1 is a
  streaming XML parse (merges, hidden rows/cols, formula cells, extent); pass
  2 streams values via openpyxl read_only.
- `extract.py` — heuristics that turn a `RawSheet` grid into a polars df +
  comments + `excel_rows`. Ported from `reference/excel_loader.py` (the frozen
  pandas prototype — do NOT edit it; it exists only as the porting reference).
  `extract_table` copies the grid, so it never mutates its `RawSheet`
  (kitchen_sink re-extracts the same one).
- `coerce.py` — vectorised text→number/date coercion as polars expressions.
- `table.py` — the immutable `Table` and `TableSet`; every verb returns a NEW
  Table + one lineage entry.
- `schema.py` — `Schema`, type constructors, two-phase cast→validate `conform`.
- `patterns.py` — regex strings + AU checksum validators + helpers.
- `sink.py` — `kitchen_sink` candidate tournament + scoring + unpivot.

## Conventions & invariants
- Polars only. `pl.Utf8` is the string dtype.
- Type constructors ALL end in `_` (`str_`, `date_`, …), even where not
  shadowing a builtin. Deliberate consistency.
- `excel_rows` tracks original 1-indexed Excel rows through the pipeline. Any
  verb that can't preserve the mapping (row count changes) sets it to `None` —
  never guess. Problems/comments cite these for the human fixing the file.
- Errors: everything under `Xl2yError`; `SchemaError` carries `.problems`.
- The library NEVER prints except `dry_run()`. Everything else is `logging`.
- Lineage is the source of truth for "what happened"; it's embedded in Parquet
  metadata (key `xl2y`) on write, and drives `dry_run()`.
- `TableSet` is keyed by ORIGINAL sheet name; `to_parquet` snake_cases only the
  output filenames / manifest keys.
- `kitchen_sink` only tries information-PRESERVING sparse modes
  (`section`, `keep`) — never `comment` — so it can't silently drop rows.

## Testing
- Fixtures are generated in code (`tests/fixtures.py`) via openpyxl — never
  binary files in git. Add pathological cases there.
- TDD: failing test first. Commit per task. No AI/tool references in commits.

## Known limitations / future
- Date coercion uses a fixed strptime list (not pandas' flexible "mixed"); some
  uncommon formats (month-year-only, `T`-separated timestamps) stay text.
- No CLI yet (`python -m xl2y`) — would be a thin wrapper over
  `load_all().kitchen_sink().to_parquet()`.
