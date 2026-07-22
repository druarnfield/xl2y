# xl2y

Excel in, a customisable and reproducible cleaning pipeline, Parquet out.
One job, done well — with a tiny, intuitive verb set that stays simple in the
easy case and flexible when you need it.

```bash
uv add git+github.com/druarnfield/xl2y   # depends only on polars + openpyxl
```

## The two ends of the spectrum

```python
import xl2y

# "I was handed a piece of shit and just want it in Parquet"
xl2y.load("mystery.xlsx").kitchen_sink().to_parquet("out.parquet")

# "I know exactly what this file should be"
schema = xl2y.Schema(
    store=xl2y.str_(nullable=False),
    date=xl2y.date_(),
    revenue=xl2y.float_(min=0),
)
(xl2y.load("report.xlsx", sheet="Q3")
    .clean()
    .conform(schema)
    .to_parquet("q3.parquet"))
```

A pipeline is just Python, so it is reproducible by re-running the script.
Every verb returns a new immutable `Table`, so a reusable pipeline is a
function: `def pipe(t): return t.clean().conform(schema)`.

## The verbs

| Verb | What it does |
|---|---|
| `xl2y.load(path, sheet=None, **hints)` | Load one table. No `sheet` → pick the sheet with the biggest table. |
| `xl2y.load_all(path, **hints)` | Load every sheet into a `TableSet`; the same chain applies to each. |
| `.clean()` | Snake_case headers, strip whitespace, drop empty rows/cols, coerce `"$1,234"` / `"15%"` / NA tokens / dates. |
| `.kitchen_sink()` | Try many interpretations of a messy sheet and keep the most competent-looking table (incl. unpivoting wide year/month tables). |
| `.cast(col="int_", ...)` | Force explicit dtypes where coercion guessed wrong. |
| `.apply(fn)` | Escape hatch: any `df -> df` callable. |
| `.conform(schema, on_error="raise")` | Cast to the schema, then validate. |
| `.dry_run()` | Print a summary + what the pipeline did so far; returns the table unchanged. |
| `.to_parquet(path)` / `.collect()` | Write Parquet (lineage embedded in metadata), or hand back the polars DataFrame. |

## Schemas and validation

A small, in-house schema: names, types, nullability, bounds, patterns,
allowed values, and a per-column `check=`. By default `conform` raises one
`SchemaError` carrying **every** problem, each citing the original Excel row
numbers (so the person fixing the spreadsheet knows where to look):

```python
from xl2y import patterns

schema = xl2y.Schema(
    email=xl2y.str_(pattern=patterns.EMAIL),
    state=xl2y.cat_("NSW", "VIC", "QLD"),
    abn=xl2y.str_(check=lambda df: df["abn"].map_elements(
        patterns.abn_valid, return_dtype=bool)),
    revenue=xl2y.float_(nullable=False, min=0),
    extra_columns="drop",
)

t = xl2y.load("f.xlsx").clean().conform(schema, on_error="quarantine")
t.rejects   # rows that failed
t.errors    # the problems, machine-readable
```

Type constructors (all with a trailing underscore, for consistency):
`str_`, `int_`, `float_`, `bool_`, `date_`, `datetime_`, `cat_`.

`patterns` ships `EMAIL`, `URL`, `UUID`, `IPV4`, `PHONE_AU`, `PHONE_E164`,
`POSTCODE_AU`, `BSB`, `DATE_ISO`, `CURRENCY`, checksum validators
`abn_valid` / `acn_valid` / `tfn_valid`, and composable helpers
`any_of` / `exact` / `digits` / `one_of`.

## Why it scales

The reader never trusts a workbook's declared dimensions (stray formatting
routinely inflates them to a million rows). A streaming XML pass finds the
true data extent and the merges/hidden-rows/formulas; values are then streamed
straight into polars columns. `load_all` processes one sheet at a time, so
peak memory is the largest sheet, not the whole workbook.

Only `.xlsx` / `.xlsm` / `.xltx` / `.xltm` are supported; `.xls` / `.xlsb`
raise a clear "convert first" error.

## Design & plans

See `docs/plans/2026-07-22-xl2y-design.md` (design) and
`docs/plans/2026-07-22-xl2y-implementation.md` (implementation plan).
