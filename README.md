# xl2y

**Messy Excel to Parquet. Sometimes it makes sense to throw the kitchen sink.**

Spreadsheets are where clean data goes to get merged cells, a title banner,
three header rows, a `Notes:` column nobody filled in, and — because someone
once bolded cell A1048576 in 2014 — a "used range" of a million rows.

`xl2y` takes that, runs a cleaning pipeline you can actually re-run tomorrow,
and hands you a Parquet file. One job. A handful of verbs. Simple when the file
is simple, flexible when it's a crime scene.

```bash
uv add git+https://github.com/druarnfield/xl2y   # only needs polars + openpyxl
```

## Pick your energy

```python
import xl2y

# Monday, 9am, someone emailed you "the numbers" and you have questions.
xl2y.load("mystery.xlsx").kitchen_sink().to_parquet("out.parquet")

# You know this file. You've been burned by this file before.
schema = xl2y.Schema(
    store=xl2y.str_(nullable=False),
    date=xl2y.date_(),
    revenue=xl2y.float_(min=0),
)
(xl2y.load("report.xlsx", sheet="Q3")
    .clean()
    .conform(schema)      # cast, validate, and complain loudly if it's wrong
    .to_parquet("q3.parquet"))
```

`kitchen_sink()` tries a bunch of interpretations of a horrible sheet and keeps
the one that looks most like a competent table — the "get it into something and
stop thinking about it" button. `clean()` + `conform()` is for when you'd
rather the file prove itself first.

A pipeline is just Python, so "reproducible" means the boring, reliable thing:
you run the script again. Every verb returns a new immutable `Table`, so a
reusable pipeline is a plain function — `def pipe(t): return t.clean().conform(schema)`.

## The whole vocabulary

| Verb | What it does |
|---|---|
| `xl2y.load(path, sheet=None, **hints)` | Load one table. No `sheet`? It picks the sheet with the biggest table and tells you which. |
| `xl2y.load_all(path, **hints)` | Every sheet into a `TableSet`; the same chain runs on each. |
| `.clean()` | snake_case the headers, trim whitespace, drop empty rows/cols, and quietly turn `"$1,234"`, `"15%"`, `"N/A"` and `03/04/2025` into things a computer respects. |
| `.kitchen_sink()` | Throw everything at it: title banners, sparse section rows, multi-row headers, even unpivoting a wide year-per-column table. Keep the best-looking result. |
| `.cast(col="int_", ...)` | For when the coercion guessed wrong and you know better. |
| `.apply(fn)` | Escape hatch. Any `df -> df` callable is welcome here. |
| `.conform(schema, on_error="raise")` | Cast to the schema, then validate. Raises by default; `"quarantine"` or `"report"` if you're feeling forgiving. |
| `.dry_run()` | Prints what you've got and what the pipeline did, then hands the table straight back. Leave it mid-chain while you fiddle; delete it when you're happy. |
| `.to_parquet(path)` / `.collect()` | Write Parquet (with a full audit trail baked into the metadata), or just give me the polars DataFrame. |

## Schemas, for the trust-issues path

A small, no-nonsense schema: names, types, nullability, bounds, patterns,
allowed values, and a `check=` hatch for anything weird. `conform` collects
**every** problem and raises one error — and it points at the *original Excel
row numbers*, because the person fixing the spreadsheet lives in Excel, not in
your DataFrame.

```python
from xl2y import patterns

schema = xl2y.Schema(
    email=xl2y.str_(pattern=patterns.EMAIL),
    state=xl2y.cat_("NSW", "VIC", "QLD"),
    abn=xl2y.str_(check=lambda df: df["abn"].map_elements(
        patterns.abn_valid, return_dtype=bool)),
    revenue=xl2y.float_(nullable=False, min=0),
    extra_columns="drop",   # columns you didn't ask for: keep | drop | error
)

t = xl2y.load("f.xlsx").clean().conform(schema, on_error="quarantine")
t.rejects   # the rows that misbehaved, set aside
t.errors    # what was wrong, machine-readable
```

```
SchemaError: 3 problem(s) in 'report.xlsx' [Q3]:
  revenue: 4 value(s) < 0 (rows 12, 40, 88, 101)
  date: 2 null value(s) (rows 55, 56)
  store_id: 1 value(s) could not be cast to int_ (rows 7)
```

Type constructors all wear a trailing underscore — `str_`, `int_`, `float_`,
`bool_`, `date_`, `datetime_`, `cat_` — for the crime of consistency.

`patterns` ships the usual suspects (`EMAIL`, `URL`, `UUID`, `IPV4`,
`PHONE_AU`, `PHONE_E164`, `POSTCODE_AU`, `BSB`, `DATE_ISO`, `CURRENCY`),
proper checksum validators for `abn_valid` / `acn_valid` / `tfn_valid`, and
little builders (`any_of`, `exact`, `digits`, `one_of`) so nobody has to
handwrite an email regex ever again.

## Why it doesn't fall over on a big ugly file

`xl2y` refuses to believe a workbook's declared dimensions — one stray bit of
formatting and Excel will happily claim a million rows. A streaming XML pass
finds the *real* extent (plus the merges, hidden rows, and uncached formulas),
then values stream straight into polars columns. `load_all` handles one sheet
at a time, so peak memory is your biggest sheet, not the whole workbook. That
"million-row" file loads in a couple of seconds and a couple of megabytes.

Supported: `.xlsx` / `.xlsm` / `.xltx` / `.xltm`. Not supported: `.xls` /
`.xlsb`, which get a polite "please convert this first" instead of a stack
trace.

## For the curious

Design and implementation notes live in `docs/plans/`. The original
pandas prototype that seeded all the extraction heuristics is frozen at
`reference/excel_loader.py`. `CLAUDE.md` has the architecture tour.
