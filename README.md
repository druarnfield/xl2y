# xl2y

**Messy Excel to Parquet. Sometimes you have to throw the kitchen sink at it.**

Spreadsheets are where clean data goes to acquire merged cells, a title banner,
three header rows, a `Notes:` column nobody ever filled in, and — because
someone bolded cell A1048576 back in 2014 — a "used range" of a million rows.

`xl2y` takes that file, runs a cleaning pipeline you can actually re-run
tomorrow, and hands you Parquet. One job, a handful of verbs. Simple when the
file is simple; flexible when it's a crime scene.

```bash
uv add git+https://github.com/druarnfield/xl2y   # just polars + openpyxl
```

## Pick your energy

```python
import xl2y

# Monday, 9am. Someone has emailed you "the numbers". You have questions.
xl2y.load("mystery.xlsx").kitchen_sink().to_parquet("out.parquet")

# You know this file. This file has hurt you before.
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

`kitchen_sink()` tries every interpretation of a horrible sheet it can think
of and keeps the one that looks most like a competent table. It is the "get
this into *something* and stop thinking about it" button. `clean()` +
`conform()` is the other mood: the file proves itself before it gets anywhere
near your pipeline.

Pipelines are plain Python, so "reproducible" means the boring, reliable
thing: you run the script again. Every verb returns a new immutable `Table`,
which makes a reusable pipeline just a function —
`def pipe(t): return t.clean().conform(schema)`.

## The whole vocabulary

| Verb | What it does |
|---|---|
| `xl2y.load(path, sheet=None, **hints)` | Load one table. No `sheet`? It picks the sheet with the biggest table and tells you which one it chose. |
| `xl2y.load_all(path, **hints)` | Every sheet into a `TableSet`; one chain, run on each. |
| `.clean()` | snake_case the headers, trim the whitespace, drop the empty rows and columns, and quietly turn `"$1,234"`, `"15%"`, `"N/A"`, and `03/04/2025` into things a computer respects. |
| `.kitchen_sink()` | Everything at once: title banners, sparse section rows, multi-row headers, even unpivoting a wide year-per-column table. Best-looking result wins. |
| `.cast(col="int_", ...)` | For when the guessing guessed wrong and you know better. |
| `.apply(fn)` | The escape hatch. Any `df -> df` callable is welcome here. |
| `.conform(schema, on_error="raise")` | Cast to the schema, then validate. Raises by default; `"quarantine"` or `"report"` if you're feeling forgiving. |
| `.dry_run()` | Prints what you've got and what the pipeline did to it, then hands the table straight back. Leave it mid-chain while you fiddle; delete it when you're happy. |
| `.to_parquet(path)` / `.collect()` | Write Parquet (full audit trail baked into the metadata), or just take the polars DataFrame and go. |

## Schemas, for people with trust issues

A small, no-nonsense schema: names, types, nullability, bounds, patterns,
allowed values, and a `check=` hatch for anything weirder. `conform` collects
**every** problem before raising a single error — and the error points at the
*original Excel row numbers*, because the person who has to fix the
spreadsheet lives in Excel, not in your DataFrame.

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
`bool_`, `date_`, `datetime_`, `cat_`. `str` was taken; the rest wear one out
of solidarity.

`patterns` ships the usual suspects (`EMAIL`, `URL`, `UUID`, `IPV4`,
`PHONE_AU`, `PHONE_E164`, `POSTCODE_AU`, `BSB`, `DATE_ISO`, `CURRENCY`),
proper checksum validators for `abn_valid` / `acn_valid` / `tfn_valid`, and
small builders (`any_of`, `exact`, `digits`, `one_of`) so that nobody,
anywhere, ever handwrites an email regex again.

## Why it doesn't fall over on a big ugly file

`xl2y` does not believe a workbook's declared dimensions, because one stray
bit of formatting is all it takes for Excel to claim a million rows. A
streaming XML pass finds the *real* extent (plus the merges, hidden rows, and
uncached formulas), then values stream straight into polars columns.
`load_all` works one sheet at a time, so peak memory is your biggest sheet,
not the whole workbook. The "million-row" file loads in a couple of seconds
and a couple of megabytes.

Supported: `.xlsx` / `.xlsm` / `.xltx` / `.xltm`. Not supported: `.xls` /
`.xlsb`, which get a polite "please convert this first" instead of a stack
trace.

## For the curious

Design and implementation notes live in `docs/plans/`. The original pandas
prototype that seeded the extraction heuristics is frozen at
`reference/excel_loader.py`, and `CLAUDE.md` has the architecture tour.
