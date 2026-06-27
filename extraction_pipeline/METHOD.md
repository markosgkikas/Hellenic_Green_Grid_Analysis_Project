# Hellenic Green Grid — ADMIE Operation & Market Data Pipeline

A two-stage data pipeline for the Greek power system: it **bulk-downloads** raw
Operation & Market files from the Greek TSO (**ADMIE / IPTO**) public API, then
**parses and unifies** the thousands of individual daily Excel files into one
tidy, analysis-ready table per file category.

The dataset spans **2019 → today** and covers system load, RES injections,
interconnection flows, unit production, reservoir levels and day-ahead RES
forecasts — the building blocks for renewable-penetration and grid-operation
analysis.

---

## Pipeline overview

```
ADMIE public API
       │   stage 1 — admie_download_all.py
       ▼
data/raw/admie/<FileCategory>/   (thousands of daily .xls / .xlsx / .zip files)
       │   stage 2 — parse_admie.py
       ▼
data/processed/<FileCategory>.csv   (one tidy long table per category)
```

---

## Stage 1 — Downloading the raw files (`admie_download_all.py`)

ADMIE does **not** expose the files directly. It exposes a JSON *index* endpoint
that, given a date range and a file category, returns metadata records — one per
file whose coverage period overlaps the range. The actual file payload lives
behind a download URL inside each record. The downloader walks a fixed list of
file categories and, for each one, follows this logic:

1. **Query the index** — a single `GET getOperationMarketFilewRange` call with
   `dateStart`, `dateEnd`, and `FileCategory` for the whole date range, returning
   a JSON list of file-metadata records.
2. **Read the download URL** — taken from each record's `file_path` field. The
   API already returns only the latest revision per date, so each record maps to
   one file to fetch.
3. **Fetch each file** in a second request and save it into the category's folder.

Other behaviours:

- **Idempotent & resumable** — a file already present on disk is skipped, so an
  interrupted run can simply be re-run to fill in the gaps.
- **Polite to the server** — a `0.5 s` pause between downloads.
- **Fault-tolerant per category/file** — the index call and each download are
  wrapped so that one failure logs an error and the run continues with the rest.

Configuration lives in constants at the top of the script — `DATE_START`,
`DATE_END`, and the `FILE_CATEGORIES` list — and it is run with no arguments:

```bash
python admie_download_all.py
```

The categories collected:

```
SystemRealizationSCADA        RealTimeSCADARES           RealTimeSCADASystemLoad
RealTimeSCADAImportsExports   ReservoirFillingRate       UnitProduction
DayAheadRESForecast (legacy)  ISP1DayAheadRESForecast    ISP2DayAheadRESForecast
ISP3IntraDayRESForecast
```

### Resulting raw layout

```
data/raw/admie/
├── SystemRealizationSCADA/        *.xls
├── RealTimeSCADARES/              *.xls
├── RealTimeSCADASystemLoad/       *.xls
├── RealTimeSCADAImportsExports/   *.xls
├── ReservoirFillingRate/          *.xls
├── DayAheadRESForecast/           *.xls    (legacy, ends 2020)
├── ISP1DayAheadRESForecast/       *.xlsx   (Target-Model successor)
├── ISP2DayAheadRESForecast/       *.xlsx
├── ISP3IntraDayRESForecast/       *.xlsx
└── UnitProduction/                *.zip → extracted to <archive>/report*.xls
```

---

## Stage 2 — Uniting the files into tidy tables (`parse_admie.py`)

Each raw file is a *human-formatted* spreadsheet: Greek labels, several header
rows, a junk metadata sheet, hours/periods laid out across columns, and a
layout that varies by category **and changes over the years**. Stage 2 reads
every file for a category and concatenates them into a single tidy long table.

> **Note on AI assistance:** the data-collection script and logic (stage 1) is my own work.
> I used **Claude (Anthropic)** *only* for stage 2 — uniting the individual
> files into one table per category — because of the genuine complexity of the
> raw data: inconsistent multi-row layouts, Greek headers, mid-history format
> changes, and a resolution switch (see below) that made a naive concatenation
> impossible. Everything was reviewed and validated end-to-end before inclusion.

### Unified output schema

Every category is normalized to the same long format:

| column | meaning |
|---|---|
| `category` | the ADMIE FileCategory |
| `date` | calendar date of the file |
| `period` | index within the day (1–24 hourly, 1–48 half-hourly, 1–96 quarter-hourly); empty for daily snapshots |
| `resolution` | `60min`, `30min`, `15min`, or `daily` |
| `timestamp` | interval **start** = `date + (period-1) × resolution` |
| `series` | the row label (load type, unit, country, zone, reservoir…) |
| `value` | numeric value |

Output is written **UTF-8 with BOM** so the Greek `series` labels render
correctly in Excel.

### How the layouts are handled

The files fall into three structural families, each with a dedicated parser:

- **matrix** — periods across columns, series down rows
  (`UnitProduction`, `SystemRealizationSCADA`, `DayAheadRESForecast`,
  `ISP1DayAheadRESForecast`).
- **realtime** — stacked blocks of *label → `Date`+periods header → values row*
  (`RealTimeSCADARES`, `RealTimeSCADASystemLoad`, `RealTimeSCADAImportsExports`).
- **reservoir** — a daily snapshot of one filling-rate value per reservoir
  (`ReservoirFillingRate`).

Real-world quirks the parser resolves:

- **Format per file** — `.xls` is read with `xlrd`, `.xlsx` with `openpyxl`,
  chosen by extension.
- **Sheet drift** — the data sheet is chosen by content, not position (some years
  put the junk `XDO_METADATA` sheet first).
- **Header gaps** — period headers are mapped column→period (tolerant of stray
  empty cells) instead of assuming a clean 1..N run.
- **DST** — the spillover 25th hour (a `0.0` placeholder on normal days) is
  dropped, while genuine 23-/25-hour DST days are preserved.
- **Resolution change** — `ISP1DayAheadRESForecast` switched from **30-minute
  (48/day)** to **15-minute (96/day)** resolution on **2025-10-01** (EU 15-minute
  Market Time Unit rollout); resolution is auto-detected per file.

```bash
# parse everything -> data/processed/<category>.csv
python parse_admie.py

# inspect one sample file per category (no full run)
python parse_admie.py --sample
```

---

## File categories & coverage

| Category | Description | Coverage | Resolution |
|---|---|---|---|
| `SystemRealizationSCADA` | Unit production & system facts | 2019 → present | 60min |
| `RealTimeSCADARES` | Real-time RES injections | 2020-05 → present | 60min |
| `RealTimeSCADASystemLoad` | System load | 2020-05 → present | 60min |
| `RealTimeSCADAImportsExports` | Net interconnection flows | 2020-05 → present | 60min |
| `ReservoirFillingRate` | Hydro reservoir filling rate | 2022-03 → present | daily |
| `DayAheadRESForecast` | Day-ahead RES forecast (legacy) | 2019 → 2020-11 | 60min |
| `ISP1DayAheadRESForecast` | Day-ahead RES forecast — ISP1 (Target Model) | 2020-10 → present | 30min → 15min |
| `ISP2DayAheadRESForecast` † | Day-ahead RES re-forecast — ISP2 | 2020-10 → present | 30min → 15min |
| `ISP3IntraDayRESForecast` † | Intraday RES forecast — ISP3 | 2020-10 → present | 30min → 15min |
| `UnitProduction` | Per-unit production (legacy) | 2019 → 2020-10 | 60min |

† **`ISP2DayAheadRESForecast` and `ISP3IntraDayRESForecast` are collected by the
downloader for completeness but are not part of the unified processed dataset.**
They are re-forecast / intraday refinements of the same RES forecast; the
analysis uses `ISP1DayAheadRESForecast`, the direct day-ahead analog of the
legacy `DayAheadRESForecast`.

> Greece launched the EU **Target Model** market on 1 November 2020, which
> retired several legacy files (`UnitProduction`, `DayAheadRESForecast`) and
> replaced them with the ISP (Integrated Scheduling Process) series.

---

## Requirements

- Python 3.10+
- `requests` (downloading)
- `xlrd` (reads `.xls`), `openpyxl` (reads `.xlsx`)

```bash
pip install requests xlrd openpyxl
```

---

## Data source & license

All data is published by **ADMIE / IPTO** (Independent Power Transmission
Operator of Greece) on their public Operation & Market Files service
(<https://www.admie.gr>). Please consult ADMIE's terms for any redistribution of
the underlying data. This repository contains the collection and processing
**code**; the raw and processed data are reproducible by running the pipeline.
