# 01 · Enterprise Excel Architecture

This module covers how to structure Excel work so it survives contact
with an organization: multi-workbook systems, linked workbooks versus
a single Data Model, naming/versioning conventions, and where Excel's
architecture starts to need a database or Power BI instead.

## 1. Worked dataset — a linked two-workbook system

Simulate a common enterprise pattern: a `Source.xlsx` workbook holding
raw data, and a `Report.xlsx` workbook that consumes it via an
external reference. On `Source.xlsx`, sheet `Data`, `A1:B4`:

| | A | B |
|---|---|---|
| **1** | Department | Headcount |
| **2** | Engineering | 42 |
| **3** | Sales | 18 |
| **4** | Support | 11 |

In `Report.xlsx`, `A1`:
`=[Source.xlsx]Data!B2` → returns `42` while `Source.xlsx` is open, and
continues to work closed as long as the file path does not move
(Excel caches the last known value and re-reads on open).
Total headcount in `A2`:
`=SUM([Source.xlsx]Data!B2:B4)` → `42+18+11=71`.

## 2. Why linked workbooks are fragile at scale

- Moving or renaming `Source.xlsx` breaks every external reference in
  every consuming workbook — there is no central registry of who
  depends on what.
- **Edit → Links → Update Values / Change Source** lets you repair a
  broken link, but this is manual, per-workbook, per-broken-link work
  — it does not scale past a handful of files.
- The enterprise-safer alternative is a single workbook with **Power
  Query connections** pulling from a shared source (a folder, database,
  or SharePoint list) rather than cell-level links between `.xlsx`
  files — Power Query re-resolves the connection on refresh rather than
  caching a brittle cell reference.

## 3. A single Data Model instead of scattered VLOOKUPs

1. Load both `Department` and a second table, `Budget` (`A1:B4`:
   `Department/Amount`, `Engineering/500000`, `Sales/200000`,
   `Support/90000`) into **Power Pivot** as two separate tables.
2. Create a relationship: `Department[Department] →
   Budget[Department]`.
3. A DAX measure, `Total Budget := SUM(Budget[Amount])`, can now be
   sliced by any field from `Department` because of the relationship
   — no `VLOOKUP` or helper column needed. This is the architectural
   shift enterprise workbooks make: relationships in a Data Model
   instead of lookup formulas copied across every report sheet.

## 4. Naming and versioning conventions

- File naming: `ProjectName_YYYYMMDD_v#.xlsx` (e.g.
  `HeadcountReport_20260831_v3.xlsx`) sorts correctly by date and makes
  the latest version unambiguous in a shared folder.
- Sheet naming: prefix internal working sheets with `_` (e.g.
  `_calc`, `_lookup`) and hide them, so only `Dashboard` and `Inputs`
  are visible to a business user opening the file.
- Named ranges over hardcoded addresses: `=SUM(Revenue)` survives a
  row insertion; `=SUM(C2:C50)` does not automatically extend.

## 5. Where Excel's architecture hits a ceiling

| Symptom | Enterprise fix |
|---|---|
| Multiple people editing the same file causes overwritten changes | Move to a shared Data Model source (SQL/SharePoint list) with each person's Excel just querying it |
| A workbook exceeds ~50MB or takes seconds to recalculate | Power Pivot's compressed columnar Data Model, or move the heavy table to a database and query it with Power Query |
| The same KPI is defined slightly differently in five workbooks | Centralize the definition as one DAX measure in one shared Power Pivot model, then have every report reference that model |

## Cheat sheet

| Concept | Purpose |
|---|---|
| External reference (`[Book.xlsx]Sheet!Cell`) | Pull a live value from another workbook |
| Power Query connection | Refreshable pull from a shared source, more robust than links |
| Power Pivot relationship | Replace VLOOKUP chains with a queryable Data Model |
| `_`-prefixed hidden sheets | Separate internal calc sheets from the user-facing view |

## Exercise

Using the `Department`/`Budget` relationship above, write a DAX measure
`Budget per Head := DIVIDE([Total Budget], [Total Headcount])` (define
`Total Headcount := SUM(Department[Headcount])` first) and confirm the
overall ratio is `790000/71 ≈ 11126.76`.
