# 04 · DAX Basics for Power Pivot

DAX (Data Analysis Expressions) is Power Pivot's formula language.
Unlike a calculated *column* (evaluated row by row), a DAX **measure**
is computed dynamically based on whatever filter context a PivotTable
applies — the same measure returns different numbers depending on
which row/column it's placed in.

## 1. Worked dataset

Reuse the `Sales`/`Products` Data Model from Module 3, or rebuild:

`Sales` (`A1:C6`): ProductID/Qty/Rep — `P1,10,Alan`; `P2,5,Priya`;
`P1,8,Sam`; `P3,12,Alan`; `P2,6,Priya`.

`Products` (`A1:B4`): ProductID/Price — `P1,20`; `P2,35`; `P3,15`.
Relationship: `Sales[ProductID]` → `Products[ProductID]`.

## 2. A basic measure

1. Power Pivot → **Measures** → New Measure. Name `Total Qty`,
   formula: `Total Qty := SUM(Sales[Qty])`.
2. Drop it into a PivotTable with Rows = `Sales[Rep]`. Manual check:
   Alan (rows 2,5) → `10+12=22`. Priya (rows 3,6) → `5+6=11`. Sam
   (row 4) → `8`. Grand total `22+11+8=41`, which also equals
   `SUM` of all six Qty values: `10+5+8+12+6=41`. ✓

## 3. A measure using RELATED-equivalent logic

1. New Measure: `Total Revenue := SUMX(Sales, Sales[Qty] *
   RELATED(Products[Price]))`
2. `SUMX` is an iterator — it walks each row of `Sales`, computes
   `Qty * RELATED(Price)` per row, then sums the results. This is the
   measure-level equivalent of the calculated column from Module 3.
3. Manual check by Rep: Alan = `10*20 + 12*15 = 200+180=380`. Priya =
   `5*35 + 6*35 = 175+210=385`. Sam = `8*20=160`. Total `380+385+160
   =925` — matches Module 3's calculated-column total exactly, which
   is the expected consistency check between the two approaches.

## 4. CALCULATE and filter context

1. New Measure: `Alan Revenue := CALCULATE([Total Revenue],
   Sales[Rep] = "Alan")`
2. `CALCULATE` overrides the current filter context with an explicit
   condition. Regardless of which row of a PivotTable this measure
   sits in, it always computes Total Revenue **as if** filtered to
   Alan only.
3. Manual check: this must always show `380` (Alan's total from
   Section 3), even when placed next to Priya's row in a PivotTable —
   that fixed value regardless of row context is exactly what
   distinguishes `CALCULATE` from a plain measure.

## 5. Average and ratio measures

1. New Measure: `Avg Deal Size := AVERAGEX(Sales, Sales[Qty] *
   RELATED(Products[Price]))`
2. Manual check: the five per-row revenues are `200, 175, 160, 180,
   210`. Sum `925`, count `5`, average `185`. The measure must return
   `185`.
3. New Measure: `% of Total := DIVIDE([Total Revenue],
   CALCULATE([Total Revenue], ALL(Sales)))`
   `DIVIDE(numerator, denominator)` is DAX's safe division — it
   returns blank instead of `#DIV/0!` if the denominator is zero.
   `ALL(Sales)` removes any row-context filter, so the denominator is
   always the grand total `925` regardless of which Rep row this sits
   in. Placed on Alan's row: `380/925 ≈ 41.1%`.

## Cheat sheet

| DAX construct | Purpose |
|---|---|
| `SUM(Table[Col])` | Simple column total |
| `SUMX(Table, expr)` | Row-by-row iterator, then sum |
| `AVERAGEX(Table, expr)` | Row-by-row iterator, then average |
| `CALCULATE(expr, filter)` | Override the current filter context |
| `ALL(Table)` | Remove filters — useful for grand-total denominators |
| `DIVIDE(num, denom)` | Division that returns blank instead of erroring on 0 |

## How It Actually Works

DAX measures are not evaluated the way worksheet formulas are — every
measure carries an implicit **filter context** (the set of rows currently
"visible" given active slicers, PivotTable rows/columns, and row filters)
and DAX's core job is to evaluate an expression against whatever rows
survive that context, using VertiPaq's columnar scan rather than a
cell-by-cell dependency graph. `CALCULATE`, the most important DAX
function, works by taking the current filter context and *modifying* it —
adding, removing, or replacing filters — before evaluating its expression
inside that new context; this is the mechanism behind seemingly-magical
patterns like year-over-year comparisons, which are really just
`CALCULATE` swapping the date filter to a different period and
re-evaluating the same base measure against it. Calculated Columns and
Measures use the engine completely differently despite both being "DAX":
a Calculated Column is computed once per row at data-refresh time and
physically stored in the columnar model (costing compression and memory
like any other column), while a Measure has no stored value at all — it is
re-evaluated fresh, from scratch, every time it appears in a visual or
PivotTable cell, against that cell's specific filter context.

## Exercise

Build `Priya % of Total := DIVIDE(CALCULATE([Total Revenue],
Sales[Rep]="Priya"), CALCULATE([Total Revenue], ALL(Sales)))` and
manually verify: `385/925 ≈ 41.6%`.
