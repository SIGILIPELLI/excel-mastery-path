# 10 · Project — Personal Budget Tracker

This capstone combines every Level 1 module into one finished workbook:
`budget-tracker.xlsx`, built incrementally since Module 1. If you followed
along module by module, most of this is already done — this project is
the checklist that ties it together and verifies every number is correct.

## Final workbook structure

Your finished `budget-tracker.xlsx` should contain four sheets:

1. **Budget** — the core Table (Module 4) with Category, Budgeted, Actual,
   Difference, and lookup-driven Type/Priority/Status/Needs Review columns
   (Modules 6–7), plus conditional formatting (Module 3) and at least one
   chart (Module 5).
2. **CategoryInfo** — the reference table of Category, Type, Priority, and
   Level (Modules 6–7).
3. **Transactions** — the flat, transaction-level log with a PivotTable
   built from it (Module 8).
4. **Messy** — the cleaning scratch sheet (Module 9), which can stay as a
   demonstration or be deleted once you've internalized the technique — it
   isn't part of the "live" budget itself.

## Build checklist

- [ ] **Budget sheet** is a real Table (`Ctrl/⌘+T`) named `Budget`, with
      headers `Category | Budgeted | Actual | Difference`, six rows of
      data (Rent, Groceries, Transport, Entertainment, Savings,
      Subscriptions), and a Total Row showing Sum.
- [ ] `Difference` column uses `=C2-B2` (relative references, fill-handle
      copied), not hardcoded values.
- [ ] A `% of Budget` column uses an absolute reference to the Budgeted
      total (`$B$8` or the Table's total row) so it doesn't break when
      copied.
- [ ] Header row is bold with a fill color; the whole table has borders;
      `Budgeted`/`Actual`/`Difference` are formatted as Currency.
- [ ] Conditional formatting on `Difference`: red fill for values greater
      than `0`, green fill for values less than `0`.
- [ ] At least one chart exists comparing Budgeted vs. Actual (clustered
      column) with a title and data labels.
- [ ] `Type` column on `Budget` is pulled via `VLOOKUP` (or `XLOOKUP`/
      `INDEX`+`MATCH`) from `CategoryInfo`.
- [ ] `Status` column uses a nested `IF` (or `IFS`) returning `Over
      Budget` / `Under Budget` / `On Track` based on `Difference`.
- [ ] `Needs Review` column uses `AND(Type="Variable", Difference>50)`.
- [ ] `CategoryInfo` sheet has `Level` computed with `IFS` from `Priority`.
- [ ] `Transactions` sheet is a real Table with a `Category` column
      restricted by a Data Validation dropdown sourced from
      `CategoryInfo!$A$2:$A$7`.
- [ ] A PivotTable exists (on its own sheet or tab) summarizing
      `Transactions` by Category, and a second view (or the same
      PivotTable's field layout changed) by Type.

## Verified final numbers

Use this table to check your finished workbook against known-correct
results (all values verified independently, matching every earlier
module):

| Metric | Expected value |
|---|---|
| Sum of Budgeted (6 categories) | `2170` |
| Sum of Actual (6 categories) | `2285` |
| Average Actual | `380.83` (`2285 / 6`) |
| Difference: Groceries | `55` (Over Budget) |
| Difference: Transport | `-20` (Under Budget) |
| Difference: Rent | `0` (On Track) |
| % of Budget: Rent | `≈55.3%` (`1200 / 2170`) |
| PivotTable: Sum of Amount, Fixed | `3050` |
| PivotTable: Sum of Amount, Variable | `1400` |
| PivotTable: Grand Total (Jan + Feb) | `4450` |
| PivotTable: Jan column total | `2285` |
| PivotTable: Feb column total | `2165` |
| Categories flagged `Needs Review` | Groceries, Entertainment (only) |

!!! info "Why Sum of Budgeted here is 2170, not 2150"
    Module 1 introduced five categories summing to `2150`; Module 4 added
    a sixth, Subscriptions (Budgeted `20`), bringing the total to `2170`.
    If your Budget sheet only has the original five categories, your Sum
    of Budgeted will correctly read `2150` instead — either is fine as
    long as it's internally consistent with which rows you built.

## Cheat sheet: full function/feature reference for Level 1

| Category | Tools covered |
|---|---|
| Navigation | Name Box, Ctrl/⌘+Arrow, ranges |
| Formulas | `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, relative/absolute refs |
| Formatting | Number formats, Conditional Formatting (Highlight/Data Bars/Color Scales) |
| Tables | `Ctrl/⌘+T`, structured references, Total Row |
| Charts | Column, Pie, Chart Elements |
| Lookups | `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, approximate-match brackets |
| Logic | `IF`, nested `IF`, `IFS`, `AND`, `OR` |
| PivotTables | Rows/Columns/Values/Filters, Refresh |
| Cleaning | `TRIM`, `PROPER`, `UPPER`/`LOWER`, Remove Duplicates, Data Validation |

## Exercise

Finish (or rebuild from scratch) `budget-tracker.xlsx` against the full
checklist above. Then extend it with one category of your own choosing —
add it to `Budget`, `CategoryInfo`, and two months of `Transactions` rows
— and confirm every formula, lookup, conditional format, and the
PivotTable all update correctly to include it without any manual formula
rewriting. This is the real test of whether the workbook was built with
proper structured references and Tables rather than one-off hardcoded
ranges.
