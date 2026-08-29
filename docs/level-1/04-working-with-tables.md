# 04 · Working with Tables

A plain range of cells with headers *looks* like a table, but Excel's
**Table** feature (Insert > Table) turns it into an actual structured
object: automatic filtering, banded rows, formulas that auto-fill down new
rows, and **structured references** — formulas that read
`=SUM(Budget[Actual])` instead of `=SUM(C2:C6)`. This module converts the
budget range into a real Table.

## 1. Creating a Table

1. Open `budget-tracker.xlsx`. Click any cell inside the data range
   (`A1:D6`, including the `Difference` column from Module 2).
2. **Insert > Table** (or **Ctrl/⌘+T**). Excel guesses the range
   (`A1:D6`) and shows a dialog with **My table has headers** checked —
   confirm it's checked, since row 1 holds labels, not data, then click
   **OK**.
3. The range now displays with a default banded-row style (alternating
   shading), a filter-arrow dropdown on each header cell, and a new
   **Table Design** tab in the ribbon.
4. Rename the Table from the default `Table1` to `Budget`: click any cell
   inside it, then **Table Design > Table Name** (top-left of that tab)
   and type `Budget`, Enter. This name is what structured references use.
5. Table names must be unique per workbook, contain no spaces, and can't
   collide with cell addresses — `Budget` is valid, `A1` would not be.

## 2. Structured references

1. Click an empty cell below the Table, e.g. `B8`, and type
   `=SUM(Budget[Budgeted])`. Press Enter — it returns `2150`, identical to
   `=SUM(B2:B6)` from Module 2, but readable without knowing which raw
   columns hold what.
2. In `C8`, type `=SUM(Budget[Actual])` — returns `2260`.
3. `Budget[Budgeted]` refers to the entire `Budgeted` column's data cells
   (not the header), automatically — if rows are added or removed from the
   Table, this reference adjusts on its own, unlike a fixed range like
   `B2:B6` which would need manual editing.
4. `Budget[#Headers]` refers to just the header row; `Budget[#All]` refers
   to headers plus data; `Budget[@Actual]` (used *inside* a formula that
   lives in the same Table row) refers to the Actual value in that specific
   row — the structured-reference equivalent of a relative reference like
   `C2`.

## 3. Auto-fill and auto-expand

1. Click into the empty cell directly below the last Table row (row 7,
   column A) and type a new category, e.g. `Subscriptions`. Press Enter —
   Excel automatically extends the Table to include this new row, complete
   with the banded formatting and filter behavior.
2. Type `20` in the new row's Budgeted cell and `25` in Actual. Because the
   `Difference` column already holds a formula (`=C2-B2` pattern) in every
   existing row, Excel automatically fills that same formula into the new
   row too — this "calculated column" behavior only applies inside a real
   Table, not a plain range.
3. The new row's Difference should read `5` — confirm it, and confirm
   `=SUM(Budget[Actual])` in `C8` (or wherever your totals sit) has updated
   to `2285` to include the new row, without editing the SUM formula
   itself.

## 4. Sorting and filtering

1. Click the filter-arrow on the `Actual` header. Choose **Sort Largest to
   Smallest** — rows reorder by Actual value, and every formula
   referencing structured references (`Budget[Actual]`, etc.) continues to
   work correctly because the references track the column, not fixed row
   numbers.
2. Click the filter-arrow on `Category`, uncheck **Select All**, then check
   only `Rent` and `Groceries` — the Table temporarily hides every other
   row (their row numbers appear in blue, signaling a filter is active)
   without deleting them.
3. **Table Design > Filter Button** toggle (checkbox in the ribbon) hides
   or shows the header filter arrows without removing the underlying
   filter/sort state.
4. Re-open the `Category` filter and click **Select All** to restore every
   row.

## 5. Total Row

1. **Table Design > Total Row** (checkbox) adds a summary row at the
   bottom of the Table.
2. Click the cell under `Actual` in that new Total Row — a dropdown
   appears with `None`, `Average`, `Count`, `Max`, `Min`, `Sum`, and more.
   Choose **Sum** — it shows `2285` (including the Subscriptions row added
   in Section 3), generated as `=SUBTOTAL(109,Budget[Actual])` rather than
   a plain `SUM`, so it automatically ignores any rows hidden by a filter.
3. Set the `Budgeted` column's Total Row to **Sum** as well — it should
   read `2170` (`2150` original + `20` Subscriptions).

## Cheat sheet

| Action | Shortcut / Location |
|---|---|
| Convert range to Table | Ctrl/⌘+T |
| Rename Table | Table Design > Table Name |
| Whole-column structured reference | `TableName[ColumnName]` |
| Same-row structured reference | `TableName[@ColumnName]` |
| Toggle Total Row | Table Design > Total Row |
| Toggle filter arrows | Table Design > Filter Button |
| Sort by column | Header filter-arrow > Sort Largest/Smallest |
| Filter to specific values | Header filter-arrow > uncheck values |

## Exercise

Convert the budget range in `budget-tracker.xlsx` into a Table named
`Budget`. Add `=SUM(Budget[Actual])` and `=SUM(Budget[Budgeted])` formulas
below it using structured references. Add a new category row
(`Subscriptions`, Budgeted `20`, Actual `25`) directly below the Table and
confirm it auto-expands with the `Difference` formula filled in
automatically. Turn on the Total Row and set both Budgeted and Actual to
Sum — confirm they read `2170` and `2285`.
