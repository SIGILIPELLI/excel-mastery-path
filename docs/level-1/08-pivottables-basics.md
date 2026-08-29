# 08 · PivotTables Basics

A PivotTable summarizes a long, row-by-row dataset — every transaction, one
row each — into a compact cross-tab: totals by category, by month, by
type, recombined however you drag the fields. This module builds a
two-month transaction log from the budget categories and pivots it three
different ways.

## 1. Building the source data

PivotTables need flat, transaction-level data — one row per fact, not a
pre-summarized table like the `Budget` sheet from earlier modules.

1. Add a new sheet named `Transactions`. Enter this table starting at
   `A1` (14 rows: 1 header + 12 data rows):

   | Category | Type | Month | Amount |
   |---|---|---|---|
   | Rent | Fixed | Jan | 1200 |
   | Groceries | Variable | Jan | 455 |
   | Transport | Variable | Jan | 130 |
   | Entertainment | Variable | Jan | 175 |
   | Savings | Fixed | Jan | 300 |
   | Subscriptions | Fixed | Jan | 25 |
   | Rent | Fixed | Feb | 1200 |
   | Groceries | Variable | Feb | 410 |
   | Transport | Variable | Feb | 140 |
   | Entertainment | Variable | Feb | 90 |
   | Savings | Fixed | Feb | 300 |
   | Subscriptions | Fixed | Feb | 25 |

2. Convert it to a Table (Module 4): select any cell inside it,
   **Ctrl/⌘+T**, confirm headers, name it `Transactions`. PivotTables built
   from a Table auto-expand their source range if rows are added later.

## 2. Creating a PivotTable

1. Click any cell inside the `Transactions` Table. **Insert > PivotTable**.
2. Confirm the table/range shows `Transactions` and choose **New
   Worksheet**, then **OK**. A blank PivotTable area appears alongside the
   **PivotTable Fields** panel (usually docked right), listing `Category`,
   `Type`, `Month`, `Amount`.
3. Drag `Category` into the **Rows** area, and `Amount` into the
   **Values** area. It defaults to **Sum of Amount**, showing each
   category's Jan+Feb combined total: Rent `2400`, Groceries `865`,
   Transport `270`, Entertainment `265`, Savings `600`, Subscriptions
   `50`, with a **Grand Total** of `4450`.
4. Rename the resulting sheet tab to `Pivot-ByCategory`.

## 3. Pivoting by a different field

1. On the same PivotTable, drag `Category` out of **Rows** and drag `Type`
   into **Rows** instead (leave `Amount` in **Values**).
2. The PivotTable now shows just two rows: `Fixed` totaling `3050`
   (Rent `2400` + Savings `600` + Subscriptions `50`) and `Variable`
   totaling `1400` (Groceries `865` + Transport `270` + Entertainment
   `265`), with the same Grand Total, `4450` — the same underlying data,
   summarized along a different dimension, took seconds instead of
   rewriting formulas.
3. Drag `Category` into the **Rows** area *below* `Type` (both `Type` and
   `Category` now stacked in Rows) — the PivotTable now nests each
   category under its Fixed/Variable group, giving both the group subtotal
   and the individual category breakdown in one view.

## 4. Adding Month as a column

1. Drag `Month` into the **Columns** area, keeping `Category` in **Rows**
   and `Amount` in **Values**.
2. The PivotTable now shows a full cross-tab: each category as a row,
   `Jan` and `Feb` as separate columns, and a `Grand Total` column on the
   right. Rent shows `1200` / `1200` (flat across both months); Groceries
   shows `455` / `410`; the `Grand Total` row at the bottom should read
   `2285` for Jan and `2165` for Feb, matching each month's total from
   Section 1.
3. Right-click any value inside the PivotTable and choose **Number Format**
   to apply Currency formatting to every value cell at once, rather than
   formatting each cell individually as you would on a plain range.

## 5. Filtering and refreshing

1. Drag `Type` into the **Filters** area (above Rows/Columns). A dropdown
   appears above the PivotTable — select `Variable` only, and the whole
   table recalculates to show just Groceries, Transport, and Entertainment.
2. Reset the filter back to **(All)** to restore every category.
3. If you edit the source `Transactions` Table (e.g. change an Amount),
   the PivotTable does **not** update automatically — right-click anywhere
   inside it and choose **Refresh**, or **PivotTable Analyze > Refresh**,
   to pull in the change. This is the single most common PivotTable
   confusion: a PivotTable is a snapshot-and-summary, not a live formula.

## Cheat sheet

| Action | Location |
|---|---|
| Insert a PivotTable | Insert > PivotTable |
| Add a summarized field | Drag field into Values (defaults to Sum) |
| Group by a dimension | Drag field into Rows or Columns |
| Restrict to specific values | Drag field into Filters |
| Change summary type (Sum/Average/Count) | Click Values field > Value Field Settings |
| Reformat all values at once | Right-click a value > Number Format |
| Pull in changed source data | Right-click > Refresh |

## Exercise

Build the `Transactions` Table and create a PivotTable summarizing Amount
by Category (confirm Rent totals `2400` and the Grand Total is `4450`).
Then rebuild it summarizing by Type instead (confirm Fixed `3050` /
Variable `1400`), and finally build a cross-tab with Category in Rows and
Month in Columns (confirm the Jan column totals `2285` and the Feb column
totals `2165`). Add a Filter on Type and confirm selecting `Variable`
narrows the table to exactly three categories.
