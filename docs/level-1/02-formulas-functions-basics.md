# 02 · Formulas & Functions Basics

Every formula starts with `=`. This module covers writing your own
arithmetic formulas, using built-in **functions** like `SUM` and `AVERAGE`,
and the single most important habit in all of Excel: knowing when a cell
reference should shift as you copy it (**relative**) and when it must stay
locked (**absolute**). We'll keep using the budget table from Module 1.

## 1. Your first formula

1. Open `budget-tracker.xlsx` from Module 1. In cell `D1`, type `Difference`
   as a header.
2. In `D2`, type `=C2-B2` and press Enter. The cell displays `0` — Actual
   (1200) minus Budgeted (1200) for Rent.
3. Every formula is just `=` followed by an expression: cell references,
   numbers, and operators (`+`, `-`, `*`, `/`, `^` for exponents). Excel
   follows standard order of operations — multiplication/division before
   addition/subtraction — and parentheses `()` override that order.
4. Click `D2` again and look at the Formula Bar — it shows `=C2-B2`, the
   formula, while the cell itself shows `0`, the calculated result. This
   distinction (stored formula vs. displayed value) matters constantly.

## 2. Copying formulas with the fill handle

1. Select `D2`, grab its fill handle (bottom-right corner square), and drag
   down to `D6`.
2. Check `D3`: it shows `55`, and its formula (Formula Bar) reads `=C3-B3`
   — not `=C2-B2`. Excel shifted the references down one row automatically
   because they were **relative references**.
3. The completed column should read: Rent `0`, Groceries `55`, Transport
   `-20`, Entertainment `75`, Savings `0` — each Actual minus its own
   Budgeted.
4. This auto-shifting is what makes the fill handle useful at all — one
   formula, written once, correctly adapts to every row it's copied into.

## 3. SUM and AVERAGE

1. In `A8`, type `Total`. In `B8`, type `=SUM(B2:B6)` and press Enter — it
   returns `2150`, the sum of all five Budgeted values.
2. Copy `B8` across to `C8` (fill handle, drag right one cell). `C8` shows
   `=SUM(C2:C6)` and evaluates to `2260`, the sum of Actual values.
3. In `A9`, type `Average`. In `B9`, type `=AVERAGE(B2:B6)` — returns `430`.
   In `C9`, `=AVERAGE(C2:C6)` returns `452`.
4. A function call is always `FUNCTIONNAME(arguments)`. `SUM` and `AVERAGE`
   both accept one or more ranges or individual cells, comma-separated —
   `=SUM(B2:B4,B6)` would sum four of the five rows, skipping `B5`.
5. `MAX(B2:B6)` and `MIN(B2:B6)` return the largest and smallest values in
   a range (`1200` and `100` for the Budgeted column) — useful for a quick
   sanity check that no value is wildly out of range. `COUNT(B2:B6)`
   returns `5`, the number of cells in the range that contain numbers.

## 4. Relative vs. absolute references

1. In `E1`, type `% of Budget`. In `E2`, type `=B2/$B$8` and press Enter —
   this returns roughly `55.8%` (format the cell as a percentage: **Home >
   Number Format > Percentage**, or Ctrl/⌘+Shift+5).
2. The `$` before both the column letter and row number in `$B$8` makes it
   an **absolute reference** — it will not shift when copied. `B2` (no
   `$`) stays **relative** and shifts normally.
3. Copy `E2` down to `E6`. Check `E3`'s formula: it reads `=B3/$B$8` — the
   numerator shifted to `B3` (relative), but the denominator stayed locked
   at `$B$8` (absolute), because every row's percentage needs to divide by
   the *same* total, not a shifting one.
4. Without the `$`, copying `=B2/B8` down to row 3 would produce
   `=B3/B9` — dividing by an empty cell, since the total only lives in
   `B8`. This is the single most common formula bug in Excel; when a copied
   formula suddenly returns `0` or `#DIV/0!`, a missing `$` on the intended
   fixed cell is the first thing to check.
5. A **mixed reference** locks only the row (`B$8`) or only the column
   (`$B8`) — useful when copying a formula both across and down a grid
   where one axis should shift and the other shouldn't. Press **F4**
   (Fn+F4 on some Mac keyboards) while a reference is selected in the
   Formula Bar to cycle through `B8` → `$B$8` → `B$8` → `$B8` → `B8`
   without retyping the `$` signs manually.
6. The five percentages read: Rent `55.8%`, Groceries `18.6%`, Transport
   `7.0%`, Entertainment `4.7%`, Savings `14.0%` — they sum to
   approximately `100%` (rounding on each displayed percentage means the
   sum can land a tenth of a point off, e.g. `100.1%`, which is expected
   and not an error).

## Cheat sheet

| Function/Syntax | Purpose | Example |
|---|---|---|
| `=A1-B1` | Basic arithmetic formula | `=C2-B2` → difference |
| `SUM(range)` | Add all values in a range | `=SUM(B2:B6)` → `2150` |
| `AVERAGE(range)` | Mean of values in a range | `=AVERAGE(C2:C6)` → `452` |
| `MAX(range)` / `MIN(range)` | Largest / smallest value | `=MAX(C2:C6)` → `1200` |
| `COUNT(range)` | Count of numeric cells | `=COUNT(B2:B6)` → `5` |
| `A1` | Relative reference — shifts when copied | |
| `$A$1` | Absolute reference — never shifts | |
| `$A1` / `A$1` | Mixed reference — column or row locked | |
| F4 | Cycle reference type in Formula Bar | |

## Exercise

In `budget-tracker.xlsx`, add a `Difference` column (`=C2-B2`, copied down),
a `Total` row using `SUM` for both Budgeted and Actual, an `Average` row
using `AVERAGE` for both, and a `% of Budget` column using an absolute
reference to the Budgeted total so it doesn't break when copied down.
Confirm the Budgeted total is `2150`, the Actual total is `2260`, and the
`% of Budget` column sums to approximately `100%`.
