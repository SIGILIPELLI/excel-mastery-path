# 09 · Auditing & Error-Checking Large Workbooks

This module covers the built-in tools for tracing formula logic,
finding inconsistencies, and catching errors before a workbook ships:
Formula Auditing, error-checking functions, and data validation as a
prevention layer.

## 1. Worked dataset — a small expense sheet with a planted bug

Build this on a sheet named `Audit`, `A1:C7`:

| | A | B | C |
|---|---|---|---|
| **1** | Item | Qty | Cost |
| **2** | Paper | 10 | 2.5 |
| **3** | Pens | 20 | 0.8 |
| **4** | Folders | 15 | 1.2 |
| **5** | Staples | 5 | 3.0 |
| **6** | Total Qty | `=SUM(B2:B4)` | |
| **7** | Total Cost | | `=SUM(C2:C4)` |

Notice `B6` and `C7` both omit row 5 (Staples) — a common copy-paste
range bug.

## 2. Tracing precedents and dependents

1. Click `B6` → **Formulas → Trace Precedents**. Blue arrows point to
   `B2:B4`, visually confirming the range stops one row short of the
   full list (`B5` is not included) — the bug is now visible, not just
   suspected.
2. Click `B2` → **Formulas → Trace Dependents**. An arrow points
   forward to `B6`, showing which totals would be affected if `B2`
   changes.
3. **Formulas → Remove Arrows** clears the trace lines when done.

## 3. Fixing and re-verifying

1. Correct `B6` to `=SUM(B2:B5)` → `10+20+15+5=50`.
2. Correct `C7` to `=SUM(C2:C5)` →
   `2.5+0.8+1.2+3.0=7.5`.
3. Re-run Trace Precedents on both — the arrows now correctly span
   `B2:B5` / `C2:C5`.

## 4. Error-checking functions

1. In `D2:D5`, compute a line extension: `D2`: `=B2*C2` →
   `10*2.5=25`.
2. In `D6`: `=SUM(D2:D5)` → `25+16+18+15=74`.
   Check: Pens `20*0.8=16`, Folders `15*1.2=18`, Staples `5*3.0=15`.
   `25+16+18+15=74`. ✓
3. Introduce a deliberate `#DIV/0!` for practice: in `E2`,
   `=D2/(B2-10)` gives `=25/0` → `#DIV/0!` (since `B2=10`).
4. Wrap it safely: `=IFERROR(D2/(B2-10),"n/a")` → returns `"n/a"`
   instead of propagating the error into any formula that references
   `E2`.
5. `=ISERROR(D2/(B2-10))` → `TRUE`, confirming the cell does contain an
   error before deciding how to handle it.

## 5. Data validation as prevention

1. Select `B2:B5` → **Data → Data Validation → Whole Number → greater
   than or equal to 0**. This stops a negative quantity from ever being
   typed in, catching the error at entry instead of during an audit.
2. Add an **Input Message** ("Enter a non-negative quantity") and an
   **Error Alert** ("Quantity cannot be negative") so the guardrail is
   self-documenting for the next person who edits the sheet.

## 6. Evaluate Formula and Watch Window

1. Select `D6` → **Formulas → Evaluate Formula** → click **Evaluate**
   repeatedly to step through `SUM(D2:D5)` expanding one operand at a
   time — useful for debugging a nested formula that returns an
   unexpected result.
2. **Formulas → Watch Window → Add Watch** on `D6` lets you monitor
   that cell's value while scrolling or editing elsewhere in a large
   workbook, without needing it visible on screen.

## Cheat sheet

| Tool | Purpose |
|---|---|
| Trace Precedents/Dependents | Visualize which cells feed into / depend on a formula |
| `IFERROR(formula, fallback)` | Replace an error result with a safe fallback |
| `ISERROR(expr)` | Test whether an expression evaluates to an error |
| Evaluate Formula | Step through a nested formula's evaluation order |
| Data Validation | Prevent bad input at the source |

## Exercise

The `Audit` sheet's `D6` total is `74`. Add a `Discount` row (`Qty=0`,
`Cost=-5`, meaning a flat $5 discount) and extend all three ranges
(`B6`, `C7`, `D6`) to include it. Use Trace Precedents to confirm all
three now cover the new row, then confirm the new `D6` total is `69`.
