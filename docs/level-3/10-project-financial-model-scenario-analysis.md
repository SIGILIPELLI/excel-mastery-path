# 10 · Project — Financial Model with Scenario Analysis

This capstone project combines Level 3's modeling skills (Module 07)
with what-if tooling — Data Tables and Scenario Manager — to test how
a financial model responds to changing assumptions, and to present
several named scenarios side by side.

## 1. Worked dataset — a 3-year product launch model

Build this on a sheet named `Model`, `A1:B7`:

| | A | B |
|---|---|---|
| **1** | Units Sold (Year 1) | 1000 |
| **2** | Annual Growth Rate | 0.15 |
| **3** | Price per Unit | 50 |
| **4** | Variable Cost per Unit | 30 |
| **5** | Fixed Costs (annual) | 8000 |
| **6** | Discount Rate | 0.10 |
| **7** | Initial Investment | -20000 |

## 2. Building the 3-year projection

`D1:G4`:

| | D | E | F | G |
|---|---|---|---|---|
| **1** | | Year 1 | Year 2 | Year 3 |
| **2** | Units | `=B1` | `=D2*(1+$B$2)` | `=E2*(1+$B$2)` |
| **3** | Revenue | `=D2*$B$3` | `=E2*$B$3` | `=F2*$B$3` |
| **4** | Profit | `=D3-D2*$B$4-$B$5` | `=E3-E2*$B$4-$B$5` | `=F3-F2*$B$4-$B$5` |

Manual check:
Year 1: Units `1000`; Revenue `1000*50=50000`; Profit
`50000-1000*30-8000=50000-30000-8000=12000`.
Year 2: Units `1000*1.15=1150`; Revenue `1150*50=57500`; Profit
`57500-1150*30-8000=57500-34500-8000=15000`.
Year 3: Units `1150*1.15=1322.5`; Revenue `1322.5*50=66125`; Profit
`66125-1322.5*30-8000=66125-39675-8000=18450`.

## 3. NPV of the project

In `B9`: `=B7+NPV(B6,D4:F4)`
Discount each profit figure: Year1 `12000/1.10=10909.09`; Year2
`15000/1.10^2=15000/1.21=12396.69`; Year3
`18450/1.10^3=18450/1.331=13860.26`. Sum `=10909.09+12396.69+13860.26=
37166.04`. Add `B7=-20000`: NPV ≈ `17166.04`.

## 4. One-variable Data Table — sensitivity to growth rate

1. In `A12:A17`, list growth rates to test: `0.05, 0.10, 0.15, 0.20,
   0.25`.
2. In `B11` (the column header cell, one row above and one column right
   of the input list per Excel's Data Table convention): `=B9` (link to
   the NPV formula).
3. Select `A11:B17` → **Data → What-If Analysis → Data Table** →
   Column input cell: `B2` (the growth rate cell) → OK. Excel
   substitutes each rate from `A12:A17` into `B2` and recalculates `B9`
   for each, filling `B12:B17` automatically.
4. Expected direction: NPV should increase monotonically as growth
   rate rises from `5%` to `25%`, since higher growth compounds units
   sold (and therefore profit) in Years 2-3.

## 5. Scenario Manager — Best/Worst/Base case

1. **Data → What-If Analysis → Scenario Manager → Add**.
2. **Base Case**: changing cells `B2,B3,B4` = `0.15, 50, 30` (the
   values already in the model).
3. **Best Case**: same changing cells = `0.25, 55, 28`.
4. **Worst Case**: same changing cells = `0.05, 45, 33`.
5. Click **Summary** → Result cell: `B9` (NPV). Excel generates a new
   sheet comparing NPV across all three scenarios side by side —
   Base Case will show the `~17166.04` computed above, Best Case
   noticeably higher (higher growth, higher price, lower cost all push
   NPV up), Worst Case noticeably lower or even negative.

## 6. Presenting the result

Add a summary cell, `B10`:
`=IF(B9>0,"Approve: NPV positive","Reject: NPV negative")`
For the Base Case NPV of `~17166.04`, this reads `"Approve: NPV
positive"` — pairing the numeric result with a plain-language decision
rule is what turns a model into a usable recommendation.

## Cheat sheet

| Tool | Purpose |
|---|---|
| Data Table (one-variable) | Recalculate one output across a range of one input |
| Scenario Manager | Save and compare named sets of multiple input changes |
| NPV + IF decision rule | Turn a numeric result into a plain-language recommendation |

## Exercise

Add a two-variable Data Table testing NPV (`B9`) across combinations of
Growth Rate (`0.05` to `0.25`, rows) and Discount Rate (`0.08` to
`0.14`, columns). Which combination pushes NPV negative first?
