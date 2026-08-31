# 01 · Advanced Formulas (SUMIFS/COUNTIFS/AVERAGEIFS)

Level 1's `SUMIF`/`COUNTIF`/`AVERAGEIF` filter on a single condition. Real
reports almost always need more than one — "total sales for Region = East
**and** Month = March." The `...IFS` family (plural) extends each function
to any number of conditions. This module also covers nested `IF`,
`IFERROR`, and combining conditions with `AND`/`OR`.

All formulas below are hand-verified against the worked dataset in Section
1 by manually tracing which rows match each condition set and adding them
by hand, then confirming the formula returns the same number.

## 1. Worked dataset

Build this table on a sheet named `Sales`, `A1:D11`:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Region | Month | Rep | Amount |
| **2** | East | Jan | Alan | 1200 |
| **3** | East | Feb | Alan | 900 |
| **4** | East | Mar | Priya | 1500 |
| **5** | West | Jan | Priya | 800 |
| **6** | West | Feb | Sam | 1100 |
| **7** | West | Mar | Sam | 1300 |
| **8** | East | Mar | Alan | 700 |
| **9** | West | Jan | Sam | 950 |
| **10** | East | Jan | Priya | 600 |
| **11** | West | Mar | Priya | 1000 |

## 2. SUMIFS — multiple AND conditions

1. In `F1` type `East Mar Total`, in `F2`:
   `=SUMIFS(D2:D11,A2:A11,"East",B2:B11,"Mar")`
2. Syntax: `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2,
   criteria2, ...])` — note the sum range comes **first**, unlike `SUMIF`
   where it comes last. This trips up almost everyone coming from `SUMIF`.
3. Manual check: rows matching East + Mar are row 4 (1500) and row 8 (700).
   `1500 + 700 = 2200`. `F2` should show `2200`.
4. Add a third condition — East, Mar, rep Alan: in `F3`:
   `=SUMIFS(D2:D11,A2:A11,"East",B2:B11,"Mar",C2:C11,"Alan")`
   Only row 8 matches → `700`.

## 3. COUNTIFS and AVERAGEIFS

1. In `F4`, count West deals over 1000:
   `=COUNTIFS(A2:A11,"West",D2:D11,">1000")`
   Matching rows: row 7 (1300), row 11 (1000 is not `>1000`, excluded) —
   only row 7 qualifies → `1`.
2. In `F5`, average deal size for Priya:
   `=AVERAGEIFS(D2:D11,C2:C11,"Priya")`
   Priya's rows: 1500, 800, 600, 1000 → sum `3900`, count `4`,
   average `975`. `F5` should show `975`.
3. `COUNTIFS`/`AVERAGEIFS` follow the same criteria-pair pattern as
   `SUMIFS`, just without a sum range (`COUNTIFS`) or with the average
   range first (`AVERAGEIFS`).

## 4. Nested IF and IFS

1. In `E2`, classify deal size with nested `IF`:
   `=IF(D2>=1200,"Large",IF(D2>=800,"Medium","Small"))`
2. Row 2 (1200) → `Large`. Copy down: row 3 (900) → `Medium`; row 5 (800)
   → `Medium`; row 8 (700) → `Small`. Trace this by hand against the raw
   `Amount` values to confirm each bucket is right.
3. The modern alternative, `IFS`, avoids nesting parentheses:
   `=IFS(D2>=1200,"Large",D2>=800,"Medium",TRUE,"Small")`
   The final `TRUE` acts as a catch-all "else" — without it, `IFS` returns
   `#N/A` when no condition matches.

## 5. IFERROR for resilient formulas

1. In `G2`, deliberately divide by a category count that might be zero:
   `=IFERROR(D2/COUNTIFS(A2:A11,"North"),"No data")`
2. There's no "North" region in the dataset, so `COUNTIFS(...)` returns
   `0`, `D2/0` raises `#DIV/0!`, and `IFERROR` catches it and returns
   `"No data"` instead of showing the raw error.
3. `IFERROR(formula, value_if_error)` wraps any formula — it's the
   standard way to keep a dashboard free of `#DIV/0!`, `#N/A`, or `#VALUE!`
   cells that would otherwise break downstream `SUM`s.

## 6. Combining conditions with AND / OR

1. In `H2`: `=IF(AND(A2="East",D2>1000),"Priority East","Standard")`
2. Row 2: East and 1200 → both true → `Priority East`. Row 3: East but 900
   (not `>1000`) → `Standard`.
3. `OR` works the same way but needs only one condition true:
   `=IF(OR(A2="West",D2>1400),"Flag","-")`
   Row 4 (East, 1500) → `OR` is true because `1500>1400` → `Flag`, even
   though region is East.

## Cheat sheet

| Function | Syntax |
|---|---|
| `SUMIFS` | `=SUMIFS(sum_range,crit_rng1,crit1,[crit_rng2,crit2,...])` |
| `COUNTIFS` | `=COUNTIFS(crit_rng1,crit1,[crit_rng2,crit2,...])` |
| `AVERAGEIFS` | `=AVERAGEIFS(avg_range,crit_rng1,crit1,...)` |
| `IFS` | `=IFS(cond1,val1,cond2,val2,...,TRUE,default)` |
| `IFERROR` | `=IFERROR(formula,value_if_error)` |
| `AND`/`OR` | `=IF(AND(c1,c2),t,f)` / `=IF(OR(c1,c2),t,f)` |

## Exercise

Using the `Sales` table above, compute: (1) total Amount for West + Feb
(expected `1100`), (2) count of deals by Sam over 900 (expected `2`: 1100
and 1300), (3) average Amount for East (expected `980` — rows 1200, 900,
1500, 700, 600 sum to `4900` over 5 rows), and (4) a nested-`IF` column
labeling every row Large/Medium/Small using the thresholds in Section 4.
Verify each result by manually listing the matching rows before trusting
the formula.
