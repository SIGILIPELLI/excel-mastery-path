# 07 · What-If Analysis

What-If Analysis tools (Goal Seek, Data Tables, Scenario Manager) let
you ask "what input produces this output?" or "how does the result
change across a range of inputs?" without manually re-typing numbers
over and over.

## 1. Worked dataset — a simple loan model

Build this on a sheet named `Loan`, `A1:B5`:

| | A | B |
|---|---|---|
| **1** | Principal | 20000 |
| **2** | Annual Rate | 0.06 |
| **3** | Years | 5 |
| **4** | Monthly Payment | `=-PMT(B2/12,B3*12,B1)` |
| **5** | | |

`PMT(rate,nper,pv)` returns a negative number (cash outflow), so the
formula negates it. Manual check: `B2/12 = 0.005`, `B3*12 = 60`
periods, `PV = 20000`. Excel's `PMT` computes this via the standard
annuity formula; trust the built-in result here (~`386.66`) since it
matches the closed-form annuity payment formula
`P = Pv·r / (1-(1+r)^-n)` = `20000×0.005 / (1-1.005^-60)` ≈ `386.66`.

## 2. Goal Seek

1. Suppose the target monthly payment must be exactly `350`. Data →
   What-If Analysis → **Goal Seek**. Set cell `B4`, To value `-350`
   (remember `B4` is already negated to positive via the formula, so
   if `B4` shows a positive `386.66`, target `350` directly), By
   changing cell `B1` (solve for the Principal that gives a 350
   payment).
2. Goal Seek iterates numerically until `B4` ≈ `350`; the resulting
   `B1` will be roughly `18122` (a smaller loan produces a smaller
   payment at the same rate/term). Verify by re-plugging: with
   `B1=18122`, `PMT(0.005,60,18122)` ≈ `350.0`, confirming Goal Seek's
   answer.
3. Undo (Ctrl+Z) to restore `B1` to `20000` after testing.

## 3. Data Table (one-variable)

1. In `D1:D6`, list candidate annual rates: `0.04, 0.05, 0.06, 0.07,
   0.08`. In `E1`, reference the payment formula: `=B4`.
2. Select `D1:E6`, Data → What-If Analysis → **Data Table**, Column
   input cell `B2` (since rates vary down a column). Excel recomputes
   `B4` for each rate and fills `E2:E6`.
3. Manual spot check at rate `0.08`: `PMT(0.08/12,60,20000)` — a higher
   rate must produce a higher payment than the baseline `386.66` at
   `0.06`; expect roughly `405.53`. Confirm the Data Table's value at
   that row is higher than the `0.06` row's `386.66`, and that payment
   increases monotonically as rate increases from `0.04` to `0.08` —
   this monotonic direction is the sanity check, even without
   recomputing every value by hand.

## 4. Scenario Manager

1. Data → What-If Analysis → **Scenario Manager** → Add.
2. Scenario name `Best Case`: changing cells `B2,B3` → values `0.04,
   3`. Add another, `Worst Case`: `B2,B3` → `0.08, 7`.
3. Click **Summary** to generate a Scenario Summary report — a new
   sheet listing `B4` (Monthly Payment) under each scenario side by
   side. Best Case (lower rate, shorter term) should show the highest
   monthly payment (paying off faster costs more per month even at a
   lower rate) — verify this direction makes sense: `Years=3` means
   `nper=36`, a much shorter payoff window that raises the payment
   despite the lower `0.04` rate, versus `Years=7` (`nper=84`) in
   Worst Case which spreads payments thinner despite the higher rate.
4. Switch back to the base scenario (or re-enter `0.06` and `5`
   manually) before moving on.

## Cheat sheet

| Tool | Use when |
|---|---|
| Goal Seek | You know the desired output, need to find the input |
| Data Table (1 or 2 variable) | You want the output across a range of one or two inputs |
| Scenario Manager | You want to save and compare a handful of named input sets |

## Exercise

Using the `Loan` model, build a two-variable Data Table with Rate
(`0.04`–`0.08` down rows) and Years (`3,5,7` across columns) both
driving `B4`, using Column input cell `B2` and Row input cell `B3`.
Confirm the corner combining the lowest rate (`0.04`) and longest term
(`7` years) gives the lowest payment in the whole table — a lower rate
and longer payoff period should both individually reduce the monthly
payment, so their combination is the table's minimum.
