# 02 · Advanced DAX & Power Pivot Modeling

This module goes beyond Level 3's DAX basics (Module 04) into time
intelligence, iterator functions, and filter-context manipulation with
`CALCULATE` — the techniques that make a Power Pivot model behave like
a real analytical engine.

## 1. Worked dataset

Load a table `Sales` into Power Pivot, `A1:C13`:

| | A | B | C |
|---|---|---|---|
| **1** | Date | Region | Amount |
| **2** | 2026-01-05 | North | 1000 |
| **3** | 2026-01-20 | South | 800 |
| **4** | 2026-02-03 | North | 1200 |
| **5** | 2026-02-15 | South | 900 |
| **6** | 2026-03-01 | North | 1100 |
| **7** | 2026-03-22 | South | 950 |
| **8** | 2027-01-10 | North | 1300 |
| **9** | 2027-01-25 | South | 1000 |
| **10** | 2027-02-05 | North | 1250 |
| **11** | 2027-02-18 | South | 1050 |
| **12** | 2027-03-02 | North | 1400 |
| **13** | 2027-03-15 | South | 1100 |

Also build a `Calendar` table with a `Date` column spanning
`2026-01-01` through `2027-12-31`, and mark it as the Date Table
(**Table → Date Table → Mark as Date Table**), then relate
`Calendar[Date] → Sales[Date]`.

## 2. Base measure

`Total Sales := SUM(Sales[Amount])`
Grand total: `1000+800+1200+900+1100+950+1300+1000+1250+1050+1400+1100
=13050`.

## 3. Time intelligence — year-over-year

`Sales PY := CALCULATE([Total Sales], SAMEPERIODLASTYEAR(Calendar[Date]))`

For **2027** in the pivot (total `1300+1000+1250+1050+1400+1100=7100`),
`Sales PY` should return the **2026** total:
`1000+800+1200+900+1100+950=5950`.

`YoY % := DIVIDE([Total Sales]-[Sales PY], [Sales PY])`
For 2027: `(7100-5950)/5950 = 1150/5950 ≈ 0.1933`, i.e. **+19.3%**
growth over the prior year.

## 4. CALCULATE — overriding filter context

`North Sales := CALCULATE([Total Sales], Sales[Region]="North")`
Regardless of what region a report row is sliced by, this measure
always recomputes as if the North filter were applied. Grand total for
North across both years:
`1000+1200+1100+1300+1250+1400=7250`.

`North Share % := DIVIDE([North Sales], [Total Sales])`
Overall: `7250/13050 ≈ 0.5556`, i.e. North is **55.6%** of total sales.

## 5. Iterator functions — SUMX

`Sales SUMX Check := SUMX(Sales, Sales[Amount])`
This produces the identical result to `[Total Sales]` (`13050`) here
because there is no per-row calculation needed, but `SUMX` becomes
essential once a measure needs a row-by-row computation before
summing, e.g. if `Sales` had a `Qty` and `UnitPrice` column instead of
a single `Amount`:
`Revenue := SUMX(Sales, Sales[Qty] * Sales[UnitPrice])`
`SUM` cannot do this — it can only total one existing column, while
`SUMX` evaluates an expression per row, then sums the results.

## 6. Ranking with RANKX

`Region Rank := RANKX(ALL(Sales[Region]), [Total Sales])`
With only two regions and North (`7250`) > South (`5800`, i.e.
`13050-7250`), North ranks `1` and South ranks `2`. `ALL(Sales[Region])`
removes the region filter just for the ranking calculation, so every
region is compared against the same full candidate set rather than
against itself.

## Cheat sheet

| Function | Purpose |
|---|---|
| `SAMEPERIODLASTYEAR(dates)` | Shift the current filter context back one year |
| `CALCULATE(expr, filter)` | Recompute an expression under a modified filter context |
| `SUMX(table, expr)` | Row-by-row expression, then sum — for calculations `SUM` alone can't do |
| `RANKX(table, expr)` | Rank a value against a set of candidates |
| `DIVIDE(num, denom)` | Division that safely returns blank/0 instead of `#DIV/0!` |

## Exercise

Add a `South PY` and `South YoY %` pair of measures mirroring Section
3. South's 2026 total is `800+900+950=2650`; its 2027 total is
`1000+1050+1100=3150`. Confirm South's YoY growth,
`(3150-2650)/2650 ≈ 0.1887` (**+18.9%**), is slightly smaller in
percentage terms than North's 19.3%, even though both regions grew in
absolute dollars.
