# 10 · Project — Interactive Sales Dashboard

This capstone combines every Level 2 skill — SUMIFS, PivotTables,
combo charts, data validation dropdowns, and a Form Control — into one
interactive one-page dashboard driven by a single dropdown selector.

## 1. Worked dataset

Build this table on a sheet named `SalesData`, `A1:D13`, and convert it
to an Excel Table named `tblSales` (Insert → Table):

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Region | Month | Rep | Amount |
| **2** | East | Jan | Alan | 1200 |
| **3** | East | Feb | Alan | 1400 |
| **4** | East | Mar | Priya | 1600 |
| **5** | West | Jan | Priya | 900 |
| **6** | West | Feb | Sam | 1100 |
| **7** | West | Mar | Sam | 1300 |
| **8** | East | Jan | Priya | 700 |
| **9** | West | Jan | Sam | 950 |
| **10** | East | Feb | Priya | 1050 |
| **11** | West | Feb | Priya | 800 |
| **12** | East | Mar | Alan | 1250 |
| **13** | West | Mar | Alan | 1000 |

## 2. Region selector (dropdown)

1. On a new sheet `Dashboard`, cell `B1`: Data Validation → List →
   Source `East,West`. Set the default to `East`.
2. In `B2`, compute the selected region's total with SUMIFS:
   `=SUMIFS(tblSales[Amount],tblSales[Region],B1)`
3. Manual check with `B1="East"`: East rows are 1200, 1400, 1600, 700,
   1050, 1250 → sum `7200`. With `B1="West"`: 900, 1100, 1300, 950,
   800, 1000 → sum `6050`.

## 3. PivotTable driven by the same selector

1. Insert a PivotTable from `tblSales` onto `Dashboard` at `D1`. Rows:
   `Month`. Values: `Sum of Amount`. Filter: `Region`.
2. Right-click the PivotTable → **PivotTable Options** isn't needed
   here; instead connect the Region filter to the `B1` dropdown using
   a **Slicer**: PivotTable Analyze → Insert Slicer → `Region`. Click
   the Slicer's East/West buttons to filter — this is more visual than
   the page filter dropdown and updates instantly.
3. Manual check with East selected: Jan `1200+700=1900`, Feb
   `1400+1050=2450`, Mar `1600+1250=2850`. These three should sum to
   `7200`, matching Section 2's SUMIFS total — cross-verifying the
   PivotTable against the formula is exactly the kind of consistency
   check a real dashboard needs.

## 4. Combo chart from the PivotTable

1. Select the PivotTable's Month/Amount data → Insert → PivotChart (or
   a regular combo chart referencing the Pivot's output range).
2. Set Amount as clustered columns. Add a second series — a 3-period
   moving flat reference line at the East average (`7200/3=2400`) by
   adding a helper column with the constant `2400` repeated for each
   month, plotted as a line — so viewers can see at a glance which
   months over/underperform the region's own average. Manual check:
   Jan (1900) is below 2400, Feb (2450) is above, Mar (2850) is above.

## 5. Keeping the dropdown, Slicer, and KPI tile consistent

1. Slicer → right-click → **Report Connections** → make sure it is
   connected to every PivotTable on the dashboard, so one click on the
   Slicer filters all Pivot-based visuals (table and chart) together.
2. The `B2` SUMIFS tile is a plain formula, not Pivot-based, so it only
   reacts to the `B1` dropdown, not the Slicer. To keep the numeric
   tile and the chart in sync, treat `B1` as the single source of
   truth: after changing the Slicer, also set `B1` to match (or skip
   the Slicer entirely and give the PivotTable a **Region** Report
   Filter set to `=B1` isn't directly supported, so instead filter the
   PivotTable manually to match `B1` each time it changes). For a
   fully single-control dashboard, prefer the `B1` dropdown driving
   both the SUMIFS tile and a manually-set PivotTable filter, and treat
   the Slicer as an optional visual convenience rather than the
   dashboard's primary control.

## Cheat sheet

| Dashboard piece | Built with |
|---|---|
| Region KPI tile | `SUMIFS` referencing the dropdown cell |
| Month breakdown | PivotTable with Region as filter/Slicer |
| Trend visualization | PivotChart or combo chart off the Pivot |
| User control | Data Validation dropdown +/or Slicer |

## Exercise

Add a second KPI tile below `B2`: average deal size for the selected
region, `=AVERAGEIFS(tblSales[Amount],tblSales[Region],B1)`. With
`B1="West"`, manually verify: West amounts 900, 1100, 1300, 950, 800,
1000 sum to `6050` over 6 deals → average `1008.33`. Confirm the
formula matches.
