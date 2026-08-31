# 08 · Advanced Dashboard Design Principles

This module covers turning raw calculations into a decision-ready
dashboard: layout, chart choice, dynamic titles, and the interactive
controls (slicers, form controls) that let a viewer filter without
touching a formula.

## 1. Worked dataset — regional sales

Build this on a sheet named `Sales`, `A1:C13`:

| | A | B | C |
|---|---|---|---|
| **1** | Region | Month | Revenue |
| **2** | North | Jan | 12000 |
| **3** | North | Feb | 13500 |
| **4** | North | Mar | 11800 |
| **5** | South | Jan | 9000 |
| **6** | South | Feb | 9600 |
| **7** | South | Mar | 10100 |
| **8** | East | Jan | 15000 |
| **9** | East | Feb | 14200 |
| **10** | East | Mar | 16000 |
| **11** | West | Jan | 8000 |
| **12** | West | Feb | 8300 |
| **13** | West | Mar | 8900 |

## 2. A dynamic title that reacts to a selection

1. Put a region selector in `E1` (type `North` for now — later this
   becomes a slicer/dropdown-driven cell).
2. In `E2`: `=SUMIFS(C2:C13,A2:A13,E1)`
   Manual check for `E1="North"`: `12000+13500+11800=37300`.
3. Chart title formula (click the chart title, then in the formula bar
   type `=Sales!$E$3`), where `E3`:
   `="Total Revenue — "&E1&": "&TEXT(E2,"$#,##0")`
   For `E1="North"`, `E3` reads `Total Revenue — North: $37,300`. This
   is the core dashboard trick: the title is a formula, not static
   text, so it updates the instant the filter cell changes.

## 3. Choosing the right chart

| Question being answered | Chart type | Why |
|---|---|---|
| How does one metric trend over time? | Line chart | Emphasizes direction/slope |
| How do categories compare at one point in time? | Bar/column chart | Easiest length comparison |
| What is the part-to-whole breakdown? | Stacked bar (avoid pie for >4 slices) | Pie charts are hard to compare past 3-4 slices |
| Is there a relationship between two numeric variables? | Scatter | Shows correlation/clusters |

Avoid 3-D charts and dual unrelated axes — both distort the visual
comparison a dashboard is supposed to make effortless.

## 4. A KPI summary row with conditional formatting

1. `E5:E8` — build one KPI per region using `SUMIFS`:
   `E5`: `=SUMIFS($C$2:$C$13,$A$2:$A$13,"North")` → `37300`
   `E6`: `=SUMIFS($C$2:$C$13,$A$2:$A$13,"South")` → `28700`
   `E7`: `=SUMIFS($C$2:$C$13,$A$2:$A$13,"East")` → `45200`
   `E8`: `=SUMIFS($C$2:$C$13,$A$2:$A$13,"West")` → `25200`
2. Select `E5:E8` → **Home → Conditional Formatting → Data Bars** to
   turn the KPI list into an instant visual ranking — East (45,200) gets
   the longest bar, West (25,200) the shortest, with no chart needed.
3. Add a rule: **Highlight Cells Rules → Less Than** `30000`,
   formatted red — this flags South and West as below-threshold
   regions at a glance.

## 5. Slicers for interactive filtering

1. Convert `A1:C13` to a Table (**Ctrl+T**).
2. Build a PivotTable from the table with `Region` in Rows and
   `Revenue` (Sum) in Values.
3. **PivotTable Analyze → Insert Slicer** → check `Region`. Clicking a
   region button in the slicer now filters the PivotTable (and any
   PivotChart built from it) instantly — this is the standard way to
   give a dashboard viewer point-and-click filtering without exposing
   any formulas.

## 6. Layout principles

- **F-pattern / Z-pattern**: put the single most important KPI top-left
  — that is where the eye lands first.
- **One screen, no scrolling**: a dashboard that requires scrolling to
  see all KPIs has failed its job; consolidate or use tabs.
- **Consistent color meaning**: pick one color for "good/on target" and
  one for "bad/off target" and use them identically in every chart and
  conditional format on the sheet — inconsistent color use is the most
  common dashboard design mistake.
- **Whitespace**: group related KPIs with spacing/borders rather than
  packing every cell — a dense sheet is a hard-to-read sheet.

## Cheat sheet

| Technique | Purpose |
|---|---|
| Formula-driven chart title | Title updates automatically with the filter |
| Data Bars | Instant visual ranking without a chart |
| Slicer | Point-and-click filter for Tables/PivotTables |
| SUMIFS-based KPI cells | Feed both text titles and conditional formats |

## Exercise

Add a `Q1 Target` value of `30000` per region in a new column, then
build a conditional format on the KPI cells (`E5:E8`) that turns the
cell green when revenue meets or exceeds its target and red otherwise.
Which regions are red with the dataset above?
