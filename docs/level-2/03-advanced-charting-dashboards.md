# 03 · Advanced Charting & Dashboards

Level 1 covered basic column/line/pie charts. This module covers
combination charts, secondary axes, sparklines, and the layout
principles that turn a handful of charts into a coherent one-page
dashboard.

## 1. Worked dataset

Build this table on a sheet named `Perf`, `A1:D7`:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | Month | Revenue | Units | Margin% |
| **2** | Jan | 12000 | 240 | 0.32 |
| **3** | Feb | 13500 | 260 | 0.30 |
| **4** | Mar | 11000 | 210 | 0.35 |
| **5** | Apr | 15500 | 300 | 0.29 |
| **6** | May | 16200 | 305 | 0.31 |
| **7** | Jun | 14800 | 275 | 0.33 |

## 2. Combination chart with a secondary axis

1. Select `A1:B7`, then hold **Ctrl** and also select `D1:D7`
   (Revenue and Margin%, skipping Units).
2. Insert → Charts → Combo Chart → **Create Custom Combo Chart**.
3. Set `Revenue` to **Clustered Column** on the Primary Axis, and
   `Margin%` to **Line** with **Secondary Axis** checked. Revenue runs
   in the thousands while Margin% is a fraction near 0.30 — on one
   axis the line would be invisible flat, so the secondary axis is
   required whenever two series differ in scale by an order of
   magnitude or more.
4. Verify visually: March has the lowest Revenue (11000) but the
   highest Margin% (0.35) — the column should dip while the line
   peaks at the same point, confirming the two axes are independent.

## 3. Sparklines

1. In `F1`, Insert → Sparklines → **Line**. Data Range `B2:B7`
   (Revenue), Location Range `F1`. This draws a small in-cell trend
   line summarizing the whole Revenue column.
2. Repeat for `F2` with Data Range `C2:C7` (Units), and `F3` with Data
   Range `D2:D7` (Margin%). Now three trend lines — Revenue, Units,
   Margin% — sit stacked for a quick visual scan.
3. Manually confirm the Revenue sparkline's shape: values rise
   12000→13500, dip to 11000, then climb to a peak of 16200 in May
   before easing to 14800 — the sparkline's line should show exactly
   that up-up-down-up-up-down zigzag.
4. This stacked-sparkline pattern (one sparkline per KPI, sharing a
   column) is the layout used in KPI dashboard headers, where a viewer
   scans trend shape without needing a full chart per metric.

## 4. Dashboard layout principles

1. **Grid alignment** — snap chart edges to cell boundaries (hold
   **Alt** while dragging) so all dashboard tiles line up into rows
   and columns, not a scattered collage.
2. **One message per chart** — a combo chart answering "did margin
   hold up as revenue grew" is stronger than one chart trying to show
   Revenue, Units, and Margin% all as equal-weight columns.
3. **Consistent color coding** — pick one color for Revenue everywhere
   in the workbook (e.g. blue) and reuse it in every chart; a viewer
   scanning multiple tiles shouldn't have to re-learn the legend each
   time.
4. **Freeze the top row** (View → Freeze Panes → Freeze Top Row) above
   dashboard tiles that reference a scrolling table below them, so
   titles stay visible while a user scrolls the raw data.

## 5. Chart formulas as titles

1. Link a chart title to a cell so it updates automatically: click the
   chart title, type `=Perf!$G$1` in the formula bar, and put a
   summary formula in `G1`, e.g.
   `="Revenue Trend — Peak "&TEXT(MAX(B2:B7),"$#,##0")`
2. Manual check: `MAX(B2:B7)` over `12000,13500,11000,15500,16200,14800`
   is `16200` (May). `TEXT(16200,"$#,##0")` formats it as `$16,200`, so
   the title reads `Revenue Trend — Peak $16,200`.

## Cheat sheet

| Task | Steps |
|---|---|
| Combo chart | Select series → Insert → Combo Chart → set type + axis per series |
| Secondary axis | Chart element → Format Data Series → Secondary Axis |
| Sparkline | Insert → Sparklines → Line/Column/Win-Loss → set Data Range |
| Linked chart title | Click title → type `=Sheet!$Cell` in formula bar |

## Exercise

Using the `Perf` table, build a combo chart with Revenue as columns
(primary axis) and Units as a line (secondary axis), then confirm by
eye that April (highest Revenue, 15500, and highest Units, 300) shows
both series peaking together — unlike Margin%, Revenue and Units move
together here, so decide whether a secondary axis is even necessary
for that pairing (it isn't, since both are large numbers on a similar
scale) and re-plot them on one shared axis instead.
