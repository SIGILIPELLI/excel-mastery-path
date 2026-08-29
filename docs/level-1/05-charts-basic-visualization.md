# 05 · Charts & Basic Visualization

A chart turns a table of numbers into a shape the eye can compare
instantly. This module builds a column chart and a pie chart from the
budget Table (Module 4), and covers the handful of chart types and
formatting options you'll reach for constantly.

## 1. Building a column chart

1. In `budget-tracker.xlsx`, select the `Category`, `Budgeted`, and
   `Actual` columns of the Table — since they aren't adjacent to each
   other on their own but `Category` is column A and `Budgeted`/`Actual`
   are B and C, selecting `A1:C7` (through the Subscriptions row added in
   Module 4) works directly since all three are contiguous.
2. **Insert > Charts > Insert Column or Bar Chart > Clustered Column**.
   Excel drops in a chart with one cluster of two bars (Budgeted, Actual)
   per category.
3. The tallest bars are both Rent's (`1200` Budgeted and Actual, matching
   exactly); Groceries and Entertainment visibly show Actual taller than
   Budgeted — the same over-budget categories flagged red by conditional
   formatting in Module 3.
4. Click the chart to reveal three floating buttons at its top-right:
   **Chart Elements** (+), **Chart Styles** (paintbrush), and **Chart
   Filters** (funnel) — these cover most day-to-day chart adjustments
   without opening a separate dialog.

## 2. Chart Elements: titles, labels, gridlines

1. Click the chart, then the **Chart Elements (+)** button.
2. Check **Chart Title** — a placeholder title appears on the chart; click
   it and type a real title, e.g. `Budget vs. Actual`.
3. Check **Data Labels** — the exact value from each bar now displays above
   it (`1200`, `1200`, `400`, `455`, etc.), removing any need to
   cross-reference the underlying table while reading the chart.
4. Uncheck **Gridlines** if the chart looks cluttered — gridlines help
   estimate values on a bar chart but add visual noise once Data Labels
   are showing the exact numbers directly.
5. Check **Legend** (usually on by default) to confirm which color
   represents Budgeted vs. Actual — click the legend and drag or use the
   Chart Elements arrow to reposition it (e.g. bottom instead of right).

## 3. Pie chart: share of total budget

1. Select just `Category` and `Budgeted` — for non-adjacent columns, select
   `A1:A7`, then hold **Ctrl/⌘** and select `B1:B7`.
2. **Insert > Charts > Insert Pie or Doughnut Chart > Pie**.
3. Each slice's size corresponds to that category's share of the total
   Budgeted amount (`2170` total from Module 4) — Rent's slice is by far
   the largest since `1200` is over half the total.
4. Add **Data Labels** (Chart Elements) and set them to show
   **Percentage** instead of the raw value: right-click any slice's data
   label, **Format Data Labels**, check **Percentage** and uncheck
   **Value**. Rent's slice should now read roughly `55%` (`1200 / 2170`),
   matching the `% of Budget` figure computed with a formula in Module 2 —
   two different ways (a formula, and a chart's own calculation) arriving
   at the same answer is a useful sanity check.
5. A pie chart only makes sense when the parts genuinely sum to a
   meaningful whole (like a total budget) — it's the wrong choice for
   comparing Budgeted vs. Actual side by side, since those aren't parts of
   one whole; that comparison is what the column chart (Section 1) is for.

## 4. Choosing a chart type

| Data shape | Best chart type |
|---|---|
| Comparing values across categories | Clustered column or bar |
| Parts of one whole (must sum to 100%) | Pie or doughnut |
| Trend over time (dates, months) | Line |
| Relationship between two numeric variables | Scatter (X Y) |
| Distribution/spread of one variable | Histogram |

1. **Insert > Recommended Charts** analyzes the selected data and suggests
   chart types likely to fit — useful when unsure, though understanding
   the table above means you rarely need to rely on the recommendation.
2. Switch an existing chart's type without rebuilding it: click the chart,
   **Chart Design > Change Chart Type**, pick a different type from the
   same dialog used to create it originally.

## 5. Moving and resizing charts

1. Click a chart's border (not its interior) and drag to reposition it
   anywhere on the sheet, or onto a different sheet entirely by cutting
   (Ctrl/⌘+X) and pasting (Ctrl/⌘+V) on the target sheet.
2. Drag a corner handle to resize proportionally; drag an edge handle to
   resize only that dimension (which can distort a pie chart into an
   oval — usually undesirable).
3. **Chart Design > Move Chart** relocates a chart to its own dedicated
   **Chart Sheet** (a full-sheet tab containing only the chart) rather than
   floating over a worksheet — useful for a chart meant to be the primary
   view of a workbook rather than a supplement to a data sheet.

## Cheat sheet

| Action | Location |
|---|---|
| Insert column/bar chart | Insert > Charts > Column or Bar |
| Insert pie chart | Insert > Charts > Pie or Doughnut |
| Add/remove title, labels, gridlines | Chart Elements (+) button |
| Change colors/style | Chart Styles (paintbrush) button |
| Filter which series/categories show | Chart Filters (funnel) button |
| Change an existing chart's type | Chart Design > Change Chart Type |
| Move chart to its own sheet | Chart Design > Move Chart |

## Exercise

From the Budget Table, build a clustered column chart comparing Budgeted
vs. Actual across all categories, with a title and data labels turned on.
Then build a separate pie chart of just the Budgeted column with
percentage labels, and confirm Rent's slice reads close to `55%`. Try
Insert > Recommended Charts on the same selection and compare what Excel
suggests to the chart types you chose yourself.
