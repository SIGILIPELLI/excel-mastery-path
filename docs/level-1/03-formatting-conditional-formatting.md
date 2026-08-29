# 03 · Formatting & Conditional Formatting

Formatting changes how a value **displays** without changing the value
itself — `1200` formatted as currency still equals `1200` in every formula
that references it. This module covers number formats, basic cell styling,
and **conditional formatting**, which changes a cell's appearance
automatically based on its value — the fastest way to make a budget's
problem rows visually obvious.

## 1. Number formats

1. Select `B2:C6` (the Budgeted and Actual columns) in `budget-tracker.xlsx`
   and open **Home > Number Format** dropdown (shows "General" by default).
2. Choose **Currency** — values display as `$1,200.00`. The underlying
   value stored in the cell is still `1200`; only the display changed,
   confirmed by the Formula Bar still showing the plain number.
3. **Ctrl/⌘+Shift+1** applies Number format with two decimals and a
   thousands separator (`1,200.00`, no `$`); **Ctrl/⌘+Shift+4** applies
   Currency; **Ctrl/⌘+Shift+5** applies Percentage — the same shortcuts
   used for the `% of Budget` column in Module 2.
4. **Increase Decimal** / **Decrease Decimal** buttons (Home tab, Number
   group) add or remove decimal places one at a time without changing the
   underlying value's precision — a formula referencing that cell still
   uses its full stored precision even if the display is rounded.
5. **Home > Number Format > More Number Formats** opens the full dialog,
   including **Custom** formats — for example `$#,##0;[Red]-$#,##0` shows
   positive currency values normally and negative ones in red with a
   leading minus, still as plain currency.

## 2. Basic cell styling

1. **Bold** (Ctrl/⌘+B), **Italic** (Ctrl/⌘+I), and **Underline**
   (Ctrl/⌘+U) apply from the Home tab or their shortcuts. Bold the header
   row (`A1:D1`) to visually separate it from the data.
2. **Fill Color** (paint-bucket icon) and **Font Color** (A with a color
   bar) are both in the Home > Font group. A light fill on the header row
   (e.g. light gray or light blue) reinforces that row 1 is structurally
   different from the data rows.
3. **Borders** (Home > Font group, grid icon) add lines between cells —
   **All Borders** for a full grid, **Bottom Border** for just a header
   underline. Select `A1:D6` and apply **All Borders** to box in the whole
   table.
4. **Merge & Center** (Home > Alignment group) combines multiple cells into
   one and centers the content — useful for a title spanning several
   columns, but avoid merging cells that hold data used in formulas,
   since a merged range behaves as a single cell for referencing purposes
   and can break range-based formulas built around it.
5. **Format Painter** (paintbrush icon, Home tab) copies all formatting
   (not values) from one cell/range to another — select the formatted
   source cell, click Format Painter, then click or drag over the target
   to apply the same look instantly.

## 3. Conditional formatting: highlight over-budget rows

Conditional formatting applies a format automatically, re-evaluating live
as values change — unlike manual formatting, which stays fixed even if the
underlying data changes.

1. Select the `Difference` column, `D2:D6`, built in Module 2.
2. **Home > Conditional Formatting > Highlight Cells Rules > Greater Than**.
3. Type `0` in the value box and choose a format (e.g. **Light Red Fill
   with Dark Red Text**), then **OK**.
4. Rows where Actual exceeds Budgeted (Groceries `55`, Entertainment `75`)
   highlight red; rows at or under budget (Rent `0`, Transport `-20`,
   Savings `0`) do not — the rule only matches strictly greater than zero.
5. Add a second rule for the opposite case: select `D2:D6` again,
   **Conditional Formatting > Highlight Cells Rules > Less Than**, value
   `0`, format **Green Fill with Dark Green Text**. Transport (`-20`, under
   budget) now highlights green.

## 4. Data bars, color scales, and icon sets

1. Select `C2:C6` (Actual). **Conditional Formatting > Data Bars** overlays
   a proportional bar inside each cell — the longest bar lands on Rent
   (`1200`, the largest Actual value), the shortest on Transport (`130`,
   the smallest).
2. **Conditional Formatting > Color Scales** shades each cell along a
   gradient (e.g. red for low values, green for high) based on its
   relative position in the selected range — useful for spotting the
   highest and lowest values in a large range at a glance, without reading
   every number.
3. **Conditional Formatting > Icon Sets** adds a small icon (arrows, traffic
   lights, stars) per cell based on thresholds you set — a 3-arrow set on
   the `Difference` column could show a red down-arrow for negative
   (under-budget), a yellow flat-arrow near zero, and a green up-arrow for
   positive (over-budget), configurable via **Manage Rules > Edit Rule**.
4. **Conditional Formatting > Manage Rules** lists every rule active on the
   worksheet, in the order they're evaluated — use **Edit Rule** to adjust
   a threshold, or the up/down arrows to reorder rules when more than one
   could apply to the same cell (the first matching rule, top to bottom,
   wins unless "Stop If True" is unchecked further down the list).
5. **Conditional Formatting > Clear Rules > Clear Rules from Selected
   Cells** (or **from Entire Sheet**) removes formatting rules without
   touching the underlying values.

## Cheat sheet

| Format | Shortcut / Location |
|---|---|
| Bold / Italic / Underline | Ctrl/⌘+B / +I / +U |
| Number format (2 decimals) | Ctrl/⌘+Shift+1 |
| Currency format | Ctrl/⌘+Shift+4 |
| Percentage format | Ctrl/⌘+Shift+5 |
| Format Painter | Home > Clipboard group |
| Highlight Cells Rules (Greater/Less Than) | Home > Conditional Formatting |
| Data Bars / Color Scales / Icon Sets | Home > Conditional Formatting |
| Manage/edit existing rules | Conditional Formatting > Manage Rules |
| Clear conditional formatting | Conditional Formatting > Clear Rules |

## Exercise

In `budget-tracker.xlsx`, format `B2:C6` as Currency, bold and fill the
header row, and add borders around the whole table (`A1:D6`). On the
`Difference` column, add two Highlight Cells rules: red fill for values
greater than 0 (over budget), green fill for values less than 0 (under
budget). Then add a Data Bars rule to the Actual column and confirm the
longest bar lands on Rent and the shortest on Transport.
