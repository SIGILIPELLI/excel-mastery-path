# 02 · PivotTables Deep Dive

Level 1 built a basic PivotTable (drag a field to Rows, one to Values).
This module goes deeper: calculated fields, grouping, multiple value
fields with different summary types, and slicers for interactive
filtering — using the same `Sales` dataset from Module 1.

## 1. Worked dataset (reused)

Same table as Module 1, on sheet `Sales`, `A1:D11` (Region, Month, Rep,
Amount). Manual totals, verified by hand-adding the rows, for reference
throughout this module: East total = `1500+900+1500... ` — recomputed
precisely below per pivot.

Full data:
East/Jan/Alan/1200, East/Feb/Alan/900, East/Mar/Priya/1500,
West/Jan/Priya/800, West/Feb/Sam/1100, West/Mar/Sam/1300,
East/Mar/Alan/700, West/Jan/Sam/950, East/Jan/Priya/600,
West/Mar/Priya/1000.

Manual region totals: East = `1200+900+1500+700+600 = 4900`.
West = `800+1100+1300+950+1000 = 5150`. Grand total = `10050`.

## 2. Building the PivotTable

1. Select `A1:D11` > **Insert > PivotTable** > New Worksheet.
2. Drag `Region` to **Rows**, `Amount` to **Values** (defaults to Sum).
3. The PivotTable should show East `4900` and West `5150`, matching the
   hand-calculated totals in Section 1 exactly — this cross-check is the
   quickest way to catch a mis-selected data range.

## 3. Multiple value fields, different summary types

1. Drag `Amount` to **Values** a second time (now two Amount columns).
2. Right-click the second one > **Value Field Settings** > change
   **Sum** to **Average**. Rename it "Avg Amount" in the same dialog.
3. For East: `4900 / 5 rows = 980`. For West: `5150 / 5 rows = 1030`.
   Confirm the PivotTable's "Avg Amount" column shows `980` and `1030`.
4. Add a third value field set to **Count** — should show `5` for both
   regions (5 transactions each), confirming row counts by hand.

## 4. Grouping rows: Month as a nested row

1. Drag `Month` to **Rows**, below `Region` (Rows box now has Region then
   Month, in that order — order matters, it controls nesting).
2. Expand `East` — it should break into Jan (`1200+600=1800`), Feb
   (`900`), Mar (`1500+700=2200`). Sum: `1800+900+2200=4900`, matching the
   East grand total from Section 2.

## 5. Calculated field

1. With the PivotTable selected: **PivotTable Analyze > Fields, Items &
   Sets > Calculated Field**.
2. Name: `Bonus`, Formula: `=Amount*0.1`.
3. Add `Bonus` to Values. For East (`4900`), Bonus should show `490.0`
   (`4900 * 0.1`). This is computed on the pivot's aggregated Amount, not
   per source row, so it is safest with `SUM`-based pivots like this one.

## 6. Slicers for interactive filtering

1. **PivotTable Analyze > Insert Slicer** > check `Rep`.
2. Click "Alan" on the slicer — the pivot filters to Alan's rows only:
   Amount total should be `1200+900+700=2800`. Click "Priya" instead —
   total should be `1500+800+600+1000=3900`. These two numbers plus Sam's
   remaining `1100+1300+950=3350` should sum back to the grand total
   `2800+3900+3350=10050`, matching Section 1 — a good sanity check that
   the slicer isn't silently excluding rows.

## Cheat sheet

| Action | Where |
|---|---|
| Add a 2nd value field | Drag field to Values again, change summary type |
| Nest rows | Drag second field below the first in Rows |
| Custom formula per pivot | PivotTable Analyze > Calculated Field |
| Interactive filter widget | PivotTable Analyze > Insert Slicer |
| Change Sum→Average/Count | Right-click value > Value Field Settings |

## Exercise

Build the PivotTable above with Region and Month nested rows, three value
fields (Sum, Average, Count of Amount), a `Bonus` calculated field at 10%,
and a `Rep` slicer. Click through each Rep on the slicer and confirm the
three totals (Alan `2800`, Priya `3900`, Sam `3350`) sum to the grand
total `10050`.
