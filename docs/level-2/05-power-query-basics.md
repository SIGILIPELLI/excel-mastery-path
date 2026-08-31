# 05 · Power Query Basics

Power Query (Data → Get & Transform) is Excel's tool for importing,
cleaning, and reshaping data through a recorded, repeatable set of
steps — instead of manual copy/paste cleanup you'd have to redo every
time the source data refreshes.

## 1. Worked dataset

Build this table on a sheet named `Raw`, `A1:C7`, then select it and
**Insert → Table** (name it `RawData`):

| | A | B | C |
|---|---|---|---|
| **1** | name | region | sales |
| **2** | alan | east | 1200 |
| **3** | PRIYA | East | 900 |
| **4** | sam | WEST | 1500 |
| **5** | Alan | east |  |
| **6** | priya | East | 800 |
| **7** | Sam | west | 1100 |

Note the inconsistent capitalization and one blank sales cell — this
is realistic messy source data.

## 2. Load into Power Query

1. Click any cell inside `RawData`, then Data → From Table/Range. This
   opens the Power Query Editor with the applied step **Source**
   already recorded.
2. Every action taken from here on is recorded as a numbered step in
   the **Applied Steps** pane on the right — this is what makes Power
   Query repeatable: re-running the query on refreshed data replays
   every step automatically.

## 3. Cleaning steps

1. Select the `name` column header → Transform → Format →
   **Capitalize Each Word** (or Lowercase then Capitalize). This
   normalizes `alan`/`Alan`/`PRIYA` inconsistencies to `Alan`, `Priya`.
2. Select the `region` column → Transform → Format → **UPPERCASE**,
   giving `EAST`/`WEST` consistently, or Lowercase for `east`/`west` —
   pick one convention and apply it everywhere so `East` and `east`
   are no longer treated as two different groups downstream.
3. Select the `sales` column → right-click header → **Replace
   Values**, replace blank/null with `0` (or use Transform → Fill →
   Down if a blank should inherit the row above's value — here a
   missing sales figure should become `0`, not inherit a neighbor's
   value, so use Replace Values, `null` → `0`).
4. Verify: after cleaning, row 5 (originally `Alan, east, <blank>`)
   should show `Alan, EAST, 0`.

## 4. Grouping

1. Select the `region` column → Transform → **Group By**. Group by
   `region`, New column name `TotalSales`, Operation `Sum`, Column
   `sales`.
2. Manual check: EAST rows are 1200, 900, 0, 800 → sum `2900`. WEST
   rows are 1500, 1100 → sum `2600`. The grouped table should show two
   rows: `EAST, 2900` and `WEST, 2600`.

## 5. Load back to Excel

1. Home → **Close & Load** (or **Close & Load To…** to choose a new
   sheet or PivotTable). This writes the cleaned/grouped result as a
   new table on a new sheet, linked to the query.
2. Refresh: edit a value back in `RawData` (e.g. change Sam's West
   sale from 1100 to 1300), then Data → **Refresh All**. The query
   re-runs every recorded step automatically and the grouped WEST
   total updates from `2600` to `2800` — this is the core value of
   Power Query over manual cleanup: the transformation logic is saved
   and replayed, not repeated by hand.

## Cheat sheet

| Task | Ribbon path |
|---|---|
| Load a table into Power Query | Data → From Table/Range |
| Fix text casing | Transform → Format → Capitalize/UPPER/lower |
| Replace blanks/errors | Right-click column → Replace Values |
| Aggregate by category | Transform → Group By |
| Send result back to a sheet | Home → Close & Load |
| Re-run after source changes | Data → Refresh All |

## Exercise

Add a `discount` column to `RawData` with values `0.1, 0.05, 0, 0.1,
0.05, 0` for rows 2–7, load it through Power Query, add a custom
column `netSales = sales * (1 - discount)`, then Group By `region`
summing `netSales`. Manually verify EAST: `1200*0.9=1080`,
`900*0.95=855`, `0*1=0`, `800*0.95=760` → sum `2695`.
