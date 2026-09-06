# 03 · Power Pivot & the Data Model

Power Pivot lets multiple tables live in one workbook's **Data Model**,
related to each other like a small database, so a single PivotTable
can pull fields from several tables at once without VLOOKUP glue
columns.

## 1. Worked dataset

`Sales` table (`A1:C6`), name it `Sales`:

| | A | B | C |
|---|---|---|---|
| **1** | ProductID | Qty | Rep |
| **2** | P1 | 10 | Alan |
| **3** | P2 | 5 | Priya |
| **4** | P1 | 8 | Sam |
| **5** | P3 | 12 | Alan |
| **6** | P2 | 6 | Priya |

`Products` table (`A1:B4`), name it `Products`:

| | A | B |
|---|---|---|
| **1** | ProductID | Price |
| **2** | P1 | 20 |
| **3** | P2 | 35 |
| **4** | P3 | 15 |

## 2. Adding tables to the Data Model

1. Click inside `Sales` → Power Pivot tab → **Add to Data Model**
   (this also converts it to an Excel Table if not already one).
   Repeat for `Products`.
2. In the Power Pivot window, Diagram View shows both tables as boxes
   with their columns listed — no relationship line yet.

## 3. Creating the relationship

1. Drag `ProductID` in `Sales` onto `ProductID` in `Products` (or Home
   → Create Relationship, picking `Sales.ProductID` →
   `Products.ProductID`).
2. This is a one-to-many relationship: one row in `Products` (e.g.
   `P1`, Price 20) relates to many rows in `Sales` (P1 appears twice:
   rows 2 and 4). The relationship must run from the "many" side
   (`Sales`) to the "one" side (`Products`) — `ProductID` is unique in
   `Products` but repeats in `Sales`.

## 4. Building a cross-table PivotTable

1. Insert → PivotTable → **Use this workbook's Data Model**. In the
   field list, both `Sales` and `Products` appear as separate groups.
2. Rows: `Sales[Rep]`. Values: sum of `Sales[Qty]`. Add a calculated
   column first (see Section 5) to bring `Price` across the
   relationship.
3. In Power Pivot window, add a calculated column in `Sales`:
   `Revenue = Sales[Qty] * RELATED(Products[Price])`
   `RELATED()` pulls a value from the "one" side of the relationship
   into each row of the "many" side — here, each Sales row looks up
   its own ProductID's Price from `Products`.
4. Manual check: row 2 (P1, Qty 10) → Price 20 → Revenue `200`. Row 3
   (P2, Qty 5) → Price 35 → Revenue `175`. Row 4 (P1, Qty 8) → Price
   20 → Revenue `160`. Row 5 (P3, Qty 12) → Price 15 → Revenue `180`.
   Row 6 (P2, Qty 6) → Price 35 → Revenue `210`.
5. Back in the PivotTable, add `Sum of Revenue` to Values, Rows still
   `Rep`. Alan: rows 2,5 → `200+180=380`. Priya: rows 3,6 →
   `175+210=385`. Sam: row 4 → `160`. Total across all reps:
   `380+385+160=925`, which should also equal the grand total shown
   at the bottom of the PivotTable.

## Cheat sheet

| Concept | How |
|---|---|
| Add a table to the Data Model | Power Pivot → Add to Data Model |
| Create a relationship | Diagram View → drag key column to key column |
| Pull a value across a relationship | `RELATED(Table[Column])` (many→one direction) |
| PivotTable from multiple tables | Insert PivotTable → Use this workbook's Data Model |

## How It Actually Works

Power Pivot's Data Model doesn't store tables the way a worksheet does — it
uses the **VertiPaq** in-memory columnar engine (the same engine underneath
Power BI), which stores each column separately and compresses it using
dictionary encoding: every distinct value in a column is stored once in a
lookup dictionary, and the column itself becomes a compact array of integer
indexes into that dictionary. This is why VertiPaq compresses low-cardinality
columns (a status flag with 3 possible values) far more aggressively than
high-cardinality ones (a unique transaction ID per row), and why sorting or
restructuring source data to group repeated values together can measurably
shrink model size even before any aggregation happens. Relationships between
tables in the Data Model are stored as explicit graph edges (much like the
worksheet dependency graph, but between tables rather than cells) and are
used at query time to perform fast indexed joins — critically, VertiPaq
relationships are almost always one-to-many and singly-directional by
default, which is precisely why a "many-to-many"-looking relationship
between two fact tables usually needs an intermediate dimension table: the
storage engine's join algorithm is built around that one-to-many shape and
isn't a general-purpose relational join engine.

## Exercise

Add a calculated column `HighValue = IF(Sales[Revenue] >= 200,
"Yes", "No")`. The five Revenue values from Section 4 are
`200, 175, 160, 180, 210` — only `200` and `210` satisfy `>=200`, so
exactly two rows should read `Yes`, the remaining three `No`.
