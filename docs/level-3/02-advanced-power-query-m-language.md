# 02 · Advanced Power Query (M Language Basics)

Every Power Query transformation is really generating code in the **M**
language, visible via the Advanced Editor. This module reads and
hand-edits M directly — custom columns, conditional logic, and merging
two queries — instead of only clicking Ribbon buttons.

## 1. Worked dataset

Two tables, both converted to Excel Tables.

`Orders` (`A1:C6`):

| | A | B | C |
|---|---|---|---|
| **1** | OrderID | CustID | Amount |
| **2** | 1 | C1 | 500 |
| **3** | 2 | C2 | 300 |
| **4** | 3 | C1 | 800 |
| **5** | 4 | C3 | 200 |
| **6** | 5 | C2 | 450 |

`Customers` (`A1:B4`):

| | A | B |
|---|---|---|
| **1** | CustID | Name |
| **2** | C1 | Alan |
| **3** | C2 | Priya |
| **4** | C3 | Sam |

## 2. Viewing the M code

1. Load `Orders` into Power Query (Data → From Table/Range). View →
   **Advanced Editor**. The base query looks like:
   ```m
   let
       Source = Excel.CurrentWorkbook(){[Name="Orders"]}[Content]
   in
       Source
   ```
2. Every UI action (a filter, a rename, a group) appends another line
   inside the `let...in` block, bound to a name, and the final `in`
   line names which step's output is returned.

## 3. Custom column with M

1. Add Column → Custom Column, name `Tier`, formula:
   `if [Amount] >= 500 then "Large" else "Small"`
2. This appends a step to the M code:
   ```m
   #"Added Custom" = Table.AddColumn(#"Changed Type", "Tier",
       each if [Amount] >= 500 then "Large" else "Small")
   ```
3. Manual check row by row: 500→Large, 300→Small, 800→Large,
   200→Small, 450→Small. Three Small, two Large.

## 4. Merging two queries

1. Load `Customers` as its own query too. In the `Orders` query,
   Home → **Merge Queries**. Match `Orders.CustID` to
   `Customers.CustID`, Join Kind **Left Outer** (keep every order, add
   customer info where it matches).
2. Underlying M: `Table.NestedJoin(Orders, {"CustID"}, Customers,
   {"CustID"}, "Customers", JoinKind.LeftOuter)`. Expand the new
   `Customers` column, keeping only `Name`.
3. Manual check: OrderID 1 (CustID C1) → Alan; OrderID 2 (C2) → Priya;
   OrderID 3 (C1) → Alan; OrderID 4 (C3) → Sam; OrderID 5 (C2) →
   Priya. Every order row must now show the correct customer name.

## 5. Editing M by hand

1. In the Advanced Editor, manually insert a step that trims
   whitespace from the Name column defensively, even though this
   dataset has none:
   ```m
   #"Trimmed Name" = Table.TransformColumns(#"Expanded Customers",
       {{"Name", Text.Trim, type text}})
   ```
   and update the final `in` line to `in #"Trimmed Name"`.
2. This demonstrates the general `Table.TransformColumns` pattern:
   `Table.TransformColumns(table, {{"ColumnName", function, type}})`
   applies `function` to every value in `ColumnName`.

## Cheat sheet

| M function | Purpose |
|---|---|
| `Table.AddColumn(table,name,each expr)` | Add a computed column |
| `Table.NestedJoin(...)` | Merge two tables (the engine behind Merge Queries) |
| `Table.TransformColumns(table,{{col,fn,type}})` | Apply a function to a column |
| `let ... in ...` | Every query's structure: named steps, final returned step |

## How It Actually Works

M, the language behind every Power Query step, is a **pure, lazily-evaluated
functional language** — every step is actually a function call that takes
the previous step's table as its argument and returns a new table, and the
query as a whole is just one composed function (`let ... in ...`)
referencing each step by name. Because M is lazy, the engine doesn't
necessarily compute step 3's entire result before running step 4 — it can
analyze the whole chain and, wherever the source supports it, use **query
folding**: translating your point-and-click steps back into the source
system's native query language (a SQL `WHERE`/`SELECT` for a database
connection, for instance) so filtering and column selection happen at the
source rather than after pulling every row into Excel. This is why step
order matters for performance even though it doesn't change the *result*:
a `Filter` step placed before a custom column that can't be folded (like
one calling a non-foldable M function) still gets pushed to the source,
but a `Filter` placed *after* that unfoldable step forces all rows to be
pulled locally first, breaking the folding chain for every step after it.
`Table.Buffer` exists specifically to force *eager* evaluation at a chosen
point, deliberately trading folding/laziness for a predictable, single
materialization when a step is expensive and would otherwise be
re-evaluated multiple times downstream.

## Exercise

Add a custom column `AmountWithTax` using M:
`each [Amount] * 1.08`. Manually verify row 1: `500*1.08=540`, row 3:
`800*1.08=864`. Confirm the query's output matches both by hand.
