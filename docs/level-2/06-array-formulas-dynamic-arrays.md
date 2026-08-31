# 06 · Array Formulas & Dynamic Arrays

Modern Excel (365 / 2021+) treats arrays as first-class values.
Functions like `FILTER`, `SORT`, `UNIQUE`, and `SEQUENCE` return
multiple results that "spill" into neighboring cells automatically —
no `Ctrl+Shift+Enter` required, unlike legacy array formulas.

## 1. Worked dataset

Build this table on a sheet named `Team`, `A1:C9`:

| | A | B | C |
|---|---|---|---|
| **1** | Name | Dept | Score |
| **2** | Alan | Sales | 82 |
| **3** | Priya | IT | 91 |
| **4** | Sam | Sales | 76 |
| **5** | Nina | IT | 88 |
| **6** | Alan | Sales | 82 |
| **7** | Priya | HR | 91 |
| **8** | Omar | HR | 69 |
| **9** | Sam | Sales | 95 |

## 2. FILTER

1. In `E1`, type `=FILTER(A2:C9,B2:B9="Sales")`. This spills 4 rows —
   every row where Dept is Sales: Alan/82, Sam/76, Alan/82, Sam/95.
2. Manual check: scanning column B for "Sales" hits rows 2, 4, 6, 9 —
   exactly 4 rows, matching the spill.
3. Combine conditions: `=FILTER(A2:C9,(B2:B9="Sales")*(C2:C9>80))`
   Multiplying two boolean arrays acts as AND. Matches: row 2
   (Sales,82>80 ✓), row 4 (Sales,76>80 ✗), row 6 (Sales,82 ✓), row 9
   (Sales,95 ✓) → 3 rows spill: Alan/82, Alan/82, Sam/95.

## 3. UNIQUE and SORT

1. In `H1`, type `=UNIQUE(B2:B9)`. Spills the distinct departments in
   first-seen order: `Sales, IT, HR` (Sales first at row 2, IT first
   at row 3, HR first at row 7).
2. In `I1`, type `=SORT(UNIQUE(B2:B9))`. Sorts that list
   alphabetically: `HR, IT, Sales`.
3. Nest with FILTER to count each department:
   `=COUNTIF(B2:B9,I1#)` — the `#` (spill range operator) refers to
   the entire spilled `SORT(UNIQUE(...))` array, so this one formula
   spills three counts aligned to `HR/IT/Sales`: HR appears in rows
   7,8 → `2`; IT in rows 3,5 → `2`; Sales in rows 2,4,6,9 → `4`.

## 4. SEQUENCE

1. In `K1`, type `=SEQUENCE(5,1,1,1)`. Spills `1,2,3,4,5` down 5 rows —
   `SEQUENCE(rows,[columns],[start],[step])`.
2. Use it to build a quick index: `=SEQUENCE(3,3,1,1)` spills a 3×3
   grid `1,2,3 / 4,5,6 / 7,8,9` reading left-to-right, top-to-bottom.

## 5. Legacy CSE array formula (for comparison)

1. Older workbooks use array formulas entered with **Ctrl+Shift+Enter**
   (shown wrapped in `{}` by Excel, not typed manually). Example, to
   sum Score only for Sales without a helper column:
   `{=SUM(IF(B2:B9="Sales",C2:C9))}`
2. Manual check: Sales rows 82, 76, 82, 95 → sum `335`. The modern
   equivalent needs no CSE: `=SUM(FILTER(C2:C9,B2:B9="Sales"))`,
   confirming the same `335`.

## Cheat sheet

| Function | Purpose |
|---|---|
| `FILTER(array,include,[if_empty])` | Rows matching a boolean condition |
| `UNIQUE(array)` | Distinct values, first-seen order |
| `SORT(array,[col],[order])` | Sort a range or spilled array |
| `SEQUENCE(rows,[cols],[start],[step])` | Generate a number sequence |
| `range#` | Spill range operator — refers to a dynamic array's full output |

## Exercise

Using the `Team` table, write one formula that lists unique names
sorted alphabetically (`=SORT(UNIQUE(A2:A9))` — expect Alan, Nina,
Omar, Priya, Sam), then write `=SUMPRODUCT((B2:B9=I1#)*C2:C9)` next to
the department SORT/UNIQUE spill from Section 3 to total Score per
department, and manually confirm Sales sums to `82+76+82+95=335`.
