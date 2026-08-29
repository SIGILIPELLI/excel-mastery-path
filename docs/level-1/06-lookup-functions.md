# 06 · Lookup Functions

Lookup functions find a value in one table based on a match in another —
the core operation behind connecting a budget's categories to reference
data that lives elsewhere in a workbook. This module covers `VLOOKUP`,
its modern replacement `XLOOKUP`, `INDEX`/`MATCH`, and approximate-match
lookups against numeric ranges (tax/commission brackets).

## 1. Building a reference table

1. In `budget-tracker.xlsx`, add a new sheet named `CategoryInfo` (right-click
   any tab > Insert, or the **+** button, then rename).
2. Enter this table starting at `A1`:

   | | A | B | C |
   |---|---|---|---|
   | **1** | Category | Type | Priority |
   | **2** | Rent | Fixed | 1 |
   | **3** | Savings | Fixed | 2 |
   | **4** | Subscriptions | Fixed | 3 |
   | **5** | Groceries | Variable | 4 |
   | **6** | Transport | Variable | 5 |
   | **7** | Entertainment | Variable | 6 |

3. This is the **lookup table** — the other lessons in this module pull
   `Type` and `Priority` into the `Budget` sheet by matching on `Category`.

## 2. VLOOKUP

1. On the `Budget` sheet, add a header `Type` in `E1` (adjust the column
   letter if your sheet layout differs from Module 4's).
2. In `E2` (the Rent row), type:
   `=VLOOKUP(A2,CategoryInfo!$A$2:$C$7,2,FALSE)`
3. This returns `Fixed`. Reading the arguments in order: **lookup value**
   (`A2`, "Rent"), **table array** (`CategoryInfo!$A$2:$C$7`, locked with
   `$` so it doesn't shift when copied), **column index number** (`2`,
   meaning the 2nd column of the table array — `Type`), and **range
   lookup** (`FALSE` for an exact match, which is what you want almost
   always when matching text).
4. Copy `E2` down through the rest of the category rows. Groceries and
   Entertainment should return `Variable`; Rent, Savings, and
   Subscriptions should return `Fixed`.
5. `VLOOKUP`'s biggest limitation: it can only look **rightward** — the
   column being returned must be to the right of the lookup column in the
   table array. It also breaks if a column is inserted into the middle of
   the table array, since the column index number (`2`) is a fixed
   position, not a column name.

## 3. XLOOKUP (Microsoft 365 / recent Excel only)

1. In `F2`, type:
   `=XLOOKUP(A2,CategoryInfo!$A$2:$A$7,CategoryInfo!$C$2:$C$7)`
2. This returns `1` (Rent's Priority). `XLOOKUP`'s arguments are
   **lookup value**, **lookup array**, **return array** — two separate
   ranges instead of one table array plus a column number, so it can
   return a column to the *left* of the lookup column just as easily as
   the right.
3. Copy `F2` down. Confirm Entertainment returns `6` and Savings returns
   `2`, matching the `CategoryInfo` table.
4. `XLOOKUP` has a built-in fourth argument for a not-found fallback:
   `=XLOOKUP(A2,CategoryInfo!$A$2:$A$7,CategoryInfo!$C$2:$C$7,"Not found")`
   returns `"Not found"` instead of the `#N/A` error `VLOOKUP` would give
   for a category that doesn't exist in `CategoryInfo`.
5. `XLOOKUP` isn't available in Excel 2019/2021 perpetual licenses without
   Microsoft 365 — if your version doesn't recognize `XLOOKUP`, use
   `INDEX`/`MATCH` (Section 4) instead, which works in every version.

## 4. INDEX and MATCH together

1. In `G2`, type:
   `=INDEX(CategoryInfo!$C$2:$C$7,MATCH(A2,CategoryInfo!$A$2:$A$7,0))`
2. `MATCH(A2,CategoryInfo!$A$2:$A$7,0)` finds "Rent"'s **position** within
   the range `A2:A7` — it's the 1st item, so `MATCH` returns `1`. The `0`
   argument means exact match (equivalent to `VLOOKUP`'s `FALSE`).
3. `INDEX(CategoryInfo!$C$2:$C$7,1)` then returns the 1st item of the
   `Priority` column — `1`.
4. Combined, `INDEX`/`MATCH` does exactly what `XLOOKUP` does in Section 3
   — look leftward or rightward freely, unaffected by inserted columns —
   but works in every Excel version back through 2007, which is why many
   older workbooks and long-time Excel users still default to it over
   `VLOOKUP` even when `XLOOKUP` isn't available.
5. Copy `G2` down and confirm it matches column `F`'s results exactly for
   every row.

## 5. Approximate-match lookup: commission brackets

Exact-match lookups (`FALSE` / `0`) find one specific value. An
**approximate-match** lookup instead finds which bracket a number falls
into — used for tax brackets, shipping-rate tiers, or (here) a sales
commission schedule.

1. On a new sheet `Commission`, build this table, **sorted ascending by
   the first column** (required for approximate-match `VLOOKUP` to work
   correctly):

   | | A | B |
   |---|---|---|
   | **1** | Sales | Rate |
   | **2** | 0 | 0% |
   | **3** | 1000 | 5% |
   | **4** | 5000 | 10% |
   | **5** | 10000 | 15% |

2. In `D2`, enter a test sales figure: `7500`. In `E2`, type:
   `=VLOOKUP(D2,A2:B5,2,TRUE)`
3. `TRUE` (or omitting the 4th argument) tells `VLOOKUP` to find the
   **largest value in column A that is less than or equal to** the lookup
   value, rather than requiring an exact match. `7500` falls between
   `5000` and `10000`, so it matches the `5000` row and returns `10%`.
4. In `F2`, compute the actual commission: `=D2*E2` → `750`.
5. Change `D2` to `12000` — `E2` should update to `15%` (the `10000`
   bracket) and `F2` to `1800`. Change it to `500` — `E2` should return
   `0%` and `F2` to `0`, since `500` falls in the `0` bracket. These three
   results (`7500→750`, `12000→1800`, `500→0`) confirm the bracket logic
   is working correctly.

## Cheat sheet

| Function | Syntax | Notes |
|---|---|---|
| `VLOOKUP` (exact) | `=VLOOKUP(value,table,col_num,FALSE)` | Looks rightward only |
| `VLOOKUP` (approximate) | `=VLOOKUP(value,table,col_num,TRUE)` | Table must be sorted ascending |
| `XLOOKUP` | `=XLOOKUP(value,lookup_array,return_array,[if_not_found])` | 365-only; any direction |
| `MATCH` | `=MATCH(value,lookup_array,0)` | Returns a **position**, not a value |
| `INDEX` | `=INDEX(array,row_num)` | Returns the value at a position |
| `INDEX`+`MATCH` | `=INDEX(return_range,MATCH(value,lookup_range,0))` | Works in every Excel version |

## Exercise

Build the `CategoryInfo` lookup table and pull `Type` into the `Budget`
sheet with `VLOOKUP`, then pull `Priority` with both `XLOOKUP` (if
available) and `INDEX`/`MATCH`, confirming all methods agree. Then build
the `Commission` bracket table and test the approximate-match `VLOOKUP`
with sales values of `500`, `7500`, and `12000`, confirming commissions of
`0`, `750`, and `1800` respectively.
