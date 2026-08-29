# 07 · Logical Functions

Logical functions turn a formula from "compute a number" into "make a
decision." This module covers `IF`, nested `IF`s, `IFS` (a cleaner
multi-branch alternative), and combining conditions with `AND`/`OR` — using
the budget's `Difference` and `CategoryInfo` data from Modules 2 and 6.

## 1. A single IF

1. On the `Budget` sheet, add a header `Status` in the next open column
   (e.g. `H1` if you've been following along with Modules 4–6's columns).
2. In `H2` (Rent, Difference `0`), type:
   `=IF(D2>0,"Over Budget","OK")`
3. `IF` takes three arguments: a **condition** (`D2>0`), the value if
   **TRUE** (`"Over Budget"`), and the value if **FALSE** (`"OK"`). Rent's
   Difference is `0`, so `0>0` is `FALSE`, and the cell shows `OK`.
4. Copy `H2` down. Groceries (`55`) and Entertainment (`75`) should show
   `Over Budget`; Rent, Transport (`-20`), and Savings should show `OK` —
   note this simple version doesn't yet distinguish "exactly on budget"
   from "under budget," which the nested version in Section 2 fixes.

## 2. Nested IF for three outcomes

1. A single `IF` only branches two ways. To get three outcomes — Over,
   Under, and On Track — nest a second `IF` inside the FALSE branch of the
   first:
   `=IF(D2>0,"Over Budget",IF(D2<0,"Under Budget","On Track"))`
2. Reading it: if Difference is greater than 0, "Over Budget"; otherwise,
   check a second condition — if Difference is less than 0, "Under
   Budget"; otherwise (meaning it's neither greater nor less than 0, so it
   must be exactly 0), "On Track".
3. Copy this down all six rows. Expected results: Rent `On Track`,
   Groceries `Over Budget`, Transport `Under Budget`, Entertainment
   `Over Budget`, Savings `On Track`, Subscriptions `Over Budget` (`5`
   from Module 4's added row).
4. Nesting more than two or three `IF`s becomes hard to read fast —
   each additional branch adds another layer of parentheses to track. This
   is exactly the problem `IFS` (Section 3) solves.

## 3. IFS for cleaner multi-branch logic

1. On the `CategoryInfo` sheet, add a header `Level` in `D1`. In `D2`
   (Rent, Priority `1`), type:
   `=IFS(C2<=2,"High",C2<=4,"Medium",TRUE,"Low")`
2. `IFS` takes pairs of (condition, result) and returns the result for the
   **first** condition that's `TRUE`, checked top to bottom — no nesting
   or extra parentheses per branch. `TRUE` as a final condition acts as a
   catch-all "else," equivalent to the innermost `ELSE` of a nested `IF`.
3. Copy `D2` down through row 7. Expected: Rent (Priority `1`) → `High`,
   Savings (`2`) → `High`, Subscriptions (`3`) → `Medium`, Groceries (`4`)
   → `Medium`, Transport (`5`) → `Low`, Entertainment (`6`) → `Low`.
4. If no condition in an `IFS` matches and there's no `TRUE` catch-all,
   the formula returns `#N/A` — always include a final catch-all branch
   unless you specifically want that error as a signal that a case was
   missed.

## 4. Combining conditions with AND / OR

1. Back on the `Budget` sheet, add a header `Needs Review` in the next
   column. Type:
   `=IF(AND(E2="Variable",D2>50),"Review","")`
   (using the `Type` column `E` pulled via `VLOOKUP` in Module 6, and
   `Difference` column `D`).
2. `AND(condition1,condition2,…)` returns `TRUE` only if **every**
   condition is `TRUE`. Here: the category must be `Variable` **and** its
   Difference must exceed `50`.
3. Copy down. Groceries (`Variable`, `55`) and Entertainment (`Variable`,
   `75`) should show `Review`; Transport (`Variable`, `-20`) does not,
   since `-20>50` is `FALSE`, failing the `AND`; Rent and Savings
   (`Fixed`) never show `Review` regardless of their Difference, since the
   `Type` condition alone already fails.
4. `OR(condition1,condition2,…)` returns `TRUE` if **any** condition is
   `TRUE`. Swap `AND` for `OR` in the same formula —
   `=IF(OR(E2="Variable",D2>50),"Review","")` — and every `Variable`
   category flags regardless of its Difference (since the first condition
   alone satisfies `OR`), plus any `Fixed` category whose Difference
   happens to exceed `50` (none do here, but the logic would catch one if
   it existed).
5. `AND`/`OR` can nest inside `IF`, or `IF` can nest inside `AND`/`OR`'s
   arguments — but as with nested `IF`s, past two or three conditions,
   naming intermediate results in helper columns (e.g. a plain `TRUE`/
   `FALSE` column for each condition) keeps a workbook easier to audit
   than one dense formula.

## Cheat sheet

| Function | Syntax | Behavior |
|---|---|---|
| `IF` | `=IF(condition,if_true,if_false)` | Two-way branch |
| Nested `IF` | `=IF(c1,r1,IF(c2,r2,r3))` | Three+ way branch, gets unwieldy fast |
| `IFS` | `=IFS(c1,r1,c2,r2,TRUE,default)` | Cleaner multi-branch, first match wins |
| `AND` | `=AND(c1,c2,…)` | TRUE only if all conditions are TRUE |
| `OR` | `=OR(c1,c2,…)` | TRUE if any condition is TRUE |

## Exercise

Add a `Status` column to the `Budget` sheet using the three-way nested
`IF` from Section 2, and confirm every row's result matches the expected
list. Add a `Level` column to `CategoryInfo` using `IFS` based on
Priority, confirming Rent/Savings are `High`, Subscriptions/Groceries are
`Medium`, and Transport/Entertainment are `Low`. Finally add a
`Needs Review` column using `AND(Type="Variable", Difference>50)` and
confirm only Groceries and Entertainment flag.
