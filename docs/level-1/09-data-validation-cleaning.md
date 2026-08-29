# 09 · Data Validation & Cleaning

Real-world data is messy: inconsistent capitalization, stray spaces,
accidental duplicate rows, and free-typed entries that should have come
from a fixed list. This module covers **data validation** (restricting
what can be typed into a cell) and **cleaning** existing messy data with
`TRIM`, `UPPER`/`LOWER`/`PROPER`, and **Remove Duplicates**.

## 1. Dropdown lists with Data Validation

1. On the `CategoryInfo` sheet (Module 6), select `A2:A7` — you'll reuse
   this as the source list for a dropdown elsewhere.
2. On the `Transactions` sheet (Module 8), select the `Category` column's
   data cells, `A2:A13`.
3. **Data > Data Validation**. In the dialog's **Settings** tab, set
   **Allow** to **List**, and in **Source**, type
   `=CategoryInfo!$A$2:$A$7` (or click the range-picker icon and select it
   on the sheet).
4. Click **OK**. Every cell in `A2:A13` now shows a small dropdown arrow
   when selected, restricted to the six categories in `CategoryInfo` —
   typing anything else (e.g. a typo like `Rennt`) triggers a rejection
   with a default error message.
5. On the **Error Alert** tab (still inside the Data Validation dialog),
   customize the **Style** (Stop / Warning / Information) and message text
   — **Stop** fully blocks invalid entries, while **Warning** allows the
   user to override with a confirmation, useful when a value is unusual
   but not necessarily wrong.
6. **Data > Data Validation > Clear All** (in the same dialog, bottom-left
   button) removes a validation rule from the selected cells without
   affecting their existing values.

## 2. Simulating messy input

1. Add a new sheet named `Messy`. Enter this exactly — note the
   inconsistent spacing and capitalization, which is intentional:

   | A |
   |---|
   | Category |
   | " rent " |
   | "GROCERIES" |
   | "  Transport" |
   | "entertainment " |
   | "Savings" |
   | "Rent" |

   (Type the values without the surrounding quote marks — they're shown
   here only to make the leading/trailing spaces visible.)

2. This mirrors what pasted or imported data (a CSV export, a form
   submission) often looks like before any dropdown restriction was in
   place to prevent it.

## 3. Cleaning text with TRIM and case functions

1. In `B2`, type `=TRIM(A2)` and copy down through `B7`. `TRIM` removes
   leading and trailing spaces and collapses any internal multiple spaces
   to single spaces — `" rent "` becomes `"rent"`, `"  Transport"` becomes
   `"Transport"`.
2. In `C2`, type `=PROPER(TRIM(A2))` and copy down. `PROPER` capitalizes
   the first letter of each word — `"rent"` becomes `"Rent"`,
   `"GROCERIES"` becomes `"Groceries"`. Combining `TRIM` inside `PROPER`
   in one formula fixes both spacing and casing issues in a single step.
3. `UPPER(A2)` would instead force everything to uppercase
   (`"RENT"`) and `LOWER(A2)` to lowercase (`"rent"`) — useful when the
   target format needs to be consistently one case rather than
   title-cased, e.g. matching a system that stores category codes in all
   caps.
4. After cleaning, column `C` should read: `Rent`, `Groceries`,
   `Transport`, `Entertainment`, `Savings`, `Rent` — matching the
   `CategoryInfo` list exactly (case and spacing both now consistent),
   which is what makes a cleaned column safe to use as a `VLOOKUP`/
   `XLOOKUP` key afterward (Module 6) — a lookup against `"GROCERIES"` with
   exact match against a table listing `"Groceries"` would otherwise fail,
   since text comparison in these lookups is case-insensitive but **not**
   whitespace-insensitive, so untrimmed spaces alone can break a match even
   when case doesn't matter.

## 4. Finding and removing duplicates

1. Notice column `C` (cleaned) now has `Rent` appearing twice (rows 2 and
   7) — a duplicate that wasn't obvious in the original messy data since
   one was `" rent "` and the other `"Rent"`.
2. Select `C1:C7`, then **Data > Data Tools > Remove Duplicates**. Confirm
   **My data has headers** is checked, then **OK**.
3. Excel reports how many duplicate values were found and removed (`1`
   duplicate found; `5` unique values remaining) — the second `Rent` row
   is deleted, leaving five unique cleaned category names.
4. **Remove Duplicates** compares entire selected rows (or just the
   selected columns, if only some columns were selected) for an exact
   match — it does not clean text first, which is exactly why Section 3's
   cleaning had to happen *before* deduplication; running Remove
   Duplicates directly on the original messy column `A` would have found
   **zero** duplicates, since `" rent "` and `"Rent"` are technically
   different text strings.
5. **Data > Data Tools > Data Validation > Circle Invalid Data** (small
   dropdown next to the Data Validation button) draws a red circle around
   any existing cell that violates an active validation rule — useful for
   auditing a sheet where invalid values were pasted in *after* a
   validation rule was already set up (validation only blocks new manual
   typing, not paste operations, by default).

## Cheat sheet

| Function/Feature | Purpose |
|---|---|
| `TRIM(text)` | Remove leading/trailing/extra spaces |
| `UPPER(text)` / `LOWER(text)` | Force upper/lower case |
| `PROPER(text)` | Capitalize first letter of each word |
| Data Validation > List | Restrict entry to a fixed dropdown list |
| Data Validation > Error Alert | Customize Stop/Warning behavior and message |
| Remove Duplicates | Delete exact-match duplicate rows |
| Circle Invalid Data | Flag existing cells that violate a validation rule |

## Exercise

Build the `Messy` sheet exactly as shown, clean it with a single
`=PROPER(TRIM(A2))` formula copied down, confirm the result matches
`CategoryInfo`'s category list (accounting for the intentional `Rent`
duplicate), then use Remove Duplicates on the cleaned column and confirm
Excel reports exactly one duplicate removed. Finally, add a Data
Validation dropdown to the `Transactions` sheet's `Category` column
sourced from `CategoryInfo!$A$2:$A$7`, and confirm typing an invalid
category is rejected.
