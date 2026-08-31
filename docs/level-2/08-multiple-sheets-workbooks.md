# 08 · Working with Multiple Sheets & Workbooks

Real workbooks split data across sheets (one per month, region, or
department) and sometimes across separate files entirely. This module
covers 3-D references, cross-sheet and cross-workbook formulas,
consolidation, and keeping links intact when files move.

## 1. Worked dataset

Create three sheets named `Jan`, `Feb`, `Mar`, each with the identical
layout `A1:B4`:

`Jan`:

| | A | B |
|---|---|---|
| **1** | Product | Sales |
| **2** | Widget | 500 |
| **3** | Gadget | 300 |
| **4** | Gizmo | 200 |

`Feb`: Widget `600`, Gadget `350`, Gizmo `150`.
`Mar`: Widget `550`, Gadget `400`, Gizmo `250`.

Add a fourth sheet `Summary`.

## 2. 3-D references (sum across sheets)

1. In `Summary!B2` (Widget total), type:
   `=SUM(Jan:Mar!B2)`
   A 3-D reference (`Jan:Mar!B2`) sums cell `B2` across every sheet
   from `Jan` through `Mar` inclusive — including any sheet later
   inserted between them.
2. Manual check: Widget sales are 500 + 600 + 550 = `1650`. `B2`
   should show `1650`.
3. Fill down for Gadget (`Jan:Mar!B3` → `300+350+400 = 1050`) and
   Gizmo (`Jan:Mar!B4` → `200+150+250 = 600`).

## 3. Cross-sheet references (single sheet)

1. In `Summary!D2`, reference just February's Widget figure directly:
   `=Feb!B2`
   This returns `600` — a plain pointer to one cell on one other
   sheet, unlike the 3-D reference which aggregates across a sheet
   range.
2. Use this pattern to build a month-by-month comparison table
   without re-typing data: `Summary!D2:F4` referencing `Jan!B2:B4`,
   `Feb!B2:B4`, `Mar!B2:B4` respectively for each product row.

## 4. Cross-workbook references

1. With both workbooks open, a formula referencing another open file
   looks like: `='[Budget2026.xlsx]Jan'!$B$2` (external workbook name
   in square brackets, sheet name, then cell).
2. If the source workbook is closed, Excel keeps the full path:
   `='C:\Reports\[Budget2026.xlsx]Jan'!$B$2` — this still works and
   updates when the source file changes, but breaks if the file is
   moved or renamed. Use **Data → Edit Links** to repoint broken links
   after a file move, rather than manually rewriting every formula.
3. Best practice: keep workbooks that link to each other in the same
   folder, and avoid renaming source files once other workbooks depend
   on them.

## 5. Grouping sheets for identical edits

1. Click the `Jan` tab, then Shift-click the `Mar` tab to select all
   three as a group (tabs turn white, title bar shows `[Group]`).
2. Any edit made now — e.g. typing a header in `A1` or applying bold
   formatting — applies identically to `Jan`, `Feb`, and `Mar`
   simultaneously. This is efficient for applying the same formatting
   or a new formula across many structurally identical sheets, but
   dangerous if forgotten: right-click a tab → **Ungroup Sheets** as
   soon as the batch edit is done, or a later "one-sheet" edit will
   silently overwrite all grouped sheets.

## Cheat sheet

| Reference type | Syntax | Behavior |
|---|---|---|
| 3-D reference | `=SUM(Jan:Mar!B2)` | Aggregates one cell across a sheet range |
| Cross-sheet | `=Feb!B2` | Points to one cell on one other sheet |
| Cross-workbook (open) | `='[File.xlsx]Sheet'!B2` | Live link while both files open |
| Cross-workbook (closed) | `='C:\path\[File.xlsx]Sheet'!B2` | Full path required |
| Fix broken links | Data → Edit Links → Change Source | Repoints after a file move/rename |

## Exercise

Add a `Summary!B5` total-of-totals: `=SUM(B2:B4)` and confirm it
equals `1650+1050+600=3300`, then cross-check that same figure with a
single 3-D reference over the full range: `=SUM(Jan:Mar!B2:B4)` — both
should return `3300`.
