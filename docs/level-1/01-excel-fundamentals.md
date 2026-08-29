# 01 · Excel Fundamentals

Excel organizes everything around three nested containers: a **workbook**
(the file) holds one or more **worksheets** (tabs), and each worksheet is a
grid of **cells** addressed by column letter and row number (e.g. `B3`).
Every lesson in this track builds on being fluent with that grid — selecting
it, naming pieces of it, and moving around it quickly. This module also
introduces the worked dataset used across the rest of Level 1: a small
personal budget.

## 1. Workbooks, worksheets, and the ribbon

1. Open Excel and choose **Blank workbook**. A new file starts with one
   worksheet tab (usually named `Sheet1`) at the bottom-left of the window.
2. Add a worksheet with the **+** button next to the last tab, or
   right-click any tab and choose **Insert**. Rename a tab by
   double-clicking it; reorder tabs by dragging them left or right.
3. The **ribbon** (top of the window) is organized into tabs — **Home**,
   **Insert**, **Formulas**, **Data**, **Review**, **View**, and more — each
   holding grouped commands. **Home** covers formatting and basic editing;
   **Formulas** and **Data** are where most of this course lives.
4. The **Name Box** (top-left, above column A) shows the address of the
   currently selected cell or range, and can be typed into to jump straight
   to any address — type `B3` and press Enter to select cell B3 instantly.
5. The **Formula Bar** (next to the Name Box) shows the actual content of
   the selected cell — a typed value, or the formula that produces a
   calculated value — even when the cell itself displays a formatted
   result.

## 2. Cells, addresses, and ranges

1. A single cell is addressed as `ColumnLetter` + `RowNumber` — `A1` is the
   top-left cell, `C5` is column C, row 5.
2. A **range** is a rectangular block of cells written as
   `TopLeft:BottomRight` — `A1:A5` is five cells stacked in column A;
   `A1:C5` is a 3-column-by-5-row block.
3. Select a range by clicking one corner cell and dragging to the opposite
   corner, or click the first cell and **Shift+click** the last cell.
   **Ctrl/⌘-click** additional cells or ranges to select several
   non-adjacent pieces at once.
4. **Ctrl/⌘+Arrow** jumps to the edge of the current block of data in that
   direction — the fastest way to move across a large sheet without
   scrolling. **Ctrl/⌘+Shift+Arrow** does the same while extending the
   selection.
5. Select an entire row by clicking its row number, or an entire column by
   clicking its column letter; click the small triangle at the top-left
   corner of the grid (where row and column headers meet) to select the
   whole sheet.

## 3. Entering and editing data

1. Click a cell and start typing to enter a value; press **Enter** to
   confirm and move down, **Tab** to confirm and move right, or **Esc** to
   cancel the edit.
2. Double-click a cell (or select it and press **F2**) to edit its existing
   content in place, with the cursor positioned inside the cell rather than
   overwriting it.
3. Excel auto-detects the type of what you type: `1200` becomes a number
   (right-aligned by default), `Rent` becomes text (left-aligned), and
   `1/15/2026`-style entries become dates — the alignment is a fast visual
   check that a column contains what you expect.
4. **Fill handle** — the small square at the bottom-right corner of a
   selected cell or range. Drag it down or across to copy a formula or
   continue a pattern (typing `Jan` and `Feb` in two cells, then dragging
   the fill handle across, continues `Mar`, `Apr`, …).
5. Undo (**Ctrl/⌘+Z**) and Redo (**Ctrl/⌘+Y**, or **⌘+Shift+Z** on Mac) are
   available for essentially every action, including formatting.

## 4. Building the Level 1 worked example: a personal budget

Every module in this level reuses the same small dataset so the numbers
stay familiar as the functions get more advanced. Build it now:

1. On `Sheet1`, enter this exactly, starting at cell `A1`:

   | | A | B | C |
   |---|---|---|---|
   | **1** | Category | Budgeted | Actual |
   | **2** | Rent | 1200 | 1200 |
   | **3** | Groceries | 400 | 455 |
   | **4** | Transport | 150 | 130 |
   | **5** | Entertainment | 100 | 175 |
   | **6** | Savings | 300 | 300 |

2. Rename the sheet tab to `Budget` (double-click the tab, type the name,
   press Enter) — later modules will refer back to this sheet by name.
3. Row 1 is a **header row** — labels describing each column, not data
   itself. Almost every Excel feature (Tables in Module 4, PivotTables in
   Module 8) assumes the first row of a data range is headers.
4. Select `A1:C6` and check the Name Box confirms the range, then look at
   the Formula Bar as you click through a few individual cells — confirm
   `B2` shows `1200` and `A1` shows `Category`.

!!! info "Why Actual isn't identical to Budgeted"
    Real budgets rarely match exactly — you'll compute the gap between
    columns B and C with formulas starting in Module 2, and use that gap to
    drive conditional formatting (Module 3) and IF-based status logic
    (Module 7).

## 5. Saving and file formats

1. **File > Save** (⌘/Ctrl+S) the first time prompts for a location and
   filename; save this as `budget-tracker.xlsx`.
2. `.xlsx` is Excel's standard modern format. `.xlsm` is the same format
   with **macros** enabled (relevant from Level 2 Module 9 onward) —
   Excel will refuse to save macro code into a plain `.xlsx` file.
3. **File > Save As** lets you export to other formats — `.csv` (plain
   comma-separated values, no formatting or formulas, just the visible
   values) is common for exporting data to other tools; `.pdf` is common
   for sharing a formatted, non-editable snapshot.
4. Excel autosaves periodically to a recovery file even before you save
   manually — **File > Info > Manage Workbook** shows any recovered
   versions if Excel or your computer crashes mid-edit.

## Cheat sheet

| Action | Shortcut |
|---|---|
| New workbook | Ctrl/⌘+N |
| Save | Ctrl/⌘+S |
| Edit active cell in place | F2 |
| Jump to edge of data block | Ctrl/⌘+Arrow |
| Extend selection to edge of data block | Ctrl/⌘+Shift+Arrow |
| Select entire sheet | Ctrl/⌘+A (or corner triangle) |
| Undo / Redo | Ctrl/⌘+Z / Ctrl/⌘+Y (⌘+Shift+Z on Mac) |
| Confirm entry, move down | Enter |
| Confirm entry, move right | Tab |
| Cancel entry | Esc |

## Exercise

Create a new workbook. Build the 5-category budget table from Section 4
exactly as shown (Category, Budgeted, Actual). Rename the sheet tab to
`Budget`. Add a sixth category row of your own choosing with a Budgeted and
Actual value, then use the Name Box to jump to cell `A1`, then to `C7` (or
whatever row your new category landed on), confirming the Formula Bar shows
the correct content each time. Save the file as `budget-tracker.xlsx` —
you'll keep building on this same file through Module 10.
