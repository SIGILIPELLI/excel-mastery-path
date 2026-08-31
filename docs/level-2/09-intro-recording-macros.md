# 09 · Intro to Recording Macros

A macro is a recorded sequence of actions that Excel can replay on
demand. The Macro Recorder writes real VBA code behind the scenes —
this module records simple macros, inspects the generated code, and
covers relative vs. absolute references and assigning a macro to a
button.

## 1. Worked dataset

Build this on a sheet named `Data`, `A1:C5`:

| | A | B | C |
|---|---|---|---|
| **1** | Item | Qty | Price |
| **2** | Pen | 10 | 2 |
| **3** | Book | 4 | 15 |
| **4** | Bag | 2 | 25 |
| **5** | | | |

## 2. Recording a formatting macro

1. Developer → **Record Macro**. Name it `HighlightHeader`, shortcut
   `Ctrl+Shift+H`, store in **This Workbook**.
2. While recording: select `A1:C1`, apply bold and a light-blue fill
   (Home → Fill Color).
3. Developer → **Stop Recording**.
4. Developer → **Macros** → select `HighlightHeader` → **Edit** to
   open the VBA editor and inspect the generated code. It will
   resemble:
   ```vb
   Sub HighlightHeader()
       Range("A1:C1").Select
       Selection.Font.Bold = True
       With Selection.Interior
           .Pattern = xlSolid
           .Color = RGB(221, 235, 247)
       End With
   End Sub
   ```
5. This confirms the recorder captured exactly two user actions
   (select range, then two format changes) as three lines of VBA.

## 3. Absolute vs. relative reference recording

1. Developer has a toggle **Use Relative References**. With it
   **off** (default), a macro recorded starting from `A1` always
   selects `A1:C1` again, regardless of which cell is active when the
   macro runs.
2. Turn **Use Relative References** on, then re-record: click `A2`
   first, then select `A2:C2` and apply italic. Stop recording.
3. Now if the active cell is `A3` and you run this macro, it applies
   italic to `A3:C3` instead — because it was recorded relative to the
   starting cell, not to the fixed address `A2`. This distinction
   matters whenever a macro should act on "the row I clicked," not a
   hardcoded row.

## 4. A calculation macro

1. Record a new macro `AddTotalColumn`: in `D1` type `Total`, in `D2`
   type `=B2*C2`, fill down to `D4`, then select `D2:D4` and apply
   currency format.
2. Manual check on the generated formulas: `D2 = 10*2 = 20`,
   `D3 = 4*15 = 60`, `D4 = 2*25 = 50`.
3. Inspect the recorded VBA — the `.FillDown` and
   `.NumberFormat = "$#,##0.00"` lines are recognizable even without
   writing VBA by hand, which is the main value of recording: it's a
   fast way to discover the property names and methods VBA uses for a
   given manual action, useful groundwork before Level 3's hand-written
   VBA.

## 5. Assigning a macro to a button

1. Developer → Insert → Form Controls → **Button**. Draw it anywhere,
   then in the Assign Macro dialog choose `HighlightHeader`.
2. Clicking the button now re-runs the recorded macro — this is the
   simplest way to give a non-technical user a one-click action
   without teaching them the Developer tab or keyboard shortcuts.

## Cheat sheet

| Task | Path |
|---|---|
| Record a macro | Developer → Record Macro |
| Stop recording | Developer → Stop Recording |
| Toggle relative recording | Developer → Use Relative References |
| View/edit recorded code | Developer → Macros → Edit |
| Run a macro | Developer → Macros → Run, or assigned shortcut/button |
| Attach to a button | Insert Form Control Button → Assign Macro |

## Exercise

Record a macro `FormatTotals` that selects `D2:D4` (after building the
Total column from Section 4), applies bold, and adds a border. Run it
via `Ctrl+Shift+` (your chosen shortcut) and confirm visually that all
three total cells (`20, 60, 50`) are now bold with a border, then open
the VBA editor and read the generated `Sub` line by line to identify
which line corresponds to the border vs. the bold.
