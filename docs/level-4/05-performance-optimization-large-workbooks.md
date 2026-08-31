# 05 · Performance Optimization for Large Workbooks

This module covers why a workbook slows down as it grows, and the
concrete fixes: calculation settings, volatile-function avoidance,
efficient lookups, and VBA screen/calculation control.

## 1. Worked dataset — a lookup-heavy sheet

Sheet `Big`, `A1:C6`:

| | A | B | C |
|---|---|---|---|
| **1** | ID | Name | Score |
| **2** | 101 | Ana | 88 |
| **3** | 102 | Ben | 92 |
| **4** | 103 | Cid | 75 |
| **5** | 104 | Dee | 95 |
| **6** | 105 | Eli | 81 |

Imagine this pattern repeated across 500,000 rows in a real workbook —
the fixes below matter precisely because they scale differently as row
count grows.

## 2. Volatile functions — the hidden cost

- `NOW()`, `TODAY()`, `RAND()`, `RANDBETWEEN()`, `OFFSET()`, and
  `INDIRECT()` are **volatile**: Excel recalculates them on *every*
  worksheet recalculation, even if nothing they depend on changed.
- A single `=OFFSET(A1,0,0,COUNTA(A:A),1)` used as a named range in a
  workbook with thousands of formulas referencing it forces a full
  recalculation cascade constantly.
- Fix: replace `OFFSET`-based dynamic ranges with Excel Tables (which
  auto-expand without a volatile formula) or `INDEX` (non-volatile):
  `=INDEX(A:A,1):INDEX(A:A,COUNTA(A:A))` returns the same dynamic range
  as the `OFFSET` version without forcing constant recalculation.

## 3. Lookup efficiency — exact match and sorted ranges

1. `=VLOOKUP(103,A2:C6,3,FALSE)` on the small table returns `75` (Cid's
   score) correctly, but on 500,000 rows, `FALSE` (exact match) is an
   unsorted linear-ish scan in the worst case — slow at scale.
2. If the lookup column is sorted, `=VLOOKUP(103,A2:C6,3,TRUE)` (or
   omitted, since `TRUE` is default) uses a binary search — dramatically
   faster on large sorted data — but returns wrong results silently on
   unsorted data, so only use it when the ID column is guaranteed
   sorted.
3. `XLOOKUP` (`=XLOOKUP(103,A2:A6,C2:C6)`) is exact-match by default,
   as fast as `VLOOKUP(...,FALSE)` on modern Excel's optimized engine,
   and doesn't require the lookup column to be the leftmost column —
   preferred where available.

## 4. Calculation mode

1. **Formulas → Calculation Options → Manual** stops Excel from
   recalculating after every single cell edit — essential when bulk-
   editing a large workbook, since Automatic mode would otherwise
   recalculate the entire dependency tree after each keystroke commit.
2. Press **F9** to recalculate manually once all edits are made, or
   **Shift+F9** to recalculate only the active sheet.
3. Remember to switch back to Automatic before handing the workbook to
   someone else — a workbook stuck in Manual mode silently shows stale
   numbers to anyone who doesn't know to press F9.

## 5. VBA performance pattern

```vba
Sub FastBulkUpdate()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim i As Long
    For i = 2 To 500000
        Cells(i, 3).Value = Cells(i, 3).Value * 1.05
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
```

Turning off screen updating and automatic calculation before a large
loop, then restoring both afterward, is the single biggest VBA
performance fix — a 500,000-row loop with screen updating left on can
be an order of magnitude slower because Excel repaints the screen on
every cell write.

## 6. Structural fixes

| Symptom | Fix |
|---|---|
| Whole-column references (`SUM(A:A)`) in many formulas | Reference a bounded range or Table column instead — a whole-column reference forces Excel to consider a million+ rows |
| Many `SUMIF`/array formulas over huge ranges | Move heavy aggregation into a PivotTable or Power Pivot measure, which uses an optimized engine instead of cell-by-cell formulas |
| A workbook with dozens of unused formatted rows/columns beyond the real data | Delete them and save — Excel tracks a "used range" that bloats file size and recalculation scope even for empty formatted cells |

## Cheat sheet

| Technique | Effect |
|---|---|
| Avoid `OFFSET`/`INDIRECT` | Removes forced full recalculation |
| `XLOOKUP` / sorted `VLOOKUP(...,TRUE)` | Faster large-scale lookups |
| Manual calculation during bulk edits | Avoids recalculating after every keystroke |
| `ScreenUpdating=False` + `Calculation=Manual` in VBA loops | Major speedup for bulk VBA writes |

## Exercise

Rewrite Section 5's macro to also avoid recalculating on every
iteration by only setting `Application.Calculation = xlCalculationAutomatic`
once at the very end (already done above) — then explain in one
sentence why calling `Range("C2:C500000").Value =
Range("C2:C500000").Value` (an array-style bulk write) would be even
faster than the per-cell loop.
