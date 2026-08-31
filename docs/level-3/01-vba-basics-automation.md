# 01 · VBA Basics for Automation

VBA (Visual Basic for Applications) is the programming language behind
Excel. This module writes hand-typed VBA — variables, loops, `If`
logic, and the Range/Cells object model — rather than relying on the
macro recorder from Level 2.

## 1. Worked dataset

Build this on a sheet named `Sheet1`, `A1:B6`:

| | A | B |
|---|---|---|
| **1** | Item | Qty |
| **2** | Pen | 12 |
| **3** | Book | 3 |
| **4** | Bag | 25 |
| **5** | Pencil | 0 |
| **6** | Eraser | 8 |

## 2. The Range and Cells objects

1. Alt+F11 opens the VBA editor. Insert → Module. Type:
   ```vb
   Sub ShowValue()
       MsgBox Range("B2").Value
   End Sub
   ```
2. Press F5 to run. A message box shows `12` — `Range("B2").Value`
   reads cell B2's contents. `Cells(2,2)` is equivalent (row 2,
   column 2), useful inside loops where row/column are variables.

## 3. Variables and a For loop

1. Type:
   ```vb
   Sub SumQty()
       Dim total As Long
       Dim i As Integer
       total = 0
       For i = 2 To 6
           total = total + Cells(i, 2).Value
       Next i
       MsgBox "Total: " & total
   End Sub
   ```
2. Manual check: `12+3+25+0+8 = 48`. The message box must read
   `Total: 48`.
3. `Dim` declares a variable's type. `Long` holds integers safely for
   sums that might exceed `Integer`'s ~32,767 limit; `Integer` is fine
   for a small loop counter like `i`.

## 4. Conditional logic — flagging low stock

1. Type:
   ```vb
   Sub FlagLowStock()
       Dim i As Integer
       For i = 2 To 6
           If Cells(i, 2).Value < 10 Then
               Cells(i, 3).Value = "Reorder"
               Cells(i, 3).Interior.Color = RGB(255, 199, 206)
           Else
               Cells(i, 3).Value = "OK"
           End If
       Next i
   End Sub
   ```
2. Manual check row by row: Pen(12)→OK, Book(3)→Reorder, Bag(25)→OK,
   Pencil(0)→Reorder, Eraser(8)→Reorder. Column `C` should read
   `OK, Reorder, OK, Reorder, Reorder` with the three Reorder cells
   shaded light red.

## 5. A user-facing Sub with InputBox

1. Type:
   ```vb
   Sub AddNewItem()
       Dim newItem As String
       Dim newQty As Variant
       Dim nextRow As Long
       newItem = InputBox("Item name:")
       newQty = InputBox("Quantity:")
       If newItem = "" Then Exit Sub
       nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
       Cells(nextRow, 1).Value = newItem
       Cells(nextRow, 2).Value = newQty
   End Sub
   ```
2. `Cells(Rows.Count, 1).End(xlUp).Row` is the standard "find the last
   used row" idiom: it starts at the bottom of the sheet and jumps up
   to the last non-empty cell in column A, so `+1` gives the first
   empty row — here, row 6 has data, so `nextRow` = `7`.
3. Run it, enter `Stapler` and `15` — row 7 should now read
   `Stapler, 15`.

## Cheat sheet

| Concept | Syntax |
|---|---|
| Read/write a cell | `Range("B2").Value` or `Cells(2,2).Value` |
| Declare a variable | `Dim name As Type` |
| Loop a fixed range | `For i = 2 To 6 ... Next i` |
| Conditional | `If cond Then ... Else ... End If` |
| Last used row | `Cells(Rows.Count, col).End(xlUp).Row` |
| Prompt the user | `InputBox("prompt text")` |

## Exercise

Write `Sub AverageQty()` that loops `B2:B6`, sums the values, divides
by the count (`5`), and shows the result in a `MsgBox`. Manually
verify: sum `48` from Section 3, divided by `5` = `9.6`. The macro's
message box must show `9.6`.
