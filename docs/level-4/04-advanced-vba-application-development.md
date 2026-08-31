# 04 · Advanced VBA Application Development

This module goes past Level 3's UserForms and error handling (Module
05) into building a small multi-procedure VBA application: a class
module, a collection of objects, and a UserForm that drives them —
the shape of a real internal tool.

## 1. Worked scenario — an expense-tracking mini app

Sheet `Data`, `A1:C1` headers `Item | Category | Amount`, initially
empty below the header — the VBA app will populate it.

## 2. A class module to represent one expense

Insert a class module named `Expense`:

```vba
Private pItem As String
Private pCategory As String
Private pAmount As Double

Public Property Let Item(v As String): pItem = v: End Property
Public Property Get Item() As String: Item = pItem: End Property
Public Property Let Category(v As String): pCategory = v: End Property
Public Property Get Category() As String: Category = pCategory: End Property
Public Property Let Amount(v As Double): pAmount = v: End Property
Public Property Get Amount() As Double: Amount = pAmount: End Property
```

Wrapping the three fields in a class (instead of three parallel arrays)
means every expense travels as one object — pass it to a function, add
it to a collection, or write it to a sheet in one call, without three
separate variables staying manually in sync.

## 3. A standard module managing a collection of them

```vba
Public allExpenses As Collection

Sub InitExpenses()
    Set allExpenses = New Collection
End Sub

Sub AddExpense(itemName As String, cat As String, amt As Double)
    Dim e As New Expense
    e.Item = itemName
    e.Category = cat
    e.Amount = amt
    allExpenses.Add e
End Sub

Function TotalByCategory(cat As String) As Double
    Dim e As Expense
    Dim total As Double
    total = 0
    For Each e In allExpenses
        If e.Category = cat Then total = total + e.Amount
    Next e
    TotalByCategory = total
End Function
```

Manual trace: after `AddExpense "Paper","Office",25`,
`AddExpense "Taxi","Travel",40`, `AddExpense "Pens","Office",8`,
calling `TotalByCategory("Office")` loops all three, matches items 1
and 3 (`Category="Office"`), and returns `25+8=33`.

## 4. Writing the collection back to the sheet

```vba
Sub DumpToSheet()
    Dim e As Expense
    Dim r As Long
    r = 2
    For Each e In allExpenses
        Sheets("Data").Cells(r, 1).Value = e.Item
        Sheets("Data").Cells(r, 2).Value = e.Category
        Sheets("Data").Cells(r, 3).Value = e.Amount
        r = r + 1
    Next e
End Sub
```

After the three `AddExpense` calls above, `DumpToSheet` writes three
rows starting at `A2`, and a plain `=SUM(C2:C4)` on the sheet
independently confirms the grand total: `25+40+8=73`.

## 5. Wiring a UserForm to the collection

A UserForm `frmAddExpense` with text boxes `txtItem`, `txtCategory`,
`txtAmount` and a button `btnAdd`:

```vba
Private Sub btnAdd_Click()
    If Not IsNumeric(txtAmount.Value) Then
        MsgBox "Amount must be a number.", vbExclamation
        Exit Sub
    End If
    AddExpense txtItem.Value, txtCategory.Value, CDbl(txtAmount.Value)
    MsgBox "Added. Office total so far: " & TotalByCategory("Office")
    Unload Me
End Sub
```

This is the pattern that makes VBA feel like a real application rather
than a macro: the UserForm is only responsible for collecting and
validating input; `AddExpense` and `TotalByCategory` (business logic)
live in reusable, independently testable procedures.

## 6. Error handling around the whole flow

```vba
Sub RunApp()
    On Error GoTo ErrHandler
    InitExpenses
    frmAddExpense.Show
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error: " & Err.Description, vbCritical
End Sub
```

`On Error GoTo ErrHandler` catches anything unanticipated (a corrupted
form, a missing sheet) so the app fails with a readable message instead
of an unhandled runtime error dialog.

## Cheat sheet

| Concept | Purpose |
|---|---|
| Class module (`Property Let/Get`) | Bundle related fields into one reusable object type |
| `Collection` | Hold a growable list of custom objects |
| `For Each ... In collection` | Iterate every object in the collection |
| UserForm calling standard-module procedures | Separate input/validation from business logic |

## Exercise

Add a `RemoveExpense(index As Integer)` procedure using
`allExpenses.Remove(index)`, then confirm that after removing item 1
(`Paper`), `TotalByCategory("Office")` correctly drops from `33` to
`8`.
