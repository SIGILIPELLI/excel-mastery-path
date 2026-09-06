# 05 · Advanced VBA (UserForms, Error Handling)

This module builds a UserForm — a custom dialog box with labeled input
fields and buttons — and adds structured error handling so a macro
fails gracefully instead of crashing with a raw VBA error.

## 1. Worked dataset

Build this on `Sheet1`, `A1:B4`:

| | A | B |
|---|---|---|
| **1** | Item | Qty |
| **2** | Pen | 12 |
| **3** | Book | 3 |
| **4** | Bag | 25 |

## 2. Building a UserForm

1. Alt+F11 → Insert → **UserForm**. From the Toolbox, add two Labels
   ("Item:", "Qty:"), two TextBoxes (`txtItem`, `txtQty` — rename via
   the Properties window's `(Name)` field), and a CommandButton
   (`cmdAdd`, Caption "Add").
2. Double-click `cmdAdd` to open its `Click` event and type:
   ```vb
   Private Sub cmdAdd_Click()
       Dim nextRow As Long
       If txtItem.Value = "" Or txtQty.Value = "" Then
           MsgBox "Both fields are required."
           Exit Sub
       End If
       nextRow = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row + 1
       Sheet1.Cells(nextRow, 1).Value = txtItem.Value
       Sheet1.Cells(nextRow, 2).Value = txtQty.Value
       txtItem.Value = ""
       txtQty.Value = ""
   End Sub
   ```
3. Add a Sub elsewhere in a regular Module to launch the form:
   ```vb
   Sub ShowAddItemForm()
       UserForm1.Show
   End Sub
   ```
4. Run `ShowAddItemForm`, enter `Stapler` / `15`, click Add. Row 5
   should now read `Stapler, 15`. Leave both fields blank and click
   Add again — the validation message must appear and no row should
   be written.

## 3. Structured error handling

1. A macro that divides by a user-entered quantity can crash on zero
   or non-numeric input. Add error handling:
   ```vb
   Sub SafeDivide()
       Dim result As Double
       On Error GoTo ErrHandler
       result = 100 / Range("B2").Value
       MsgBox "Result: " & result
       Exit Sub
   ErrHandler:
       MsgBox "Error " & Err.Number & ": " & Err.Description
   End Sub
   ```
2. With `B2 = 12`: `100/12 = 8.3333...`, message shows `Result:
   8.33333333333333`. Manually change `B2` to `0` and re-run: `100/0`
   raises runtime error 11 ("Division by zero"), and `On Error GoTo
   ErrHandler` routes execution to the handler instead of crashing —
   the message box should read `Error 11: Division by zero`.
3. `On Error Resume Next` is a looser alternative that skips the
   failing line and continues — useful only when a failure on one line
   is genuinely safe to ignore; `On Error GoTo Label` is preferred
   whenever the error should be reported, not silently swallowed.

## 4. Cleaning up with `On Error GoTo 0`

1. Reset error handling after a risky block so later errors in the
   same Sub aren't silently caught by a handler meant for one specific
   line:
   ```vb
   Sub PartialProtection()
       On Error GoTo ErrHandler
       Dim x As Double
       x = 100 / Range("B2").Value
       On Error GoTo 0   ' error handling off from here on
       Dim y As Double
       y = 50 / Range("B3").Value   ' if this errors, it's unhandled
       MsgBox x & " / " & y
       Exit Sub
   ErrHandler:
       MsgBox "Error in first calculation: " & Err.Description
       Resume Next
   End Sub
   ```
2. With `B2=12` (from Section 3) and `B3=3`: `x=8.3333`, `y=50/3=
   16.6667`. Manual check confirms both values before trusting the
   macro's message box.

## Cheat sheet

| Construct | Purpose |
|---|---|
| `UserForm1.Show` | Display a custom dialog |
| `TextBox.Value` | Read/write a form field |
| `On Error GoTo Label` | Route runtime errors to a handler |
| `Err.Number` / `Err.Description` | Inspect the error that was caught |
| `On Error Resume Next` | Skip the failing line (use sparingly) |
| `On Error GoTo 0` | Disable the active error handler from this point on |

## How It Actually Works

VBA's error handling is built around a single per-procedure error state,
not exceptions in the try/catch sense — `On Error Goto Label` doesn't wrap
a block of code the way a try/catch does; it sets a standing instruction
telling the VBA runtime "if any statement in this procedure from this
point forward raises a runtime error, jump execution to this label," which
stays active until explicitly cleared (`On Error Goto 0`) or the procedure
ends. This is why a single `On Error Resume Next` at the top of a long
routine is dangerous: it doesn't just guard the next risky line, it
silently swallows *every* subsequent runtime error for the rest of that
procedure until something resets it, which is why disciplined VBA wraps
just the risky lines and immediately restores normal error propagation
afterward. UserForms are a separate COM subsystem again — each control
(TextBox, ComboBox) is its own object firing its own events (`Change`,
`Click`) on the VBA message loop, and a form is **modal** by default,
meaning it takes over the application's event loop entirely and no other
Excel code runs until it's closed — which is exactly why a modal form left
open by a bug can look like Excel has frozen: the rest of the application
genuinely is blocked, waiting for that one form's message loop to release
control.

## Exercise

Add a `Cancel` CommandButton to the UserForm from Section 2 whose
`Click` event runs `Unload Me`. Manually confirm: opening the form,
typing text, then clicking Cancel closes the form without writing
anything to `Sheet1` — the row count in column A should be unchanged
from before the form was opened.
