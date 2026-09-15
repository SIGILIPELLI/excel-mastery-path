---
description: "Building Excel-Based Tools for Teams — This module covers packaging a workbook so a team can use it safely and consistently: protected structure, an…"
---

# 07 · Building Excel-Based Tools for Teams

This module covers packaging a workbook so a team can use it safely and
consistently: protected structure, an Add-in for shared custom
functions, and a co-authoring-friendly layout.

## 1. Worked dataset — a shared budget request form

Sheet `Form`, `A1:B5`:

| | A | B |
|---|---|---|
| **1** | Requestor | (blank, to be filled in) |
| **2** | Department | (blank) |
| **3** | Amount Requested | (blank) |
| **4** | Approved? | `=IF(B3<=5000,"Auto-approved","Needs manager sign-off")` |
| **5** | | |

With `B3=3200`, `B4` reads `"Auto-approved"`; with `B3=7500`, `B4`
reads `"Needs manager sign-off"` — the rule cell is the shared piece
of logic every requester relies on without needing to know the
threshold.

## 2. Protecting structure while keeping inputs editable

1. Select `B1:B3` (the only cells a team member should edit) →
   **Format Cells → Protection → uncheck Locked**.
2. **Review → Protect Sheet** (optionally with a password) locks every
   *other* cell, including the `B4` formula — a user can fill in the
   form but cannot accidentally overwrite the approval rule.
3. This is the standard shared-tool pattern: unlock only the input
   cells before protecting the sheet, since every cell is Locked by
   default.

## 3. Data validation for consistent input

1. `B2` (Department): **Data → Data Validation → List** →
   `Engineering,Sales,Support,Finance` — restricts entries to a fixed
   set so downstream `SUMIFS`/PivotTables aggregating by department
   never split "Sales" and "sales" into two categories.
2. `B3` (Amount): **Data Validation → Decimal → greater than 0** —
   prevents a blank or negative submission from silently producing a
   nonsensical `Approved?` result.

## 4. Packaging shared logic as an Add-in

1. Write a reusable custom function in a standard VBA module, e.g.:
   ```vba
   Function ApprovalStatus(amount As Double) As String
       If amount <= 5000 Then
           ApprovalStatus = "Auto-approved"
       Else
           ApprovalStatus = "Needs manager sign-off"
       End If
   End Function
   ```
2. Save the workbook containing this module as an **Excel Add-in
   (.xlam)** (**File → Save As → Excel Add-In**).
3. Once loaded (**File → Options → Add-ins → Manage: Excel Add-ins →
   Browse**) by every team member, `=ApprovalStatus(B3)` is available
   in *any* workbook they open — this is how a team standardizes a
   business rule once instead of pasting the same `IF` formula into
   every new request form, and a future rule change (say, raising the
   threshold to `7500`) only needs to be made in the one `.xlam` file.

## 5. Co-authoring considerations

- Store the shared workbook in OneDrive/SharePoint, not a local drive,
  so **Co-authoring** allows multiple people to edit simultaneously
  with per-cell change tracking, instead of the older "locked for
  editing by another user" exclusive-lock behavior.
- Structured References (Tables) co-author more reliably than raw cell
  ranges, since a Table's formulas auto-adjust as rows are added by
  different users at the same time.
- Avoid VBA macros in a co-authored file when possible — VBA project
  changes do not merge between simultaneous editors the way cell edits
  do, and can silently desync.

## Cheat sheet

| Technique | Purpose |
|---|---|
| Unlock inputs, then Protect Sheet | Let users fill in data without touching the logic |
| Data Validation (List/Decimal) | Force consistent, valid input across every user |
| `.xlam` Add-in | Ship one shared custom function to every team member's Excel |
| OneDrive/SharePoint + Tables | Reliable simultaneous co-authoring |

## How It Actually Works

Shared, co-authored Excel files (via OneDrive/SharePoint) work through
**operational transformation-style conflict resolution**, not a single
shared in-memory dependency graph — each collaborator's Excel client
maintains its own local copy of the workbook and its own local dependency
graph, and co-authoring syncs a stream of individual *cell-level edit
operations* between clients rather than syncing the whole file or a shared
live calculation session. This is why two people editing genuinely
different cells merge seamlessly (their edit operations simply both apply,
each recalculating locally) while two people editing the *same* cell
resolve as a last-write-wins conflict with no true merge — there's no
concept of merging two conflicting formula edits the way version control
merges text lines, because a formula is treated as one atomic value, not
composable text. Version History (also OneDrive/SharePoint-backed) works
by the storage layer keeping full snapshots of the file at each save point,
not by tracking incremental cell diffs — restoring an old version replaces
the entire workbook wholesale, including its calculation chain, which is
why restoring from history is safe (any half-finished edit is discarded
completely) but coarse (there's no way to restore just one changed range).

## Exercise

Change the Add-in's `ApprovalStatus` threshold from `5000` to `7500`,
reload the Add-in, and confirm a request of `B3=6000` now reads
`"Auto-approved"` instead of `"Needs manager sign-off"` — without
editing the request-form workbook itself.
