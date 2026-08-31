# 08 · Data Governance in Spreadsheet-Heavy Organizations

This module covers the governance layer that keeps spreadsheet-based
reporting trustworthy at scale: a single source of truth, access
controls, audit trails, and version control practices for `.xlsx`
files.

## 1. Worked dataset — the "many copies" problem

Imagine three regional managers each maintain their own copy of the
same `Revenue` figures:

| | A | B |
|---|---|---|
| **1** | Region | Q1 Revenue |
| **2** | North (Manager A's copy) | 120000 |
| **3** | North (Manager B's copy, one week stale) | 115000 |
| **4** | North (Finance's copy, includes a since-corrected error) | 121500 |

All three claim to be "the North Q1 number," and a report combining
data from all three managers will silently pick whichever copy each
contributor pasted in — the classic spreadsheet governance failure.

## 2. Single source of truth

1. Designate one workbook (or better, one Power Pivot Data Model /
   database table) as authoritative, and have every other report pull
   from it via Power Query rather than copy-pasting values.
2. `=SUMIFS('[MasterRevenue.xlsx]Data'!$B:$B,...)` style external
   references, or a Power Query connection to the same file/table, mean
   a correction made once in the master (fixing `115000` → `120000`)
   propagates everywhere on refresh instead of requiring every
   downstream copy to be manually fixed.

## 3. Access control

| Control | What it prevents |
|---|---|
| **Protect Workbook** (structure) | Sheets can't be added, deleted, renamed, or reordered |
| **Protect Sheet** with unlocked input cells | Formulas can't be overwritten; only designated inputs are editable |
| File-level permissions (SharePoint/OneDrive "Can View" vs "Can Edit") | Only authorized people can even open the file for editing |
| **Mark as Final** | A soft warning (not real protection) that a document is meant to be read-only |

`Protect Sheet`/`Protect Workbook` passwords are deterrents, not
cryptographic security — anyone with real access to the file's
contents can eventually remove them. Real access control must happen
at the file/sharing-permission level (SharePoint/OneDrive), not by
relying on a sheet-protection password alone.

## 4. Audit trail

1. **Review → Track Changes** (legacy `.xls`-era feature; in modern
   `.xlsx`/365 use **Version History** in OneDrive/SharePoint instead)
   shows who changed what and when.
2. A dedicated `_ChangeLog` sheet with columns `Date | User | Field
   Changed | Old Value | New Value`, filled in manually or via a VBA
   `Worksheet_Change` event handler, gives a durable record even after
   version history is eventually pruned:
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       Dim ws As Worksheet
       Set ws = ThisWorkbook.Sheets("_ChangeLog")
       ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Now
       ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = Application.UserName
       ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = Target.Address
       ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 3).Value = Target.Value
   End Sub
   ```

## 5. Version control for spreadsheets

- OneDrive/SharePoint **Version History** keeps every saved version and
  lets you restore a prior one — the practical equivalent of `git log`
  / `git checkout` for a binary `.xlsx` file, which does not diff
  cleanly in traditional git.
- For workbooks that do live in git (e.g. alongside code), store the
  underlying data as `.csv` where possible (which *does* diff readably)
  and regenerate the `.xlsx` presentation layer from it, rather than
  committing the binary `.xlsx` directly.

## 6. A governance checklist

1. Is there exactly one authoritative source for each metric?
2. Can every consumer trace their number back to that source (a
   formula/connection), rather than a pasted value?
3. Is edit access restricted to the people who should be changing raw
   inputs?
4. Is there a way to see who changed what, and when, after the fact?
5. Is there a recovery path (version history/backup) if a change turns
   out to be wrong?

## Cheat sheet

| Governance need | Excel/365 mechanism |
|---|---|
| Single source of truth | Power Query connection to one master file/table |
| Access control | SharePoint/OneDrive sharing permissions + Protect Sheet/Workbook |
| Audit trail | Version History, or a `_ChangeLog` sheet via `Worksheet_Change` |
| Recoverability | OneDrive/SharePoint Version History |

## Exercise

Using the "many copies" scenario in Section 1, write the Power Query
`M` step (or plain formula) each regional manager's report should use
instead of a pasted number, so that fixing the master value once (from
`115000`/`121500` confusion to the correct `120000`) automatically
corrects every manager's report on refresh.
