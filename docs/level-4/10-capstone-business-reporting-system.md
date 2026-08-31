# 10 · Capstone — Full Business Reporting System in Excel

This capstone combines the path's Level 4 modules — architecture,
DAX, VBA, performance, integration, and governance — into one
end-to-end system: raw data in, a governed Data Model in the middle,
and a protected, automated dashboard out.

## 1. Worked dataset — monthly regional sales

Sheet `RawSales` (the single source of truth, per Module 08), `A1:C13`:

| | A | B | C |
|---|---|---|---|
| **1** | Month | Region | Revenue |
| **2** | Jan | North | 22000 |
| **3** | Jan | South | 17000 |
| **4** | Feb | North | 24500 |
| **5** | Feb | South | 18200 |
| **6** | Mar | North | 21000 |
| **7** | Mar | South | 19100 |
| **8** | Apr | North | 26000 |
| **9** | Apr | South | 20500 |
| **10** | May | North | 27500 |
| **11** | May | South | 21800 |
| **12** | Jun | North | 25000 |
| **13** | Jun | South | 22900 |

## 2. Architecture — one Table, one Data Model (Module 01)

1. Convert `A1:C13` to a Table named `SalesTbl` (**Ctrl+T**). Every
   dashboard element below references this Table, never a raw range —
   new months appended to it flow through automatically.
2. Load `SalesTbl` into Power Pivot as the single Data Model backing
   every visual in the workbook (no duplicate copies per report tab).

## 3. DAX measures (Module 02)

`Total Revenue := SUM(SalesTbl[Revenue])`
Grand total: `22000+17000+24500+18200+21000+19100+26000+20500+27500+
21800+25000+22900 = 265500`.

`North Revenue := CALCULATE([Total Revenue], SalesTbl[Region]="North")`
North total: `22000+24500+21000+26000+27500+25000=146000`.

`North Share % := DIVIDE([North Revenue],[Total Revenue])`
`146000/265500 ≈ 0.5499`, i.e. North is **~55%** of total revenue.

## 4. A VBA refresh-and-report button (Module 04, 05)

```vba
Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ThisWorkbook.Connections("SalesTbl").Refresh
    ThisWorkbook.RefreshAll

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed. Total Revenue: " & _
        Format(Application.WorksheetFunction.Sum(Range("RawSales!C2:C13")), "$#,##0")
End Sub
```

Wrapping the refresh in `ScreenUpdating=False`/manual calculation
(Module 05's performance pattern) keeps a one-click refresh fast even
as `RawSales` grows well beyond 12 rows.

## 5. Governance and protection (Module 08, 07)

1. Unlock only the input area of `RawSales` (new rows appended at the
   bottom) and **Protect Sheet** on every other tab — the DAX measures
   and dashboard visuals cannot be accidentally edited by a report
   viewer.
2. Add a `_ChangeLog` sheet (Module 08 pattern) so any edit to
   `RawSales` is timestamped and attributed.
3. Document, on a `README` tab, that `RawSales` is the single
   authoritative source — any other report referencing this data must
   pull from the Data Model, never re-paste values.

## 6. The finished dashboard

- A KPI row: `Total Revenue` (`$265,500`), `North Share %` (`~55.0%`),
  built as formula-driven cells per Module 08 of Level 3.
- A line chart of monthly `Total Revenue` trend by Region, using a
  PivotChart off the Data Model — North should visibly run above South
  every month in this dataset, consistent with its ~55% overall share.
- A slicer on `Region` (Level 3 Module 08 technique) so a viewer can
  isolate North or South without touching any formula.
- A `RefreshDashboard` button (Section 4) wired to a shape via
  **Assign Macro**, so refreshing after new data is appended is one
  click, not a manual multi-step process.

## Cheat sheet — how the whole system fits together

| Layer | Module this capstone draws from |
|---|---|
| One Table, one Data Model | 01 · Enterprise Excel Architecture |
| DAX measures for KPIs | 02 · Advanced DAX & Power Pivot Modeling |
| One-click refresh, performance-safe | 04/05 · VBA + Performance Optimization |
| Protected structure, change log | 07/08 · Team Tools + Data Governance |
| Slicers, formula-driven titles | Level 3 · 08 Advanced Dashboard Design |

## Exercise

Append a `Jul` row for both regions (North `28000`, South `23500`),
run `RefreshDashboard`, and confirm `Total Revenue` updates to
`265500+28000+23500=317000` and `North Share %` shifts to
`(146000+28000)/317000 = 174000/317000 ≈ 0.5489` (~54.9%) — a slight
dip from ~55.0% since South grew proportionally faster this month.
