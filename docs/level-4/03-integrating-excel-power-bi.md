# 03 · Integrating Excel with Power BI

This module covers moving a Power Pivot model built in Excel into
Power BI, and the reverse flow — analyzing a Power BI dataset from
Excel — so the two tools share one model instead of duplicating logic.

## 1. Worked dataset

Reuse the `Sales` table from Module 02, `A1:C13`, already loaded into
Excel's Power Pivot Data Model with its DAX measures (`Total Sales`,
`North Sales`, `YoY %`).

## 2. Excel → Power BI: bringing the model over

1. In Power BI Desktop, **Get Data → Excel Workbook** → point to the
   `.xlsx` file. If the workbook's Data Model contains tables and
   measures, Power BI imports the Data Model directly (not just the
   raw sheet cells), including the relationships and DAX measures
   already defined — `Total Sales` arrives in Power BI still equal to
   `13050`, unchanged.
2. Alternatively, **File → Options → Export** the Power Pivot model as
   a `.bism`/via **Analyze in Excel** round trip is the inverse
   direction (Section 3) — for a one-time full transfer, opening the
   `.xlsx` directly in Power BI Desktop is simplest.

## 3. Power BI → Excel: "Analyze in Excel"

1. From a Power BI dataset published to the Power BI Service, use
   **Analyze in Excel** (or **Export → Analyze in Excel** from the
   dataset's context menu) to download an `.xlsx` file containing a
   PivotTable connected live to the Power BI dataset via a data
   connection — not a static copy.
2. Any DAX measure defined in the Power BI dataset (e.g. `YoY %`) shows
   up as a value field exactly as it would in a native Power BI report,
   and refreshing the PivotTable (**Data → Refresh All**) pulls the
   latest numbers from the Power BI Service.

## 4. Why do this instead of keeping two separate models?

| Problem with two separate models | Fixed by a shared model |
|---|---|
| The DAX for `YoY %` gets written slightly differently in Excel vs. Power BI, so the two tools disagree | One dataset, one measure definition, both tools query it |
| Refreshing Excel's copy doesn't update Power BI's copy (or vice versa) | A single source refreshes once; every consumer sees the update |
| Row-level security rules defined in Power BI don't apply to the Excel copy | Users connecting via Analyze in Excel inherit the same RLS rules as the Power BI report |

## 5. Publishing an Excel workbook that already has a Data Model

1. **File → Publish → Publish to Power BI** (requires a Power BI
   account with a workspace) uploads the workbook's Data Model and any
   worksheets as a Power BI dataset/report pairing, preserving the
   measures built in Excel.
2. From there, further visuals can be built natively in Power BI, while
   the Excel file remains usable standalone — the two are now two
   front ends on the same measures, not two independent calculations.

## 6. A practical decision rule

- Keep it in Excel if the audience is a handful of people who need to
  poke at the numbers directly (add a column, test a scenario).
- Move to Power BI once the audience grows past what email/shared-drive
  distribution of a workbook can sanely support, or once row-level
  security per viewer is required.
- Either way, define each business metric's DAX **once** and reuse it
  everywhere — the integration exists specifically so `Total Sales` and
  `YoY %` never have to be redefined by hand a second time.

## Cheat sheet

| Action | Direction |
|---|---|
| Get Data → Excel Workbook | Excel Data Model → Power BI Desktop |
| Analyze in Excel | Power BI Service dataset → live Excel PivotTable |
| Publish to Power BI | Excel workbook → Power BI dataset/report |

## Exercise

Using Analyze in Excel against a published `Sales` dataset, build a
PivotTable with `Region` in Rows and the `YoY %` measure in Values.
Confirm North's `YoY %` still reads `~19.3%` and South's reads
`~18.9%`, matching the values computed by hand in Module 02 — proof the
measure logic did not change in the round trip.
