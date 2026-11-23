---
title: "Excel Cheat Sheet: Quick Reference, Beginner to Advanced"
description: "A free Excel cheat sheet: 40 quick-reference tables from every level of the Excel Mastery Path, each linked to its full lesson."
---

# Excel Cheat Sheet

Quick-reference tables from **40 lessons** of the free Excel Mastery Path, in course order from beginner to advanced. Each section links to the full lesson with explanations and runnable examples.

## Level 1 · Entry

### [01 · Excel Fundamentals](level-1/01-excel-fundamentals.md)

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

### [02 · Formulas &amp; Functions Basics](level-1/02-formulas-functions-basics.md)

| Function/Syntax | Purpose | Example |
|---|---|---|
| `=A1-B1` | Basic arithmetic formula | `=C2-B2` → difference |
| `SUM(range)` | Add all values in a range | `=SUM(B2:B6)` → `2150` |
| `AVERAGE(range)` | Mean of values in a range | `=AVERAGE(C2:C6)` → `452` |
| `MAX(range)` / `MIN(range)` | Largest / smallest value | `=MAX(C2:C6)` → `1200` |
| `COUNT(range)` | Count of numeric cells | `=COUNT(B2:B6)` → `5` |
| `A1` | Relative reference — shifts when copied | |
| `$A$1` | Absolute reference — never shifts | |
| `$A1` / `A$1` | Mixed reference — column or row locked | |
| F4 | Cycle reference type in Formula Bar | |

### [03 · Formatting &amp; Conditional Formatting](level-1/03-formatting-conditional-formatting.md)

| Format | Shortcut / Location |
|---|---|
| Bold / Italic / Underline | Ctrl/⌘+B / +I / +U |
| Number format (2 decimals) | Ctrl/⌘+Shift+1 |
| Currency format | Ctrl/⌘+Shift+4 |
| Percentage format | Ctrl/⌘+Shift+5 |
| Format Painter | Home > Clipboard group |
| Highlight Cells Rules (Greater/Less Than) | Home > Conditional Formatting |
| Data Bars / Color Scales / Icon Sets | Home > Conditional Formatting |
| Manage/edit existing rules | Conditional Formatting > Manage Rules |
| Clear conditional formatting | Conditional Formatting > Clear Rules |

### [04 · Working with Tables](level-1/04-working-with-tables.md)

| Action | Shortcut / Location |
|---|---|
| Convert range to Table | Ctrl/⌘+T |
| Rename Table | Table Design > Table Name |
| Whole-column structured reference | `TableName[ColumnName]` |
| Same-row structured reference | `TableName[@ColumnName]` |
| Toggle Total Row | Table Design > Total Row |
| Toggle filter arrows | Table Design > Filter Button |
| Sort by column | Header filter-arrow > Sort Largest/Smallest |
| Filter to specific values | Header filter-arrow > uncheck values |

### [05 · Charts &amp; Basic Visualization](level-1/05-charts-basic-visualization.md)

| Action | Location |
|---|---|
| Insert column/bar chart | Insert > Charts > Column or Bar |
| Insert pie chart | Insert > Charts > Pie or Doughnut |
| Add/remove title, labels, gridlines | Chart Elements (+) button |
| Change colors/style | Chart Styles (paintbrush) button |
| Filter which series/categories show | Chart Filters (funnel) button |
| Change an existing chart's type | Chart Design > Change Chart Type |
| Move chart to its own sheet | Chart Design > Move Chart |

### [06 · Lookup Functions](level-1/06-lookup-functions.md)

| Function | Syntax | Notes |
|---|---|---|
| `VLOOKUP` (exact) | `=VLOOKUP(value,table,col_num,FALSE)` | Looks rightward only |
| `VLOOKUP` (approximate) | `=VLOOKUP(value,table,col_num,TRUE)` | Table must be sorted ascending |
| `XLOOKUP` | `=XLOOKUP(value,lookup_array,return_array,[if_not_found])` | 365-only; any direction |
| `MATCH` | `=MATCH(value,lookup_array,0)` | Returns a **position**, not a value |
| `INDEX` | `=INDEX(array,row_num)` | Returns the value at a position |
| `INDEX`+`MATCH` | `=INDEX(return_range,MATCH(value,lookup_range,0))` | Works in every Excel version |

### [07 · Logical Functions](level-1/07-logical-functions.md)

| Function | Syntax | Behavior |
|---|---|---|
| `IF` | `=IF(condition,if_true,if_false)` | Two-way branch |
| Nested `IF` | `=IF(c1,r1,IF(c2,r2,r3))` | Three+ way branch, gets unwieldy fast |
| `IFS` | `=IFS(c1,r1,c2,r2,TRUE,default)` | Cleaner multi-branch, first match wins |
| `AND` | `=AND(c1,c2,…)` | TRUE only if all conditions are TRUE |
| `OR` | `=OR(c1,c2,…)` | TRUE if any condition is TRUE |

### [08 · PivotTables Basics](level-1/08-pivottables-basics.md)

| Action | Location |
|---|---|
| Insert a PivotTable | Insert > PivotTable |
| Add a summarized field | Drag field into Values (defaults to Sum) |
| Group by a dimension | Drag field into Rows or Columns |
| Restrict to specific values | Drag field into Filters |
| Change summary type (Sum/Average/Count) | Click Values field > Value Field Settings |
| Reformat all values at once | Right-click a value > Number Format |
| Pull in changed source data | Right-click > Refresh |

### [09 · Data Validation &amp; Cleaning](level-1/09-data-validation-cleaning.md)

| Function/Feature | Purpose |
|---|---|
| `TRIM(text)` | Remove leading/trailing/extra spaces |
| `UPPER(text)` / `LOWER(text)` | Force upper/lower case |
| `PROPER(text)` | Capitalize first letter of each word |
| Data Validation > List | Restrict entry to a fixed dropdown list |
| Data Validation > Error Alert | Customize Stop/Warning behavior and message |
| Remove Duplicates | Delete exact-match duplicate rows |
| Circle Invalid Data | Flag existing cells that violate a validation rule |

### [10 · Project — Personal Budget Tracker](level-1/10-project-personal-budget-tracker.md)

| Category | Tools covered |
|---|---|
| Navigation | Name Box, Ctrl/⌘+Arrow, ranges |
| Formulas | `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, relative/absolute refs |
| Formatting | Number formats, Conditional Formatting (Highlight/Data Bars/Color Scales) |
| Tables | `Ctrl/⌘+T`, structured references, Total Row |
| Charts | Column, Pie, Chart Elements |
| Lookups | `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, approximate-match brackets |
| Logic | `IF`, nested `IF`, `IFS`, `AND`, `OR` |
| PivotTables | Rows/Columns/Values/Filters, Refresh |
| Cleaning | `TRIM`, `PROPER`, `UPPER`/`LOWER`, Remove Duplicates, Data Validation |

## Level 2 · Intermediate

### [01 · Advanced Formulas (SUMIFS/COUNTIFS/AVERAGEIFS)](level-2/01-advanced-formulas.md)

| Function | Syntax |
|---|---|
| `SUMIFS` | `=SUMIFS(sum_range,crit_rng1,crit1,[crit_rng2,crit2,...])` |
| `COUNTIFS` | `=COUNTIFS(crit_rng1,crit1,[crit_rng2,crit2,...])` |
| `AVERAGEIFS` | `=AVERAGEIFS(avg_range,crit_rng1,crit1,...)` |
| `IFS` | `=IFS(cond1,val1,cond2,val2,...,TRUE,default)` |
| `IFERROR` | `=IFERROR(formula,value_if_error)` |
| `AND`/`OR` | `=IF(AND(c1,c2),t,f)` / `=IF(OR(c1,c2),t,f)` |

### [02 · PivotTables Deep Dive](level-2/02-pivottables-deep-dive.md)

| Action | Where |
|---|---|
| Add a 2nd value field | Drag field to Values again, change summary type |
| Nest rows | Drag second field below the first in Rows |
| Custom formula per pivot | PivotTable Analyze > Calculated Field |
| Interactive filter widget | PivotTable Analyze > Insert Slicer |
| Change Sum→Average/Count | Right-click value > Value Field Settings |

### [03 · Advanced Charting &amp; Dashboards](level-2/03-advanced-charting-dashboards.md)

| Task | Steps |
|---|---|
| Combo chart | Select series → Insert → Combo Chart → set type + axis per series |
| Secondary axis | Chart element → Format Data Series → Secondary Axis |
| Sparkline | Insert → Sparklines → Line/Column/Win-Loss → set Data Range |
| Linked chart title | Click title → type `=Sheet!$Cell` in formula bar |

### [04 · Data Validation Advanced &amp; Form Controls](level-2/04-data-validation-advanced-form-controls.md)

| Task | How |
|---|---|
| List validation | Data Validation → List → Source range or comma list |
| Cascading dropdown | Named ranges per category + `=INDIRECT(cell)` as Source |
| Custom rule | Data Validation → Custom → formula returning TRUE/FALSE |
| No duplicates | `=COUNTIF($range,cell)=1` as Custom formula |
| Combo Box (Form Control) | Input range + Cell link (returns position number) |
| Position → text | `=INDEX(range,cell_link)` |

### [05 · Power Query Basics](level-2/05-power-query-basics.md)

| Task | Ribbon path |
|---|---|
| Load a table into Power Query | Data → From Table/Range |
| Fix text casing | Transform → Format → Capitalize/UPPER/lower |
| Replace blanks/errors | Right-click column → Replace Values |
| Aggregate by category | Transform → Group By |
| Send result back to a sheet | Home → Close & Load |
| Re-run after source changes | Data → Refresh All |

### [06 · Array Formulas &amp; Dynamic Arrays](level-2/06-array-formulas-dynamic-arrays.md)

| Function | Purpose |
|---|---|
| `FILTER(array,include,[if_empty])` | Rows matching a boolean condition |
| `UNIQUE(array)` | Distinct values, first-seen order |
| `SORT(array,[col],[order])` | Sort a range or spilled array |
| `SEQUENCE(rows,[cols],[start],[step])` | Generate a number sequence |
| `range#` | Spill range operator — refers to a dynamic array's full output |

### [07 · What-If Analysis](level-2/07-what-if-analysis.md)

| Tool | Use when |
|---|---|
| Goal Seek | You know the desired output, need to find the input |
| Data Table (1 or 2 variable) | You want the output across a range of one or two inputs |
| Scenario Manager | You want to save and compare a handful of named input sets |

### [08 · Working with Multiple Sheets &amp; Workbooks](level-2/08-multiple-sheets-workbooks.md)

| Reference type | Syntax | Behavior |
|---|---|---|
| 3-D reference | `=SUM(Jan:Mar!B2)` | Aggregates one cell across a sheet range |
| Cross-sheet | `=Feb!B2` | Points to one cell on one other sheet |
| Cross-workbook (open) | `='[File.xlsx]Sheet'!B2` | Live link while both files open |
| Cross-workbook (closed) | `='C:\path\[File.xlsx]Sheet'!B2` | Full path required |
| Fix broken links | Data → Edit Links → Change Source | Repoints after a file move/rename |

### [09 · Intro to Recording Macros](level-2/09-intro-recording-macros.md)

| Task | Path |
|---|---|
| Record a macro | Developer → Record Macro |
| Stop recording | Developer → Stop Recording |
| Toggle relative recording | Developer → Use Relative References |
| View/edit recorded code | Developer → Macros → Edit |
| Run a macro | Developer → Macros → Run, or assigned shortcut/button |
| Attach to a button | Insert Form Control Button → Assign Macro |

### [10 · Project — Interactive Sales Dashboard](level-2/10-project-interactive-sales-dashboard.md)

| Dashboard piece | Built with |
|---|---|
| Region KPI tile | `SUMIFS` referencing the dropdown cell |
| Month breakdown | PivotTable with Region as filter/Slicer |
| Trend visualization | PivotChart or combo chart off the Pivot |
| User control | Data Validation dropdown +/or Slicer |

## Level 3 · Advanced

### [01 · VBA Basics for Automation](level-3/01-vba-basics-automation.md)

| Concept | Syntax |
|---|---|
| Read/write a cell | `Range("B2").Value` or `Cells(2,2).Value` |
| Declare a variable | `Dim name As Type` |
| Loop a fixed range | `For i = 2 To 6 ... Next i` |
| Conditional | `If cond Then ... Else ... End If` |
| Last used row | `Cells(Rows.Count, col).End(xlUp).Row` |
| Prompt the user | `InputBox("prompt text")` |

### [02 · Advanced Power Query (M Language Basics)](level-3/02-advanced-power-query-m-language.md)

| M function | Purpose |
|---|---|
| `Table.AddColumn(table,name,each expr)` | Add a computed column |
| `Table.NestedJoin(...)` | Merge two tables (the engine behind Merge Queries) |
| `Table.TransformColumns(table,{{col,fn,type}})` | Apply a function to a column |
| `let ... in ...` | Every query's structure: named steps, final returned step |

### [03 · Power Pivot &amp; the Data Model](level-3/03-power-pivot-data-model.md)

| Concept | How |
|---|---|
| Add a table to the Data Model | Power Pivot → Add to Data Model |
| Create a relationship | Diagram View → drag key column to key column |
| Pull a value across a relationship | `RELATED(Table[Column])` (many→one direction) |
| PivotTable from multiple tables | Insert PivotTable → Use this workbook's Data Model |

### [04 · DAX Basics for Power Pivot](level-3/04-dax-basics-power-pivot.md)

| DAX construct | Purpose |
|---|---|
| `SUM(Table[Col])` | Simple column total |
| `SUMX(Table, expr)` | Row-by-row iterator, then sum |
| `AVERAGEX(Table, expr)` | Row-by-row iterator, then average |
| `CALCULATE(expr, filter)` | Override the current filter context |
| `ALL(Table)` | Remove filters — useful for grand-total denominators |
| `DIVIDE(num, denom)` | Division that returns blank instead of erroring on 0 |

### [05 · Advanced VBA (UserForms, Error Handling)](level-3/05-advanced-vba-userforms-error-handling.md)

| Construct | Purpose |
|---|---|
| `UserForm1.Show` | Display a custom dialog |
| `TextBox.Value` | Read/write a form field |
| `On Error GoTo Label` | Route runtime errors to a handler |
| `Err.Number` / `Err.Description` | Inspect the error that was caught |
| `On Error Resume Next` | Skip the failing line (use sparingly) |
| `On Error GoTo 0` | Disable the active error handler from this point on |

### [06 · Building Custom Functions with LAMBDA](level-3/06-custom-functions-lambda.md)

| Construct | Syntax |
|---|---|
| Inline LAMBDA | `=LAMBDA(params,formula)(args)` |
| Named custom function | Define Name → `=LAMBDA(params,formula)`, then call `=Name(args)` |
| Recursive LAMBDA | Reference the function's own Defined Name inside its formula |
| LET inside LAMBDA | `LAMBDA(x,LET(name,expr,...,result))` |

### [07 · Financial Modeling Techniques](level-3/07-financial-modeling-techniques.md)

| Function | Purpose |
|---|---|
| `NPV(rate, values)` | Present value of a series of *future* cash flows |
| `IRR(values)` | Discount rate that makes NPV = 0 |
| `PMT(rate,nper,pv)` | Fixed periodic payment for a loan/annuity |
| Amortization | Interest = balance × rate; Principal = Payment − Interest |

### [08 · Advanced Dashboard Design Principles](level-3/08-advanced-dashboard-design-principles.md)

| Technique | Purpose |
|---|---|
| Formula-driven chart title | Title updates automatically with the filter |
| Data Bars | Instant visual ranking without a chart |
| Slicer | Point-and-click filter for Tables/PivotTables |
| SUMIFS-based KPI cells | Feed both text titles and conditional formats |

### [09 · Auditing &amp; Error-Checking Large Workbooks](level-3/09-auditing-error-checking-workbooks.md)

| Tool | Purpose |
|---|---|
| Trace Precedents/Dependents | Visualize which cells feed into / depend on a formula |
| `IFERROR(formula, fallback)` | Replace an error result with a safe fallback |
| `ISERROR(expr)` | Test whether an expression evaluates to an error |
| Evaluate Formula | Step through a nested formula's evaluation order |
| Data Validation | Prevent bad input at the source |

### [10 · Project — Financial Model with Scenario Analysis](level-3/10-project-financial-model-scenario-analysis.md)

| Tool | Purpose |
|---|---|
| Data Table (one-variable) | Recalculate one output across a range of one input |
| Scenario Manager | Save and compare named sets of multiple input changes |
| NPV + IF decision rule | Turn a numeric result into a plain-language recommendation |

## Level 4 · Master

### [01 · Enterprise Excel Architecture](level-4/01-enterprise-excel-architecture.md)

| Concept | Purpose |
|---|---|
| External reference (`[Book.xlsx]Sheet!Cell`) | Pull a live value from another workbook |
| Power Query connection | Refreshable pull from a shared source, more robust than links |
| Power Pivot relationship | Replace VLOOKUP chains with a queryable Data Model |
| `_`-prefixed hidden sheets | Separate internal calc sheets from the user-facing view |

### [02 · Advanced DAX &amp; Power Pivot Modeling](level-4/02-advanced-dax-power-pivot-modeling.md)

| Function | Purpose |
|---|---|
| `SAMEPERIODLASTYEAR(dates)` | Shift the current filter context back one year |
| `CALCULATE(expr, filter)` | Recompute an expression under a modified filter context |
| `SUMX(table, expr)` | Row-by-row expression, then sum — for calculations `SUM` alone can't do |
| `RANKX(table, expr)` | Rank a value against a set of candidates |
| `DIVIDE(num, denom)` | Division that safely returns blank/0 instead of `#DIV/0!` |

### [03 · Integrating Excel with Power BI](level-4/03-integrating-excel-power-bi.md)

| Action | Direction |
|---|---|
| Get Data → Excel Workbook | Excel Data Model → Power BI Desktop |
| Analyze in Excel | Power BI Service dataset → live Excel PivotTable |
| Publish to Power BI | Excel workbook → Power BI dataset/report |

### [04 · Advanced VBA Application Development](level-4/04-advanced-vba-application-development.md)

| Concept | Purpose |
|---|---|
| Class module (`Property Let/Get`) | Bundle related fields into one reusable object type |
| `Collection` | Hold a growable list of custom objects |
| `For Each ... In collection` | Iterate every object in the collection |
| UserForm calling standard-module procedures | Separate input/validation from business logic |

### [05 · Performance Optimization for Large Workbooks](level-4/05-performance-optimization-large-workbooks.md)

| Technique | Effect |
|---|---|
| Avoid `OFFSET`/`INDIRECT` | Removes forced full recalculation |
| `XLOOKUP` / sorted `VLOOKUP(...,TRUE)` | Faster large-scale lookups |
| Manual calculation during bulk edits | Avoids recalculating after every keystroke |
| `ScreenUpdating=False` + `Calculation=Manual` in VBA loops | Major speedup for bulk VBA writes |

### [06 · Excel + Python Integration (Overview)](level-4/06-excel-python-integration.md)

| Tool | Runs where | Best for |
|---|---|---|
| `=PY()` | Inside the open workbook, sandboxed | Exploratory analysis, in-cell charts |
| `openpyxl` | Standalone Python | Precise cell writes, preserving formatting |
| `pandas` | Standalone Python | Heavy tabular transforms, at the cost of formatting fidelity |

### [07 · Building Excel-Based Tools for Teams](level-4/07-excel-based-tools-for-teams.md)

| Technique | Purpose |
|---|---|
| Unlock inputs, then Protect Sheet | Let users fill in data without touching the logic |
| Data Validation (List/Decimal) | Force consistent, valid input across every user |
| `.xlam` Add-in | Ship one shared custom function to every team member's Excel |
| OneDrive/SharePoint + Tables | Reliable simultaneous co-authoring |

### [08 · Data Governance in Spreadsheet-Heavy Organizations](level-4/08-data-governance-spreadsheet-organizations.md)

| Governance need | Excel/365 mechanism |
|---|---|
| Single source of truth | Power Query connection to one master file/table |
| Access control | SharePoint/OneDrive sharing permissions + Protect Sheet/Workbook |
| Audit trail | Version History, or a `_ChangeLog` sheet via `Worksheet_Change` |
| Recoverability | OneDrive/SharePoint Version History |

### [09 · Career Growth: Excel Power User to Analytics Lead](level-4/09-career-growth-analytics-lead.md)

| Career step | This path's equivalent module |
|---|---|
| Prove ROI of your own work | Section 2's time-saved calculation pattern |
| Build a portfolio | Package each level's capstone with a business-question framing |
| Level up beyond Excel | Add SQL, a second BI tool, and statistics fundamentals |

### [10 · Capstone — Full Business Reporting System in Excel](level-4/10-capstone-business-reporting-system.md)

| Layer | Module this capstone draws from |
|---|---|
| One Table, one Data Model | 01 · Enterprise Excel Architecture |
| DAX measures for KPIs | 02 · Advanced DAX & Power Pivot Modeling |
| One-click refresh, performance-safe | 04/05 · VBA + Performance Optimization |
| Protected structure, change log | 07/08 · Team Tools + Data Governance |
| Slicers, formula-driven titles | Level 3 · 08 Advanced Dashboard Design |

---

Go deeper with the full [Excel Mastery Path](index.md) curriculum or test yourself with the [interview questions](interview-questions.md).
