---
title: "Excel Interview Questions and Answers (180 Questions)"
description: "180 Excel interview questions with answers and explanations, from beginner to advanced, organised by level. Free, from the Excel Mastery Path."
---

# Excel Interview Questions and Answers

**180 questions** with the correct answer and a short explanation, organised from beginner (Level 1) to advanced (Level 4). Every question comes from a lesson in the free Excel Mastery Path, so you can open the lesson to go deeper.

!!! tip "Practise, don't just read"
    Each lesson also ends with an interactive quiz, and the [final quiz](quiz.md) tests all four levels. Try to answer each question below before reading the answer.

## Level 1 · Entry

### [01 · Excel Fundamentals](level-1/01-excel-fundamentals.md)

**Q1.** How are the three nested containers in Excel ordered?

**Answer:** A workbook holds worksheets, and each worksheet is a grid of cells

The workbook is the file, worksheets are tabs, and cells are addressed by column letter and row number.

**Q2.** What does the Name Box let you do besides showing the selected address?

**Answer:** Type an address such as B3 and press Enter to jump to it

It's the box above column A.

**Q3.** What does A1:C5 represent?

**Answer:** A block 3 columns wide by 5 rows tall

A range is written TopLeft:BottomRight.

**Q4.** What does Ctrl/Cmd+Arrow do?

**Answer:** Jumps to the edge of the current block of data in that direction

Adding Shift also extends the selection.

### [02 · Formulas &amp; Functions Basics](level-1/02-formulas-functions-basics.md)

**Q5.** What does every formula start with?

**Answer:** =

Every formula is = followed by an expression.

**Q6.** You fill D2 (=C2-B2) down to D3. What does D3 contain?

**Answer:** =C3-B3, because the references are relative

Relative references shift as you copy them.

**Q7.** What does the Formula Bar show for a cell with a formula?

**Answer:** The formula, while the cell shows the calculated result

The stored formula versus displayed value distinction matters constantly.

**Q8.** What does =SUM(B2:B4,B6) do?

**Answer:** Sums four of the five rows, skipping B5

SUM accepts multiple ranges or cells, comma-separated.

### [03 · Formatting &amp; Conditional Formatting](level-1/03-formatting-conditional-formatting.md)

**Q9.** When you format 1200 as currency, what changes?

**Answer:** Only how it displays; the stored value is still 1200

Every formula referencing it still uses 1200.

**Q10.** What does conditional formatting do?

**Answer:** Changes a cell's appearance automatically based on its value

It makes a budget's problem rows visually obvious.

**Q11.** What happens to a cell's precision when you use Decrease Decimal?

**Answer:** Only the display is rounded; formulas use the full stored precision

Display rounding doesn't change the underlying value.

**Q12.** What does the custom format $#,##0;[Red]-$#,##0 do?

**Answer:** Shows positives normally and negatives in red with a minus

It's a Custom format from More Number Formats.

### [04 · Working with Tables](level-1/04-working-with-tables.md)

**Q13.** What does Insert &gt; Table (Ctrl/Cmd+T) give you over a plain range?

**Answer:** Filtering, banded rows, auto-filling formulas and structured references

It turns the range into a structured object.

**Q14.** What does =SUM(Budget[Actual]) have over =SUM(C2:C6)?

**Answer:** It adjusts automatically when rows are added or removed

A fixed range would need manual editing.

**Q15.** What does Budget[@Actual] refer to?

**Answer:** The Actual value in that specific row

It is the structured-reference equivalent of a relative reference.

**Q16.** Which table name is valid?

**Answer:** Budget

Table names must be unique, contain no spaces and can't collide with cell addresses.

### [05 · Charts &amp; Basic Visualization](level-1/05-charts-basic-visualization.md)

**Q17.** Which chart type does the lesson insert first for Budgeted vs Actual per category?

**Answer:** Clustered Column

It shows one cluster of two bars (Budgeted, Actual) per category.

**Q18.** What do the three floating buttons on a selected chart control?

**Answer:** Chart Elements, Chart Styles and Chart Filters

They cover most day-to-day chart adjustments without a separate dialog.

**Q19.** Why might you uncheck Gridlines once Data Labels are showing?

**Answer:** Gridlines add visual noise when exact numbers are already displayed

Gridlines help estimate values only when labels are absent.

**Q20.** How do you select non-adjacent columns such as Category and Budgeted for a pie chart?

**Answer:** Select A1:A7, then hold Ctrl/Cmd and select B1:B7

Ctrl/Cmd-click adds another range to the selection.

### [06 · Lookup Functions](level-1/06-lookup-functions.md)

**Q21.** In =VLOOKUP(A2,CategoryInfo!$A$2:$C$7,2,FALSE), what does 2 mean?

**Answer:** The 2nd column of the table array

The column index number is a fixed position, not a column name.

**Q22.** What does FALSE as the last VLOOKUP argument request?

**Answer:** An exact match

It is what you want almost always when matching text.

**Q23.** What is VLOOKUP's biggest limitation?

**Answer:** It can only look rightward from the lookup column

It also breaks if a column is inserted in the middle of the table array.

**Q24.** How does XLOOKUP take its range arguments?

**Answer:** A lookup array and a separate return array

Its arguments are lookup value, lookup array and return array.

### [07 · Logical Functions](level-1/07-logical-functions.md)

**Q25.** What are IF's three arguments?

**Answer:** A condition, the value if TRUE and the value if FALSE

=IF(D2&gt;0,"Over Budget","OK") shows the pattern.

**Q26.** Rent's Difference is 0. What does =IF(D2&gt;0,"Over Budget","OK") show?

**Answer:** OK

0&gt;0 is FALSE.

**Q27.** How does IFS differ from nested IFs?

**Answer:** It takes condition/result pairs and returns the first TRUE result, without nesting parentheses

Using TRUE as the final condition acts as a catch-all else.

**Q28.** Why nest a second IF in the FALSE branch?

**Answer:** To get three outcomes (Over, Under, On Track) from one formula

A single IF only branches two ways.

### [08 · PivotTables Basics](level-1/08-pivottables-basics.md)

**Q29.** What kind of source data does a PivotTable need?

**Answer:** Flat, transaction-level data with one row per fact

The Transactions table has one row per category per month.

**Q30.** What does converting the source to a Table give a PivotTable?

**Answer:** It auto-expands its source range when rows are added

PivotTables built from a Table pick up new rows.

**Q31.** Dragging Category to Rows and Amount to Values defaults to what?

**Answer:** Sum of Amount per category

Rent totals 2400 across Jan and Feb.

**Q32.** After swapping Category for Type in Rows, how many data rows appear?

**Answer:** Two: Fixed (3050) and Variable (1400)

Fixed = Rent 2400 + Savings 600 + Subscriptions 50.

### [09 · Data Validation &amp; Cleaning](level-1/09-data-validation-cleaning.md)

**Q33.** How do you restrict a cell to a dropdown of allowed values?

**Answer:** Data &gt; Data Validation, Allow: List, with a Source range

The lesson uses =CategoryInfo!$A$2:$A$7 as the source.

**Q34.** What is the difference between Stop and Warning error alert styles?

**Answer:** Stop blocks invalid entries; Warning lets the user override after confirming

Warning is useful when a value is unusual but not necessarily wrong.

**Q35.** What does TRIM do?

**Answer:** Removes leading and trailing spaces and collapses internal multiple spaces to single spaces

It's the first step in cleaning pasted or imported text.

**Q36.** Does Clear All in the Data Validation dialog change existing cell values?

**Answer:** No, it removes only the rule

The existing values stay as they were.

### [10 · Project — Personal Budget Tracker](level-1/10-project-personal-budget-tracker.md)

**Q37.** Which formula pattern should the Difference column use?

**Answer:** =C2-B2 with relative references, not hardcoded values

Relative references fill down correctly.

**Q38.** Why does the % of Budget column use an absolute reference to the Budgeted total?

**Answer:** So it doesn't break when copied

For example $B$8 or the Table's total row.

**Q39.** What does the Needs Review column use?

**Answer:** AND(Type="Variable", Difference&gt;50)

It flags variable categories that are more than 50 over budget.

**Q40.** Which four sheets make up the finished workbook?

**Answer:** Budget, CategoryInfo, Transactions and Messy

Messy is the cleaning scratch sheet and can be deleted later.

## Level 2 · Intermediate

### [01 · Advanced Formulas (SUMIFS/COUNTIFS/AVERAGEIFS)](level-2/01-advanced-formulas.md)

**Q41.** Where does the sum range go in SUMIFS compared with SUMIF?

**Answer:** First in SUMIFS, last in SUMIF

The lesson notes this trips up almost everyone coming from SUMIF.

**Q42.** Using the Sales data, what does =SUMIFS(D2:D11,A2:A11,"East",B2:B11,"Mar") return?

**Answer:** 2200

Rows 4 (1500) and 8 (700) match East + Mar.

**Q43.** =COUNTIFS(A2:A11,"West",D2:D11,"&gt;1000") returns 1. Why isn't the 1000 row counted?

**Answer:** 1000 is not greater than 1000

Only row 7 (1300) qualifies.

**Q44.** What does =AVERAGEIFS(D2:D11,C2:C11,"Priya") return?

**Answer:** 975

Priya's deals 1500, 800, 600, 1000 sum to 3900 over 4 rows.

### [02 · PivotTables Deep Dive](level-2/02-pivottables-deep-dive.md)

**Q45.** What does the lesson recommend to catch a mis-selected data range?

**Answer:** Cross-check the PivotTable totals against hand-calculated totals

East 4900 and West 5150 matched the hand sums.

**Q46.** How do you get an average alongside the sum?

**Answer:** Drag Amount to Values a second time and change its Value Field Settings to Average

East's average is 4900/5 = 980.

**Q47.** What controls nesting when you drag Month below Region in Rows?

**Answer:** The order of fields in the Rows box

Order matters: Region then Month nests months under each region.

**Q48.** What does a calculated field Bonus = Amount*0.1 show for East's 4900?

**Answer:** 490

4900 * 0.1 = 490.

### [03 · Advanced Charting &amp; Dashboards](level-2/03-advanced-charting-dashboards.md)

**Q49.** Why use a secondary axis for Margin% alongside Revenue?

**Answer:** The scales differ by an order of magnitude, so the line would look flat on one axis

Revenue runs in thousands while Margin% is a fraction near 0.30.

**Q50.** What is a sparkline?

**Answer:** A small in-cell trend line summarizing a column of values

One sparkline per KPI, stacked, gives a quick visual scan in dashboard headers.

**Q51.** Which month has the lowest Revenue but the highest Margin% in the dataset?

**Answer:** Mar

March: 11000 revenue and 0.35 margin.

**Q52.** How do you snap chart edges to cell boundaries when dragging?

**Answer:** Hold Alt while dragging

This is the grid alignment principle for dashboard tiles.

### [04 · Data Validation Advanced &amp; Form Controls](level-2/04-data-validation-advanced-form-controls.md)

**Q53.** What does =INDIRECT(G1) do in the cascading dropdown?

**Answer:** Turns the text in G1 (e.g. "Fruit") into a reference to the named range of that name

G2's list changes to the items for whichever category G1 holds.

**Q54.** What does the custom validation =AND(H2&gt;0,H2&lt;=100) allow?

**Answer:** Numbers above 0 up to and including 100

150 and 0 are rejected; 100 is accepted.

**Q55.** How does =COUNTIF($I$2:$I$10,I2)=1 block duplicates?

**Answer:** If the value already appears, the count becomes 2, returning FALSE and rejecting it

Only a first occurrence gives a count of 1.

**Q56.** What do Form Controls like combo boxes and check boxes let a sheet do?

**Answer:** Respond to user clicks without VBA

First enable the Developer tab in Customize Ribbon.

### [05 · Power Query Basics](level-2/05-power-query-basics.md)

**Q57.** What makes Power Query repeatable?

**Answer:** Every action is recorded as a numbered Applied Step that replays on refresh

Manual copy/paste cleanup would have to be redone each time.

**Q58.** Why Replace Values (null → 0) rather than Fill Down for a blank sales cell?

**Answer:** A missing figure should become 0, not inherit a neighbor's value

Row 5 becomes Alan, EAST, 0.

**Q59.** After cleaning, what does Group By region with Sum of sales show?

**Answer:** EAST 2900 and WEST 2600

East = 1200+900+0+800; West = 1500+1100.

**Q60.** Why normalize region case to one convention?

**Answer:** So East and east aren't treated as separate groups downstream

Pick one convention and apply it everywhere.

### [06 · Array Formulas &amp; Dynamic Arrays](level-2/06-array-formulas-dynamic-arrays.md)

**Q61.** What does 'spill' mean for dynamic array functions?

**Answer:** Results automatically fill neighboring cells, with no Ctrl+Shift+Enter needed

FILTER, SORT, UNIQUE and SEQUENCE return multiple results.

**Q62.** How do you combine two conditions inside FILTER as AND?

**Answer:** Multiply the boolean arrays, e.g. (B2:B9="Sales")*(C2:C9&gt;80)

Multiplication acts as AND and returns 3 rows in the example.

**Q63.** What does =SORT(UNIQUE(B2:B9)) return for the Team data?

**Answer:** HR, IT, Sales

UNIQUE gives first-seen order Sales, IT, HR; SORT orders it alphabetically.

**Q64.** What does the # operator in I1# refer to?

**Answer:** The entire spilled array

=COUNTIF(B2:B9,I1#) spills counts aligned to the sorted departments.

### [07 · What-If Analysis](level-2/07-what-if-analysis.md)

**Q65.** What does Goal Seek do in the loan model?

**Answer:** Iterates to find the input (e.g. Principal) that produces a target output (e.g. a 350 payment)

It solves for B1 so that B4 reaches 350, roughly 18122.

**Q66.** In a one-variable Data Table with rates listed down a column, which input cell do you specify?

**Answer:** Column input cell B2 (the rate)

Rates vary down a column, so B2 is the column input cell.

**Q67.** What sanity check does the lesson give for the Data Table results?

**Answer:** Payments rise monotonically as the rate increases from 0.04 to 0.08

At 0.08 expect roughly 405.53, higher than 386.66 at 0.06.

**Q68.** Why does the formula negate PMT?

**Answer:** PMT returns a negative number for a cash outflow

=-PMT(B2/12,B3*12,B1) shows the payment as a positive amount.

### [08 · Working with Multiple Sheets &amp; Workbooks](level-2/08-multiple-sheets-workbooks.md)

**Q69.** What does =SUM(Jan:Mar!B2) compute?

**Answer:** B2 summed across every sheet from Jan through Mar, including sheets inserted between

Widget is 500 + 600 + 550 = 1650.

**Q70.** How does =Feb!B2 differ from a 3-D reference?

**Answer:** It points to one cell on one sheet, rather than aggregating across a range of sheets

It returns 600, February's Widget figure.

**Q71.** A linked source workbook is moved. What breaks, and what fixes it?

**Answer:** The links break; use Data &gt; Edit Links to repoint them

Closed-workbook references store the full path.

**Q72.** What is the best practice for workbooks that link to each other?

**Answer:** Keep them in the same folder and avoid renaming source files

This reduces broken links.

### [09 · Intro to Recording Macros](level-2/09-intro-recording-macros.md)

**Q73.** What does the Macro Recorder write behind the scenes?

**Answer:** Real VBA code

You can inspect it via Developer &gt; Macros &gt; Edit.

**Q74.** What happens with Use Relative References OFF (the default)?

**Answer:** The macro always acts on the recorded fixed cells, like A1:C1

It selects A1:C1 regardless of which cell is active.

**Q75.** A macro recorded relative to A2 is run with A3 active. Which row does it format?

**Answer:** A3:C3

It acts relative to the starting cell, useful for 'the row I clicked'.

**Q76.** What are the totals the AddTotalColumn macro should generate for D2:D4?

**Answer:** 20, 60 and 50

D2 = 10*2, D3 = 4*15, D4 = 2*25.

### [10 · Project — Interactive Sales Dashboard](level-2/10-project-interactive-sales-dashboard.md)

**Q77.** With the Region selector set to East, what should the SUMIFS total be?

**Answer:** 7200

East rows: 1200+1400+1600+700+1050+1250.

**Q78.** Why cross-check the PivotTable against the SUMIFS total?

**Answer:** A consistency check a real dashboard needs: East months 1900+2450+2850 = 7200

Both methods should agree.

**Q79.** What does a Slicer provide compared with a page filter dropdown?

**Answer:** A more visual East/West button filter that updates instantly

Insert it via PivotTable Analyze &gt; Insert Slicer &gt; Region.

**Q80.** What drives the whole dashboard?

**Answer:** A single dropdown selector for Region

It combines SUMIFS, PivotTables, combo charts, validation dropdowns and a Form Control.

## Level 3 · Advanced

### [01 · VBA Basics for Automation](level-3/01-vba-basics-automation.md)

**Q81.** How do you open the VBA editor?

**Answer:** Alt+F11

Then Insert &gt; Module to type a Sub.

**Q82.** Why declare total As Long rather than Integer in the SumQty macro?

**Answer:** Long holds sums that may exceed Integer's ~32,767 limit

Integer is fine for a small loop counter like i.

**Q83.** What does Cells(2,2) refer to?

**Answer:** Row 2, column 2, equivalent to B2

It is useful in loops where row and column are variables.

**Q84.** In FlagLowStock, which items are flagged Reorder (Qty &lt; 10)?

**Answer:** Book, Pencil and Eraser

Book(3), Pencil(0) and Eraser(8) are under 10; Pen(12) and Bag(25) are OK.

### [02 · Advanced Power Query (M Language Basics)](level-3/02-advanced-power-query-m-language.md)

**Q85.** What does every Power Query transformation generate?

**Answer:** Code in the M language, visible in the Advanced Editor

Each UI action appends a step inside the let...in block.

**Q86.** What does the final in line of an M query name?

**Answer:** Which step's output is returned

Steps are bound to names inside let.

**Q87.** What Join Kind keeps every order and adds customer info where it matches?

**Answer:** Left Outer

Underlying M: Table.NestedJoin(..., JoinKind.LeftOuter).

**Q88.** Using the rule Amount &gt;= 500 -&gt; Large, how many orders are Large?

**Answer:** Two (500 and 800)

500→Large, 300→Small, 800→Large, 200→Small, 450→Small.

### [03 · Power Pivot &amp; the Data Model](level-3/03-power-pivot-data-model.md)

**Q89.** What does the Data Model let a PivotTable do?

**Answer:** Pull fields from several related tables at once without VLOOKUP glue columns

Tables relate like a small database.

**Q90.** In the Sales/Products relationship, which side is the 'one' side?

**Answer:** Products, where ProductID is unique

ProductID repeats in Sales (the many side).

**Q91.** What does RELATED(Products[Price]) do inside a Sales calculated column?

**Answer:** Pulls the matching Price from the one side into each Sales row

Each Sales row looks up its own ProductID's Price.

**Q92.** What is the Revenue for row 3 (P2, Qty 5, Price 35)?

**Answer:** 175

5 × 35 = 175.

### [04 · DAX Basics for Power Pivot](level-3/04-dax-basics-power-pivot.md)

**Q93.** How does a DAX measure differ from a calculated column?

**Answer:** A measure is computed dynamically from the filter context; a column is evaluated row by row

The same measure returns different numbers depending on where it's placed in the PivotTable.

**Q94.** What does SUMX(Sales, Sales[Qty] * RELATED(Products[Price])) do?

**Answer:** Iterates each Sales row, computes the product, then sums the results

SUMX is an iterator.

**Q95.** What does CALCULATE do?

**Answer:** Overrides the current filter context with an explicit condition

Alan Revenue stays 380 even next to Priya's row.

**Q96.** What is Alan's total revenue in the lesson's data?

**Answer:** 380

10×20 + 12×15 = 200 + 180.

### [05 · Advanced VBA (UserForms, Error Handling)](level-3/05-advanced-vba-userforms-error-handling.md)

**Q97.** What does On Error GoTo ErrHandler do?

**Answer:** Routes execution to the handler label instead of crashing on a runtime error

With B2 = 0, error 11 (Division by zero) is shown by the handler.

**Q98.** Why is Exit Sub placed before the ErrHandler label?

**Answer:** So the handler only runs when an error occurs, not after normal completion

Without it the code would fall through into the handler.

**Q99.** What does the Add button's validation do when both fields are blank?

**Answer:** Shows a 'Both fields are required' message and writes no row

It uses MsgBox and Exit Sub.

**Q100.** How does On Error Resume Next differ?

**Answer:** It skips the failing line and continues, a looser alternative

It's useful only when a failure on one line is acceptable.

### [06 · Building Custom Functions with LAMBDA](level-3/06-custom-functions-lambda.md)

**Q101.** What does LAMBDA let you do?

**Answer:** Define a reusable custom function in worksheet formulas, with no VBA

Name it via Name Manager so it behaves like a built-in function.

**Q102.** What does the trailing (B2,C2) do in =LAMBDA(cost,pct,cost*(1+pct))(B2,C2)?

**Answer:** Immediately calls the LAMBDA with those arguments

For Pen, 10*(1+0.20) = 12.

**Q103.** What does FactorialFn(5) return?

**Answer:** 120

5*4*3*2*1 = 120; a named LAMBDA can call itself.

**Q104.** What is LET useful for in a LAMBDA?

**Answer:** Naming intermediate values within one formula for readability

For Margin, profit and margin are named steps.

### [07 · Financial Modeling Techniques](level-3/07-financial-modeling-techniques.md)

**Q105.** Why is the upfront investment added outside the NPV function?

**Answer:** NPV assumes the first value is one period away, but the investment is at time 0

The formula is =B1+NPV(B6,B2:B5).

**Q106.** What does a positive NPV indicate?

**Answer:** The investment is expected to create value at that discount rate

Here NPV ≈ 2180.18 at 10%.

**Q107.** What is IRR?

**Answer:** The discount rate at which NPV would be exactly zero

Since NPV at 10% is positive, the IRR must be higher than 10%.

**Q108.** What happens to interest and principal as a loan amortizes?

**Answer:** Interest shrinks and principal grows each year as the balance declines

Year 1 interest is 1200; Year 2 is about 987.17.

### [08 · Advanced Dashboard Design Principles](level-3/08-advanced-dashboard-design-principles.md)

**Q109.** What is the dynamic title trick?

**Answer:** Make the chart title a formula referencing a cell, so it updates when the filter changes

E3 builds a string such as Total Revenue — North: $37,300.

**Q110.** Which chart best shows how one metric trends over time?

**Answer:** Line chart

It emphasizes direction and slope.

**Q111.** Why avoid pie charts past 3–4 slices?

**Answer:** They're hard to compare

Use a stacked bar for part-to-whole breakdowns.

**Q112.** Why avoid 3-D charts and dual unrelated axes?

**Answer:** They distort the visual comparison a dashboard should make effortless

Keep charts honest and simple.

### [09 · Auditing &amp; Error-Checking Large Workbooks](level-3/09-auditing-error-checking-workbooks.md)

**Q113.** What bug is planted in the Audit sheet?

**Answer:** Totals use B2:B4 and C2:C4, omitting row 5 (Staples)

It is a common copy-paste range bug.

**Q114.** What does Trace Precedents show?

**Answer:** Arrows to the cells a formula depends on

For B6 it shows B2:B4, revealing the range stops one row short; Trace Dependents shows the reverse.

**Q115.** What does =IFERROR(D2/(B2-10),"n/a") return when B2 = 10?

**Answer:** "n/a"

It prevents the error from propagating into dependent formulas.

**Q116.** How does data validation help prevent errors?

**Answer:** It stops invalid entries such as a negative quantity at entry rather than during an audit

Add an Input Message and Error Alert to guide users.

### [10 · Project — Financial Model with Scenario Analysis](level-3/10-project-financial-model-scenario-analysis.md)

**Q117.** What is Year 1 profit in the projection?

**Answer:** 12000

50000 revenue − 1000×30 variable cost − 8000 fixed cost.

**Q118.** How does the model compute Year 2 units?

**Answer:** Year 1 units × (1 + growth rate), giving 1150

The growth rate is 0.15 and referenced absolutely as $B$2.

**Q119.** What should happen to NPV as growth rate rises from 5% to 25%?

**Answer:** It increases monotonically

Higher growth compounds units and profit in Years 2–3.

**Q120.** What does Scenario Manager add to the model?

**Answer:** Several named scenarios (best, worst, base) presented side by side

It's used alongside the Data Table for what-if analysis.

## Level 4 · Master

### [01 · Enterprise Excel Architecture](level-4/01-enterprise-excel-architecture.md)

**Q121.** Why are linked workbooks fragile at scale?

**Answer:** Moving or renaming the source breaks every external reference, with no central registry of dependencies

Edit Links repairs are manual, per workbook and per broken link.

**Q122.** What is the enterprise-safer alternative to cell-level links between files?

**Answer:** Power Query connections to a shared source such as a folder, database or SharePoint list

Power Query re-resolves the connection on refresh instead of caching a brittle reference.

**Q123.** What does a relationship in a Data Model replace?

**Answer:** VLOOKUPs and helper columns copied across report sheets

A DAX measure like Total Budget can then be sliced by any related field.

**Q124.** What do Report.xlsx's cached values do when Source.xlsx is closed?

**Answer:** Excel keeps the last known value and re-reads on open, as long as the path doesn't move

The link still works unless the file moves.

### [02 · Advanced DAX &amp; Power Pivot Modeling](level-4/02-advanced-dax-power-pivot-modeling.md)

**Q125.** What does SAMEPERIODLASTYEAR enable?

**Answer:** Year-over-year comparison, e.g. Sales PY

For 2027 it returns the 2026 total of 5950.

**Q126.** What is the YoY % for 2027 (7100 vs 5950)?

**Answer:** About +19.3%

(7100−5950)/5950 ≈ 0.1933.

**Q127.** Why must the Calendar table be marked as a Date Table?

**Answer:** So time intelligence functions work correctly across a continuous date range

It spans 2026-01-01 to 2027-12-31 and relates to Sales[Date].

**Q128.** What does North Share % show?

**Answer:** North's sales as a share of total, about 55.6%

7250 / 13050 ≈ 0.5556.

### [03 · Integrating Excel with Power BI](level-4/03-integrating-excel-power-bi.md)

**Q129.** What does Power BI import when you open an Excel workbook that has a Data Model?

**Answer:** The Data Model directly, including relationships and DAX measures

Total Sales arrives still equal to 13050.

**Q130.** What does Analyze in Excel produce?

**Answer:** An .xlsx with a PivotTable connected live to the Power BI dataset

Refresh All pulls the latest numbers from the Power BI Service.

**Q131.** What problem does a shared model solve?

**Answer:** One measure definition, so Excel and Power BI don't disagree and refresh from a single source

Users via Analyze in Excel also inherit the dataset's row-level security.

**Q132.** For a one-time full transfer from Excel to Power BI, what is simplest?

**Answer:** Open the .xlsx directly in Power BI Desktop via Get Data

The reverse direction uses Analyze in Excel.

### [04 · Advanced VBA Application Development](level-4/04-advanced-vba-application-development.md)

**Q133.** Why wrap an expense's fields in a class instead of three parallel arrays?

**Answer:** Each expense travels as one object, with no separate variables to keep in sync

You can pass it to functions, add it to a collection or write it to a sheet in one call.

**Q134.** What does TotalByCategory("Office") return after adding Paper 25 Office, Taxi 40 Travel, Pens 8 Office?

**Answer:** 33

It loops all three and adds the two Office items: 25 + 8.

**Q135.** What do Property Let and Property Get do in the Expense class?

**Answer:** Set and read the private fields

Private pItem, pCategory and pAmount are exposed through properties.

**Q136.** What does the DumpToSheet procedure do?

**Answer:** Loops the collection and writes each expense's fields into the Data sheet rows

It starts at row 2 below the headers.

### [05 · Performance Optimization for Large Workbooks](level-4/05-performance-optimization-large-workbooks.md)

**Q137.** What is a volatile function?

**Answer:** One Excel recalculates on every worksheet recalculation, even if its inputs didn't change

Examples: NOW, TODAY, RAND, RANDBETWEEN, OFFSET and INDIRECT.

**Q138.** What does the lesson suggest instead of OFFSET-based dynamic ranges?

**Answer:** Excel Tables or INDEX, which are non-volatile

=INDEX(A:A,1):INDEX(A:A,COUNTA(A:A)) returns the same range without constant recalculation.

**Q139.** What is the risk of VLOOKUP with TRUE (approximate match)?

**Answer:** It uses binary search, which is fast but silently wrong on unsorted data

Only use it when the lookup column is guaranteed sorted.

**Q140.** Why switch Calculation Options to Manual when bulk-editing?

**Answer:** Automatic mode would recalculate the entire dependency tree after each edit

Switch back and recalculate when done.

### [06 · Excel + Python Integration (Overview)](level-4/06-excel-python-integration.md)

**Q141.** What can a =PY() cell not do?

**Answer:** Reach the local file system or network, so it can't pull in an external CSV or API directly

Python in Excel runs in a secure, isolated container.

**Q142.** When is openpyxl or pandas the better choice than =PY()?

**Answer:** Batch-processing many workbooks unattended, such as a folder of 200 monthly reports

=PY() works inside a single open workbook.

**Q143.** What does pandas do better than a Python loop per cell?

**Answer:** Vectorized operations over large tables, such as df["Units"] * 1.1

For Widget (120), Units_x1.1 = 132.0.

**Q144.** What is the total of the Units column in the worked dataset?

**Answer:** 525

120+95+140+60+110 = 525.

### [07 · Building Excel-Based Tools for Teams](level-4/07-excel-based-tools-for-teams.md)

**Q145.** How do you let team members edit only the input cells on a protected form?

**Answer:** Uncheck Locked on the input cells, then Protect Sheet

Every cell is Locked by default, so unlock the inputs before protecting.

**Q146.** Why restrict Department to a validation list?

**Answer:** So SUMIFS and PivotTables never split 'Sales' and 'sales' into two categories

Consistent input keeps downstream aggregation clean.

**Q147.** How do you make =ApprovalStatus(B3) available in any workbook?

**Answer:** Save the module's workbook as an Excel Add-in (.xlam) and load it

This is how a team standardizes a business rule.

**Q148.** With B3 = 7500, what does the approval formula show?

**Answer:** Needs manager sign-off

The rule auto-approves amounts up to 5000.

### [08 · Data Governance in Spreadsheet-Heavy Organizations](level-4/08-data-governance-spreadsheet-organizations.md)

**Q149.** What is the 'many copies' problem?

**Answer:** Several copies of the same figure disagree, and a report silently picks whichever was pasted

North's Q1 revenue appeared as 120000, 115000 and 121500.

**Q150.** How does a single source of truth fix an error once?

**Answer:** Other reports pull from it via Power Query or links, so a correction propagates on refresh

Fix 115000 to 120000 once in the master.

**Q151.** Are sheet protection passwords real security?

**Answer:** No, they are deterrents; real access control belongs at the file/sharing-permission level

Use SharePoint/OneDrive permissions such as Can View vs Can Edit.

**Q152.** What does Mark as Final provide?

**Answer:** A soft warning that a document is meant to be read-only, not real protection

It's only advisory.

### [09 · Career Growth: Excel Power User to Analytics Lead](level-4/09-career-growth-analytics-lead.md)

**Q153.** What does the lesson say marks the jump from Level 3 to Level 4?

**Answer:** Going from 'I can build a great model' to 'I can be trusted to design the reporting system a team depends on'

Governance, architecture and integration are what that trust is built on.

**Q154.** A tool saves 8 people 15 minutes a week. How many hours per year at 48 working weeks?

**Answer:** 96

8 × 15 = 120 min = 2 hours a week; 2 × 48 = 96.

**Q155.** At $45/hour, what is that annual value?

**Answer:** $4,320

96 × 45 = 4,320.

**Q156.** Which extra skill does the lesson say helps move from reporting what happened to explaining why?

**Answer:** Statistics fundamentals

It also lists SQL, a BI tool beyond Power BI's integration, and stakeholder communication.

### [10 · Capstone — Full Business Reporting System in Excel](level-4/10-capstone-business-reporting-system.md)

**Q157.** Why does every dashboard element reference the Table SalesTbl rather than a raw range?

**Answer:** New months appended to the Table flow through automatically

The Table is the single source of truth.

**Q158.** What is the grand total revenue in the dataset?

**Answer:** 265500

The sum of all twelve monthly region figures.

**Q159.** What does North Share % come to?

**Answer:** About 55%

146000 / 265500 ≈ 0.5499.

**Q160.** Why does RefreshDashboard turn off ScreenUpdating and set manual calculation?

**Answer:** It keeps a one-click refresh fast as data grows

It restores automatic calculation afterward.

## Mixed questions across all levels

**Q161.** You copy `=C2-B2` from D2 down to D6 using the fill handle. What does D3 contain?

**Answer:** =C3-B3

Relative references shift as you copy.

**Q162.** Your % of Budget formula breaks when copied down. What is the usual fix?

**Answer:** Lock the total with an absolute reference such as $B$8

An absolute reference stays fixed when copied.

**Q163.** A manager wants over-budget rows to stand out automatically. Which feature fits?

**Answer:** Conditional formatting on the Difference column

It changes appearance based on cell values.

**Q164.** You write `=SUM(C2:C6)`, then add a new row of data below. What is a better formula?

**Answer:** Convert the data to a Table and use =SUM(Budget[Actual])

Structured references adjust as rows are added or removed.

**Q165.** You must pull a category's Type from another sheet by exact text match. Which formula pattern is right?

**Answer:** =VLOOKUP(A2, CategoryInfo!$A$2:$C$7, 2, FALSE)

FALSE requests an exact match, and the table array is locked with $ so it doesn't shift.

**Q166.** Your VLOOKUP breaks after someone inserts a column in the lookup table. Which function avoids that and can look left?

**Answer:** XLOOKUP (or INDEX/MATCH)

XLOOKUP takes separate lookup and return arrays, so it doesn't rely on a fixed column number.

**Q167.** Sales data has Region, Month, Rep, Amount. You need total sales for East in March. Which formula?

**Answer:** =SUMIFS(D2:D11, A2:A11, "East", B2:B11, "Mar")

SUMIFS takes the sum range first, then criteria pairs.

**Q168.** You want a summary of a flat transaction log by category and month that you can re-pivot freely. What do you build?

**Answer:** A PivotTable from a Table

PivotTables summarize transaction-level data and auto-expand when sourced from a Table.

**Q169.** A dashboard title must read "Total Revenue — North: $37,300" and change with the selected region. What do you do?

**Answer:** Link the chart title to a formula cell that builds the text

A formula-driven title updates the instant the filter cell changes.

**Q170.** Two series on a chart differ by an order of magnitude (revenue in thousands, margin near 0.3). What should you do?

**Answer:** Put margin on a secondary axis

Otherwise the small series looks flat.

**Q171.** Messy imported data has stray spaces and mixed capitalization, and you want a repeatable cleanup that replays on refresh. Which tool?

**Answer:** Power Query

Power Query records every step in Applied Steps.

**Q172.** Your loan model has a target payment of 350 and you need the principal that achieves it. Which tool?

**Answer:** Goal Seek

Goal Seek solves for one input to reach a target output.

**Q173.** You must compare best/base/worst assumption sets side by side in a model. Which tool?

**Answer:** Scenario Manager

It stores named scenarios of input cells.

**Q174.** A shared form should let colleagues edit only the input cells and not overwrite the formula. What do you do?

**Answer:** Unlock the input cells, then Protect Sheet

Every cell is Locked by default; protection applies once inputs are unlocked.

**Q175.** An organization's workbooks link to each other with cell-level external references and keep breaking when files move. What is the better architecture?

**Answer:** Power Query connections to one shared source

Power Query re-resolves the connection on refresh instead of caching brittle cell references.

**Q176.** A workbook with thousands of formulas recalculates constantly because of OFFSET-based named ranges. What is the fix?

**Answer:** Replace them with Tables or INDEX, which are non-volatile

Volatile functions recalculate on every worksheet recalculation.

**Q177.** You must find an error in a total that omits the last row. Which built-in tool shows what the formula depends on?

**Answer:** Trace Precedents

The blue arrows show the range stops one row short.

**Q178.** A DAX measure must show North revenue regardless of which region a PivotTable row is sliced by. Which function overrides the filter context?

**Answer:** CALCULATE

CALCULATE changes the filter context with an explicit condition.

**Q179.** You want year-over-year growth in a Power Pivot model. What must exist in the model first?

**Answer:** A Date Table marked as the Date Table, related to your facts

Time-intelligence functions like SAMEPERIODLASTYEAR need a proper date table.

**Q180.** Several colleagues keep their own copies of the same revenue figures and the numbers disagree. What principle fixes this?

**Answer:** Single source of truth that reports pull from

Correct once in the master and it propagates on refresh.

---

More free practice: the [Excel Mastery Path](index.md) curriculum, the [Interview Prep Pack](https://learnmasterypath.com/interview-pack/) and the [free resume ATS check](https://learnmasterypath.com/resume/).
