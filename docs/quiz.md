---
title: "Final Quiz"
description: "Test what you have learned across all four levels of Excel Mastery Path with a 20-question final quiz."
---

# Final Quiz

20 questions drawn from all four levels of this course. You need **70% (14 out of 20)** to pass —
you can retake it as many times as you like. Your best result is saved in this browser.

<div id="mp-quiz-root"></div>

<script type="application/json" id="mp-quiz-data">
[
  {
    "question": "You copy `=C2-B2` from D2 down to D6 using the fill handle. What does D3 contain?",
    "options": [
      "=C2-B2",
      "=C3-B2",
      "=$C$2-$B$2",
      "=C3-B3"
    ],
    "correctIndex": 3,
    "explanation": "Relative references shift as you copy."
  },
  {
    "question": "Your % of Budget formula breaks when copied down. What is the usual fix?",
    "options": [
      "Lock the total with an absolute reference such as $B$8",
      "Use a bigger number format",
      "Convert to text",
      "Use Merge & Center"
    ],
    "correctIndex": 0,
    "explanation": "An absolute reference stays fixed when copied."
  },
  {
    "question": "A manager wants over-budget rows to stand out automatically. Which feature fits?",
    "options": [
      "Freeze panes",
      "Data validation",
      "Conditional formatting on the Difference column",
      "Bold on the header"
    ],
    "correctIndex": 2,
    "explanation": "It changes appearance based on cell values."
  },
  {
    "question": "You write `=SUM(C2:C6)`, then add a new row of data below. What is a better formula?",
    "options": [
      "Convert the data to a Table and use =SUM(Budget[Actual])",
      "=SUM(C2:C6,0)",
      "=C2+C3+C4+C5+C6",
      "=SUM(C:C6)"
    ],
    "correctIndex": 0,
    "explanation": "Structured references adjust as rows are added or removed."
  },
  {
    "question": "You must pull a category's Type from another sheet by exact text match. Which formula pattern is right?",
    "options": [
      "=SUM(CategoryInfo!B:B)",
      "=VLOOKUP(A2, CategoryInfo!$A$2:$C$7, 2, FALSE)",
      "=IF(A2, CategoryInfo)",
      "=VLOOKUP(A2, CategoryInfo!A2:C7, 2, TRUE)"
    ],
    "correctIndex": 1,
    "explanation": "FALSE requests an exact match, and the table array is locked with $ so it doesn't shift."
  },
  {
    "question": "Your VLOOKUP breaks after someone inserts a column in the lookup table. Which function avoids that and can look left?",
    "options": [
      "XLOOKUP (or INDEX/MATCH)",
      "TRIM",
      "SUMIF",
      "ROUND"
    ],
    "correctIndex": 0,
    "explanation": "XLOOKUP takes separate lookup and return arrays, so it doesn't rely on a fixed column number."
  },
  {
    "question": "Sales data has Region, Month, Rep, Amount. You need total sales for East in March. Which formula?",
    "options": [
      "=SUMIF(A2:A11, \"East\", \"Mar\")",
      "=COUNTIFS(D2:D11, \"East\")",
      "=SUMIFS(D2:D11, A2:A11, \"East\", B2:B11, \"Mar\")",
      "=AVERAGE(D2:D11)"
    ],
    "correctIndex": 2,
    "explanation": "SUMIFS takes the sum range first, then criteria pairs."
  },
  {
    "question": "You want a summary of a flat transaction log by category and month that you can re-pivot freely. What do you build?",
    "options": [
      "A PivotTable from a Table",
      "A chart on raw rows",
      "A VLOOKUP grid",
      "A macro"
    ],
    "correctIndex": 0,
    "explanation": "PivotTables summarize transaction-level data and auto-expand when sourced from a Table."
  },
  {
    "question": "A dashboard title must read \"Total Revenue — North: $37,300\" and change with the selected region. What do you do?",
    "options": [
      "Use conditional formatting",
      "Use a text box with a picture",
      "Link the chart title to a formula cell that builds the text",
      "Type the title manually each time"
    ],
    "correctIndex": 2,
    "explanation": "A formula-driven title updates the instant the filter cell changes."
  },
  {
    "question": "Two series on a chart differ by an order of magnitude (revenue in thousands, margin near 0.3). What should you do?",
    "options": [
      "Use a pie chart",
      "Delete one series",
      "Use 3-D",
      "Put margin on a secondary axis"
    ],
    "correctIndex": 3,
    "explanation": "Otherwise the small series looks flat."
  },
  {
    "question": "Messy imported data has stray spaces and mixed capitalization, and you want a repeatable cleanup that replays on refresh. Which tool?",
    "options": [
      "Freeze Panes",
      "Power Query",
      "Conditional formatting",
      "Goal Seek"
    ],
    "correctIndex": 1,
    "explanation": "Power Query records every step in Applied Steps."
  },
  {
    "question": "Your loan model has a target payment of 350 and you need the principal that achieves it. Which tool?",
    "options": [
      "Scenario Manager",
      "Pivot slicer",
      "Goal Seek",
      "Flash Fill"
    ],
    "correctIndex": 2,
    "explanation": "Goal Seek solves for one input to reach a target output."
  },
  {
    "question": "You must compare best/base/worst assumption sets side by side in a model. Which tool?",
    "options": [
      "Data validation",
      "Sparklines",
      "Remove Duplicates",
      "Scenario Manager"
    ],
    "correctIndex": 3,
    "explanation": "It stores named scenarios of input cells."
  },
  {
    "question": "A shared form should let colleagues edit only the input cells and not overwrite the formula. What do you do?",
    "options": [
      "Unlock the input cells, then Protect Sheet",
      "Hide the formulas",
      "Rename the sheet",
      "Mark as Final"
    ],
    "correctIndex": 0,
    "explanation": "Every cell is Locked by default; protection applies once inputs are unlocked."
  },
  {
    "question": "An organization's workbooks link to each other with cell-level external references and keep breaking when files move. What is the better architecture?",
    "options": [
      "Email copies",
      "Power Query connections to one shared source",
      "Rename files often",
      "More links"
    ],
    "correctIndex": 1,
    "explanation": "Power Query re-resolves the connection on refresh instead of caching brittle cell references."
  },
  {
    "question": "A workbook with thousands of formulas recalculates constantly because of OFFSET-based named ranges. What is the fix?",
    "options": [
      "Replace them with Tables or INDEX, which are non-volatile",
      "Use INDIRECT",
      "Turn off the computer",
      "Add more OFFSETs"
    ],
    "correctIndex": 0,
    "explanation": "Volatile functions recalculate on every worksheet recalculation."
  },
  {
    "question": "You must find an error in a total that omits the last row. Which built-in tool shows what the formula depends on?",
    "options": [
      "Goal Seek",
      "Spelling",
      "Trace Precedents",
      "Remove Arrows"
    ],
    "correctIndex": 2,
    "explanation": "The blue arrows show the range stops one row short."
  },
  {
    "question": "A DAX measure must show North revenue regardless of which region a PivotTable row is sliced by. Which function overrides the filter context?",
    "options": [
      "RELATED",
      "SUM",
      "CALCULATE",
      "TRIM"
    ],
    "correctIndex": 2,
    "explanation": "CALCULATE changes the filter context with an explicit condition."
  },
  {
    "question": "You want year-over-year growth in a Power Pivot model. What must exist in the model first?",
    "options": [
      "A form control",
      "A chart",
      "A VBA macro",
      "A Date Table marked as the Date Table, related to your facts"
    ],
    "correctIndex": 3,
    "explanation": "Time-intelligence functions like SAMEPERIODLASTYEAR need a proper date table."
  },
  {
    "question": "Several colleagues keep their own copies of the same revenue figures and the numbers disagree. What principle fixes this?",
    "options": [
      "Single source of truth that reports pull from",
      "More sheets",
      "More protection passwords",
      "Larger fonts"
    ],
    "correctIndex": 0,
    "explanation": "Correct once in the master and it propagates on refresh."
  }
]
</script>
