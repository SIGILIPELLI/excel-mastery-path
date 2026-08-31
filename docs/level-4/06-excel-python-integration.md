# 06 · Excel + Python Integration (Overview)

This module covers the two current ways to combine Excel and Python:
the built-in `=PY()` formula (Python in Excel), and driving Excel files
from standalone Python via `openpyxl`/`pandas` — and when to reach for
each.

## 1. Worked dataset

Sheet `Data`, `A1:B6`:

| | A | B |
|---|---|---|
| **1** | Product | Units |
| **2** | Widget | 120 |
| **3** | Gadget | 95 |
| **4** | Gizmo | 140 |
| **5** | Sprocket | 60 |
| **6** | Doohickey | 110 |

## 2. Python in Excel — `=PY()`

1. In a cell, `=PY(` opens a Python-enabled formula cell (available on
   Microsoft 365 builds with Python in Excel enabled). Reference sheet
   ranges as pandas objects using `xl()`:
   `df = xl("A1:B6", headers=True)`
   `df["Units"].sum()`
   This returns `120+95+140+60+110=525`, computed by the Python runtime
   but displayed as a normal cell result.
2. Python in Excel runs in a secure, isolated container — it can use
   pandas, NumPy, and Matplotlib, but it cannot reach the local file
   system or network, so a `=PY()` cell can analyze the workbook's own
   data but cannot pull in an external CSV or API directly.
3. A quick chart from the same cell:
   `df.plot(kind="bar", x="Product", y="Units")`
   renders as an image output in the cell — useful for exploratory
   analysis that would otherwise need a full chart object.

## 3. Driving Excel from standalone Python — `openpyxl`

```python
import openpyxl

wb = openpyxl.load_workbook("Data.xlsx")
ws = wb["Data"]

total = 0
for row in ws.iter_rows(min_row=2, max_row=6, min_col=2, max_col=2, values_only=True):
    total += row[0]

ws["B8"] = "Total"
ws["B9"] = total
wb.save("Data.xlsx")
```

Manual check: `total` sums the same five values as Section 2,
`120+95+140+60+110=525`, and the script writes `525` into `B9` of the
saved file — this is the pattern for batch-processing many workbooks
unattended (a folder of 200 monthly reports, for example), which
`=PY()` inside a single open workbook cannot do.

## 4. `pandas` for heavier transforms

```python
import pandas as pd

df = pd.read_excel("Data.xlsx", sheet_name="Data")
df["Units_x1.1"] = df["Units"] * 1.1
df.to_excel("Data_updated.xlsx", index=False)
```

For `Widget` (`120`), `Units_x1.1` computes `120*1.1=132.0`; for
`Gizmo` (`140`), `140*1.1=154.0`. `pandas` is preferred over `openpyxl`
for row-wise numeric transforms across large tables because its
vectorized operations avoid a Python-level loop per cell.

## 5. Choosing between the three approaches

| Situation | Best tool |
|---|---|
| One-off analysis inside a workbook someone else will open and read normally | `=PY()` — stays inside the familiar Excel file |
| Automating the same edit across many `.xlsx` files on a schedule, no Excel install required | Standalone `openpyxl`/`pandas` script |
| Heavy numeric transforms (joins, groupby, reshaping) on data already in Excel | `pandas`, then write back with `to_excel` |
| Reading a value or writing a result but preserving existing formatting/charts exactly | `openpyxl` (cell-level control) over `pandas` (which rewrites the whole sheet) |

## 6. A caution on round-tripping

`pandas.to_excel` and even `openpyxl` can strip conditional formatting,
data validation, or chart objects that pandas doesn't understand when
it rewrites a sheet. For a workbook with dashboards/formatting that
must survive the edit, prefer `openpyxl`'s in-place cell writes (as in
Section 3) over `pandas.to_excel`'s full-sheet rewrite (Section 4).

## Cheat sheet

| Tool | Runs where | Best for |
|---|---|---|
| `=PY()` | Inside the open workbook, sandboxed | Exploratory analysis, in-cell charts |
| `openpyxl` | Standalone Python | Precise cell writes, preserving formatting |
| `pandas` | Standalone Python | Heavy tabular transforms, at the cost of formatting fidelity |

## Exercise

Write an `openpyxl` script that opens `Data.xlsx`, computes the average
units (`525/5=105`), writes it to `B10` labeled "Average" in `A10`, and
saves the file — then open it in Excel and confirm `B10` reads `105`.
