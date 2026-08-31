# 04 · Data Validation Advanced & Form Controls

Level 1 covered basic dropdown lists. This module covers dependent
(cascading) dropdowns, custom validation formulas, and interactive
Form Controls (option buttons, checkboxes, combo boxes) that let a
sheet respond to user clicks without VBA.

## 1. Worked dataset

Build this table on a sheet named `Catalog`, `A1:B7`:

| | A | B |
|---|---|---|
| **1** | Category | Item |
| **2** | Fruit | Apple |
| **3** | Fruit | Banana |
| **4** | Fruit | Cherry |
| **5** | Veg | Carrot |
| **6** | Veg | Potato |
| **7** | Veg | Onion |

## 2. Named ranges for dependent dropdowns

1. Build two named lists: select `B2:B4` (Apple, Banana, Cherry) and
   name it `Fruit` (Formulas → Define Name). Select `B5:B7` (Carrot,
   Potato, Onion) and name it `Veg`.
2. Type `Fruit` in `D1` and `Veg` in `E1` as the two category labels
   the first dropdown will offer.
3. In `G1`, set Data Validation → List → Source `=D1:E1` (Fruit, Veg).
   This is the first dropdown.
4. In `G2`, set Data Validation → List → Source
   `=INDIRECT(G1)`. `INDIRECT` turns the *text* `"Fruit"` selected in
   `G1` into a reference to the named range `Fruit`, so `G2`'s list
   changes to Apple/Banana/Cherry when `G1="Fruit"`, and to
   Carrot/Potato/Onion when `G1="Veg"`. This is the standard
   cascading-dropdown pattern.
5. Manual check: set `G1` to `Veg`. `INDIRECT("Veg")` resolves to the
   named range `Veg` = `B5:B7` = Carrot, Potato, Onion — the `G2`
   dropdown should offer exactly those three, not the fruits.

## 3. Custom validation formulas

1. Select `H2`, Data Validation → Custom, Formula:
   `=AND(H2>0,H2<=100)`
   This restricts entry to numbers strictly between 0 and 100
   inclusive of 100 — typing `150` or `0` is rejected, `100` is
   accepted. Trace by hand: `H2=100` → `AND(100>0,100<=100)` →
   `AND(TRUE,TRUE)` → `TRUE`, allowed. `H2=0` → `AND(FALSE,TRUE)` →
   `FALSE`, rejected.
2. Select `I2`, Data Validation → Custom, Formula:
   `=COUNTIF($I$2:$I$10,I2)=1`
   This blocks duplicate entries anywhere in `I2:I10` — if the value
   being typed already appears once elsewhere in the range, the count
   becomes `2` and the formula returns `FALSE`, rejecting the entry.

## 4. Form Controls — Combo Box and Check Box

1. Enable the Developer tab (File → Options → Customize Ribbon →
   check **Developer**).
2. Developer → Insert → Form Controls → **Combo Box**. Draw it near
   `K1`. Right-click → Format Control → Input range `B2:B7`, Cell
   link `K2`. Selecting an item in the combo box now writes that
   item's **position number** (1–6) into `K2`, not the text itself —
   this is the key gotcha with Form Control combo boxes.
3. In `K3`, retrieve the actual text: `=INDEX(B2:B7,K2)`. If the combo
   box shows "Potato" selected (the 5th item in `B2:B7`), `K2` holds
   `5` and `K3` becomes `=INDEX(B2:B7,5)` = `Potato`. Confirm by
   counting: Apple(1), Banana(2), Cherry(3), Carrot(4), Potato(5),
   Onion(6) — position 5 is indeed Potato.
4. Developer → Insert → Form Controls → **Check Box**. Right-click →
   Format Control → Cell link `L1`. Checking the box writes `TRUE` to
   `L1`, unchecking writes `FALSE` — use `=IF(L1,"Included","Excluded")`
   in `L2` to turn that boolean into a readable label.

## Cheat sheet

| Task | How |
|---|---|
| List validation | Data Validation → List → Source range or comma list |
| Cascading dropdown | Named ranges per category + `=INDIRECT(cell)` as Source |
| Custom rule | Data Validation → Custom → formula returning TRUE/FALSE |
| No duplicates | `=COUNTIF($range,cell)=1` as Custom formula |
| Combo Box (Form Control) | Input range + Cell link (returns position number) |
| Position → text | `=INDEX(range,cell_link)` |

## Exercise

Build the cascading dropdown in Section 2, then set `G1` to `Fruit`
and confirm `G2` only offers Apple/Banana/Cherry. Add a custom
validation rule on a new cell `M2` that only accepts even numbers:
`=MOD(M2,2)=0` — verify by hand that `4` is accepted (`MOD(4,2)=0`)
and `7` is rejected (`MOD(7,2)=1`).
