# 06 · Building Custom Functions with LAMBDA

`LAMBDA` lets you define a reusable custom function entirely in
worksheet formulas — no VBA required — and give it a name via the Name
Manager so it behaves like a built-in function everywhere in the
workbook.

## 1. Worked dataset

Build this on a sheet named `Prices`, `A1:C5`:

| | A | B | C |
|---|---|---|---|
| **1** | Item | Cost | MarkupPct |
| **2** | Pen | 10 | 0.20 |
| **3** | Book | 50 | 0.15 |
| **4** | Bag | 200 | 0.10 |
| **5** | Gizmo | 80 | 0.25 |

## 2. A simple LAMBDA

1. In `D2`, without naming it yet, test the raw expression:
   `=LAMBDA(cost,pct,cost*(1+pct))(B2,C2)`
   A `LAMBDA` defines parameters then a formula using them; the
   trailing `(B2,C2)` immediately calls it with those arguments.
2. Manual check: `10*(1+0.20)=12`. `D2` should show `12`.

## 3. Naming it as a reusable function

1. Formulas → Define Name. Name: `SellPrice`. Refers to:
   `=LAMBDA(cost,pct,cost*(1+pct))`
2. Now in `D2:D5`, use it like a built-in function:
   `=SellPrice(B2,C2)`
3. Manual check all four rows: Pen `10*1.20=12`; Book `50*1.15=57.5`;
   Bag `200*1.10=220`; Gizmo `80*1.25=100`.

## 4. LAMBDA with conditional logic

1. Define a second named LAMBDA, `PriceTier`:
   `=LAMBDA(price,IF(price>=150,"Premium",IF(price>=50,"Standard","Budget")))`
2. Apply it: `=PriceTier(D2)` for each row. Using the sell prices from
   Section 3 (12, 57.5, 220, 100): Pen(12)→Budget, Book(57.5)→Standard,
   Bag(220)→Premium, Gizmo(100)→Standard.

## 5. Recursive LAMBDA (advanced)

1. `LAMBDA` can call itself by name if defined with that name via
   Define Name, enabling recursion without VBA. Define `FactorialFn`:
   `=LAMBDA(n,IF(n<=1,1,n*FactorialFn(n-1)))`
2. Test with `=FactorialFn(5)`. Manual check: `5! = 5*4*3*2*1 = 120`.
   Trace the recursion: `FactorialFn(5)=5*FactorialFn(4)
   =5*4*FactorialFn(3)=5*4*3*FactorialFn(2)=5*4*3*2*FactorialFn(1)
   =5*4*3*2*1=120`.

## 6. Combining LAMBDA with LET for readability

1. `LET` names intermediate values inside one formula, which pairs
   well with LAMBDA for clarity. Define `Margin`:
   `=LAMBDA(cost,price,LET(profit,price-cost,margin,profit/price,margin))`
2. Test `=Margin(B2,D2)` for Pen: cost `10`, price `12` (from Section
   3). `profit=12-10=2`, `margin=2/12≈0.1667` (16.67%).

## Cheat sheet

| Construct | Syntax |
|---|---|
| Inline LAMBDA | `=LAMBDA(params,formula)(args)` |
| Named custom function | Define Name → `=LAMBDA(params,formula)`, then call `=Name(args)` |
| Recursive LAMBDA | Reference the function's own Defined Name inside its formula |
| LET inside LAMBDA | `LAMBDA(x,LET(name,expr,...,result))` |

## Exercise

Define a named LAMBDA `RoundedMargin` that wraps `Margin` from Section
6 and rounds the result to 1 decimal place as a percentage:
`=LAMBDA(cost,price,ROUND(Margin(cost,price)*100,1))`. Test on Bag
(cost `200`, sell price `220` from Section 3): margin `= (220-200)/220
≈ 0.0909` → `9.1` after rounding to 1 decimal on the percentage.
