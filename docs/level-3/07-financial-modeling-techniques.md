# 07 · Financial Modeling Techniques

This module covers the core building blocks of a financial model: NPV,
IRR, amortization schedules, and structuring a model with clearly
separated inputs, calculations, and outputs — the standard layout
professional analysts use.

## 1. Worked dataset — an investment decision

Build this on a sheet named `Model`, `A1:B6`:

| | A | B |
|---|---|---|
| **1** | Initial Investment | -10000 |
| **2** | Year 1 Cash Flow | 3000 |
| **3** | Year 2 Cash Flow | 4000 |
| **4** | Year 3 Cash Flow | 5000 |
| **5** | Year 4 Cash Flow | 3500 |
| **6** | Discount Rate | 0.10 |

## 2. NPV

1. In `B8`: `=B1+NPV(B6,B2:B5)`
   `NPV()` discounts a series of *future* cash flows (it assumes the
   first value in the range occurs one period from now), so the
   upfront investment in `B1` (which happens at time 0) is added
   separately, outside the NPV function.
2. Manual check, discounting each flow by `1.10^n`:
   Year1: `3000/1.10 = 2727.27`
   Year2: `4000/1.10^2 = 4000/1.21 = 3305.79`
   Year3: `5000/1.10^3 = 5000/1.331 = 3756.57`
   Year4: `3500/1.10^4 = 3500/1.4641 = 2390.55`
   Sum of discounted flows: `2727.27+3305.79+3756.57+2390.55 =
   12180.18`. Add `B1 = -10000`: NPV ≈ `2180.18`. A positive NPV means
   the investment is expected to create value at a 10% discount rate.

## 3. IRR

1. In `B9`: `=IRR(B1:B5)`
   IRR is the discount rate at which NPV would be exactly zero.
2. Sanity check: since NPV at 10% is positive (`+2180.18`), the true
   IRR must be **higher** than 10% (a higher discount rate shrinks NPV
   toward zero). Excel's IRR should return roughly `21%` — confirm the
   returned rate is indeed above 10%, consistent with the NPV sign.

## 4. Amortization schedule

Build a loan schedule for a `20000` loan at `6%` annual, 5 years,
`E1:H6`:

| | E | F | G | H |
|---|---|---|---|---|
| **1** | Year | Payment | Interest | Principal |
| **2** | 1 | `=-PMT($B$6,5,20000)` | `=20000*$B$6` | `=E2payment-G2` |

Simplify with concrete formulas instead, `E1:H6`:

1. `F2`: `=-PMT(0.06,5,20000)` ≈ `4747.11` (annual payment).
2. `G2` (Year 1 interest): `=20000*0.06 = 1200`.
3. `H2` (Year 1 principal repaid): `=F2-G2 ≈ 4747.11-1200=3547.11`.
4. Remaining balance after Year 1: `20000-3547.11=16452.89`. Year 2
   interest: `=16452.89*0.06≈987.17`. Year 2 principal:
   `≈4747.11-987.17=3759.94`. Notice interest shrinks and principal
   grows each year as the balance declines — the hallmark shape of an
   amortization schedule.

## 5. Sensitivity — separating inputs from calculations

1. Best practice: keep every raw assumption (discount rate, growth
   rate, loan rate) in a clearly labeled "Inputs" block (like `B1:B6`
   here), and have every calculation formula reference those cells —
   never hardcode a rate inside a calculation formula. This is why
   Section 2's NPV formula references `$B$6` rather than typing `0.10`
   directly: changing one input cell should flow through the entire
   model.
2. Test this: change `B6` from `0.10` to `0.08` and confirm `B8`'s NPV
   recalculates upward (a lower discount rate increases the present
   value of future cash flows) without editing any other formula.

## Cheat sheet

| Function | Purpose |
|---|---|
| `NPV(rate, values)` | Present value of a series of *future* cash flows |
| `IRR(values)` | Discount rate that makes NPV = 0 |
| `PMT(rate,nper,pv)` | Fixed periodic payment for a loan/annuity |
| Amortization | Interest = balance × rate; Principal = Payment − Interest |

## Exercise

Recompute NPV at a discount rate of `15%` instead of `10%` and confirm
it drops compared to Section 2's `2180.18` (a higher discount rate
always lowers NPV, all else equal) — then check whether it goes
negative, which would mean the investment fails to clear a 15% hurdle
rate.
