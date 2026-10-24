# Battery Storage Project Finance & Capacity Pricing Model

An Excel project finance model for a 400 MW portfolio of front-of-the-meter battery storage and backup generation. It turns capex, O&M, degradation and capacity revenue into monthly and annual cash flows. It then sizes a senior and mezzanine capital stack to target DSCRs, and reports NPV, project and equity IRR, and the **capacity price needed to hit the sponsor's target equity return**.

Every number in the workbook is a live Excel formula driven by input cells, with no pasted values. An independent Python implementation recalculates the model, and a script checks that the two agree for every scenario.

> **Illustrative assumptions only.** The portfolio, prices and costs are realistic but hypothetical. They are not data from a real project, and this is not investment advice. Each input is logged in the workbook's *Assumption Log* with an owner and the source it should be checked against.

📥 **[Download the workbook](model/Battery_Storage_Project_Finance_Model.xlsx)** (open in Microsoft Excel; data tables recalculate on open, or press F9)

---

## Key outputs (base case)

| Metric | Value |
|---|---|
| Total capex (incl. 5% contingency) | $526.6m |
| Senior debt / mezzanine / sponsor equity | $292.7m / $28.9m / $204.9m |
| Total gearing | 61.1% |
| **Equity IRR** (levered, after-tax) | **11.56%** vs. 11.0% target |
| Project IRR (unlevered, after-tax) | 7.46% |
| WACC | 7.38% |
| Project NPV @ WACC | $2.0m |
| Min senior / combined DSCR | 1.35x / 1.20x |
| ITC cash received (30% credit sold at $0.92) | $125.8m |
| Equity payback year | 2029 |
| **Required capacity price** for 11% equity IRR | **$10.87/kW-month** (vs. $11.00 assumed) |

### Scenarios

| | Base | Downside | Upside |
|---|---|---|---|
| Drivers | n/a | capex +10%, price −10%, 6-month COD delay, rates +100 bps, O&M +10% | capex −5%, price +5%, rates −50 bps, O&M −5% |
| Total debt ($m) | 321.6 | 257.1 | 352.2 |
| Gearing | 61.1% | 44.4% | 70.4% |
| Project IRR | 7.46% | 4.25% | 8.97% |
| **Equity IRR** | **11.56%** | **2.67%** | **18.54%** |
| Project NPV ($m) | 2.0 | (117.0) | 59.3 |
| Min senior DSCR | 1.35x | 1.35x | 1.47x (gearing-capped) |

### Sensitivity: equity IRR, capex flex (rows) × capacity price flex (columns)

| Capex \ Price | −20% | −10% | 0% | +10% | +20% |
|---|---|---|---|---|---|
| **−20%** | 11.3% | 17.4% | 23.6% | 26.9% | 29.8% |
| **−10%** | 6.8% | 11.4% | 16.9% | 22.5% | 25.9% |
| **0%** | 3.8% | 7.3% | **11.6%** | 16.5% | 21.4% |
| **+10%** | 1.7% | 4.5% | 7.8% | 11.7% | 16.1% |
| **+20%** | 0.1% | 2.4% | 5.1% | 8.1% | 11.7% |

Equity returns move far more with capacity price and capex than with financing terms. A 12-month construction delay costs about 3.7 points of equity IRR, and +100 bps on both debt tranches costs about 1.7 points.

---

## Portfolio

| Asset | Technology | MW | MWh | Capex ($m) | Base COD |
|---|---|---|---|---|---|
| Alpha | BESS, 4h | 100 | 400 | 138.6 | Dec-2027 |
| Bravo | BESS, 4h | 75 | 300 | 105.5 | Jan-2028 |
| Charlie | BESS, 2h | 50 | 100 | 41.0 | Apr-2028 |
| Delta | BESS, 4h | 125 | 500 | 170.6 | Jul-2028 |
| Echo | Backup generation | 50 | n/a | 70.9 | Jul-2028 |
| **Total** | | **400** | **1,300** | **526.6** | |

## Workbook structure

| Sheet | What it does |
|---|---|
| **Cover** | Purpose, units, conventions, methodology, sheet guide |
| **Summary** | Returns, sources & uses, required capacity price, debt metrics, EBITDA vs. debt service and debt balance charts |
| **Inputs** | Market, tax and financing assumptions; asset-level table (MW, duration, EPC cost, build time, COD, accreditation, degradation, O&M, margin, lease, ITC eligibility, MACRS class); MACRS tables |
| **Scenarios** | Base / Downside / Upside selector, sensitivity flex cells, Excel **Data Tables** (scenario, capex, price, delay, rates, 2-D capex × price) and the required-price solver |
| **Monthly** | Asset-by-asset operating flags, degradation-adjusted accredited MW, capacity and energy revenue, opex and construction capex, Jan-2027 to Dec-2049 |
| **Annual** | Annual roll-up, ITC, MACRS depreciation, sculpted senior and mezzanine debt schedules, tax, project and equity cash flows |
| **Assumption Log** | Every input with category, owner, rationale and validation source; values link live to Inputs |
| **Checks** | 13 integrity checks (debt fully repaid, DSCR floors met, sources = uses, monthly-to-annual tie-outs, depreciation = basis), rolled up to the Summary |

## Methodology

- **Revenue**: capacity payments on accredited MW (nameplate × accreditation × (1 − degradation)^operating year), plus an energy and ancillary margin. Both escalate at 2%/yr.
- **Costs**: fixed O&M, site lease and G&A escalate at 2.5%/yr. Insurance and property tax are a percentage of capex. Capex is spread evenly over each asset's construction period.
- **Tax**: 30% ITC on storage capex, sold under transferability at $0.92 per $1 and received the year after COD. Depreciable basis is reduced by half the ITC. Assets use 5-year MACRS (storage) or 15-year MACRS (generation). Tax losses are assumed usable by the sponsor.
- **Debt sizing**: both tranches are drawn at term conversion (the latest COD) with a 15-year tenor. Senior debt service is sculpted to CFADS ÷ 1.35x, and the senior amount is the lower of the PV of that debt service and 60% of capex. Mezzanine takes the gap down to a 1.20x combined DSCR, capped at 75% total gearing. CFADS for sizing is pre-tax EBITDA, so the model has no circular references.
- **Returns**: project IRR on unlevered after-tax cash flows; equity IRR on levered after-tax cash flows, net of financing fees. Project NPV is discounted at WACC.
- **Required capacity price**: a 41-point data table flexes the capacity price ±40%, and the solver interpolates where equity IRR crosses the target. Goal Seek (Equity IRR → target by changing `Inputs!CapPrice`) gives the same answer.

## Verification

`tools/reference_model.py` is a separate NumPy implementation of the same logic. `tools/recalc_and_verify.sh` recalculates the workbook headless in LibreOffice once per scenario, then compares capex, debt sizing, gearing, WACC, IRRs, NPVs, DSCRs, payback, the integrity checks, every evaluated data-table cell and the required price against the Python model:

```
OK  EqIRR        workbook      0.115619593570377  reference    0.11561959357037344
OK  ProjNPV      workbook       2039.87975551279  reference     2039.8797555127792
...
OK  Data tables: 81 cells match reference
OK  RequiredPrice workbook 10.8656  reference 10.8666
ALL MATCH
```

LibreOffice evaluates only the first two output columns of a multi-column one-variable data table. Excel evaluates all of them, so the script skips the cells LibreOffice leaves blank.

## Reproduce

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python tools/build_workbook.py              # writes model/Battery_Storage_Project_Finance_Model.xlsx
(cd tools && ../.venv/bin/python reference_model.py)  # Python results + required price
tools/recalc_and_verify.sh                            # needs LibreOffice (brew install --cask libreoffice)
```

Assumptions live in `tools/assumptions.py` and feed both the workbook builder and the reference model.

---

Built by [Kshama Upadhyay](https://github.com/Kshama16).
