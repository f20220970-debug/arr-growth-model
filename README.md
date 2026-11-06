# ARR Growth Model (Excel + Python)

[![tests](https://github.com/f20220970-debug/arr-growth-model/actions/workflows/tests.yml/badge.svg)](https://github.com/f20220970-debug/arr-growth-model/actions/workflows/tests.yml)

A driver-based plan for a brokerage that earns **commission on insurance premiums**. It answers the founder's questions:

- **When do we reach $10M ARR?**
- **Does the cash last until then?**
- **Which driver matters most?**

The model is built twice, and the two versions are tested to agree to the dollar:

- **`ARR_Growth_Model.xlsx`:** 336 live formulas. Change any blue input, or flip the scenario switch between Bear, Base and Bull, and every number updates.
- **`model.py`:** a Python version of the same model, used for 5,000 Monte Carlo runs and the sensitivity analysis.

> **All inputs are illustrative** (6,000 leads a month, 15% lead-to-bind, $2,300 average premium, 13% commission, 81% renewal, $5M starting cash). None of them are company data. Change them in the workbook.

Run it: `pip install -r requirements.txt && python model.py && python build_excel.py && pytest -q`. After rebuilding, recalculate the workbook in Excel or LibreOffice.

## How the model works

```
leads(m)        = leads(m-1) x (1 + lead growth)
new policies    = leads x lead-to-bind
churned         = active(m-1) x (1 - renewal^(1/12))
active policies = active(m-1) + new - churned
ARR             = active policies x avg premium x commission rate
costs           = leads x cost per lead + ROUNDUP(leads / leads per agent) x agent cost + fixed opex
cash(m)         = cash(m-1) + ARR/12 - costs
```

## Results

| | Bear | **Base** | Bull |
| --- | --- | --- | --- |
| ARR at month 12 | $3.2M | **$5.4M** | $8.9M |
| ARR at month 24 | $6.3M | **$14.4M** | $31.6M |
| $10M ARR reached in | not within 24 months | **month 20** | month 14 |
| Cash runs out in | month 13 | **month 19** | never |
| Extra cash needed | $4.9M | **$1.3M** | 0 |
| LTV / CAC | 2.5× | **5.9×** | 13.2× |
| CAC payback | 19 months | **10.7 months** | 6.5 months |

**The headline:** in the Base case, the plan reaches $10M ARR in month 20 but **runs out of cash in month 19**. Growth is healthy (LTV/CAC 5.9×), but marketing and agent costs scale with leads every month, while commission arrives gradually as the book builds. The decision is either to raise about **$1.3M plus a buffer by month 15**, or to bring cost per lead down. Both levers can be tested in the workbook.

**Monte Carlo (5,000 runs around Base):**
- ARR at month 24: P10 / P50 / P90 = **$9.6M / $14.3M / $21.5M**.
- Chance of reaching $10M ARR by month 18: **33%**. By month 24: **88%**.
- **Chance cash goes below zero: 63%.**

**Sensitivity (ARR at month 18, each driver moved ±20%):** in order of impact, revenue per policy (premium × commission) moves it most, then lead-to-bind conversion, then lead growth. Renewal rate matters least over an 18-month horizon, though it dominates long-term LTV.

## Files

```
ARR_Growth_Model.xlsx   Inputs (scenario switch) · Summary · Model (24 months) · Risk
model.py                Python mirror + Monte Carlo + sensitivity -> results.json
build_excel.py          writes the workbook (formulas, colour-coded inputs)
tests/                  Excel = Python for the Base case; scenarios ordered; book never negative
```

Companion repos: [soma-brokerage-ops-lab](https://github.com/f20220970-debug/soma-brokerage-ops-lab) and [revenue-pipeline-forecaster](https://github.com/f20220970-debug/revenue-pipeline-forecaster).
