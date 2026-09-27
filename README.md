# Property Underwriting Model — Project Meridian Ridge

A real-estate development underwriting model built as a portfolio methodology demonstration
for Data Analyst / Business Analyst applications. It models a fictional **40-unit weekend-home
development called "Project Meridian Ridge"**, set in a generic Indian hill-station /
weekend-home micro-market.

**All figures in this workbook are invented and illustrative.** No real project, developer,
client, or employer is represented here — this is a standalone, synthetic methodology
demonstration built from scratch, not a real deal or real internal data from any organization.

## What's in this repo

```
model/property_underwriting_model.xlsx   the model (Excel workbook, live formulas)
build_model.py                           the script that generates the workbook (openpyxl)
README.md                                this file
```

`build_model.py` is included so the model is fully reproducible and auditable — every number
in the workbook traces back to a formula or a clearly labeled input cell in the script, not a
manually pasted value.

## Modeling approach

The workbook is built so that changing any assumption (land cost, unit mix, sale price,
absorption pace, financing terms, discount rate) flows through every downstream tab via live
Excel formulas — nothing is hardcoded as a computed result. Blue text marks input/assumption
cells; black is a same-sheet formula; green is a formula linking to another tab.

### Tabs

1. **Summary Dashboard** — headline metrics (total revenue, total development cost, gross
   margin, base-case IRR/NPV, peak funding requirement) and a scenario comparison table, all
   pulled live from the other tabs. Includes two native charts:
   - Base-case cashflow over time (net quarterly cashflow + cumulative cash balance)
   - IRR sensitivity by price multiplier and absorption delay (a bar-chart view of the
     Sensitivity tab's grid; the Sensitivity tab itself also carries a color-scale heatmap)

2. **Assumptions** — the single source of truth: unit count and mix (3 unit types), sale
   prices per sqft, land area and rate, construction cost per sqft, soft costs, contingency,
   marketing cost, phasing/timeline (land closing, construction start/duration, sales launch),
   financing terms (equity share, construction debt rate), the target discount rate, and the
   base-case quarterly absorption curve (units sold per quarter, summing to 40).

3. **Cost Buildup** — rolls up land cost, land approvals, base construction cost, amenities,
   soft costs, contingency, interest-during-construction (IDC, approximated using the average
   -outstanding-balance method — documented in a cell comment), and marketing cost into a
   **Total Development Cost**, plus gross margin, all as formulas referencing Assumptions.

4. **Absorption Schedule** — a 16-quarter (4-year) sales realization schedule: units sold per
   quarter (pulled from the Assumptions absorption curve via `INDEX`, offset by the sales-launch
   quarter), cumulative units, revenue recognized per quarter, and cumulative revenue.

5. **Cashflow** — the full base-case development cashflow: land cost (at closing), construction
   cost (straight-line over the build period), IDC, and marketing cost as outflows; sales
   revenue as inflows; net cashflow and running cumulative cash balance by quarter. Also
   computes the base-case project IRR (via Excel's `IRR()`, quarterly then annualized) and NPV
   (via `NPV()` at the target discount rate) and the peak funding requirement.

6. **Exit Scenarios** — three scenarios (Downside / Base / Upside), each with its own
   price multiplier, absorption-timing delay, and cost-overrun multiplier as adjustable levers.
   Each scenario re-derives its own 16-quarter units-sold, revenue, cost, and net-cashflow rows
   from the same Assumptions/Cashflow building blocks, then computes its own IRR and NPV live.

7. **Sensitivity** — a 2-way grid of annualized IRR (and a companion NPV grid) across a price
   multiplier axis (0.90x–1.10x) and an absorption-pace axis (−1 to +3 quarters faster/slower
   than base). Since the LibreOffice-based recalculation used to verify this workbook does not
   reliably support Excel's native What-If Data Table feature, the grid is built explicitly:
   25 hidden helper cashflow blocks (one per price/delay combination) each feed a live
   `IRR()`/`NPV()` formula into the visible grid. A color-scale conditional format on the IRR
   grid gives it a heatmap look.

### Key modeling simplifications (documented so a reviewer can see the judgment calls)

- Revenue in the base absorption schedule assumes a constant blended unit mix each quarter
  (rather than simulating exactly which unit type sells in which quarter).
- Construction and IDC costs are phased straight-line over the build period rather than an
  S-curve.
- IDC is an average-outstanding-balance approximation, not a full draw-down schedule.
- Exit scenarios hold cost overruns and price/absorption shocks as independent single
  multipliers/offsets rather than a fully joint Monte Carlo distribution.
- The absorption curve itself (units sold per quarter in the base case) is a hardcoded
  assumption input, not a derived output — that is intentional; it represents the underwriting
  team's assumed sales pace.

## Verification

The workbook was built with `openpyxl` (Python), which writes formulas as live formula strings
but does not compute cached values on save. To verify the formulas actually evaluate correctly
(no `#REF!`, `#VALUE!`, `#NAME?`, etc.) in this headless build environment, the file was
recalculated with LibreOffice Calc in headless mode
(`soffice --headless --convert-to xlsx`), which forces a full formula recalculation and writes
the results back into the same cells as cached values. The recalculated file was then
re-opened with `openpyxl` and every cell across all 7 tabs was scanned for error strings:
**zero formula errors found.** Sample outputs were also spot-checked for reasonableness
(e.g., base-case gross margin ≈ 19%, base-case annualized IRR ≈ 38%, Downside/Upside IRRs of
≈ 7% / ≈ 123% bracketing it sensibly, and a Sensitivity grid that moves monotonically with
price and absorption pace as expected).

**This confirms the formulas are correct and syntactically valid, and that LibreOffice's
formula engine evaluates them without error.** It does not by itself guarantee Excel's own
formula engine will produce identical results in every edge case (Excel and LibreOffice can
differ on rare function behaviors), so **you should still open the file in Excel and confirm
it recalculates as expected** (Excel will do this automatically on open; if not, press F9).
The delivered workbook already carries LibreOffice-computed cached values, so it will render
correctly even before Excel's first recalculation.

## Reproducing the workbook

```bash
pip install openpyxl
python build_model.py
```

This regenerates `model/property_underwriting_model.xlsx` from scratch with fresh formulas
(no cached values — recalculate in Excel, or via LibreOffice headless, before viewing outside
Excel).
