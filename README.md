# Open ITR Suite — ITR-1 / ITR-2 / ITR-4, AY 2026-27

Income-tax return preparation software for **ITR-1 (SAHAJ), ITR-2 and ITR-4 (SUGAM),
Assessment Year 2026-27 (Previous Year 2025-26)**, built directly from the official
CBDT artefacts in this repository — with **automatic form detection**, an advanced
salary engine (perquisites, terminal benefits, arrears → Form 10E), a full
capital-gains module with Excel import/export, and import pipelines for AIS, Form
26AS and previous-year return JSONs.

| File | Role in the software |
|---|---|
| `ITR-1_2026_Main_V1.1.json` | Official JSON Schema (draft-04) — generated ITR-1 is validated against it, field for field |
| `ITR-2_2026_Main_V1.1.json` / `ITR-4_2026_Main_V1.1.json` / `ITR-3_2026_Main_V1.1.json` | Schemas driving the ITR-2 / ITR-4 generators (and eligibility analysis for ITR-3) |
| `CBDT_*Validation Rules*AY 2026-27*.pdf` | Source of the Category A / B / D validation rules implemented by the rule engines |
| `ITR 1_Schema change document_AY2026-27_V1.1.pdf` | Confirms the schema version in force (V1.1, 30-06-2026) |
| `ITR1_AY_26-27_V1.2.xlsm` / `ITR2_…V1.3.xlsm` (`CSV_112A.csv`, `CSV_TDS*.csv`, `CSV_IT.csv`) / `ITR3…` / `ITR4…` | Department utilities — behavioural reference; their CSV layouts are understood by the importers |

## What it does

* **Guided web wizard** (9 steps: eligibility → personal → salary → house property →
  other income → deductions → taxes paid → bank → review & file) that produces a
  ready-to-upload ITR-1 JSON.
* **Live computation panel** — updates while you type: slab-wise tax, rebate, cess,
  interest & fee, refund/balance for the chosen regime, a both-regimes comparison
  strip, and **live error detection** (CBDT rule violations and auto-adjustments are
  flagged in the panel as you fill).
* **Mandatory-field mapping per step** — every wizard step knows its required fields
  (PAN, DOB, address, PIN, mobile, email, bank account, verification details, …).
  Missing/invalid fields are listed live at the top of the step, the offending inputs
  are highlighted red, and you **cannot move to the next step, jump ahead on the
  stepper, or download the JSON/PDF/Word until they are fixed**. Conditional
  requirements are checked too (revised-return acknowledgement, representative
  details, alternate address, complete-or-empty challan rows, TDS claimed ≤
  deducted, co-owner share % between 1–100, let-out rent, …).
* **Errors name the exact field** — no more bare Python messages like
  `invalid literal for int() with base 10: ''`. A mistyped value produces e.g.
  `Standard deduction override (Salary income): 'abc' is not a whole number (rupees)
  — enter digits only (commas are fine, e.g. 75000 or 75,000).` The web API also
  returns the wizard field path so the UI highlights the exact input. Indian comma
  grouping (`12,00,000`), spaces, `₹`/`Rs.` marks and decimal text are all accepted
  and normalised automatically.
* **Detailed computation view + PDF / Word export** — a full *Computation of Total
  Income and Tax Liability* sheet (Parts A/B/C: income heads, Chapter VI-A with
  claimed-vs-allowable, slab-wise tax, surcharge & 87A notes, 234A/B/C workings with
  instalment detail, taxes paid, refund/payable) for **both regimes**, visible in the
  Review tab and downloadable as **PDF** or **Word (.docx)**.
* **Dual-regime tax engine** — computes liability under the new regime (default,
  u/s 115BAC) and the old regime side-by-side, and recommends the cheaper one.
* **Full statutory computation**: slab rates FY 2025-26, age-based old-regime slabs,
  rebate u/s 87A (₹60,000 new / ₹12,500 old) with new-regime **marginal relief**
  (threshold ₹12,70,590 per CBDT rule 191), surcharge with marginal relief, 4% H&E
  cess, rounding u/s 288A/288B.
* **Interest & fees**: auto-computed 234A / 234B (from actual challan dates & filing
  date), 234C instalment deferment, late fee 234F (₹1,000/₹5,000), revised-return fee
  234-I — all user-overridable.
* **Every schedule of ITR-1**: salary with exempt allowances & HRA schedule 10(13A),
  up to 2 house properties (co-ownership, 24(b) loan schedule), other sources incl.
  dividend quarterly breakup & family-pension 57(iia), LTCG 112A up to ₹1.25 lakh,
  exempt income u/s 10, all Chapter VI-A deductions (80C, 80CCC, 80CCD(1/1B/2), 80D,
  80DD/80U, 80DDB, 80E/80EE/80EEA/80EEB, 80G/80GGA/80GGC, 80GG, 80TTA/80TTB, 80CCH),
  TDS-1/2/3, TCS, advance-tax/self-assessment challans, bank accounts, verification.
* **Regime-aware eligibility capping** — e.g. old regime applies 80C+80CCC+80CCD(1) ≤
  ₹1.5 lakh, 24(b) self-occupied interest ≤ ₹2 lakh, HRA least-of-three; new regime
  keeps only 80CCD(2) & 80CCH, standard deduction ₹75,000, and disallows HP loss
  set-off — exactly per the CBDT validation rules.
* **Two-layer validation before download**:
  1. JSON-Schema validation against `ITR-1_2026_Main_V1.1.json`;
  2. ~110 implemented CBDT Category A (blocking) / B (139(9) notice) / D (claim-forms)
     rules, each tagged with its rule number from the official PDF.
* **CLI** for building, regime comparison, and validating any official ITR-1 JSON file.

## Multi-form suite (v2)

Everything in the ITR-1 list stays true; v2 adds a complete multi-form filing suite:

* **Auto form detection** — the Start tab offers a radio group (Auto / ITR-1 / ITR-2 /
  ITR-4 / ITR-3) and the server engine recommends the right form from your income
  profile with CBDT-aligned reasons (capital gains beyond 112A limit → ITR-2,
  presumptive business → ITR-4, books-of-account business → ITR-3 notice, RNOR /
  foreign assets / directorship → ITR-2/3). Eligibility per form is shown as badges.
* **ITR-2 generator** — complete Part-A GEN, Schedule S (per-employer salary with
  nature-wise earnings/perquisites/profits), Schedule HP (multiple properties),
  full capital-gains schedules: **111A @ 20%, 112A @ 12.5% with ₹1.25L exemption and
  31-01-2018 grandfathering per scrip, 12.5%-without-indexation vs 20%-indexed for
  property (auto lower-tax choice per the 2024 amendment), quarter-wise accrual for
  234C, Schedule 112A with ISIN scrip rows, §54/54EC/54F exemption machinery**, CYLA,
  SI, EI, 80D schedule, relief §89/90/91 blocks — every generated document validated
  with zero errors against the official ITR-2 schema.
* **ITR-4 generator** — presumptive 44AD (6% digital / 8% cash), 44ADA (50%), 44AE
  (₹7,500/vehicle·month, ₹1,000/ton·month heavy), with turnover ceilings (₹3 cr /
  ₹75 lakh), max(declared, presumptive) logic, business/profession code catalogues,
  Schedule BP financial particulars, 112A LTCG capped at ₹1.25L, 10-IEA flag handling.
* **Advanced salary engine** — employer block, payslip earnings grid (component
  catalogue: BASIC/DA/HRA/LTA/CEA/SPECIAL/ARREARS/… with statutory §10 exemptions
  auto-limited per regime), **perquisites (17(2)) with a rule-3 valuation catalogue**, terminal
  benefits with **least-of-three exemption workings** (gratuity 10(10), leave
  encashment 10(10AA) ₹25L cap, commuted pension 10(10A), VRS 10(10C) ₹5L cap,
  retrenchment 10(10B)), **arrears with automatic relief u/s 89 + Form-10E working
  table**, and a **monthly salary grid** (double-click the 17(1) box) that aggregates
  into the annual figures.
* **Capital-gains module with Excel round-trip** — five sections (equity ST, equity LT
  with grandfathering, land & building, other ST, other LT) + exemptions; export the
  workbook to `.xlsx`, re-import it (or the department's `CSV_112A.csv` layout), and
  download a JSON template.
* **Imports** —
  * previous-year **ITR JSON** (any form/AY): prefill personal/bank/verification
    (empty fields only), capture **brought-forward losses** into the UI, and drive a
    **year-on-year comparison** table;
  * **AIS** JSON or the password-protected ZIP (password = PAN lowercase + DOB
    ddmmyyyy — **auto-detected** from your profile), with TDS/TCS/salary/interest/
    dividend/rent classification, an **“as per AIS vs as per return” comparison**,
    and **click-to-add income** straight into the return;
  * **Form 26AS** `.txt` or password-protected PDF (password = DOB ddmmyyyy,
    auto-detected; AES-256 supported), parsing Part-A1/A2/B/C into tax credits with
    dedupe-safe merge.
* **Computation sheets for every form** — the Review tab's detailed computation
  (and the downloadable **PDF / Word**) works for ITR-1/2/4 with annexures: salary
  break-up, terminal-benefit least-of-three tables, capital-gain worksheet +
  transaction list + quarterly accrual, presumptive financials, and the Form-10E
  relief table.
* **Multi-form validator** — `validate_all(doc, form=…)` runs the official schema of
  the chosen form plus a generic CBDT rule set (identity, verification, refund
  banking, 80CCE/80TTA/80TTB ceilings, 87A guard rails per regime, tax-paid
  arithmetic, the liability-minus-paid balance equation) — clean builds yield zero
  issues; tampered documents fail loudly.

## Quick start

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

### Web app (guided wizard)

```bash
python run_web.py            # → http://127.0.0.1:5000/
```

Click **“Load sample data”** in the app to explore a pre-filled ITR-1 return, then
**“Download ITR JSON”** on the Review tab. Two richer worked examples ship in
`samples/` — `sample_itr2.json` (salary with perks & terminal benefits, house
property, equity + property gains, §54EC) and `sample_itr4.json` (44AD shop + 44ADA
tuition + 44AE goods vehicles): loading them via the API exercise every v2 feature.

### Command line

```bash
# compare both tax regimes for a taxpayer
python itr_cli.py compare samples/sample_taxpayer.json

# build the official e-filing JSON (schema- and rule-validated)
python itr_cli.py build samples/sample_taxpayer.json --out ITR1_AY2026-27.json

# validate ANY official ITR-1 JSON (e.g. one you edited by hand)
python itr_cli.py validate ITR1_AY2026-27.json
```

### Tests

```bash
pytest tests/                    # 340+ tests: tax engine, per-form builders (schema
                                 # conformance on rich & minimal samples), rule engine,
                                 # imports (AIS zip / 26AS AES-PDF / ITR JSON), CG
                                 # Excel round-trip, computation reports, web API,
                                 # multi-form validator, utility field/NOB/residential/
                                 # special-rate parity
node tests/wizard_steps_check.js # wizard mandatory-field mapping (52 checks)
```

## Input data model

The builder consumes a friendly JSON document (see `samples/sample_taxpayer.json` for
a complete, commented example). Every derived figure — gross totals, eligible
deductions, schedule totals, tax, refund/balance — is *computed*, so the output is
internally consistent with the CBDT rules by construction. Highlights of the model:

```jsonc
{
  "personal":   { "first_name": "…", "pan": "…", "dob": "1991-06-15",
                  "employer_category": "OTH", "address": { … } },
  "filing":     { "return_section": 11,            // 139(1); 12 = 139(4), 17 = revised…
                  "opt_out_new_regime": false,      // false → new regime (default)
                  "filing_date": "2026-07-15" },
  "salary":     { "salary_17_1": 960000, "hra": { … }, "exempt_allowances": [ … ] },
  "house_properties": [ { "occupancy": "L", "loans": [ … ] } ],   // max 2
  "other_sources":    [ { "nature": "SAV", "amount": 8500 } ],
  "ltcg112a":   { "sale_consideration": 250000, "cost_of_acquisition": 150000 },
  "deductions": { "80c": {…}, "80d": {…}, "80g": {…}, "80ccd2": 96000, … },
  "tax_paid":   { "tds_on_salary": […], "tds_others": […], "tcs": […],
                  "challans": [ { "bsr_code": "…", "date": "…", "amount": 0 } ] },
  "bank_accounts":    [ { "ifsc": "…", "account_no": "…", "use_for_refund": true } ],
  "verification": { "name": "…", "father_name": "…", "place": "…", "capacity": "S" }
}
```

## Architecture

```
itr1/
  constants.py    AY 2026-27 statutory parameters (slabs, caps, dates, enums)
  tax_engine.py   slabs, surcharge & marginal relief, 87A rebate, cess, rounding,
                  234A/234B/234C interest (with workings), 234F/234-I fees
  builder.py      friendly input  →  official ITR-1 JSON (all derived values computed)
  itr2_builder.py friendly input  →  official ITR-2 JSON (Schedule S/HP/CG/112A/CYLA/
                  SI/EI/80D, relief §89, per-quarter accrual)
  itr4_builder.py friendly input  →  official ITR-4 JSON (44AD/44ADA/44AE, Schedule BP)
  forms.py        ITR-form eligibility radar + build dispatcher
  salaryx.py      salary-slip break-up, perks, terminal benefits, arrears → Form 10E
  capgains.py     capital-gains engine (111A/112A/12.5%/20%-indexed, set-off,
                  exemptions, quarterly accrual) + xlsx/CSV import-export
  imports.py      previous-year ITR JSON, AIS (json/zip), 26AS (txt/AES-256 PDF),
                  AIS↔return comparison, YoY
  validator.py    jsonschema layer + CBDT Category A/B/D rule engine (per form)
  schemakit.py    multi-form official-schema loader
  report.py       verbose line-by-line computation (both regimes, all forms)
  compfile.py     PDF (reportlab) + Word (python-docx) computation-sheet renderers
  webapp.py       Flask API + wizard host (summary/generate/download per form,
                  computation json/pdf/docx, imports, cg xlsx, salary helpers…)
web/              single-page wizard (vanilla JS, no build step)
itr_cli.py        CLI: build / compare / validate
tools/            extract_xlsm.py — decompile the official .xlsm utilities
extracted/xlsm/   full VBA sources + workbook structure of every official
                  ITR-1/2/3/4 Excel utility (see extracted/README.md)
samples/          worked examples: sample_taxpayer (ITR-1), sample_itr2, sample_itr4
tests/            340+ unit & integration tests + wizard node harness
```

## Official Excel utilities: macros & structure extracted

Every official `.xlsm` preparation utility in this repo has been decompiled
read-only into `extracted/xlsm/<FORM>/` — **all 465 VBA modules
(~6.5 lakh lines) plus a complete workbook inventory** (65→21→24→66 sheets
per utility incl. hidden reference sheets, 2,800+ dropdown validation rules,
15,000+ defined names, protection flags, companion CSV templates). Re-run
any time with `python tools/extract_xlsm.py` (needs `oletools`).

Findings we folded back into the software (`tests/test_official_parity.py`
pins them):

* **Identity block parity** — `CreationInfo`/`Form_ITRn` now mirror the
  department utility byte-for-byte for AY 2026-27: `SWVersionNo`
  R3/R4/R4/R2 per form, `IntermediaryCity` Delhi, AY `2026`, `Ver1.0`,
  `Digest "-"`, and the utilities' own form-description wording
  (`itr1/__init__.py` keeps the extracted constants;
  `SW90002627` remains the *department's* id — this software ships with its
  own `SW…` placeholder, replace it with your registered SW number).
* **Emission shape** — the exact JSON key order and which schedules the
  utilities always emit vs. emit only for updated returns (139(8A)) or for
  the new regime (ITR-4 deduction schedules).
* **Mechanics of `Digest`** — iterative HMAC-SHA-256 injected at publish
  time; the portal-facing JSON carries `-`, which is what we emit too.
* **All sheets' cell contents dumped** (hidden *and* visible:
  `extracted/xlsm/*/sheets/`, ~1.13M values incl. every formula): the
  utilities' dropdown masters — 88.5k/124.3k-row ISIN lists, 40.4k IFSC
  codes, bank/state/country/nature-of-business lists — plus `dropdowns.json`
  with every dropdown rule resolved to its actual values. Cross-checking our
  code against those masters **caught and fixed 5 wrong
  nature-of-business codes** in the ITR-4 44AD pick-list (pinned forever by
  `tests/test_master_list_parity.py`).
* **Field maps for every form** (`fields.json` + readable `fields.md` per
  utility — 376/2,729/3,874/494 fields for ITR-1/2/3/4): every input cell
  with the utility's own defined name (`sheet1.PAN`…), its on-screen label
  and its dropdown values — the same addresses the official VBA reads.
* **Proved end-to-end mapping** (`itr1/fieldmap.py` +
  `tests/test_field_mapping.py`, 114 tests): for every field on the
  utilities' personal-information & filing-status pages, the crosswalk pins
  utility defined-name → our input key → official JSON path, and a **sentinel
  flow test** injects a value at the input and asserts it lands at the
  documented JSON node. Per form: **ITR-1 36 mapped** (+1 partial),
  **ITR-2 35 mapped** (+3 partial), **ITR-4 31 mapped** (+3 partial);
  every exemption (landline, foreign ZIP, RNOR day-count grids, ward/circle,
  Goa's Portuguese-Code 5A, Form-10IEA history grid…) is explicit with its
  reason. Doing this audit surfaced and fixed four real gaps: ITR-2
  residential status (was hard-coded RES), ITR-2/ITR-4 secondary address,
  ITR-2/ITR-4 secondary mobile/email, ITR-2 Aadhaar — plus the wizard now
  collects Status and Residential status.
* **Special-rate income (Schedule SI)** (`itr1/special_rates.py`): the full
  `SecCode` table from the official schema (111A / LTCG / 115BB / 115BBJ /
  115BBE / 115A-family / FII / DTAA / PTI) with the OS-family buckets wired
  end-to-end in ITR-2 — lottery 115BB @30%, online-games 115BBJ @30%,
  unexplained 115BBE @60% **plus the flat 25% surcharge on 115BBE tax the
  utilities' TaxCalc sheets compute separately**. Unsupported codes raise a
  helpful error instead of producing a wrong document.
* **Residential status now drives computation exactly like the macros**:
  the ITR-4 utility's `Lock44AD` (locks & clears the 44AD/44ADA cells for
  'NR') → our builder refuses NRI + 44AD/44ADA; RNOR stays eligible (it is
  resident u/s 6). Rebate u/s 87A is 0 for non-residents and HUFs (utility
  formula `Rebate87Aformula_new`, portal rules 533/534), senior-citizen
  slabs apply only to RES/NOR (utility `resStatus` checks), ITR-1 rejects
  NRI/NOR filers (ROR-only), and the eligibility radar recommends ITR-2/3
  accordingly. HUF + 44ADA is rejected (portal rule 212).
* **Full official nature-of-business masters** (`itr1/business_codes.py`,
  generated by `tools/build_business_codes.py` from the utilities' own
  dropdown lists): 356-code master, **315 codes for 44AD**, **38 for
  44ADA** (incl. 16021 Social Media Influencers), **7 for 44AE** — any code
  outside the official dropdown is rejected at build time, and the wizard
  shows searchable pick-lists fed from the same data.

See `extracted/README.md` for the full field-level tables and notes.

## Tax computation notes (AY 2026-27)

* **New regime slabs** (u/s 115BAC): nil to ₹4L; 5% ₹4–8L; 10% ₹8–12L; 15% ₹12–16L;
  20% ₹16–20L; 25% ₹20–24L; 30% above ₹24L. Standard deduction ₹75,000. Rebate u/s
  87A to ₹60,000 ⇒ **zero tax up to ₹12 lakh** taxable income (₹12.75L for salaried);
  marginal relief covers the ₹12L–₹12,70,590 band.
* **Old regime slabs**: 2.5L/5L/10L with 5%/20%/30% (senior citizens: ₹3L basic
  exemption; super senior: ₹5L); standard deduction ₹50,000; rebate ₹12,500 up to
  ₹5 lakh total income.
* LTCG u/s 112A in ITR-1 is limited to ₹1,25,000 (exempt); gains beyond need ITR-2.
* ITR-1 eligibility: resident individual; total income ≤ ₹50 lakh + LTCG ≤ ₹1.25L;
  agriculture income ≤ ₹5,000; no business/profession, capital gains (other than the
  above), foreign income/assets, etc.

## Compliance notes

* `CreationInfo.SWCreatedBy` carries the placeholder `SW00000001`. Third-party software
  providers integrating with the e-Filing portal must substitute the **SW registration
  number issued by the e-Filing project** (pattern `SW` + 8 digits) before production
  use; keep the pattern valid or schema validation will fail.
* The software computes returns per the published schema and validation rules. The
  e-Filing portal additionally cross-verifies data with PAN/26AS/AIS masters at upload
  time (name vs PAN, TDS credits, Aadhaar linkage), which no offline tool can check.
* **ITR-3** is supported as a *recommendation + import* target only: auto-detection
  explains exactly when it is required (business with books of account, companies
  partnership, etc.) and previous-year ITR-3 JSONs parse fine for prefill/comparison,
  but generation (balance-sheet, P&L) is intentionally left to the department utility.
* Brought-forward losses are captured from imported files and shown in the Imports tab
  for the CFL schedule; automatic set-off against the current year's heads is **not**
  wired into the computation yet (flagged in the Review notes).
* Interest u/s 234A/234B/234C uses documented statutory logic with challan dates; the
  portal's authoritative computation may differ in edge cases — overrides are provided
  (`interest_override` or the "Taxes paid" tab).

## Disclaimer

This is an assistive return-preparation tool and not an official Income-tax Department
product. Verify figures against Form 16/16A/26AS/AIS before filing. Consult a tax
professional for advice.
