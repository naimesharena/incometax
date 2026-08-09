# Official Excel utility extraction (`extracted/xlsm/`)

The four official CBDT Excel preparation utilities checked into this
repository were opened **read-only** and fully decompiled here:

| Utility (source artefact) | Worksheets | VBA modules | Lines of VBA | Dropdown rules | Formulas |
|---|---|---|---|---|---|
| `ITR1_AY_26-27_V1.2.xlsm` | 21 (5 visible) | 73 | ~73,600 | 387 | 1,018 |
| `ITR2_AY_26-27_V1.3/…V1.3.xlsm` | 66 (25 visible) | 155 | ~236,700 | 1,122 | 6,196 |
| `ITR3_AY_26-27_V1.2/…V1.2.xlsm` | 66 (40 visible) | 158 | ~240,600 | 1,257 | 6,449 |
| `ITR4_AY_26-27_V1.1.xlsm` | 24 (7 visible) | 79 | ~96,900 | 522 | 1,545 |

Per utility directory (`extracted/xlsm/<FORM>/`):

- **`vba/`** — every VBA module decompiled from `xl/vbaProject.bin`
  (`.bas` modules, `.cls` sheet/workbook classes, `.frm` forms), verbatim.
- **`structure.json`** — machine-readable workbook inventory: sheet names,
  visibility states, used-range dimensions, formula/merge counts per sheet,
  cell-protection flags, all 2,800+ data-validation (dropdown) rules with
  their source ranges/lists, all 15,000+ defined names, customXml payloads,
  and the companion CSV import templates that ship next to each utility.
- **`sheets/`** — **full cell contents of every worksheet** (hidden *and*
  visible: 21/66/66/24 dumps with ~1.13 million stored values and every
  formula extracted alongside into
  `.formulas.tsv`): the dropdown master lists the utilities hide from users —
  `ISIN List` (88.5k / 124.3k rows in ITR-2/ITR-3), `IFSC` (40.4k), `DB` /
  `BankCode` / `DropDownValues` (nature-of-business, state, country, employer
  category, section codes…), `Nature Of Business`, `HelpCSV`, `ITold`, plus
  all hidden calc schedules (`TaxCalc`, `Tax Calculated`, `BP`, `DPM - DOA`,
  `80G`…). Tab-separated grid (first column = row number); dumps larger than
  ~1 MB are stored as `.tsv.gz`.
- **`dropdowns.json`** — every list-typed data-validation rule resolved to its
  actual values (inline lists, defined names, hidden-sheet ranges); the few
  conditionally-dynamic `IF(...) / INDIRECT(...)` rules are kept verbatim with
  their formulas. Non-list numeric bounds are recorded as `constraint` rules.
- **`fields.json` / `fields.md`** — the **field map of every visible form
  sheet** (376 fields in ITR-1, 2,729 in ITR-2, 3,874 in ITR-3, 494 in ITR-4):
  each input addressed four ways — cell (`AK8`), the utility's own defined
  name (`sheet1.PAN`), the on-screen label (`PAN`), and its resolved dropdown
  values when it is a pick-list. `fields.md` renders each sheet exactly in the
  on-screen layout (e.g. ITR-3 `Part A Gen`: First/Middle/Last Name, PAN,
  Status, Date of Commencement of Business, Date of Birth/Formation,
  Residential Status with `RES/NRI/NOR`, addresses, `Filed u/s`…). This is the
  same information the utility's VBA reads via named ranges.
- **`summary.md`** — the same layout as readable tables.

Regenerate everything (official binaries are never modified):

```bash
python tools/extract_xlsm.py            # all four utilities
python tools/extract_xlsm.py <file.xlsm>  # a single utility
```

Requires `oletools` (`pip install oletools`); the rest is the Python
standard library.

## Identity block extracted from the utilities' VBA

These exact values are written by the utilities into every generated JSON,
and this software now **mirrors them** (see `itr1/__init__.py` and
`tests/test_official_parity.py`):

| Field (`CreationInfo` / `Form_ITRn`) | ITR-1 (V1.2) | ITR-2 (V1.3) | ITR-3 (V1.2) | ITR-4 (V1.1) |
|---|---|---|---|---|
| `SWVersionNo` | R3 | R4 | R4 | R2 |
| `SWCreatedBy` / `JSONCreatedBy` | SW90002627 | SW90002627 | SW90002627 | SW90002627 |
| `IntermediaryCity` | Delhi | Delhi | Delhi | Delhi |
| `Digest` (pre-upload) | `-` | `-` | `-` | `-` |
| `AssessmentYear` | 2026 | 2026 | 2026 | 2026 |
| `SchemaVer` / `FormVer` | Ver1.0 | Ver1.0 | Ver1.0 | Ver1.0 |
| `Form_ITRn.Description` | For Indls having Income from Salary, Pension, family pension and Interest | For Individuals and HUFs not having income from profits and gains of business or profession | *(bug: utility reuses ITR-1 wording; we use the correct ITR-3 wording)* | For Individuals having Income From Presemptive Business *(verbatim, incl. spelling)* |

Notes on what the VBA revealed:

- **JSON envelope**: root `ITR` → `<ITRn>` → `CreationInfo`, `Form_ITRn`,
  then the schedules (key order we now match): ITR-1 emits
  PersonalInfo, FilingStatus, (`PartA_139_8A`, `PartB-ATI` *only for updated
  returns u/s 139(8A)*), ITR1_IncomeDeductions, ITR1_TaxComputation,
  TDSonSalaries, TDSonOthThanSals, ScheduleTDS3Dtls, ScheduleTCS,
  TaxPayments, TaxReturnPreparer, Refund, Verification.
  ITR-2 additionally always emits ScheduleS/HP/CGFor23/112A/115AD/OS/CYLA/
  BFLA/CFL/PartB-TI/PartB_TTI/IT/TDS1-3/TCS/VDA/80-series schedules, and
  conditionally ScheduleSPI, ScheduleTR1, ScheduleFA and ScheduleAL
  (income > ₹50L). ITR-4 emits the 80C/80D/… detail schedules **only for the
  new (115BAC) regime** — matching that schema's `UsrVIA` design, which our
  builders already follow.
- **Digest**: computed at the very end as an iterative HMAC-SHA-256 over the
  assembled JSON (`getHashKey()`/`getHashIteration()` in GenerateJson),
  then injected as `CreationInfo.Digest` (44-char base64). The schema
  accepts `-` for it, which is what the utility emits until its
  publish/upload step — this software emits `-` as well.
- **ITR-2 and ITR-3 share one workbook/VBA codebase**, switched by
  `getITRNo`; the ITR-2 utility simply keeps the ITR-3-only schedules
  (Part A-BS, Manufacturing/Trading/Profit & Loss accounts, OI, QD, DPM/DOA/
  DEP/DCG, ESR, UD, ICDS, TPSA, GST) hidden and their JSON emission
  commented out.
- **Hidden reference sheets** hold every dropdown source list
  (`ISIN List`, `Nature Of Business`, `ITold`, `HelpCSV`, pincode/IFSC style
  masters…) — 41 hidden of 66 in ITR-2, 16 of 21 in ITR-1, 17 of 24 in ITR-4 —
  and their full contents are now dumped under `sheets/`.
- **The extraction caught a real bug**: our ITR-4 44AD business-code
  pick-list had 5 of 6 wrong codes (e.g. `05001` labelled "Civil construction"
  is actually electricity generation; `06011` doesn't exist — cafés are
  `10007/10008` restaurant codes). All 44AD/44ADA/44AE catalogs in
  `itr1/itr4_builder.py` are now verified code-by-code against the official
  hidden masters and pinned by `tests/test_master_list_parity.py`.
- Sheets and the VBA project are protected; `PWD.bas` is merely an
  interactive protect/unprotect helper (no embedded password).
- `ImportExcel.bas` is the utility-side importer for the companion CSVs
  (`CSV_112A.csv` is the one our capital-gains Excel/CSV import is
  compatible with; 115AD/TDS/TCS/IT templates are catalogued in each
  `structure.json → companionCSVs`).
- `PreFillJson.bas`/`ImportJson.bas` are the utility's own prefill/import
  paths — the same shapes as the portal-downloaded JSON that our
  `itr1/imports.py` (`/api/import/itr-json`) already ingests.
- u/s 139(8A) *updated return* blocks (`PartA_139_8A`, `PartB-ATI`) exist in
  all four utilities but are emitted only for updated returns; this software
  currently supports original/revised filings (139(1)/(5)) — see the main
  README scope notes.

## How the extraction maps to this software

| Utility piece | This software |
|---|---|
| `GenerateJson.bas` (ITR-1) | `itr1/builder.py` |
| `GenerateJson.bas` (ITR-2/ITR-3 shared) | `itr1/itr2_builder.py` (ITR-3: out of scope, see main README) |
| `GenerateJSON.bas` (ITR-4) | `itr1/itr4_builder.py` + `itr1/forms.py` |
| Salary sheets / `ModuleEA10_13A` | `itr1/salaryx.py` |
| `ScheduleCGFor23`/`Schedule112A` + `CSV_112A.csv` | `itr1/capgains.py` (`/api/cg/*` xlsx/CSV round-trip) |
| `ImportJson` / `PreFillJson` | `itr1/imports.py` (`/api/import/itr-json`, YoY compare) |
| Dropdown validity rules (`structure.json → dataValidations`) | enforced via the official JSON schemas in `itr1/validator.py` |
| Identity constants (table above) | `itr1/__init__.py` + `tests/test_official_parity.py` |
| Field maps (`fields.json`) — identity & filing pages | `itr1/fieldmap.py` crosswalk + `tests/test_field_mapping.py` sentinel flows (utility field → our input → official JSON path, all provable) |
| NOB dropdowns (`NOB44AD`/`NOB44ADA`/`NOB` + ITR-3 `Nature_of_Business`) | `tools/build_business_codes.py` → `itr1/business_codes.py` (356-code master, 315/38/7 per-section lists, verbatim labels) + `tests/test_business_codes.py` |
| Residential-status macros (`Lock44AD`, `resStatus` slab gating, `Rebate87Aformula_new`) | `itr1/itr4_builder.py` & `itr1/builder.py` build-time gates; `itr1/tax_engine.py` (`residential_status`/`status` threading: 87A = 0 for NRI/HUF, senior slabs RES/NOR-only); `itr1/forms.py` radar; `tests/test_residential_status.py` |
| Special-rate TaxCalc annotations (115BBE flat 25% surcharge, SI rates) | `itr1/special_rates.py` (`SecCode` = schema-verbatim SI codes incl. DTAA/PTI; 115BB/115BBJ/115BBE wired end-to-end in ITR-2) + `tests/test_special_rates.py` |
