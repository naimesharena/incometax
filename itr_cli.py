#!/usr/bin/env python3
"""ITR-1 (Sahaj) command-line tool for AY 2026-27.

Examples
--------
Build a return JSON from a taxpayer data file (see samples/sample_taxpayer.json)::

    python itr_cli.py build samples/sample_taxpayer.json --out ITR1_AY2026-27.json

Compare the two tax regimes for a taxpayer::

    python itr_cli.py compare samples/sample_taxpayer.json

Validate an existing official ITR-1 JSON against the CBDT schema + rules::

    python itr_cli.py validate ITR1_AY2026-27.json
"""

from __future__ import annotations

import argparse
import json
import sys

from itr1.builder import build_return, compute_both_regimes
from itr1.errors import DataError
from itr1.validator import validate_all, summarize


def _fmt(n):
    return f"{int(round(n or 0)):,}"


def _regime_row(label, rep):
    tax = rep["tax"]
    tp = tax["TaxesPaid"]
    return [
        f"  {label}",
        f"    Gross total income (incl. LTCG) : Rs. {_fmt(rep['gross_total_income_incl_ltcg'])}",
        f"    Chapter VI-A deductions       : Rs. {_fmt(rep['deductions_eligible']['TotalChapVIADeductions'])}",
        f"    Total income (rounded 288A)   : Rs. {_fmt(rep['total_income'])}",
        f"    Tax on total income           : Rs. {_fmt(tax['TotalTaxPayable'])}",
        f"      (-) Rebate u/s 87A          : Rs. {_fmt(tax['Rebate87A'])}",
        f"      (+) H&E cess @ 4%           : Rs. {_fmt(tax['EducationCess'])}",
        f"    Gross tax liability           : Rs. {_fmt(tax['GrossTaxLiability'])}",
        f"      (-) Relief u/s 89           : Rs. {_fmt(tax['Section89'])}",
        f"      (+) Interest & fee 234*     : Rs. {_fmt(tax['TotalIntrstPay'])}"
        f"  [A:{_fmt(tax['IntrstPay']['IntrstPayUs234A'])} B:{_fmt(tax['IntrstPay']['IntrstPayUs234B'])}"
        f" C:{_fmt(tax['IntrstPay']['IntrstPayUs234C'])} F:{_fmt(tax['IntrstPay']['LateFilingFee234F'])}"
        f" I:{_fmt(tax['IntrstPay'].get('FeeFurnish234I', 0))}]",
        f"    Total tax + interest          : Rs. {_fmt(tax['TotTaxPlusIntrstPay'])}",
        f"    Taxes paid (TDS {_fmt(tp['TDS'])} + TCS {_fmt(tp['TCS'])} + "
        f"AT {_fmt(tp['AdvanceTax'])} + SAT {_fmt(tp['SelfAssessmentTax'])})"
        f" : Rs. {_fmt(tp['TotalTaxesPaid'])}",
        f"    Balance tax payable           : Rs. {_fmt(tax['BalTaxPayable'])}",
        f"    Refund due                    : Rs. {_fmt(tax['RefundDue'])}",
        "",
    ]


def cmd_build(args):
    with open(args.input, encoding="utf-8") as fh:
        data = json.load(fh)
    result = build_return(data)
    doc, report = result["json"], result["report"]

    issues = validate_all(doc)
    summ = summarize(issues)

    print("=" * 70)
    print("ITR-1 (SAHAJ) - AY 2026-27 | regime:", report["regime"].upper())
    print("=" * 70)
    for line in _regime_row("Chosen regime computation", report):
        print(line)

    if report["warnings"]:
        print("Build notes:")
        for w in report["warnings"]:
            print("  *", w)
        print()

    if issues:
        print("Validation findings:")
        for i in issues:
            print(f"  [{i.category}] {i.rule}: {i.message}" +
                  (f" ({i.path})" if i.path else ""))
    else:
        print("Validation: PASS - JSON schema valid; no CBDT rule defects detected.")

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, indent=2)
        print(f"\nWritten: {args.out}")
    elif not args.no_stdout:
        print("\n" + json.dumps(doc, indent=2))

    if not summ["ok"]:
        return 2 if not args.force else 0
    return 0


def cmd_compare(args):
    with open(args.input, encoding="utf-8") as fh:
        data = json.load(fh)
    both = compute_both_regimes(data)
    print("=" * 74)
    print("REGIME COMPARISON - AY 2026-27")
    print("=" * 74)
    for line in _regime_row("NEW regime (default, u/s 115BAC)", both["new"]["report"]):
        print(line)
    for line in _regime_row("OLD regime (with Chapter VI-A deductions & exemptions)", both["old"]["report"]):
        print(line)
    tnew = both["new"]["report"]["tax"]["TotTaxPlusIntrstPay"]
    told = both["old"]["report"]["tax"]["TotTaxPlusIntrstPay"]
    better, saving = ("OLD", tnew - told) if told < tnew else ("NEW", told - tnew)
    print(f"Recommendation: {better} regime saves Rs. {_fmt(abs(saving))}.")
    for name in ("new", "old"):
        for w in both[name]["report"]["warnings"]:
            print(f"  * [{name}] {w}")
    return 0


def cmd_validate(args):
    with open(args.input, encoding="utf-8") as fh:
        doc = json.load(fh)
    issues = validate_all(doc)
    if not issues:
        print("Validation: PASS - JSON schema valid; no CBDT rule defects detected.")
        return 0
    print("Validation findings:")
    for i in issues:
        print(f"  [{i.category}] {i.rule}: {i.message}" + (f" ({i.path})" if i.path else ""))
    summ = summarize(issues)
    print(f"\nSummary: schema={summ['SCHEMA']} category-A={summ['A']} "
          f"category-B={summ['B']} category-D={summ['D']}")
    return 2 if not summ["ok"] else 0


def main(argv=None):
    parser = argparse.ArgumentParser(prog="itr_cli", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build", help="Build the official ITR-1 JSON from taxpayer data")
    p_build.add_argument("input", help="taxpayer data JSON (see samples/)")
    p_build.add_argument("--out", help="output file for the official ITR JSON")
    p_build.add_argument("--no-stdout", action="store_true", help="do not print the JSON")
    p_build.add_argument("--force", action="store_true",
                         help="exit 0 even if validations fail")
    p_build.set_defaults(func=cmd_build)

    p_cmp = sub.add_parser("compare", help="Compare tax under the new and old regimes")
    p_cmp.add_argument("input", help="taxpayer data JSON")
    p_cmp.set_defaults(func=cmd_compare)

    p_val = sub.add_parser("validate", help="Validate an official ITR-1 JSON document")
    p_val.add_argument("input", help="official ITR JSON file")
    p_val.set_defaults(func=cmd_validate)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except DataError as exc:                      # pinpoint the bad field
        print(f"\nInput error — please fix this field and retry:\n  {exc}",
              file=sys.stderr)
        if getattr(exc, "path", ""):
            print(f"  (field: {exc.field} · model path: {exc.path})", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
