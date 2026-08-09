"""Detailed computation report ("Computation of Total Income & Tax").

Builds a verbose, line-by-line breakdown for BOTH tax regimes from the
taxpayer input data.  The same structure is rendered by:

* the web UI (Review tab, "Detailed computation" section),
* the PDF export  (:func:`itr1.compfile.computation_pdf`),
* the Word export (:func:`itr1.compfile.computation_docx`).

Every line is ``{"label": str, "note": str, "amount": int|None}``; a ``None``
amount is a header/spacer row.  Sections keep their subtotals next to them so
layout engines stay simple.
"""

from __future__ import annotations

import copy
import datetime as _dt

from . import ASSESSMENT_YEAR, PREVIOUS_YEAR, __version__
from . import constants as C
from . import tax_engine as T
from .builder import build_return


def _row(label, amount=None, note="", level=0):
    return {"label": label, "amount": amount, "note": note, "level": level}


def _inr(n):
    return f"Rs. {int(round(n or 0)):,}"


# Chapter VI-A line catalogue shared by the ITR-1 and multi-form paths.
_VIA_META = [
    ("Section80C", "80C", "life insurance, PPF, ELSS, EPF etc. (80C+80CCC+80CCD(1) capped at Rs. 1,50,000)"),
    ("Section80CCC", "80CCC", "pension fund contribution"),
    ("Section80CCDEmployeeOrSE", "80CCD(1)", "own NPS contribution (10% of salary / 20% of GTI)"),
    ("Section80CCD1B", "80CCD(1B)", "additional NPS (max Rs. 50,000)"),
    ("Section80CCDEmployer", "80CCD(2)", "employer NPS contribution (10%/14% of salary; allowed in both regimes)"),
    ("Section80D", "80D", "health insurance & medical expenditure (Schedule 80D)"),
    ("Section80DD", "80DD", "dependent with disability (Rs. 75,000 / 1,25,000)"),
    ("Section80DDB", "80DDB", "specified disease treatment (Rs. 40,000 / 1,00,000)"),
    ("Section80E", "80E", "education loan interest"),
    ("Section80EE", "80EE", "first-home loan interest (max Rs. 50,000)"),
    ("Section80EEA", "80EEA", "affordable housing loan interest (max Rs. 1,50,000)"),
    ("Section80EEB", "80EEB", "electric vehicle loan interest (max Rs. 1,50,000)"),
    ("Section80G", "80G", "donations (Schedule 80G; qualifying limit = 10% of adjusted GTI)"),
    ("Section80GG", "80GG", "rent paid where no HRA (max Rs. 60,000 / 25% of adjusted income)"),
    ("Section80GGA", "80GGA", "donations for scientific research / rural development"),
    ("Section80GGC", "80GGC", "contributions to political parties (other mode only)"),
    ("Section80U", "80U", "self with disability (Rs. 75,000 / 1,25,000)"),
    ("Section80TTA", "80TTA", "savings account interest (max Rs. 10,000)"),
    ("Section80TTB", "80TTB", "interest for senior citizens (max Rs. 50,000)"),
    ("AnyOthSec80CCH", "80CCH", "Agniveer Corpus Fund (both regimes)"),
]


def _regime_computation(data: dict, opt_out: bool) -> dict:
    cp = copy.deepcopy(data or {})
    cp.setdefault("filing", {})["opt_out_new_regime"] = opt_out
    built = build_return(cp)
    doc, rep = built["json"], built["report"]
    itr1 = doc["ITR"]["ITR1"]
    inc = itr1["ITR1_IncomeDeductions"]
    tax = rep["tax"]
    personal = cp.get("personal", {}) or {}
    dob = personal.get("dob") or ""

    age_note = ""
    if T.is_super_senior_citizen(dob):
        age_note = "Super senior citizen (80+) slab rates applied."
    elif T.is_senior_citizen(dob):
        age_note = "Senior citizen (60+) slab rates applied."

    # ---------------------------------------------------------------- A: income
    salary_rows = [
        _row("Salary as per section 17(1)", inc.get("Salary")),
        _row("Value of perquisites u/s 17(2)", inc.get("PerquisitesValue")),
        _row("Profits in lieu of salary u/s 17(3)", inc.get("ProfitsInSalary")),
        _row("Less: allowances exempt u/s 10", -inc["AllwncExemptUs10"].get("TotalAllwncExemptUs10", 0),
             "includes HRA exemption as per Schedule 10(13A), where claimed"),
        _row("Net salary", inc.get("NetSalary"), level=1),
        _row("Less: standard deduction u/s 16(ia)", -inc.get("DeductionUs16ia", 0),
             f"maximum {_inr(C.STD_DEDUCTION_16IA_OLD if opt_out else C.STD_DEDUCTION_16IA_NEW)}"),
        _row("Less: entertainment allowance u/s 16(ii)", -inc.get("EntertainmentAlw16ii", 0)),
        _row("Less: professional tax u/s 16(iii)", -inc.get("ProfessionalTaxUs16iii", 0)),
    ]
    exempt_rows = [{
        "label": f"    u/s {r.get('SalNatureDesc')}",
        "amount": r.get("SalOthAmount"),
    } for r in inc["AllwncExemptUs10"].get("AllwncExemptUs10Dtls", []) or []]

    hp_sections = []
    for idx, p in enumerate(inc.get("PropertyDetails", []) or [], start=1):
        rd = p["Rentdetails"]
        occ = {"L": "Let out", "D": "Deemed let out", "S": "Self occupied"}.get(p.get("ifLetOut"))
        share = float(p.get("AsseseeShareProperty", 100.0))
        rows = [
            _row("Gross rent / annual lettable value", rd.get("AnnualLetableValue")),
            _row("Less: unrealised rent + municipal taxes", -rd.get("TotalUnrealizedAndTax", 0)),
            _row("Balance annual value", rd.get("BalanceALV"), level=1),
        ]
        if share < 100:
            rows.append(_row(f"Assessee's share ({share:g}%)", rd.get("AnnualOfPropOwned"),
                             "co-owned property", 1))
        rows += [
            _row("Less: 30% standard deduction u/s 24(a)", -rd.get("ThirtyPercentOfBalance", 0)),
            _row("Less: interest on borrowed capital u/s 24(b)", -rd.get("IntOnBorwCap", 0)),
            _row("Add: arrears / unrealised rent (70% taxable)", rd.get("ArrearsUnrealizedRentRcvd", 0)),
        ]
        hp_sections.append({
            "title": f"Property {idx} ({occ})",
            "rows": rows,
            "income": rd.get("IncomeOfHP"),
        })
    hp_note = ""
    if not opt_out and rep["hp_income"] == 0 and hp_sections:
        hp_note = ("Under the new regime no interest u/s 24(b) is allowed for "
                   "self-occupied property and a house-property loss cannot be set off.")

    os_rows = []
    for r in (cp.get("other_sources") or []):
        if not r.get("amount"):
            continue
        os_rows.append(_row(C.OTHER_SOURCE_NATURES.get(r.get("nature"), r.get("nature")),
                            int(r.get("amount", 0) or 0),
                            r.get("description", "")))
    if inc.get("DeductionUs57iia"):
        cap = C.FAMILY_PENSION_DED_CAP_OLD if opt_out else C.FAMILY_PENSION_DED_CAP_NEW
        os_rows.append(_row("Less: deduction u/s 57(iia) on family pension",
                            -inc.get("DeductionUs57iia"),
                            f"lower of 1/3rd of family pension or {_inr(cap)}"))

    # ---------------------------------------------------------------- VIA
    usr = inc["UsrDeductUndChapVIA"]
    ded = inc["DeductUndChapVIA"]
    via_rows = []
    for key, label, note in _VIA_META:
        u = usr.get(key, 0) or 0
        e = ded.get(key, 0) or 0
        if not u and not e:
            continue
        n = note
        if e < u:
            n = (note + f" — restricted: claimed {_inr(u)}, "
                 f"allowable {_inr(e)}")
        via_rows.append({"section": label, "note": n, "claimed": u, "eligible": e})
    if not opt_out:
        kept = [r for r in via_rows if r["eligible"] > 0]
        if len(via_rows) > len(kept):
            via_caption = "Under the new regime only 80CCD(2) and 80CCH are admissible; the rest are shown at NIL."
        else:
            via_caption = "Under the new regime only 80CCD(2) and 80CCH are admissible."
    else:
        via_caption = ""

    # ---------------------------------------------------------------- B: tax
    ti = rep["total_income"]
    ltcg = rep["ltcg112a"]
    working = tax.get("Working", {})
    slab_rows = T.slab_breakup(max(0, ti - ltcg), opt_out, dob or "1990-01-01")
    surcharge = tax.get("SurchargeComputed", 0)
    if surcharge > 0:
        sur_note = ("10% surcharge as total income exceeds Rs. 50 lakh; marginal "
                    "relief limits (tax + surcharge) to tax on Rs. 50 lakh plus the "
                    "excess income.")
    else:
        sur_note = "No surcharge (total income within Rs. 50 lakh)."
    interest_rows = [
        ("IntrstPayUs234A", "Interest u/s 234A (delay in filing)"),
        ("IntrstPayUs234B", "Interest u/s 234B (default in advance tax)"),
        ("IntrstPayUs234C", "Interest u/s 234C (deferment of instalments)"),
        ("LateFilingFee234F", "Fee u/s 234F (late filing)"),
        ("FeeFurnish234I", "Fee u/s 234-I (revised return after 31-12-2026)"),
    ]
    interest = []
    detail = tax.get("InterestDetail", {})
    for key, label in interest_rows:
        d = detail.get(key, {})
        interest.append({"label": label, "amount": tax["IntrstPay"].get(key, 0),
                         "note": d.get("note", "")})
        if key == "IntrstPayUs234C" and d.get("instalments"):
            interest[-1]["instalments"] = d["instalments"]

    # ---------------------------------------------------------------- C: paid
    paid = tax["TaxesPaid"]
    result_note = ""
    if tax["BalTaxPayable"] > 0:
        result_note = (f"Balance tax payable: {_inr(tax['BalTaxPayable'])} — pay as "
                       f"self-assessment tax before filing to avoid further interest.")
    elif tax["RefundDue"] > 0:
        result_note = f"Refund due: {_inr(tax['RefundDue'])} — ensure the refund bank account is correct."

    return {
        "regime": "old" if opt_out else "new",
        "regime_title": ("Old tax regime (with Chapter VI-A deductions & exemptions)"
                         if opt_out else
                         "New tax regime (default, u/s 115BAC)"),
        "age_note": age_note,
        "salary": {
            "rows": salary_rows,
            "exempt_rows": exempt_rows,
            "income": rep["salary_income"],
        },
        "house_properties": hp_sections,
        "hp_note": hp_note,
        "hp_income": rep["hp_income"],
        "other_sources": {"rows": os_rows, "income": rep["other_sources_income"]},
        "gti_excl_ltcg": rep["gross_total_income_excl_ltcg"],
        "ltcg": ltcg,
        "ltcg_note": ("LTCG u/s 112A up to Rs. 1,25,000 is exempt and reported for "
                      "information." if ltcg else ""),
        "gti_incl_ltcg": rep["gross_total_income_incl_ltcg"],
        "via_rows": via_rows,
        "via_caption": via_caption,
        "via_total": rep["deductions_eligible"]["TotalChapVIADeductions"],
        "total_income": ti,
        "total_income_note": "rounded off to the nearest multiple of ten u/s 288A",
        "slab_rows": slab_rows,
        "slab_tax": working.get("slab_tax", 0),
        "ltcg_tax": working.get("ltcg_tax", 0),
        "surcharge": surcharge,
        "surcharge_note": sur_note,
        "tax_payable": tax["TotalTaxPayable"],
        "rebate": tax["Rebate87A"],
        "rebate_note": working.get("rebate_note", ""),
        "tax_after_rebate": tax["TaxPayableOnRebate"],
        "cess": tax["EducationCess"],
        "cess_note": "Health & Education Cess @ 4% on tax after rebate",
        "gross_tax_liability": tax["GrossTaxLiability"],
        "relief89": tax["Section89"],
        "relief89_note": ("relief on salary arrears — furnish Form 10E" if tax["Section89"] else ""),
        "net_tax_liability": tax["NetTaxLiability"],
        "interest": interest,
        "total_interest": tax["TotalIntrstPay"],
        "tot_tax_plus_interest": tax["TotTaxPlusIntrstPay"],
        "taxes_paid": {
            "tds": paid["TDS"], "tcs": paid["TCS"],
            "advance_tax": paid["AdvanceTax"],
            "self_assessment_tax": paid["SelfAssessmentTax"],
            "total": paid["TotalTaxesPaid"],
        },
        "balance_payable": tax["BalTaxPayable"],
        "refund_due": tax["RefundDue"],
        "result_note": result_note,
        "warnings": rep["warnings"],
    }


# ---------------------------------------------------------------------------
# Multi-form path (ITR-2 / ITR-4) — same renderer contract plus annexures:
# salary-slip workings, capital-gain schedules, presumptive business and the
# Form-10E relief table.
# ---------------------------------------------------------------------------

_CG_SPAN = {"Q1": "1 Apr – 15 Jun", "Q2": "16 Jun – 15 Sep", "Q3": "16 Sep – 15 Dec",
            "Q4": "16 Dec – 15 Mar", "Q5": "16 – 31 Mar"}
_CG_KIND = {"111a": "STCG u/s 111A", "112a": "LTCG u/s 112A",
            "apprate_st": "STCG (other assets)", "ltcg_rate": "LTCG (non-112A)"}


def _num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def _via_rows_generic(usr_via: dict, ded_via: dict, opt_out: bool):
    rows = []
    for key, label, note in _VIA_META:
        u = _num(usr_via.get(key))
        e = _num(ded_via.get(key))
        if not u and not e:
            continue
        n = note + (f" — restricted: claimed {_inr(u)}, allowable {_inr(e)}"
                    if e < u else "")
        rows.append({"section": label, "note": n, "claimed": int(u), "eligible": int(e)})
    return rows


def _hp_sections_from_doc(property_details):
    """Rentdetails rows — identical layout in the ITR-1 and ITR-2 schemas."""
    out = []
    for idx, p in enumerate(property_details or [], start=1):
        rd = p.get("Rentdetails") or {}
        occ = {"L": "Let out", "D": "Deemed let out", "S": "Self occupied"}.get(
            p.get("ifLetOut"), p.get("ifLetOut", ""))
        share = float(p.get("AsseseeShareProperty", 100.0) or 100.0)
        rows = [
            _row("Gross rent / annual lettable value", rd.get("AnnualLetableValue")),
            _row("Less: unrealised rent + municipal taxes",
                 -rd.get("TotalUnrealizedAndTax", 0)),
            _row("Balance annual value", rd.get("BalanceALV"), level=1),
        ]
        if share < 100:
            rows.append(_row(f"Assessee's share ({share:g}%)",
                             rd.get("AnnualOfPropOwned"), "co-owned property", 1))
        rows += [
            _row("Less: 30% standard deduction u/s 24(a)",
                 -rd.get("ThirtyPercentOfBalance", 0)),
            _row("Less: interest on borrowed capital u/s 24(b)",
                 -rd.get("IntOnBorwCap", 0)),
            _row("Add: arrears / unrealised rent (70% taxable)",
                 rd.get("ArrearsUnrealizedRentRcvd", 0)),
        ]
        out.append({"title": f"Property {idx} ({occ})", "rows": rows,
                    "income": rd.get("IncomeOfHP")})
    return out


def _os_rows_generic(data: dict, opt_out: bool, ded57_reported: int):
    rows = []
    for r in data.get("other_sources") or []:
        if not r.get("amount"):
            continue
        rows.append(_row(C.OTHER_SOURCE_NATURES.get(r.get("nature"),
                                                    r.get("nature")),
                         int(r.get("amount", 0) or 0), r.get("description", "")))
    if ded57_reported:
        cap = C.FAMILY_PENSION_DED_CAP_OLD if opt_out else C.FAMILY_PENSION_DED_CAP_NEW
        rows.append(_row("Less: deduction u/s 57(iia) on family pension",
                         -ded57_reported,
                         f"lower of 1/3rd of family pension or {_inr(cap)}"))
    return rows


def _cg_details(cg):
    """One annexure line per transaction."""
    lines = []
    for row in (cg.get("rows") or {}).get("equity_lt") or []:
        lines.append(_row(
            f"{row.get('scrip', 'Scrip')} · LTCG 112A",
            row.get("gain"),
            f"sold {row.get('sale_date', '')} at {_inr(row.get('sale'))}; cost base "
            f"{_inr(row.get('cost'))}{'; grandfathered (FMV 31-01-2018 applied)' if row.get('note') else ''}"))
    for row in (cg.get("rows") or {}).get("equity_st") or []:
        lines.append(_row(
            f"{row.get('scrip', 'Scrip')} · STCG {'111A' if row.get('stt') else 'at slab rates'}",
            row.get("gain"), f"{row.get('buy_date', '')} → {row.get('sale_date', '')}"))
    for row in (cg.get("rows") or {}).get("land_building") or []:
        lines.append(_row(
            f"{row.get('description', 'Land / building')} · LTCG",
            row.get("gain"),
            (row.get("note") or "") + f" · sale {row.get('sale_date', '')}"))
    for kind in ("other_st", "other_lt"):
        for row in (cg.get("rows") or {}).get(kind) or []:
            lines.append(_row(
                f"{row.get('description', 'Other asset')} · "
                f"{'STCG' if kind == 'other_st' else 'LTCG'}",
                row.get("gain"), f"{row.get('buy_date', '')} → {row.get('sale_date', '')}"))
    return lines


def _cg_section(cg):
    return {
        "rows": cg.get("worksheet") or [],
        "details": _cg_details(cg),
        "quarterly": [
            _row(f"{_CG_KIND.get(k, k)} — {_CG_SPAN[q]}", amt)
            for k, buckets in (cg.get("quarterly") or {}).items()
            for q, amt in sorted(buckets.items()) if amt],
        "total": cg["total_cg"],
        "total_incl_exempt112a": cg["total_cg_incl_exempt112a"],
        "note": ("112A exemption of Rs. 1,25,000 against long-term equity gains is "
                 "built in; carry-forward CG losses: ST Rs. "
                 f"{int(cg.get('cg_loss_carry_st', 0) or 0):,} / LT Rs. "
                 f"{int(cg.get('cg_loss_carry_lt', 0) or 0):,}"),
    }


def _business_section(pres):
    rows = []
    for r in pres.get("44ad_rows") or []:
        rows.append(_row(
            f"44AD — {r.get('name', '')} ({r.get('code', '')})", r.get("income"),
            f"turnover {_inr(r.get('gross_turnover'))} · presumptive "
            f"{_inr(r.get('presumptive_income'))} (6% digital / 8% cash of "
            f"{_inr(r.get('turnover_bank'))} + {_inr(r.get('turnover_cash'))} cash)"))
    for r in pres.get("44ada_rows") or []:
        rows.append(_row(
            f"44ADA — {r.get('name', '')} ({r.get('code', '')})", r.get("income"),
            f"receipts {_inr(r.get('gross_receipts'))} · 50% presumptive = "
            f"{_inr(r.get('presumptive_income'))}"))
    for r in pres.get("44ae_rows") or []:
        rows.append(_row(
            f"44AE — vehicle {r.get('reg_no', '')}", r.get("income"),
            f"{r.get('months', 0)} month(s) · {'heavy goods vehicle' if r.get('heavy') else 'other vehicle'}"))
    rows.append(_row("Income from business / profession", pres.get("total_income"),
                     "higher of presumptive income vs declared profit is applied", 1))
    fin = pres.get("financials") or {}
    fin_map = (("bank_balance", "Balance with banks"), ("cash_in_hand", "Cash in hand"),
               ("sundry_debtors", "Sundry debtors"), ("inventories", "Inventories"),
               ("fixed_assets", "Fixed assets"), ("total_assets", "Total assets"),
               ("sundry_creditors", "Sundry creditors"),
               ("partner_capital", "Partners' / proprietor's capital"),
               ("total_capital_liabilities", "Total capital & liabilities"))
    fin_rows = [_row(label, int(_num(fin.get(k)))) for k, label in fin_map
                if _num(fin.get(k))]
    return {"rows": rows, "income": pres.get("total_income"),
            "financials_rows": fin_rows}


def _salary_workings(exp):
    hra = exp.get("hra") or {}
    hra_rows = []
    if hra:
        for k, label in (("hra_received", "HRA actually received"),
                         ("basic_salary", "Basic salary (period)"),
                         ("da", "DA forming part of retirement benefits"),
                         ("rent_paid", "Rent paid (period)"),
                         ("excess_rent", "Rent paid − 10% of salary"),
                         ("fifty_40", "50%/40% of salary (metro / non-metro)"),
                         ("exempt", "HRA exempt u/s 10(13A) — least of the three")):
            if _num(hra.get(k)):
                hra_rows.append(_row(label, int(_num(hra.get(k)))))
    return {
        "components": exp.get("component_breakdown") or [],
        "perquisites": exp.get("perquisites_rows") or [],
        "profits": exp.get("profits_rows") or [],
        "terminal": exp.get("terminal_workings") or [],
        "hra_rows": hra_rows,
        "arrears_total": exp.get("arrears_total") or 0,
        "monthly_totals": exp.get("monthly_totals") or {},
    }


def _regime_computation_generic(data: dict, opt_out: bool, form: str) -> dict:
    from .forms import build_for_form

    cp = copy.deepcopy(data or {})
    cp.setdefault("filing", {})["opt_out_new_regime"] = opt_out
    built = build_for_form(cp, form)
    doc, rep = built["json"], built["report"]
    node = (doc.get("ITR") or {}).get(form) or {}
    tax = rep["tax"]
    working = tax.get("Working", {})
    personal = cp.get("personal", {}) or {}
    dob = personal.get("dob") or ""

    age_note = ""
    if T.is_super_senior_citizen(dob):
        age_note = "Super senior citizen (80+) slab rates applied."
    elif T.is_senior_citizen(dob):
        age_note = "Senior citizen (60+) slab rates applied."

    # ------------------------------------------------ salary (per-regime, from the built schedule)
    exp = rep.get("salary_expansion") or {}
    rich = bool(exp and (exp.get("component_breakdown") or exp.get("terminal_workings")
                         or exp.get("perquisites_rows") or exp.get("profits_rows")))
    salary_workings = _salary_workings(exp) if rich else None
    if form == "ITR2":
        sp = node.get("ScheduleS") or {}
        s_rows = sp.get("Salaries") or []
        s17_1 = sum(_num((r.get("Salarys") or {}).get("Salary")) for r in s_rows)
        s17_2 = sum(_num((r.get("Salarys") or {}).get("ValueOfPerquisites")) for r in s_rows)
        s17_3 = sum(_num((r.get("Salarys") or {}).get("ProfitsinLieuOfSalary")) for r in s_rows)
        ex10 = _num(sp.get("AllwncExtentExemptUs10"))
        ex_dtls = (sp.get("AllwncExemptUs10") or {}).get("AllwncExemptUs10Dtls") or []
        std16ia = sp.get("DeductionUnderSection16ia") or 0
        ea16ii = sp.get("EntertainmntalwncUs16ii") or 0
        pt16iii = sp.get("ProfessionalTaxUs16iii") or 0
        net_sal = sp.get("NetSalary")
    else:  # ITR-4
        ided = node.get("IncomeDeductions") or {}
        s17_1 = _num(ided.get("Salary"))
        s17_2 = _num(ided.get("PerquisitesValue"))
        s17_3 = _num(ided.get("ProfitsInSalary"))
        ex10 = _num((ided.get("AllwncExemptUs10") or {}).get("TotalAllwncExemptUs10"))
        ex_dtls = (ided.get("AllwncExemptUs10") or {}).get("AllwncExemptUs10Dtls") or []
        std16ia = ided.get("DeductionUs16ia") or 0
        ea16ii = ided.get("EntertainmntalwncUs16ii") or 0
        pt16iii = ided.get("ProfessionalTaxUs16iii") or 0
        net_sal = ided.get("NetSalary")
    if _num(rep.get("salary_income")) or s17_1:
        salary_rows = [
            _row("Salary as per section 17(1)", int(s17_1),
                 "pay, allowances & arrears as per the salary-slip break-up"),
            _row("Value of perquisites u/s 17(2)", int(s17_2)),
            _row("Profits in lieu of salary u/s 17(3)", int(s17_3),
                 "taxable parts of terminal benefits where reported there"),
            _row("Less: allowances exempt u/s 10 (incl. HRA)", -int(ex10),
                 "only exemptions permitted under the chosen regime are applied"),
            _row("Net salary", net_sal, level=1),
            _row("Less: standard deduction u/s 16(ia)", -std16ia,
                 f"maximum {_inr(C.STD_DEDUCTION_16IA_OLD if opt_out else C.STD_DEDUCTION_16IA_NEW)}"),
            _row("Less: entertainment allowance u/s 16(ii)", -ea16ii),
            _row("Less: professional tax u/s 16(iii)", -pt16iii),
        ]
        exempt_rows = [{"label": f"    u/s {r.get('SalNatureDesc')}",
                        "amount": r.get("SalOthAmount")} for r in ex_dtls]
        salary_income = rep["salary_income"]
    else:
        salary_rows, exempt_rows, salary_income = [], [], rep["salary_income"]

    # ------------------------------------------------ house property
    if form == "ITR2":
        hp_sections = _hp_sections_from_doc(
            (node.get("ScheduleHP") or {}).get("PropertyDetails"))
    else:
        hp_sections = ([{"title": "Income from house property",
                         "rows": [_row("Net income from house property "
                                       "(self-occupied / let-out computation)",
                                       rep["hp_income"], "as per ITR-4 head computation")],
                         "income": rep["hp_income"]}]
                       if _num(rep.get("hp_income")) else [])

    # ------------------------------------------------ other sources
    ded57 = 0
    if form == "ITR4":
        os_node = node.get("IncomeDeductions") or {}
        # DeductionUs57iia may live in the OS schedule; use reported value if any
        ded57 = int(_num(os_node.get("DeductionUs57iia")))
    os_rows = _os_rows_generic(cp, opt_out, ded57)

    # ------------------------------------------------ capital gains (ITR-2)
    cg_sec = None
    cg = rep.get("capital_gains")
    if cg and cg.get("has_any_cg"):
        cg_sec = _cg_section(cg)

    # ------------------------------------------------ presumptive business (ITR-4)
    biz_sec = None
    if rep.get("presumptive"):
        biz_sec = _business_section(rep["presumptive"])

    # ------------------------------------------------ Chapter VI-A
    if form == "ITR4":
        inc_ded = node.get("IncomeDeductions") or {}
        usr_via = inc_ded.get("UsrDeductUndChapVIA") or {}
        ded_via = inc_ded.get("DeductUndChapVIA") or {}
    else:
        via_node = node.get("ScheduleVIA") or {}
        usr_via = via_node.get("UsrDeductUndChapVIA") or {}
        ded_via = via_node.get("DeductUndChapVIA") or {}
    via_rows = _via_rows_generic(usr_via, ded_via, opt_out)
    if not opt_out:
        kept = [r for r in via_rows if r["eligible"] > 0]
        via_caption = ("Under the new regime only 80CCD(2) and 80CCH are admissible; "
                       "the rest are shown at NIL." if len(via_rows) > len(kept) else
                       "Under the new regime only 80CCD(2) and 80CCH are admissible.")
    else:
        via_caption = ""

    # ------------------------------------------------ totals
    if form == "ITR4":
        gti_excl = rep.get("gross_total_income_excl_ltcg", 0)
        gti_incl = rep.get("gross_total_income_incl_ltcg", 0)
        ltcg_taxable = int(_num(rep.get("ltcg112a")))
        gti_rows = [
            _row("Gross total income (salary + house property + business + other sources)",
                 gti_excl),
            _row("LTCG u/s 112A (within the Rs. 1,25,000 exemption)", ltcg_taxable,
                 "capped at Rs. 1,25,000 for ITR-4; fully exempt, no tax"),
            _row("Gross total income", gti_incl, level=1),
        ]
        ltcg_note = ""
    else:
        gti_excl = rep.get("gross_total_income_excl_cg", 0)
        gti_incl = rep.get("gross_total_income_incl_cg", 0)
        cg_total_gross = int(_num((rep.get("capital_gains") or {})
                                  .get("total_cg_incl_exempt112a")))
        gti_rows = [
            _row("Gross total income (before capital gains)", gti_excl),
            _row("Net capital gains (after exemptions u/s 54/54F/54EC)", cg_total_gross,
                 "including the exempt first Rs. 1,25,000 of 112A gains"),
            _row("Gross total income", gti_incl, level=1),
        ]
        ltcg_taxable = int(_num((rep.get("capital_gains") or {}).get("ltcg_112a_taxable")))
        ltcg_note = ""
    via_total = int(_num((rep.get("deductions_eligible") or {})
                         .get("TotalChapVIADeductions")))
    ti = rep["total_income"]

    # ------------------------------------------------ tax
    slab_base = _num(working.get("income_excl_ltcg", ti))
    slab_rows = T.slab_breakup(max(0, slab_base), opt_out, dob or "1990-01-01")
    surcharge = tax.get("SurchargeComputed", 0)
    sur_note = ("10% surcharge with marginal relief as applicable."
                if surcharge > 0 else "No surcharge (total income within Rs. 50 lakh).")
    special_rows = []
    interest_rows = [
        ("IntrstPayUs234A", "Interest u/s 234A (delay in filing)"),
        ("IntrstPayUs234B", "Interest u/s 234B (default in advance tax)"),
        ("IntrstPayUs234C", "Interest u/s 234C (deferment of instalments)"),
        ("LateFilingFee234F", "Fee u/s 234F (late filing)"),
        ("FeeFurnish234I", "Fee u/s 234-I (revised return after 31-12-2026)"),
    ]
    interest = []
    detail = tax.get("InterestDetail", {})
    for key, label in interest_rows:
        d = detail.get(key, {})
        interest.append({"label": label, "amount": tax["IntrstPay"].get(key, 0),
                         "note": d.get("note", "")})
        if key == "IntrstPayUs234C" and d.get("instalments"):
            interest[-1]["instalments"] = d["instalments"]

    paid = tax["TaxesPaid"]
    result_note = ""
    if tax["BalTaxPayable"] > 0:
        result_note = (f"Balance tax payable: {_inr(tax['BalTaxPayable'])} — pay as "
                       "self-assessment tax before filing to avoid further interest.")
    elif tax["RefundDue"] > 0:
        result_note = (f"Refund due: {_inr(tax['RefundDue'])} — ensure the refund "
                       "bank account is pre-validated.")

    special_tax = int(_num(working.get("tax_111a")) + _num(working.get("tax_ltcg_125"))
                      + _num(working.get("tax_ltcg_idx20")))
    return {
        "regime": "old" if opt_out else "new",
        "regime_title": ("Old tax regime (with Chapter VI-A deductions & exemptions)"
                         if opt_out else "New tax regime (default, u/s 115BAC)"),
        "age_note": age_note,
        "salary": {"rows": salary_rows, "exempt_rows": exempt_rows,
                   "income": salary_income},
        "salary_workings": salary_workings,
        "business": biz_sec,
        "house_properties": hp_sections,
        "hp_note": "",
        "hp_income": rep["hp_income"],
        "other_sources": {"rows": os_rows, "income": rep["other_sources_income"]},
        "capital_gains": cg_sec,
        "gti_excl_ltcg": gti_excl,
        "gti_rows": gti_rows,
        "ltcg": ltcg_taxable,
        "ltcg_note": ltcg_note,
        "gti_incl_ltcg": gti_incl,
        "via_rows": via_rows,
        "via_caption": via_caption,
        "via_total": via_total,
        "total_income": ti,
        "total_income_note": "rounded off to the nearest multiple of ten u/s 288A",
        "form10e_rows": (rep.get("form10e") or {}).get("rows") or [],
        "form10e_meta": {k: v for k, v in (rep.get("form10e") or {}).items()
                         if k != "rows"},
        "slab_rows": slab_rows,
        "slab_tax": working.get("slab_tax", 0),
        "special_tax": special_tax,
        "special_tax_note": ("111A @ 20% · non-112A LTCG @ 12.5%/20% — see the "
                             "capital-gains worksheet"),
        "ltcg_tax": working.get("ltcg_tax", 0),
        "surcharge": surcharge,
        "surcharge_note": sur_note,
        "tax_payable": tax["TotalTaxPayable"],
        "rebate": tax["Rebate87A"],
        "rebate_note": working.get("rebate_note", ""),
        "tax_after_rebate": tax["TaxPayableOnRebate"],
        "cess": tax["EducationCess"],
        "cess_note": "Health & Education Cess @ 4% on tax after rebate",
        "gross_tax_liability": tax["GrossTaxLiability"],
        "relief89": tax["Section89"],
        "relief89_note": ("relief on salary arrears — Form 10E working below"
                          if tax["Section89"] else ""),
        "net_tax_liability": tax["NetTaxLiability"],
        "interest": interest,
        "total_interest": tax["TotalIntrstPay"],
        "tot_tax_plus_interest": tax["TotTaxPlusIntrstPay"],
        "taxes_paid": {
            "tds": paid["TDS"], "tcs": paid["TCS"],
            "advance_tax": paid["AdvanceTax"],
            "self_assessment_tax": paid["SelfAssessmentTax"],
            "total": paid["TotalTaxesPaid"],
        },
        "balance_payable": tax["BalTaxPayable"],
        "refund_due": tax["RefundDue"],
        "result_note": result_note,
        "warnings": rep.get("warnings", []),
    }


def build_computation(data: dict, form: str | None = None) -> dict:
    """Full computation report for both regimes plus the recommendation.

    ``form`` picks the target form (ITR1 default; ITR2/ITR4 produce the same
    renderer contract enriched with the multi-form annexures)."""
    form = (form or "ITR1").upper()
    per = ((lambda d, opt_out: _regime_computation_generic(d, opt_out, form))
           if form in ("ITR2", "ITR3", "ITR4", "ITR-2", "ITR-4") else
           _regime_computation)
    new_c = per(data, opt_out=False)
    old_c = per(data, opt_out=True)

    if new_c["tot_tax_plus_interest"] < old_c["tot_tax_plus_interest"]:
        recommended, saving = "new", old_c["tot_tax_plus_interest"] - new_c["tot_tax_plus_interest"]
    elif old_c["tot_tax_plus_interest"] < new_c["tot_tax_plus_interest"]:
        recommended, saving = "old", new_c["tot_tax_plus_interest"] - old_c["tot_tax_plus_interest"]
    else:
        recommended, saving = "new", 0

    personal = (data or {}).get("personal", {}) or {}
    filing = (data or {}).get("filing", {}) or {}
    first = personal.get("first_name", "")
    last = personal.get("last_name", "")
    form_label = {"ITR1": "ITR-1 (SAHAJ)", "ITR2": "ITR-2",
                  "ITR4": "ITR-4 (SUGAM)"}.get(form.replace("-", ""), form)
    return {
        "meta": {
            "title": "Computation of Total Income and Tax Liability",
            "form": form_label,
            "assessment_year": ASSESSMENT_YEAR,
            "previous_year": PREVIOUS_YEAR,
            "name": f"{first} {personal.get('middle_name', '')} {last}".replace("  ", " ").strip(),
            "pan": (personal.get("pan") or "").upper(),
            "dob": personal.get("dob", ""),
            "generated_on": str(_dt.date.today()),
            "filing_date": filing.get("filing_date", ""),
            "return_section": C.RETURN_SECTIONS.get(int(filing.get("return_section", 11) or 11), ""),
            "chosen_regime": "old" if filing.get("opt_out_new_regime") else "new",
            "software": f"Open ITR Suite v{__version__}",
        },
        "new": new_c,
        "old": old_c,
        "comparison": {
            "new_total": new_c["tot_tax_plus_interest"],
            "old_total": old_c["tot_tax_plus_interest"],
            "recommended": recommended,
            "saving": saving,
        },
    }
