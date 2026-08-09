"""Build an official ITR-1 e-filing JSON (AY 2026-27, schema V1.1) from a
friendly taxpayer data dictionary.

The builder performs *all* derived arithmetic itself (gross totals, eligible
deductions restricted by the applicable caps, tax computation, refund /
balance payable, schedule totals, ...) so that the generated JSON is
internally consistent with the CBDT validation rules by construction.

Input model (``data``) - see ``samples/sample_taxpayer.json`` for a fully
worked example.  Every monetary key is optional and defaults to zero.
"""

from __future__ import annotations

import copy
import datetime as _dt
import re

from . import (SW_CREATED_BY, UTILITY_SW_VERSION_BY_FORM,
               UTILITY_FORM_DESCRIPTION)
from . import constants as C
from . import tax_engine as T
from . import salaryx as SX
from .errors import DataError, field_scope, parse_float, parse_int, reset_scopes

TODAY = _dt.date.today

# Friendly names for the most common taxpayer-facing keys, used to build
# precise error messages ("TDS deducted (Taxes paid › TDS on salary row 2)")
# instead of raw Python tracebacks.
KEY_LABELS = {
    "amount": "Amount", "tds": "TDS deducted", "tds_claimed": "TDS claimed",
    "tds_deducted": "TDS deducted", "tcs_claimed": "TCS claimed",
    "tcs_collected": "TCS collected", "amount_subject": "Amount on which TCS collected",
    "amount_paid": "Amount paid/credited", "salary_income": "Salary income chargeable",
    "challan_no": "Challan serial no.", "gross_receipt": "Gross receipt",
    "salary_17_1": "Salary as per section 17(1)",
    "perquisites_17_2": "Value of perquisites u/s 17(2)",
    "profits_17_3": "Profits in lieu of salary u/s 17(3)",
    "hra_received": "HRA received", "rent_paid": "Rent paid",
    "basic_salary": "Basic salary", "da": "Dearness allowance",
    "entertainment_allowance": "Entertainment allowance u/s 16(ii)",
    "professional_tax": "Professional tax u/s 16(iii)",
    "standard_deduction_override": "Standard deduction override",
    "annual_lettable_value": "Annual lettable value / rent received",
    "rent_not_realized": "Unrealized rent", "municipal_taxes": "Municipal taxes paid",
    "interest_on_borrowed_capital": "Interest on borrowed capital u/s 24(b)",
    "arrears_unrealized_rent_received": "Arrears/unrealized rent received",
    "interest": "Interest", "loan_amount": "Loan amount",
    "outstanding": "Outstanding loan amount", "premium": "Premium",
    "medical_exp": "Medical expenditure", "preventive": "Preventive health check-up",
    "share_pct": "Percentage share in property",
    "assessee_share_pct": "Your share in the property (%)",
    "sale_consideration": "Sale consideration (LTCG 112A)",
    "cost_of_acquisition": "Cost of acquisition (LTCG 112A)",
    "cash": "Donation paid in cash", "other": "Donation other than cash",
    "relief_89": "Relief u/s 89", "return_section": "Return filed under section",
    "pin_code": "PIN code",
    "IntrstPayUs234A": "Interest u/s 234A override",
    "IntrstPayUs234B": "Interest u/s 234B override",
    "IntrstPayUs234C": "Interest u/s 234C override",
    "LateFilingFee234F": "Late filing fee u/s 234F override",
    "FeeFurnish234I": "Fee u/s 234-I override",
}


def _g(mapping, key, default=0, label=None):
    """Get ``key`` from ``mapping`` as a whole-rupee int.

    Blank/None → ``default``; comma-grouped or decimal text is accepted;
    anything else raises :class:`~itr1.errors.DataError` naming the field.
    """
    if mapping is None:
        return default
    if key not in mapping:
        return default
    return parse_int(mapping.get(key), default=default,
                     key_label=label or KEY_LABELS.get(key, str(key)),
                     key_path=str(key))


def _s(mapping, key, default=""):
    if mapping is None:
        return default
    v = mapping.get(key, default)
    return default if v is None else str(v)


def _num(mapping, key, default=0.0, label=None):
    if mapping is None:
        return default
    if key not in mapping:
        return default
    return parse_float(mapping.get(key), default=default,
                       key_label=label or KEY_LABELS.get(key, str(key)),
                       key_path=str(key))


# ---------------------------------------------------------------------------
# Salary
# ---------------------------------------------------------------------------

def _build_hra(hra: dict | None) -> tuple[dict | None, int]:
    """Return (ScheduleEA10_13A block, eligible HRA exemption)."""
    if not hra:
        return None, 0
    with field_scope("HRA exemption u/s 10(13A)", "hra"):
        hra_received = _g(hra, "hra_received")
        rent_paid = _g(hra, "rent_paid")
        basic = _g(hra, "basic_salary")
        da = _g(hra, "da")
    if max(hra_received, rent_paid) <= 0:
        return None, 0
    salary_171 = basic + da
    is_metro = str(hra.get("place", "1")) in ("1", "metro", "True", "true")
    block = {
        "Placeofwork": "1" if is_metro else "2",
        "ActlHRARecv": hra_received,
        "ActlRentPaid": rent_paid,
        "DtlsSalUsSec171": salary_171,
        "BasicSalary": basic,
        "ActlRentPaid10Per": max(0, rent_paid - salary_171 // 10),
        "Sal40Or50Per": (salary_171 // 2) if is_metro else (salary_171 * 2) // 5,
    }
    if da:
        block["DearnessAllwnc"] = da
    # rule 263: least of three
    eligible = min(block["ActlHRARecv"], block["ActlRentPaid10Per"], block["Sal40Or50Per"])
    block["EligbleExmpAllwncUs13A"] = max(0, eligible)
    return block, block["EligbleExmpAllwncUs13A"]


def _build_salary(data: dict, opt_out: bool, warnings: list) -> dict:
    sal = data.get("salary", {}) or {}
    # expand slip components / perquisites / terminal benefits / monthly grid /
    # arrears into the classic totals (keeps every downstream rule working)
    exp = SX.expand_salary(data) if SX.has_rich_salary(sal) else None
    if exp is not None:
        warnings.extend(exp["warnings"])
        data["_salary_expansion"] = exp
    with field_scope("Salary income", "salary"):
        if exp is not None:
            s171, s172, s173 = exp["salary_17_1"], exp["perquisites_17_2"], exp["profits_17_3"]
        else:
            s171 = _g(sal, "salary_17_1")
            s172 = _g(sal, "perquisites_17_2")
            s173 = _g(sal, "profits_17_3")
        gross = s171 + s172 + s173                       # rule 59

        hra_input = (exp["hra"] if exp is not None else None) or sal.get("hra")
        hra_block, hra_eligible = _build_hra(hra_input)
        exempt_rows = []
        total_exempt = 0
        if hra_block and not opt_out:
            hra_block, hra_eligible = None, 0            # rule 165: no HRA in new regime
            warnings.append("HRA exemption is not available under the new tax regime; removed.")
        new_regime_banned = {"10(5)", "10(13A)", "10(14)(i)", "10(14)(ii)", "10(17)", "EIC"}
        exempt_source = ([{"section": r["section"], "amount": r["amount"], "_note": r.get("note", "")}
                          for r in exp["exempt_rows"]] if exp is not None
                         else (sal.get("exempt_allowances", []) or []))
        for ridx, row in enumerate(exempt_source):
            with field_scope(f"Exempt allowance row {ridx + 1}",
                             f"exempt_allowances.{ridx}"):
                sec = _s(row, "section")
                amt = _g(row, "amount")
            if amt <= 0:
                continue
            if opt_out and sec in ("10(14)(i)(115BAC)", "10(14)(ii)(115BAC)"):
                # rule 150: the (115BAC) variants exist only in the new regime
                warnings.append(f"Allowance {sec} is allowed only in the new regime; removed.")
                continue
            if not opt_out and sec in new_regime_banned:
                # rules 149/161/164-167/301: not exempt under the new regime
                warnings.append(f"Allowance u/s {sec} is not exempt under the new tax regime; "
                                f"removed (rules 149/161/164-167).")
                continue
            if not opt_out and sec == "10(14)(ii)(115BAC)" and amt > 38_400:
                amt = 38_400     # rule 148
                warnings.append("Allowance u/s 10(14)(ii)(115BAC) capped at Rs. 38,400 (rule 148).")
            exempt_rows.append({"SalNatureDesc": sec, "SalOthAmount": amt})
            total_exempt += amt
        if hra_block:
            exempt_rows.append({"SalNatureDesc": "10(13A)", "SalOthAmount": hra_eligible})
            total_exempt += hra_eligible

        net = max(0, gross - total_exempt)               # rule 60

        # standard deduction u/s 16(ia)
        override = (exp["standard_deduction_override"] if exp is not None
                    else _g(sal, "standard_deduction_override", None))
        cap = C.STD_DEDUCTION_16IA_OLD if opt_out else C.STD_DEDUCTION_16IA_NEW
        std = min(cap, net) if override is None else min(cap, net, override)

        ent_claimed = (exp["entertainment_allowance"] if exp is not None
                       else _g(sal, "entertainment_allowance"))
        prof_tax = (exp["professional_tax"] if exp is not None
                    else _g(sal, "professional_tax"))
        employer_cat = _s(data.get("personal", {}), "employer_category", "NA")
        ent = 0
        if opt_out and ent_claimed > 0:
            if employer_cat in ("CGOV", "SGOV", "PSU"):
                ent = min(ent_claimed, C.ENTERTAINMENT_ALLOW_16II_MAX, s171 // 5)  # rule 57
            else:
                warnings.append("Entertainment allowance u/s 16(ii) is allowed only to "
                                "Government/PSU employees; removed (rule 58).")
        if opt_out:
            prof = min(prof_tax, C.PROFESSIONAL_TAX_16III_MAX)
        else:
            prof = 0                                     # rules 163/168

        ded16 = std + ent + prof                         # rule 61
        income_from_sal = max(0, net - ded16)            # rule 62
        return {
            "GrossSalary": gross,
            "Salary": s171,
            "PerquisitesValue": s172,
            "ProfitsInSalary": s173,
            "AllwncExemptUs10": ({"AllwncExemptUs10Dtls": exempt_rows,
                                  "TotalAllwncExemptUs10": total_exempt}
                                 if exempt_rows else {"TotalAllwncExemptUs10": 0}),
            "NetSalary": net,
            "DeductionUs16": ded16,
            "DeductionUs16ia": std,
            "EntertainmentAlw16ii": ent,
            "ProfessionalTaxUs16iii": prof,
            "IncomeFromSal": income_from_sal,
        }, hra_block


# ---------------------------------------------------------------------------
# House property
# ---------------------------------------------------------------------------

def _build_house_property(props: list[dict], opt_out: bool, warnings: list) -> tuple[list, int]:
    details = []
    total_hp = 0
    for idx, p in enumerate((props or [])[: C.HP_MAX_PROPERTIES], start=1):
        with field_scope(f"House property {idx}", f"house_properties.{idx - 1}"):
            occupancy = _s(p, "occupancy", "L")
            co_owned = bool(p.get("co_owned"))
            share = _num(p, "assessee_share_pct", 100.0)
            if not co_owned:
                share = 100.0                            # rule 334
            own = share / 100.0

            alv = _g(p, "annual_lettable_value")
            unrealized = _g(p, "rent_not_realized")
            muni_tax = _g(p, "municipal_taxes")
            if occupancy == "S":
                alv = unrealized = muni_tax = 0          # rules 45/49
            if unrealized > alv:
                unrealized = alv                         # rule 336
            total_unrl_tax = unrealized + muni_tax       # rule 46 (i - ii split)
            balance = max(0, alv - total_unrl_tax)
            annual_owned = round(balance * own)
            thirty_pct = round(annual_owned * C.HP_STD_DEDUCTION_RATE)  # rule 43

            interest_claimed = _g(p, "interest_on_borrowed_capital")
            loans = p.get("loans") or []
            loan_rows = []
            loan_interest_total = 0
            for ln in loans:
                interest = _g(ln, "interest")
                loan_interest_total += interest
                loan_rows.append({
                    "LoanTknFrom": _s(ln, "lender_type", "B"),
                    "BankOrInstnName": _s(ln, "lender_name"),
                    "LoanAccNoOfBankOrInstnRefNo": _s(ln, "loan_account_no"),
                    "DateofLoan": _s(ln, "loan_date"),
                    "TotalLoanAmt": _g(ln, "loan_amount"),
                    "LoanOutstndngAmt": _g(ln, "outstanding"),
                    "InterestUs24B": interest,
                })
            if loan_rows:
                interest_claimed = loan_interest_total   # rule 240
            interest_own = round(interest_claimed * own)

            if occupancy == "S":
                if opt_out:
                    interest_own = min(interest_own, C.HP_SELF_OCCUPIED_INTEREST_CAP_OLD)  # rule 48
                else:
                    if interest_own:
                        warnings.append("Interest u/s 24(b) on self-occupied property is not "
                                        "allowed in the new regime; removed (rule 162).")
                    interest_own = 0                      # rule 162 / 253
                    if loan_rows:
                        warnings.append("Schedule 24(b) details for self-occupied property are "
                                        "not admissible in the new regime; removed (rule 253).")
                        loan_rows = []

            total_deduct = thirty_pct + interest_own
            arrears_received = _g(p, "arrears_unrealized_rent_received")
            arrears_taxable = round(arrears_received * C.HP_ARREARS_TAXABLE_FRACTION)

            income_hp = annual_owned - total_deduct + arrears_taxable  # rule 47

            addr = {
                "AddrDetail": _s(p, "address_line"),
                "CityOrTownOrDistrict": _s(p, "city"),
                "StateCode": _s(p, "state_code"),
                "CountryCode": _s(p, "country_code", "91"),
                "PinCode": _g(p, "pin_code"),
            }
            block = {
                "HPSNo": idx,
                "AddressDetailWithZipCode": addr,
                "PropertyOwner": _s(p, "owner", "SE"),
                "PropCoOwnedFlg": "YES" if co_owned else "NO",
                "ifLetOut": occupancy,
                "Rentdetails": {
                    "AnnualLetableValue": alv,
                    "RentNotRealized": unrealized,
                    "LocalTaxes": muni_tax,
                    "TotalUnrealizedAndTax": total_unrl_tax,
                    "BalanceALV": balance,
                    "AnnualOfPropOwned": annual_owned,
                    "ThirtyPercentOfBalance": thirty_pct,
                    "IntOnBorwCap": interest_own,
                    "TotalDeduct": total_deduct,
                    "ArrearsUnrealizedRentRcvd": arrears_taxable,
                    "IncomeOfHP": income_hp,
                },
            }
            if p.get("owner") == "OT" and _s(p, "owner_other"):
                block["PropertyOwnerOther"] = _s(p, "owner_other")
            if co_owned:
                block["AsseseeShareProperty"] = share
                co_rows = []
                for cidx, co in enumerate(p.get("co_owners") or [], start=1):
                    row = {"CoOwnersSNo": cidx, "NameCoOwner": _s(co, "name")}
                    if _s(co, "pan"):
                        row["PAN_CoOwner"] = _s(co, "pan")
                    if _s(co, "aadhaar"):
                        row["Aadhaar_CoOwner"] = _s(co, "aadhaar")
                    if "share_pct" in co:
                        row["PercentShareProperty"] = _num(co, "share_pct", 0.0)
                    co_rows.append(row)
                if co_rows:
                    block["CoOwners"] = co_rows
            tenant_rows = []
            for tidx, tn in enumerate(p.get("tenants") or [], start=1):
                row = {"TenantSNo": tidx, "NameofTenant": _s(tn, "name")}
                if _s(tn, "pan"):
                    row["PANofTenant"] = _s(tn, "pan")
                if _s(tn, "aadhaar"):
                    row["AadhaarofTenant"] = _s(tn, "aadhaar")
                if _s(tn, "pan_tan"):
                    row["PANTANofTenant"] = _s(tn, "pan_tan")
                tenant_rows.append(row)
            if tenant_rows:
                block["TenantDetails"] = tenant_rows
            if loan_rows:
                block["Rentdetails"]["Section24B"] = {
                    "Section24BDtls": loan_rows,
                    "TotalInterestUs24B": loan_interest_total,
                }
            details.append(block)
            total_hp += income_hp

    # New regime: house-property loss cannot be set off (rules 160/174)
    chargeable = total_hp if opt_out else max(0, total_hp)
    if not opt_out and total_hp < 0:
        warnings.append("House property loss cannot be set off under the new regime "
                        "(rule 160); it is ignored in Gross Total Income.")
    return details, chargeable


# ---------------------------------------------------------------------------
# Other sources
# ---------------------------------------------------------------------------

def _build_other_sources(rows: list[dict], opt_out: bool, warnings: list) -> tuple[dict, int, int]:
    items = []
    total = 0
    family_pension = 0
    seen = set()
    for ridx, r in enumerate(rows or []):
        with field_scope(f"Other income row {ridx + 1}", f"other_sources.{ridx}"):
            nature = _s(r, "nature", "OTH")
            amt = _g(r, "amount")
        if nature in ("SAV", "IFD", "TAX", "FAP") and nature in seen:
            warnings.append(f"'{C.OTHER_SOURCE_NATURES.get(nature, nature)}' selected more "
                            f"than once under Income from Other Sources (rules 50/51/55/56).")
        seen.add(nature)
        if amt <= 0:
            continue
        row = {"OthSrcNatureDesc": nature, "OthSrcOthAmount": amt}
        desc = _s(r, "description")
        if nature == "OTH" and desc:
            row["OthSrcOthNatOfInc"] = desc
        if nature == "DIV":
            q = r.get("dividend_quarters") or {}
            with field_scope("Dividend quarter-wise breakup",
                             f"other_sources.{ridx}.dividend_quarters"):
                row["DividendInc"] = {"DateRange": {
                    "Upto15Of6": _g(q, "upto_15_06"),
                    "Upto15Of9": _g(q, "upto_15_09"),
                    "Up16Of9To15Of12": _g(q, "16_09_15_12"),
                    "Up16Of12To15Of3": _g(q, "16_12_15_03"),
                    "Up16Of3To31Of3": _g(q, "16_03_31_03"),
                }}
        items.append(row)
        total += amt
        if nature == "FAP":
            family_pension += amt

    ded_57iia = 0
    if family_pension > 0:
        cap = C.FAMILY_PENSION_DED_CAP_OLD if opt_out else C.FAMILY_PENSION_DED_CAP_NEW
        ded_57iia = min(family_pension // 3, cap)    # rules 54/214
    others_block = {"OthersIncDtlsOthSrc": items} if items else {}
    return others_block, total - ded_57iia, ded_57iia


# ---------------------------------------------------------------------------
# Exempt income u/s 10 (non-salary)
# ---------------------------------------------------------------------------

def _build_exempt_income(rows: list[dict], warnings: list) -> dict:
    items = []
    total = 0
    seen = set()
    for ridx, r in enumerate(rows or []):
        with field_scope(f"Exempt income row {ridx + 1}", f"exempt_income.{ridx}"):
            amt = _g(r, "amount")
        sub = _s(r, "sub_category")
        if amt <= 0:
            continue
        if sub in seen:
            warnings.append(f"Exempt income nature '{sub}' selected more than once (rule 184).")
        seen.add(sub)
        if sub == "10(1)" and amt > C.MAX_AGRICULTURE_INCOME:
            warnings.append("Agricultural income exempt in ITR-1 cannot exceed Rs. 5,000 (rule 29).")
        row = {"Category": _s(r, "category", "OTH"), "OthAmount": amt}
        if sub:
            row["SubCategory"] = sub
        if _s(r, "description"):
            row["Description"] = _s(r, "description")
        items.append(row)
        total += amt
    block = {"ExemptIncAgriOthUs10Total": total}
    if items:
        block["ExemptIncAgriOthUs10Dtls"] = items
    return block


# ---------------------------------------------------------------------------
# Schedule 80G
# ---------------------------------------------------------------------------

def _donation_row(d: dict, eligible_pct: float, warnings: list) -> dict:
    cash = _g(d, "cash")
    other = _g(d, "other")
    if cash > C.CASH_DONATION_LIMIT:
        warnings.append(f"Cash donation of Rs. {cash} to {_s(d, 'name')} is not eligible "
                        f"(above Rs. 2,000) for 80G (rule 88).")
        cash_eligible = 0
    else:
        cash_eligible = cash
    total = cash + other
    eligible = round((cash_eligible + other) * eligible_pct)
    row = {
        "DoneeWithPanName": _s(d, "name"),
        "DoneePAN": _s(d, "pan"),
        "AddressDetail": {
            "AddrDetail": _s((d.get("address") or {}), "addr_detail") or _s((d.get("address") or {}), "line"),
            "CityOrTownOrDistrict": _s((d.get("address") or {}), "city"),
            "StateCode": _s((d.get("address") or {}), "state_code"),
            "PinCode": _g((d.get("address") or {}), "pin_code"),
        },
        "DonationAmtCash": cash,
        "DonationAmtOtherMode": other,
        "DonationAmt": total,
        "EligibleDonationAmt": eligible,
    }
    if _s(d, "arn"):
        row["ArnNbr"] = _s(d, "arn")
    if other > 0:
        if _s(d, "txn_ref"):
            row["TransactionRefNum"] = _s(d, "txn_ref")
        if _s(d, "ifsc"):
            row["IFSCCode"] = _s(d, "ifsc")
    return row, cash_eligible, other, total, eligible


def _build_80g(g: dict | None, adjusted_gti: int, warnings: list) -> tuple[dict, int]:
    zero = {"TotalDonationsUs80GCash": 0, "TotalDonationsUs80GOtherMode": 0,
            "TotalDonationsUs80G": 0, "TotalEligibleDonationsUs80G": 0}
    if not g:
        return zero, 0
    sub_blocks = {
        "Don100Percent": ("don_100", 1.0, "TotDon100PercentCash", "TotDon100PercentOtherMode",
                          "TotDon100Percent", "TotEligibleDon100Percent"),
        "Don50PercentNoApprReqd": ("don_50", 0.5, "TotDon50PercentNoApprReqdCash",
                                   "TotDon50PercentNoApprReqdOtherMode", "TotDon50PercentNoApprReqd",
                                   "TotEligibleDon50Percent"),
        "Don100PercentApprReqd": ("don_100_ql", 1.0, "TotDon100PercentApprReqdCash",
                                  "TotDon100PercentApprReqdOtherMode", "TotDon100PercentApprReqd",
                                  "TotEligibleDon100PercentApprReqd"),
        "Don50PercentApprReqd": ("don_50_ql", 0.5, "TotDon50PercentApprReqdCash",
                                 "TotDon50PercentApprReqdOtherMode", "TotDon50PercentApprReqd",
                                 "TotEligibleDon50PercentApprReqd"),
    }
    out = dict(zero)
    grand_cash = grand_other = grand_total = grand_eligible = 0
    ql_eligible = 0
    noql_eligible = 0
    for blk, (key, pct, ck, ok, tk, ek) in sub_blocks.items():
        rows = []
        cash_t = other_t = tot_t = elig_t = 0
        for d in g.get(key) or []:
            row, cash_el, other, total, eligible = _donation_row(d, pct, warnings)
            rows.append(row)
            cash_t += row["DonationAmtCash"]
            other_t += other
            tot_t += total
            elig_t += eligible
        if not rows:
            continue
        if blk.endswith("ApprReqd"):
            ql_eligible += elig_t
        else:
            noql_eligible += elig_t
        out[blk] = {"DoneeWithPan": rows, ck: cash_t, ok: other_t, tk: tot_t, ek: elig_t}
        grand_cash += cash_t
        grand_other += other_t
        grand_total += tot_t
    # qualifying limit = 10% of adjusted GTI
    qualifying_limit = max(0, adjusted_gti // 10)
    ql_allowed = min(ql_eligible, qualifying_limit)
    grand_eligible = noql_eligible + ql_allowed
    # restrict eligible to qualifying limit inside the blocks themselves
    out.update({
        "TotalDonationsUs80GCash": grand_cash,
        "TotalDonationsUs80GOtherMode": grand_other,
        "TotalDonationsUs80G": grand_total,
        "TotalEligibleDonationsUs80G": grand_eligible,
    })
    return out, grand_eligible


# ---------------------------------------------------------------------------
# Schedule 80GGA / 80GGC
# ---------------------------------------------------------------------------

def _build_80gga(g: dict | None, warnings: list) -> tuple[dict, int]:
    zero = {"TotalDonationAmtCash80GGA": 0, "TotalDonationAmtOtherMode80GGA": 0,
            "TotalDonationsUs80GGA": 0, "TotalEligibleDonationAmt80GGA": 0}
    if not g:
        return zero, 0
    rows = []
    cash_t = other_t = tot_t = elig_t = 0
    for d in g.get("donations") or []:
        cash = _g(d, "cash")
        other = _g(d, "other")
        if cash > C.CASH_DONATION_LIMIT:
            warnings.append("80GGA: cash donation above Rs. 2,000 is not deductible (rule 143).")
            cash_el = 0
        else:
            cash_el = cash
        total = cash + other
        eligible = cash_el + other
        row = {
            "RelevantClauseUndrDedClaimed": _s(d, "clause", "80GGA2b"),
            "NameOfDonee": _s(d, "name"),
            "AddressDetail": {
                "AddrDetail": _s((d.get("address") or {}), "addr_detail") or _s((d.get("address") or {}), "line"),
                "CityOrTownOrDistrict": _s((d.get("address") or {}), "city"),
                "StateCode": _s((d.get("address") or {}), "state_code"),
                "PinCode": _g((d.get("address") or {}), "pin_code"),
            },
            "DonationAmtCash": cash,
            "DonationAmtOtherMode": other,
            "DonationAmt": total,
            "EligibleDonationAmt": eligible,
        }
        if _s(d, "pan"):
            row["DoneePAN"] = _s(d, "pan")
        rows.append(row)
        cash_t += cash
        other_t += other
        tot_t += total
        elig_t += eligible
    if not rows:
        return zero, 0
    return {"DonationDtlsSciRsrchRuralDev": rows,
            "TotalDonationAmtCash80GGA": cash_t,
            "TotalDonationAmtOtherMode80GGA": other_t,
            "TotalDonationsUs80GGA": tot_t,
            "TotalEligibleDonationAmt80GGA": elig_t}, elig_t


def _build_80ggc(g: dict | None, warnings: list) -> tuple[dict, int]:
    zero = {"TotalDonationAmtCash80GGC": 0, "TotalDonationAmtOtherMode80GGC": 0,
            "TotalDonationsUs80GGC": 0, "TotalEligibleDonationAmt80GGC": 0}
    if not g:
        return zero, 0
    rows = []
    cash_t = other_t = tot_t = elig_t = 0
    for d in g.get("donations") or []:
        cash = _g(d, "cash")
        other = _g(d, "other")
        if cash:
            warnings.append("80GGC: cash contributions are not deductible; only other-mode "
                            "contributions are eligible.")
        total = cash + other
        eligible = other                             # rule 194
        row = {
            "DonationDate": _s(d, "date"),
            "DonationAmtCash": cash,
            "DonationAmtOtherMode": other,
            "DonationAmt": total,
            "EligibleDonationAmt": eligible,
        }
        if _s(d, "txn_ref"):
            row["TransactionRefNum"] = _s(d, "txn_ref")
        if _s(d, "ifsc"):
            row["IFSCCode"] = _s(d, "ifsc")
        if _s(d, "party_name"):
            row["PoliticalPartyName"] = _s(d, "party_name")
        if _s(d, "party_pan"):
            row["PoliticalPartyPAN"] = _s(d, "party_pan")
        rows.append(row)
        cash_t += cash
        other_t += other
        tot_t += total
        elig_t += eligible
    if not rows:
        return zero, 0
    return {"Schedule80GGCDetails": rows,
            "TotalDonationAmtCash80GGC": cash_t,
            "TotalDonationAmtOtherMode80GGC": other_t,
            "TotalDonationsUs80GGC": tot_t,
            "TotalEligibleDonationAmt80GGC": elig_t}, elig_t


# ---------------------------------------------------------------------------
# Schedule 80D (health insurance)
# ---------------------------------------------------------------------------

def _build_80d(d: dict | None, warnings: list) -> tuple[dict, int]:
    def ins_block(band):
        policies = list(band.get("policies") or [])
        # convenience: a single policy can be supplied at band level
        if not policies and (_s(band, "insurer") or _s(band, "policy_no")):
            policies = [{"insurer": _s(band, "insurer"), "policy_no": _s(band, "policy_no"),
                         "amount": _g(band, "premium")}]
        rows = [{"InsurerName": _s(p, "insurer"), "PolicyNo": _s(p, "policy_no"),
                 "HealthInsAmt": _g(p, "amount")} for p in policies if _g(p, "amount") > 0]
        return rows, sum(r["HealthInsAmt"] for r in rows)

    if not d:
        return None, 0
    self_flag = _s(d, "self_family_senior_citizen", "S")
    par_flag = _s(d, "parents_senior_citizen", "P")
    block = {"SeniorCitizenFlag": self_flag, "ParentsSeniorCitizenFlag": par_flag}

    total = 0
    prev_total = 0

    if self_flag == "N":
        sf = d.get("self_family") or {}
        premium = _g(sf, "premium")
        policies, pol_total = ins_block(sf)
        if policies:
            premium = pol_total                      # rule 234
        prev = _g(sf, "preventive")
        prev_total += prev
        amt = min(premium + prev, C.CAP_80D_SELF_FAMILY)   # rule 127
        block.update({"SelfAndFamily": amt, "HealthInsPremSlfFam": premium,
                      "PrevHlthChckUpSlfFam": prev})
        if policies:
            block["Sec80DSelfFamHIDtls"] = {"Sch80DInsDtls": policies, "TotalPayments": pol_total}
        total += amt
    elif self_flag == "Y":
        sf = d.get("self_family_sr") or {}
        premium = _g(sf, "premium")
        policies, pol_total = ins_block(sf)
        if policies:
            premium = pol_total
        prev = _g(sf, "preventive")
        prev_total += prev
        med = _g(sf, "medical_exp")
        amt = min(premium + prev + med, C.CAP_80D_SELF_FAMILY_SR)  # rule 130
        block.update({"SelfAndFamilySeniorCitizen": amt, "HlthInsPremSlfFamSrCtzn": premium,
                      "PrevHlthChckUpSlfFamSrCtzn": prev, "MedicalExpSlfFamSrCtzn": med})
        if policies:
            block["Sec80DSelfFamSrCtznHIDtls"] = {"Sch80DInsDtls": policies, "TotalPayments": pol_total}
        total += amt

    if par_flag == "N":
        pf = d.get("parents") or {}
        premium = _g(pf, "premium")
        policies, pol_total = ins_block(pf)
        if policies:
            premium = pol_total
        prev = _g(pf, "preventive")
        prev_total += prev
        amt = min(premium + prev, C.CAP_80D_PARENTS)       # rule 132
        block.update({"Parents": amt, "HlthInsPremParents": premium,
                      "PrevHlthChckUpParents": prev})
        if policies:
            block["Sec80DParentsHIDtls"] = {"Sch80DInsDtls": policies, "TotalPayments": pol_total}
        total += amt
    elif par_flag == "Y":
        pf = d.get("parents_sr") or {}
        premium = _g(pf, "premium")
        policies, pol_total = ins_block(pf)
        if policies:
            premium = pol_total
        prev = _g(pf, "preventive")
        prev_total += prev
        med = _g(pf, "medical_exp")
        amt = min(premium + prev + med, C.CAP_80D_PARENTS_SR)  # rule 134
        block.update({"ParentsSeniorCitizen": amt, "HlthInsPremParentsSrCtzn": premium,
                      "PrevHlthChckUpParentsSrCtzn": prev, "MedicalExpParentsSrCtzn": med})
        if policies:
            block["Sec80DParentsSrCtznHIDtls"] = {"Sch80DInsDtls": policies, "TotalPayments": pol_total}
        total += amt

    if prev_total > C.CAP_80D_PREV_CHECKUP:
        warnings.append(f"Preventive health check-up capped at Rs. {C.CAP_80D_PREV_CHECKUP} "
                        f"combined (rule 129).")
    eligible = min(total, C.CAP_80D_TOTAL)           # rule 136
    block["EligibleAmountOfDedn"] = eligible
    return {"Sec80DSelfFamSrCtznHealth": block}, eligible


# ---------------------------------------------------------------------------
# Loan schedules (80E / 80EE / 80EEA / 80EEB)
# ---------------------------------------------------------------------------

def _loan_rows(block: dict | None, interest_key: str, date_lo: str | None = None,
               date_hi: str | None = None, warnings: list | None = None) -> tuple[list, int]:
    rows = []
    total_interest = 0
    for ln in (block or {}).get("loans") or []:
        dl = _s(ln, "loan_date")
        interest = _g(ln, "interest")
        row = {
            "LoanTknFrom": _s(ln, "lender_type", "B"),
            "BankOrInstnName": _s(ln, "lender_name"),
            "LoanAccNoOfBankOrInstnRefNo": _s(ln, "loan_account_no"),
            "DateofLoan": dl,
            "TotalLoanAmt": _g(ln, "loan_amount"),
            "LoanOutstndngAmt": _g(ln, "outstanding"),
            interest_key: interest,
        }
        rows.append(row)
        total_interest += interest
        if date_lo and dl and not (date_lo <= dl <= date_hi) and warnings is not None:
            warnings.append(f"Loan sanctioned on {dl} is outside the permitted window "
                            f"{date_lo} to {date_hi}.")
    return rows, total_interest


# ---------------------------------------------------------------------------
# Chapter VI-A deductions
# ---------------------------------------------------------------------------

def _build_vi_a(data: dict, opt_out: bool, gti_excl_ltcg: int, salary_171: int,
                employer_cat: str, dob: str, interest_income: int, savings_interest: int,
                family_pension_exists: bool, warnings: list) -> tuple[dict, dict, dict, int]:
    """Return (UsrDeductUndChapVIA, DeductUndChapVIA, schedule_blocks, eligible_total)."""
    ded = data.get("deductions", {}) or {}
    senior = T.is_senior_citizen(dob)

    usr = {k: 0 for k in (
        "Section80C", "Section80CCC", "Section80CCDEmployeeOrSE", "Section80CCD1B",
        "Section80CCDEmployer", "Section80D", "Section80DD", "Section80DDB",
        "Section80E", "Section80EE", "Section80EEA", "Section80EEB", "Section80G",
        "Section80GG", "Section80GGA", "Section80GGC", "Section80U", "Section80TTA",
        "Section80TTB", "AnyOthSec80CCH")}

    with field_scope("Deduction amounts", "deductions"):
        usr["Section80C"] = _g((ded.get("80c") or {}), "amount",
                               label="Deduction u/s 80C")
        usr["Section80CCC"] = _g((ded.get("80ccc") or {}), "amount",
                                 label="Deduction u/s 80CCC")
        usr["Section80CCDEmployeeOrSE"] = _g(ded, "80ccd1",
                                             label="Deduction u/s 80CCD(1)")
        usr["Section80CCD1B"] = _g(ded, "80ccd1b", label="Deduction u/s 80CCD(1B)")
        usr["Section80CCDEmployer"] = _g(ded, "80ccd2",
                                         label="Employer contribution u/s 80CCD(2)")
        usr["Section80DDB"] = _g((ded.get("80ddb") or {}), "amount",
                                 label="Medical treatment u/s 80DDB")
        usr["Section80GG"] = _g((ded.get("80gg") or {}), "amount",
                                label="Rent paid u/s 80GG")
        usr["Section80TTA"] = _g(ded, "80tta", label="Savings interest u/s 80TTA")
        usr["Section80TTB"] = _g(ded, "80ttb", label="Interest u/s 80TTB (senior citizens)")
        usr["AnyOthSec80CCH"] = _g(ded, "80cch", label="Agniveer corpus u/s 80CCH")

    sched: dict = {}
    elig = dict.fromkeys(usr, 0)

    # ---- 80C (with rule 224 detail rows) ----
    c80 = ded.get("80c") or {}
    items = []
    for iidx, it in enumerate(c80.get("items") or []):
        with field_scope(f"80C investment row {iidx + 1}", f"deductions.80c.items.{iidx}"):
            if _g(it, "amount") <= 0:
                continue
            items.append({"IdentificationNo": _s(it, "identification_no"),
                          "Amount": _g(it, "amount")})
    usr_c80 = usr["Section80C"]
    if items:
        usr_c80 = sum(i["Amount"] for i in items)    # rules 241/247
        sched["Schedule80C"] = {"Schedule80CDtls": items, "TotalAmt": usr_c80}
    elig["Section80C"] = min(usr_c80, C.CAP_80C_CCC_CCD1)

    # ---- 80CCC with pension fund details (rule 337) ----
    ccc = ded.get("80ccc") or {}
    funds = []
    for fidx, f in enumerate(ccc.get("funds") or []):
        with field_scope(f"80CCC pension fund row {fidx + 1}", f"deductions.80ccc.funds.{fidx}"):
            if _g(f, "amount") <= 0:
                continue
            funds.append({"TypeofIdentifier": _s(f, "type", "PRAN"),
                          "NameofIdentifier": _s(f, "name"),
                          "Amount": _g(f, "amount")})
    elig["Section80CCC"] = min(usr["Section80CCC"], C.CAP_80C_CCC_CCD1)

    # ---- 80CCD(1): 10% of salary or 20% of GTI (rules 2/3) ----
    if employer_cat in C.PENSIONER_CATEGORIES or employer_cat == "NA":
        cap_ccd1 = int(gti_excl_ltcg * C.CAP_80CCD1_PCT_GTI_OTHERS)
    else:
        cap_ccd1 = int(salary_171 * C.CAP_80CCD1_PCT_SALARIED)
    elig["Section80CCDEmployeeOrSE"] = min(usr["Section80CCDEmployeeOrSE"], cap_ccd1)

    # ---- combined 80C + 80CCC + 80CCD(1) <= 1,50,000 (rule 1) ----
    trio = elig["Section80C"] + elig["Section80CCC"] + elig["Section80CCDEmployeeOrSE"]
    if trio > C.CAP_80C_CCC_CCD1:
        excess = trio - C.CAP_80C_CCC_CCD1
        for k in ("Section80CCDEmployeeOrSE", "Section80CCC", "Section80C"):
            cut = min(excess, elig[k])
            elig[k] -= cut
            excess -= cut
        warnings.append("Combined deduction u/s 80C + 80CCC + 80CCD(1) restricted to "
                        "Rs. 1,50,000 (rule 1).")

    # ---- 80CCD(1B) (rule 115) & PRAN (rule 226) ----
    elig["Section80CCD1B"] = min(usr["Section80CCD1B"], C.CAP_80CCD1B)
    prans_in = ded.get("prans") or []
    if isinstance(prans_in, str):
        prans_in = [p.strip() for p in prans_in.split(",")]
    prans = [{"PRANNum": str(p)} for p in prans_in if str(p).strip()]

    # ---- 80CCD(2): employer contribution (rules 4/120/216; allowed in both regimes) ----
    if employer_cat in C.PENSIONER_CATEGORIES or employer_cat == "NA":
        cap_ccd2 = 0                                  # rule 116
        if usr["Section80CCDEmployer"]:
            warnings.append("80CCD(2) is not admissible for pensioners / NA employer "
                            "category (rule 116); removed.")
    elif opt_out:
        pct = C.CAP_80CCD2_PCT_CG_SG_OLD if employer_cat in ("CGOV", "SGOV") \
            else C.CAP_80CCD2_PCT_OTHERS_OLD
        cap_ccd2 = int(salary_171 * pct)
    else:
        cap_ccd2 = int(salary_171 * C.CAP_80CCD2_PCT_NEW)   # rule 216
    elig["Section80CCDEmployer"] = min(usr["Section80CCDEmployer"], cap_ccd2)

    # ---- 80D schedule ----
    with field_scope("Deduction u/s 80D (mediclaim)", "deductions.80d"):
        sched_80d, elig_80d = _build_80d(ded.get("80d"), warnings)
    if sched_80d:
        sched["Schedule80D"] = sched_80d
    usr["Section80D"] = _g(ded, "80d_manual") or elig_80d
    elig["Section80D"] = elig_80d
    if not sched_80d and _g(ded, "80d_manual"):
        elig["Section80D"] = min(_g(ded, "80d_manual"), C.CAP_80D_TOTAL)
        warnings.append("Schedule 80D details are required to claim 80D (rule 138).")

    # ---- 80DD / 80U ----
    dd = ded.get("80dd")
    if dd:
        nature = _s(dd, "nature", "1")
        amount = C.CAP_80DD_SEVERE if nature == "2" else C.CAP_80DD
        row = {"NatureOfDisability": nature, "TypeOfDisability": _s(dd, "type", "2"),
               "DeductionAmount": amount, "DependentType": _s(dd, "dependent_type", "2")}
        if _s(dd, "dependent_pan"):
            row["DependentPan"] = _s(dd, "dependent_pan")
        if _s(dd, "dependent_aadhaar"):
            row["DependentAadhaar"] = _s(dd, "dependent_aadhaar")
        if _s(dd, "form10ia_ack"):
            row["Form10IAAckNum"] = _s(dd, "form10ia_ack")
        if _s(dd, "udid"):
            row["UDIDNum"] = _s(dd, "udid")
        sched["Schedule80DD"] = row
        usr["Section80DD"] = amount
        elig["Section80DD"] = amount                   # rules 203/204

    uu = ded.get("80u")
    if uu:
        nature = _s(uu, "nature", "1")
        amount = C.CAP_80U_SEVERE if nature == "2" else C.CAP_80U
        row = {"NatureOfDisability": nature, "TypeOfDisability": _s(uu, "type", "2"),
               "DeductionAmount": amount}
        if _s(uu, "form10ia_ack"):
            row["Form10IAAckNum"] = _s(uu, "form10ia_ack")
        if _s(uu, "udid"):
            row["UDIDNum"] = _s(uu, "udid")
        sched["Schedule80U"] = row
        usr["Section80U"] = amount
        elig["Section80U"] = amount                    # rules 200/201

    # ---- 80DDB (rules 5/7) ----
    ddb = ded.get("80ddb") or {}
    if _g(ddb, "amount", label="Medical treatment u/s 80DDB") > 0:
        cap = C.CAP_80DDB_SR if _s(ddb, "type") == "2" else C.CAP_80DDB_NORMAL
        elig["Section80DDB"] = min(_g(ddb, "amount"), cap)

    # ---- loan-linked deductions ----
    e80 = ded.get("80e")
    if e80:
        with field_scope("Education loan u/s 80E", "deductions.80e.loans"):
            rows, tot = _loan_rows(e80, "Interest80E")
        if rows:
            sched["Schedule80E"] = {"Schedule80EDtls": rows, "TotalInterest80E": tot}
        usr["Section80E"] = tot
        elig["Section80E"] = tot

    ee80 = ded.get("80ee")
    if ee80:
        with field_scope("Home loan u/s 80EE", "deductions.80ee.loans"):
            rows, tot = _loan_rows(ee80, "Interest80EE", "2016-04-01", "2017-03-31", warnings)
        if rows:
            sched["Schedule80EE"] = {"Schedule80EEDtls": rows, "TotalInterest80EE": tot}
        usr["Section80EE"] = tot
        elig["Section80EE"] = min(tot, C.CAP_80EE)   # rule 121
        if any(_g(l, "loan_amount") > C.CAP_80EE_LOAN_AMOUNT for l in (ee80.get("loans") or [])):
            warnings.append("80EE admissible only for loans up to Rs. 35 lakh (rule 227).")
            elig["Section80EE"] = 0

    eea = ded.get("80eea")
    if eea:
        sdv = _g(eea, "stamp_duty_value")
        with field_scope("Affordable housing loan u/s 80EEA", "deductions.80eea.loans"):
            rows, tot = _loan_rows(eea, "Interest80EEA", "2019-04-01", "2022-03-31", warnings)
        block = {"PropStmpDtyVal": sdv}
        if rows:
            block["Schedule80EEADtls"] = rows
        block["TotalInterest80EEA"] = tot
        sched["Schedule80EEA"] = block
        usr["Section80EEA"] = tot
        elig["Section80EEA"] = min(tot, C.CAP_80EEA)  # rule 122
        if sdv > C.CAP_80EEA_STAMP_DUTY_VALUE:
            warnings.append("80EEA allowed only where stamp-duty value <= Rs. 45 lakh (rule 229).")
            elig["Section80EEA"] = 0
        if elig["Section80EEA"] > 0 and elig["Section80EE"] > 0:
            warnings.append("Only one of 80EE / 80EEA is allowed (rule 123); keeping 80EEA.")
            elig["Section80EE"] = 0

    eeb = ded.get("80eeb")
    if eeb:
        with field_scope("Electric vehicle loan u/s 80EEB", "deductions.80eeb.loans"):
            rows, tot = _loan_rows(eeb, "Interest80EEB", "2019-04-01", "2023-03-31", warnings)
        if rows:
            sched["Schedule80EEB"] = {"Schedule80EEBDtls": rows, "TotalInterest80EEB": tot}
        usr["Section80EEB"] = tot
        elig["Section80EEB"] = min(tot, C.CAP_80EEB)  # rule 124

    # ---- 80TTA / 80TTB (rules 11-16) ----
    if senior:
        usr["Section80TTA"] = 0
        elig["Section80TTB"] = min(usr["Section80TTB"], interest_income, C.CAP_80TTB)
    else:
        usr["Section80TTB"] = 0
        elig["Section80TTA"] = min(usr["Section80TTA"], savings_interest, C.CAP_80TTA)  # rule 12

    # ---- 80GG (rule 114) ----
    gg = ded.get("80gg") or {}
    if _g(gg, "amount"):
        if _s((data.get("salary") or {}).get("hra", {}), "hra_received") not in ("", "0") \
                and _g((data.get("salary") or {}).get("hra", {}), "hra_received") > 0:
            warnings.append("80GG is not admissible when HRA is claimed (rule 119).")
            elig["Section80GG"] = 0
        else:
            adjusted = max(0, gti_excl_ltcg - (elig["Section80C"] + elig["Section80CCC"]
                                               + elig["Section80CCDEmployeeOrSE"]
                                               + elig["Section80CCD1B"]))
            cap = min(C.CAP_80GG, int(adjusted * C.CAP_80GG_PCT_OF_ADJ_TI))
            rent_paid = _g(gg, "rent_paid")
            if rent_paid:
                cap = min(cap, max(0, rent_paid - adjusted // 10))
            elig["Section80GG"] = min(_g(gg, "amount"), cap)
            usr["Section80GG"] = _g(gg, "amount")

    # ---- 80CCH (Agniveer) — available in both regimes (rules 186/187) ----
    cap_cch = min(int(salary_171 * C.CAP_80CCH_PCT_OF_SALARY), C.CAP_80CCH_SCHEMA_MAX)
    if employer_cat != "CGOV":
        if usr["AnyOthSec80CCH"]:
            warnings.append("80CCH is admissible only to Central Government employees "
                            "enrolled under Agnipath (rule 187).")
        elig["AnyOthSec80CCH"] = 0
    else:
        elig["AnyOthSec80CCH"] = min(usr["AnyOthSec80CCH"], cap_cch)

    # ---- 80G / 80GGA / 80GGC ----
    other_deductions = sum(v for k, v in elig.items() if k not in ("Section80G",))
    adjusted_gti_80g = max(0, gti_excl_ltcg - other_deductions)
    with field_scope("Deduction u/s 80G (donations)", "deductions.80g"):
        sched_80g, elig_80g = _build_80g(ded.get("80g"), adjusted_gti_80g, warnings)
    sched["Schedule80G"] = sched_80g
    usr["Section80G"] = max(elig_80g, _g(ded, "80g_manual"))
    elig["Section80G"] = elig_80g

    with field_scope("Deduction u/s 80GGA", "deductions.80gga"):
        sched_80gga, elig_80gga = _build_80gga(ded.get("80gga"), warnings)
    sched["Schedule80GGA"] = sched_80gga
    usr["Section80GGA"] = elig_80gga
    elig["Section80GGA"] = elig_80gga

    with field_scope("Deduction u/s 80GGC (political party)", "deductions.80ggc"):
        sched_80ggc, elig_80ggc = _build_80ggc(ded.get("80ggc"), warnings)
    sched["Schedule80GGC"] = sched_80ggc
    usr["Section80GGC"] = elig_80ggc
    elig["Section80GGC"] = elig_80ggc
    if elig_80ggc > 0 and _g(ded, "80ggc_manual"):
        usr["Section80GGC"] = _g(ded, "80ggc_manual")

    # ---- new-regime restriction (rule 146 and related) ----
    if not opt_out:
        allowed_new = {"Section80CCDEmployer", "AnyOthSec80CCH"}
        stripped = [k for k, v in elig.items() if v and k not in allowed_new]
        for k in stripped:
            elig[k] = 0
        for k in list(usr):
            if k != "TotalChapVIADeductions" and k not in allowed_new:
                usr[k] = 0
        if stripped:
            warnings.append("Under the new tax regime, Chapter VI-A deductions other than "
                            "80CCD(2) and 80CCH are not allowed; they are ignored "
                            "(rules 153-175).")
        # schedules not permitted under new regime (rule 255)
        for key in ("Schedule80C", "Schedule80D", "Schedule80DD", "Schedule80U",
                    "Schedule80E", "Schedule80EE", "Schedule80EEA", "Schedule80EEB",
                    "Schedule80G", "Schedule80GGA", "Schedule80GGC"):
            if key == "Schedule80G":
                sched[key] = {"TotalDonationsUs80GCash": 0, "TotalDonationsUs80GOtherMode": 0,
                              "TotalDonationsUs80G": 0, "TotalEligibleDonationsUs80G": 0}
            elif key == "Schedule80GGA":
                sched[key] = {"TotalDonationAmtCash80GGA": 0, "TotalDonationAmtOtherMode80GGA": 0,
                              "TotalDonationsUs80GGA": 0, "TotalEligibleDonationAmt80GGA": 0}
            elif key == "Schedule80GGC":
                sched[key] = {"TotalDonationAmtCash80GGC": 0, "TotalDonationAmtOtherMode80GGC": 0,
                              "TotalDonationsUs80GGC": 0, "TotalEligibleDonationAmt80GGC": 0}
            else:
                sched.pop(key, None)
        # HRA schedule also not applicable
        data.pop("_hra_block", None)

    usr_total = sum(usr.values())
    elig_total = min(sum(elig.values()), gti_excl_ltcg)   # rules 17/18

    usr["TotalChapVIADeductions"] = usr_total
    elig["TotalChapVIADeductions"] = elig_total

    usr_block = dict(usr)
    if funds:
        usr_block["PensionContribution80CCC"] = funds
    if prans:
        usr_block["PRANDtls"] = prans
    ddb = ded.get("80ddb") or {}
    if _s(ddb, "type"):
        usr_block["Section80DDBUsrType"] = _s(ddb, "type")
    if _s(ddb, "disease"):
        usr_block["NameOfSpecDisease80DDB"] = _s(ddb, "disease")
    if _s((ded.get("80gg") or {}), "form10ba_ack"):
        usr_block["Form10BAAckNum"] = _s(ded.get("80gg"), "form10ba_ack")

    return usr_block, elig, sched, elig_total


# ---------------------------------------------------------------------------
# Taxes paid
# ---------------------------------------------------------------------------

def _build_tax_paid(tp: dict | None) -> tuple[dict, dict, dict, dict, dict, int, int, list]:
    tp = tp or {}
    tds1_rows = []
    for ridx, r in enumerate(tp.get("tds_on_salary") or []):
        with field_scope(f"TDS on salary row {ridx + 1}", f"tax_paid.tds_on_salary.{ridx}"):
            tds = _g(r, "tds")
            if tds <= 0:
                continue
            tds1_rows.append({
                "EmployerOrDeductorOrCollectDetl": {
                    "TAN": _s(r, "tan"), "EmployerOrDeductorOrCollecterName": _s(r, "employer_name")},
                "IncChrgSal": _g(r, "salary_income"),
                "TotalTDSSal": tds,
            })
    total_tds1 = sum(r["TotalTDSSal"] for r in tds1_rows)

    tds2_rows = []
    for ridx, r in enumerate(tp.get("tds_others") or []):
        with field_scope(f"TDS (other than salary) row {ridx + 1}", f"tax_paid.tds_others.{ridx}"):
            if _g(r, "tds_claimed") <= 0:
                continue
            claimed = min(_g(r, "tds_claimed"), _g(r, "tds_deducted"))  # rule 98
            tds2_rows.append({
                "EmployerOrDeductorOrCollectDetl": {
                    "TAN": _s(r, "tan"), "EmployerOrDeductorOrCollecterName": _s(r, "deductor_name")},
                "AmtForTaxDeduct": _g(r, "amount_paid"),
                "DeductedYr": _s(r, "deduction_year", "2025"),
                "TotTDSOnAmtPaid": _g(r, "tds_deducted"),
                "ClaimOutOfTotTDSOnAmtPaid": claimed,
                "TDSSection": _s(r, "section", "94A"),
            })
    total_tds2 = sum(r["ClaimOutOfTotTDSOnAmtPaid"] for r in tds2_rows)

    tds3_rows = []
    for ridx, r in enumerate(tp.get("tds3") or []):
        with field_scope(f"TDS on rent (26QC) row {ridx + 1}", f"tax_paid.tds3.{ridx}"):
            if _g(r, "tds_claimed") <= 0:
                continue
            tds3_rows.append({
                "PANofTenant": _s(r, "tenant_pan"),
                "NameOfTenant": _s(r, "tenant_name"),
                "GrsRcptToTaxDeduct": _g(r, "gross_receipt"),
                "DeductedYr": _s(r, "deduction_year", "2025"),
                "TDSDeducted": _g(r, "tds_deducted"),
                "TDSClaimed": min(_g(r, "tds_claimed"), _g(r, "tds_deducted")),
                "TDSSection": _s(r, "section", "4-IB"),
            })
        if _s(r, "tenant_aadhaar"):
            tds3_rows[-1]["AadhaarofTenant"] = _s(r, "tenant_aadhaar")
    total_tds3 = sum(r["TDSClaimed"] for r in tds3_rows)

    tcs_rows = []
    for ridx, r in enumerate(tp.get("tcs") or []):
        with field_scope(f"TCS row {ridx + 1}", f"tax_paid.tcs.{ridx}"):
            claimed = _g(r, "tcs_claimed")
            if claimed <= 0:
                continue
            claimed = min(claimed, _g(r, "tcs_collected"))                # rule 96
            tcs_rows.append({
                "EmployerOrDeductorOrCollectDetl": {
                    "TAN": _s(r, "tan"), "EmployerOrDeductorOrCollecterName": _s(r, "collector_name")},
                "AmtTaxCollected": _g(r, "amount_subject"),
                "CollectedYr": _s(r, "collection_year", "2025"),
                "TotalTCS": _g(r, "tcs_collected"),
                "AmtTCSClaimedThisYear": claimed,
            })
    total_tcs = sum(r["AmtTCSClaimedThisYear"] for r in tcs_rows)

    challan_rows = []
    for cidx, c in enumerate(tp.get("challans") or []):
        with field_scope(f"Advance/Self-assessment tax challan {cidx + 1}",
                         f"tax_paid.challans.{cidx}"):
            if _g(c, "amount") <= 0:
                continue
            challan_rows.append({
                "BSRCode": _s(c, "bsr_code"),
                "DateDep": _s(c, "date"),
                "SrlNoOfChaln": _g(c, "challan_no"),
                "Amt": _g(c, "amount"),
            })
    total_challans = sum(c["Amt"] for c in challan_rows)

    tds_total = total_tds1 + total_tds2 + total_tds3                  # rule 108
    sched_tds1 = {"TotalTDSonSalaries": total_tds1}
    if tds1_rows:
        sched_tds1["TDSonSalary"] = tds1_rows
    sched_tds2 = {"TotalTDSonOthThanSals": total_tds2}
    if tds2_rows:
        sched_tds2["TDSonOthThanSal"] = tds2_rows
    sched_tds3 = {"TotalTDS3Details": total_tds3}
    if tds3_rows:
        sched_tds3["TDS3Details"] = tds3_rows
    sched_tcs = {"TotalSchTCS": total_tcs}
    if tcs_rows:
        sched_tcs["TCS"] = tcs_rows
    sched_it = {"TotalTaxPayments": total_challans}
    if challan_rows:
        sched_it["TaxPayment"] = challan_rows

    return sched_tds1, sched_tds2, sched_tds3, sched_tcs, sched_it, tds_total, total_tcs, challan_rows


# ---------------------------------------------------------------------------
# Public builder
# ---------------------------------------------------------------------------

def build_return(data: dict, *, creation_info: dict | None = None) -> dict:
    """Build the complete official ITR-1 JSON plus a computation report.

    Returns ``{"json": <official JSON>, "report": {...}}``.
    ``report`` carries the computed numbers (used by the UI and tests).
    """
    reset_scopes()
    data = copy.deepcopy(data or {})
    warnings: list[str] = []

    personal = data.get("personal", {}) or {}
    filing = data.get("filing", {}) or {}
    opt_out = bool(filing.get("opt_out_new_regime"))
    dob = _s(personal, "dob")
    if not re.match(r"^[12]\d{3}-\d{2}-\d{2}$", dob):
        if dob:
            warnings.append(f"Date of birth '{dob}' is not a valid YYYY-MM-DD date; "
                            "computation assumes age below 60. Fix it before filing.")
        else:
            with field_scope("Salary income", "salary"):
                has_income = any(_g(data.get("salary", {}), k) for k in
                                 ("salary_17_1", "perquisites_17_2", "profits_17_3"))
            if has_income or data.get("other_sources"):
                warnings.append("Date of birth is not filled yet — computation assumes age "
                                "below 60 (required before filing).")
        dob = "1990-01-01"
    employer_cat = _s(personal, "employer_category", "NA")

    # Residential-status gate: ITR-1 (Sahaj) is only for an individual who
    # is "a resident (other than not ordinarily resident)" — the official
    # utility states this verbatim (EfilingCommon.bas eligibility message).
    # NRI and RNOR filers must use ITR-2/ITR-3.
    residential_status = (_s(personal, "residential_status")
                          or _s(filing, "residential_status") or "RES").upper()[:3]
    if residential_status in ("NRI", "NOR"):
        raise DataError(
            f"residential status is {'Non-Resident' if residential_status == 'NRI' else 'Resident but Not Ordinarily Resident'}: "
            "ITR-1 is only for individuals who are resident (other than not "
            "ordinarily resident).  Please file ITR-2 (or ITR-3 if you have "
            "business/professional income).  The generated JSON will not be "
            "produced so a wrong form cannot be filed by mistake.")

    # ---- income heads ----
    salary_block, hra_block = _build_salary(data, opt_out, warnings)
    if hra_block:
        data["_hra_block"] = hra_block

    hp_blocks, hp_total = _build_house_property(data.get("house_properties"), opt_out, warnings)
    others_block, oth_src_total, ded_57iia = _build_other_sources(
        data.get("other_sources"), opt_out, warnings)
    exempt_block = _build_exempt_income(data.get("exempt_income"), warnings)

    salary_171 = salary_block["Salary"]

    interest_income = sum(
        _g(r, "amount") for r in (data.get("other_sources") or [])
        if _s(r, "nature") in ("SAV", "IFD", "TAX"))
    savings_interest = sum(
        _g(r, "amount") for r in (data.get("other_sources") or [])
        if _s(r, "nature") == "SAV")

    gti_excl_ltcg = salary_block["IncomeFromSal"] + hp_total + oth_src_total  # rules 22/160

    ltcg_in = data.get("ltcg112a") or {}
    with field_scope("LTCG u/s 112A", "ltcg112a"):
        ltcg = max(0, _g(ltcg_in, "sale_consideration") - _g(ltcg_in, "cost_of_acquisition"))
        ltcg = min(ltcg, C.MAX_LTCG_112A)                 # rule 217
        if (_g(ltcg_in, "sale_consideration") - _g(ltcg_in, "cost_of_acquisition")
                > C.MAX_LTCG_112A):
            warnings.append("LTCG u/s 112A above Rs. 1.25 lakh cannot be reported in ITR-1; "
                            "please use ITR-2. The excess is ignored here.")
    gti_incl_ltcg = gti_excl_ltcg + ltcg              # rule 292

    # ---- deductions ----
    usr_via, elig_via, ded_scheds, via_total = _build_vi_a(
        data, opt_out, gti_excl_ltcg, salary_171, employer_cat, dob,
        interest_income, savings_interest, ded_57iia > 0, warnings)

    # 80DD/80U/80E etc. schedules collected above; note VIA total already capped at GTI
    total_income = T.round_off_income(max(0, gti_incl_ltcg - via_total))  # rule 24

    with field_scope("LTCG u/s 112A", "ltcg112a"):
        ltcg_block = {
            "TotSaleCnsdrn": _g(ltcg_in, "sale_consideration"),
            "TotCstAcqisn": _g(ltcg_in, "cost_of_acquisition"),
            "LongCap112A": ltcg,
        }

    # ---- tax paid ----
    sched_tds1, sched_tds2, sched_tds3, sched_tcs, sched_it, tds_total, tcs_total, challans = \
        _build_tax_paid(data.get("tax_paid"))

    # ---- tax computation ----
    filing_date = _s(filing, "filing_date") or str(TODAY())
    with field_scope("Filing status", "filing"):
        return_section = _g(filing, "return_section", 11)
    section89 = _g(data, "relief_89")
    exp_sal = data.get("_salary_expansion")
    if exp_sal and exp_sal.get("arrears_rows") and not section89:
        dob_year = int(dob[:4]) if dob and dob[:4].isdigit() else None
        reg = "old" if opt_out else "new"
        w10e = SX.relief_89_working(total_income, exp_sal["arrears_rows"],
                                    regime=reg, dob_year=dob_year)
        data["_form10e"] = w10e
        if w10e["relief89"] > 0:
            section89 = w10e["relief89"]
            warnings.append(
                f"Relief u/s 89 computed automatically from your arrear breakup "
                f"(Form 10E workings): Rs. {section89:,}. File Form 10E online before/at filing.")
        elif w10e["total_arrears"] > 0:
            warnings.append("Arrears declared, but no relief u/s 89 arises (the arrears are "
                            "taxed at or below the slab rates of the years they belong to).")
    tax = T.compute_tax(
        total_income=total_income, ltcg112a=ltcg, opt_out_new_regime=opt_out, dob=dob,
        section89=section89, filing_date=filing_date,
        return_section=return_section, tds=tds_total, tcs=tcs_total,
        advance_tax_challans=(data.get("tax_paid") or {}).get("challans"),
        interest_override=data.get("interest_override"),
        residential_status=residential_status,
        status=_s(personal, "status") or "I")

    # ---- assemble income/deductions block ----
    inc_ded = {
        "GrossSalary": salary_block["GrossSalary"],
        "Salary": salary_block["Salary"],
        "PerquisitesValue": salary_block["PerquisitesValue"],
        "ProfitsInSalary": salary_block["ProfitsInSalary"],
        "AllwncExemptUs10": salary_block["AllwncExemptUs10"],
        "NetSalary": salary_block["NetSalary"],
        "DeductionUs16": salary_block["DeductionUs16"],
        "DeductionUs16ia": salary_block["DeductionUs16ia"],
        "EntertainmentAlw16ii": salary_block["EntertainmentAlw16ii"],
        "ProfessionalTaxUs16iii": salary_block["ProfessionalTaxUs16iii"],
        "IncomeFromSal": salary_block["IncomeFromSal"],
        "TotalIncomeChargeableUnHP": hp_total,
        "IncomeOthSrc": oth_src_total,
        "DeductionUs57iia": ded_57iia,
        "GrossTotIncome": gti_excl_ltcg,
        "GrossTotIncomeIncLTCG112A": gti_incl_ltcg,
        "UsrDeductUndChapVIA": usr_via,
        "DeductUndChapVIA": elig_via,
        "TotalIncome": total_income,
        "ExemptIncAgriOthUs10": exempt_block,
    }
    if hp_blocks:
        inc_ded["PropertyDetails"] = hp_blocks
    if others_block:
        inc_ded["OthersInc"] = others_block

    tax_comp = {
        "TotalTaxPayable": tax["TotalTaxPayable"],
        "Rebate87A": tax["Rebate87A"],
        "TaxPayableOnRebate": tax["TaxPayableOnRebate"],
        "EducationCess": tax["EducationCess"],
        "GrossTaxLiability": tax["GrossTaxLiability"],
        "Section89": tax["Section89"],
        "NetTaxLiability": tax["NetTaxLiability"],
        "TotalIntrstPay": tax["TotalIntrstPay"],
        "IntrstPay": tax["IntrstPay"],
        "TotTaxPlusIntrstPay": tax["TotTaxPlusIntrstPay"],
    }

    # ---- personal info ----
    personal_scope = field_scope("Personal details & address", "personal.address")
    personal_scope.__enter__()
    addr_in = personal.get("address", {}) or {}
    addr = {
        "ResidenceNo": _s(addr_in, "residence_no"),
        "LocalityOrArea": _s(addr_in, "locality_or_area"),
        "CityOrTownOrDistrict": _s(addr_in, "city_or_town_or_district"),
        "StateCode": _s(addr_in, "state_code"),
        "CountryCode": _s(addr_in, "country_code", "91"),
        "PinCode": _g(addr_in, "pin_code"),
        "CountryCodeMobile": _g(addr_in, "country_code_mobile", 91),
        "MobileNo": _g(addr_in, "mobile_no"),
        "EmailAddress": _s(addr_in, "email_address"),
    }
    if _s(addr_in, "residence_name"):
        addr["ResidenceName"] = _s(addr_in, "residence_name")
    if _s(addr_in, "road_or_street"):
        addr["RoadOrStreet"] = _s(addr_in, "road_or_street")
    if _g(addr_in, "mobile_no_sec"):
        addr["CountryCodeMobileNoSec"] = _g(addr_in, "country_code_mobile_sec", 91)
        addr["MobileNoSec"] = _g(addr_in, "mobile_no_sec")
    if _s(addr_in, "email_address_sec"):
        addr["EmailAddressSec"] = _s(addr_in, "email_address_sec")

    secondary_same = personal.get("secondary_address_same", True)
    personal_block = {
        "AssesseeName": {
            "FirstName": _s(personal, "first_name"),
            "MiddleName": _s(personal, "middle_name"),
            "SurNameOrOrgName": _s(personal, "last_name") or _s(personal, "first_name"),
        },
        "PAN": _s(personal, "pan"),
        "Address": addr,
        "SecondaryAdd": "Y" if secondary_same else "N",
        "DOB": dob,
        "EmployerCategory": employer_cat,
    }
    if not secondary_same:
        alt_in = personal.get("alternate_address", {}) or {}
        personal_block["AlternateAddress"] = {
            "ResidenceNo": _s(alt_in, "residence_no"),
            "LocalityOrArea": _s(alt_in, "locality_or_area"),
            "CityOrTownOrDistrict": _s(alt_in, "city_or_town_or_district"),
            "StateCode": _s(alt_in, "state_code"),
            "CountryCode": _s(alt_in, "country_code", "91"),
            "PinCode": _g(alt_in, "pin_code"),
        }
        if _s(alt_in, "residence_name"):
            personal_block["AlternateAddress"]["ResidenceName"] = _s(alt_in, "residence_name")
        if _s(alt_in, "road_or_street"):
            personal_block["AlternateAddress"]["RoadOrStreet"] = _s(alt_in, "road_or_street")
    if _s(personal, "aadhaar"):
        personal_block["AadhaarCardNo"] = _s(personal, "aadhaar")

    # ---- filing status ----
    filing_block = {
        "ReturnFileSec": return_section,
        "OptOutNewTaxRegime": "Y" if opt_out else "N",
        "AsseseeRepFlg": "Y" if filing.get("is_representative") else "N",
        "ItrFilingDueDate": _s(filing, "due_date", C.DUE_DATE_FILING),
    }
    if filing.get("seventh_provisio"):
        filing_block["SeventhProvisio139"] = "Y"
        filing_block["IncrExpAggAmt2LkTrvFrgnCntryFlg"] = \
            "Y" if filing.get("foreign_travel_exceeds_2l") else "N"
        filing_block["IncrExpAggAmt1LkElctrctyPrYrFlg"] = \
            "Y" if filing.get("electricity_exceeds_1l") else "N"
        filing_block["clauseiv7provisio139i"] = "Y" if filing.get("deposit_exceeds_1cr") else "N"
    else:
        filing_block["SeventhProvisio139"] = "N"
        filing_block["IncrExpAggAmt2LkTrvFrgnCntryFlg"] = "N"
        filing_block["IncrExpAggAmt1LkElctrctyPrYrFlg"] = "N"
        filing_block["clauseiv7provisio139i"] = "N"
    if return_section == 17:
        if _s(filing, "original_receipt_no"):
            filing_block["ReceiptNo"] = _s(filing, "original_receipt_no")
        if _s(filing, "original_filing_date"):
            filing_block["OrigRetFiledDate"] = _s(filing, "original_filing_date")
    if return_section in (13, 14, 16, 18):
        if _s(filing, "notice_no"):
            filing_block["NoticeNo"] = _s(filing, "notice_no")
        if _s(filing, "notice_date"):
            filing_block["NoticeDateUnderSec"] = _s(filing, "notice_date")
    if filing.get("is_representative"):
        rep = filing.get("representative", {}) or {}
        filing_block["AssesseeRep"] = {
            "RepName": _s(rep, "name"),
            "RepEmailID": _s(rep, "email"),
            "CountryCodeRepMobileNo": _g(rep, "country_code", 91),
            "RepMobileNo": _g(rep, "mobile"),
        }

    # ---- verification ----
    ver = data.get("verification", {}) or {}
    ver_block = {
        "Declaration": {
            "AssesseeVerName": _s(ver, "name"),
            "FatherName": _s(ver, "father_name"),
            "AssesseeVerPAN": _s(ver, "pan") or _s(personal, "pan"),
        },
        "Capacity": _s(ver, "capacity", "S"),
        "Place": _s(ver, "place"),
    }

    personal_scope.__exit__(None, None, None)

    # ---- bank accounts ----
    bank_rows = []
    for bidx, b in enumerate(data.get("bank_accounts", []) or []):
        row = {
            "IFSCCode": _s(b, "ifsc"),
            "BankName": _s(b, "bank_name"),
            "BankAccountNo": _s(b, "account_no"),
            "AccountType": _s(b, "account_type", "SB"),
            "UseForRefund": "true" if b.get("use_for_refund") else "false",
        }
        bank_rows.append(row)
    if not bank_rows:
        warnings.append("At least one bank account is required for the return.")
    elif tax["RefundDue"] > 0 and not any(b["UseForRefund"] == "true" for b in bank_rows):
        warnings.append("A refund is due: nominate one bank account for refund (UseForRefund).")

    # ---- creation info ----
    ci = dict(creation_info or {})
    json_created_by = ci.get("sw_created_by", SW_CREATED_BY)
    creation_block = {
        "SWVersionNo": str(ci.get("sw_version", UTILITY_SW_VERSION_BY_FORM["ITR1"])),
        "SWCreatedBy": json_created_by,
        "JSONCreatedBy": json_created_by,
        "JSONCreationDate": str(ci.get("creation_date", TODAY())),
        "IntermediaryCity": _s(ci, "intermediary_city", "Delhi"),
        "Digest": ci.get("digest", "-"),
    }

    itr1 = {
        "CreationInfo": creation_block,
        "Form_ITR1": {
            "FormName": "ITR-1",
            "Description": UTILITY_FORM_DESCRIPTION["ITR1"],
            "AssessmentYear": "2026",
            "SchemaVer": "Ver1.0",
            "FormVer": "Ver1.0",
        },
        "PersonalInfo": personal_block,
        "FilingStatus": filing_block,
        "ITR1_IncomeDeductions": inc_ded,
        "ITR1_TaxComputation": tax_comp,
        "TaxPaid": {
            "TaxesPaid": tax["TaxesPaid"],
            "BalTaxPayable": tax["BalTaxPayable"],
        },
        "Refund": {
            "RefundDue": tax["RefundDue"],
            "BankAccountDtls": {"AddtnlBankDetails": bank_rows},
        },
        "TDSonSalaries": sched_tds1,
        "TDSonOthThanSals": sched_tds2,
        "ScheduleTDS3Dtls": sched_tds3,
        "ScheduleTCS": sched_tcs,
        "TaxPayments": sched_it,
        "LTCG112A": ltcg_block,
        "Verification": ver_block,
    }
    if data.get("_hra_block"):
        itr1["ScheduleEA10_13A"] = data["_hra_block"]

    # deduction schedules (only include meaningful ones)
    for key, value in ded_scheds.items():
        if key in ("Schedule80G", "Schedule80GGA", "Schedule80GGC"):
            # include only when gross donations are non-zero or the schedule has rows
            totals = [v for k, v in value.items() if isinstance(v, int)]
            has_rows = any(isinstance(v, list) and v for v in value.values()) or \
                any("Donee" in k or "Dtls" in k or "Details" in k for k in value)
            if sum(totals) > 0 or has_rows:
                itr1[key] = value
            continue
        itr1[key] = value

    doc = {"ITR": {"ITR1": itr1}}

    report = {
        "regime": "old" if opt_out else "new",
        "gross_total_income_excl_ltcg": gti_excl_ltcg,
        "gross_total_income_incl_ltcg": gti_incl_ltcg,
        "ltcg112a": ltcg,
        "hp_income": hp_total,
        "other_sources_income": oth_src_total,
        "salary_income": salary_block["IncomeFromSal"],
        "deductions_eligible": elig_via,
        "total_income": total_income,
        "tax": tax,
        "warnings": warnings,
        # advanced salary material (None when only classic boxes are used)
        "salary_expansion": data.get("_salary_expansion"),
        "form10e": data.get("_form10e"),
    }
    return {"json": doc, "report": report}


def compute_both_regimes(data: dict) -> dict:
    """Compute the liability under both regimes for comparison (helper for UI)."""
    cur = copy.deepcopy(data or {})
    filing = cur.setdefault("filing", {})
    filing["opt_out_new_regime"] = False
    new_res = build_return(cur)
    filing["opt_out_new_regime"] = True
    old_res = build_return(cur)
    return {"new": new_res, "old": old_res}
