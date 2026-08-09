"""Validation of an official ITR-1 JSON document.

Two layers, mirroring the e-Filing portal:

1. **Schema validation** against ``ITR-1_2026_Main_V1.1.json`` (JSON Schema
   draft-04) bundled in this repository.
2. **CBDT validation rules** for AY 2026-27.  Each check references the rule
   number from ``CBDT_e-Filing_ITR 1_Validation Rules_AY 2026-27.pdf`` and is
   categorised exactly like the portal:
     - **A** - return will not be allowed to be uploaded (hard error)
     - **B** - upload allowed, defect notice u/s 139(9) likely (warning)
     - **D** - claim may be disallowed without supporting forms (info)
"""

from __future__ import annotations

import json
import os
import re

from . import constants as C

SCHEMA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           os.pardir, "ITR-1_2026_Main_V1.1.json")

_schema_cache = None


def load_schema(path: str | None = None) -> dict:
    global _schema_cache
    if path or _schema_cache is None:
        with open(path or SCHEMA_PATH, "r", encoding="utf-8") as fh:
            schema = json.load(fh)
        if path is None:
            _schema_cache = schema
        return schema
    return _schema_cache


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _dig(node, *path, default=0):
    cur = node
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return default
        cur = cur[p]
    return cur


def _i(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


class Issue:
    __slots__ = ("category", "rule", "message", "path")

    def __init__(self, category: str, rule: str, message: str, path: str = ""):
        self.category = category  # "A" | "B" | "D" | "SCHEMA"
        self.rule = rule
        self.message = message
        self.path = path

    def as_dict(self):
        return {"category": self.category, "rule": self.rule,
                "message": self.message, "path": self.path}

    def __repr__(self):
        return f"[{self.category}] {self.rule}: {self.message}"


# ---------------------------------------------------------------------------
# Layer 1 - JSON Schema
# ---------------------------------------------------------------------------

def validate_schema(doc: dict, schema: dict | None = None) -> list[Issue]:
    try:
        import jsonschema
    except ImportError:  # pragma: no cover
        return [Issue("SCHEMA", "-", "package 'jsonschema' is not installed")]
    schema = schema or load_schema()
    validator_cls = jsonschema.validators.validator_for(schema)
    validator_cls.check_schema(schema)
    validator = validator_cls(schema)
    issues = []
    for err in sorted(validator.iter_errors(doc), key=lambda e: list(e.absolute_path)):
        path = "$" + "".join(f".{p}" if isinstance(p, str) else f"[{p}]"
                             for p in err.absolute_path)
        issues.append(Issue("SCHEMA", "schema", err.message, path))
    return issues


# ---------------------------------------------------------------------------
# Layer 2 - CBDT validation rules
# ---------------------------------------------------------------------------

def validate_rules(doc: dict) -> list[Issue]:
    issues = []

    def A(rule, msg, cond=True):
        if cond:
            issues.append(Issue("A", f"Rule {rule}", msg))

    def B(rule, msg, cond=True):
        if cond:
            issues.append(Issue("B", f"Rule {rule}", msg))

    def D(rule, msg, cond=True):
        if cond:
            issues.append(Issue("D", f"Rule {rule}", msg))

    itr1 = _dig(doc, "ITR", "ITR1", default={})
    inc = itr1.get("ITR1_IncomeDeductions", {})
    tc = itr1.get("ITR1_TaxComputation", {})
    pi = itr1.get("PersonalInfo", {})
    fs = itr1.get("FilingStatus", {})
    paid = itr1.get("TaxPaid", {}).get("TaxesPaid", {})
    refund = itr1.get("Refund", {})
    ver = itr1.get("Verification", {})

    opt_out = _dig(fs, "OptOutNewTaxRegime", default="N") == "Y"
    regime_name = "old" if opt_out else "new"
    dob = _dig(pi, "DOB", default="2000-01-01")
    senior = str(dob) <= C.SENIOR_CITIZEN_DOB_ON_OR_BEFORE
    employer_cat = _dig(pi, "EmployerCategory", default="NA")

    usr = inc.get("UsrDeductUndChapVIA", {})
    ded = inc.get("DeductUndChapVIA", {})
    gti = _i(inc.get("GrossTotIncome"))
    gti_inc = _i(inc.get("GrossTotIncomeIncLTCG112A"))
    ti = _i(inc.get("TotalIncome"))
    ltcg = _i(_dig(doc, "ITR", "ITR1", "LTCG112A", "LongCap112A"))
    sal = _i(inc.get("IncomeFromSal"))
    hp = _i(inc.get("TotalIncomeChargeableUnHP"))
    os_inc = _i(inc.get("IncomeOthSrc"))

    D_close = lambda a, b, tol=1: abs(_i(a) - _i(b)) <= tol

    # -- Rule 1: 80C + 80CCC + 80CCD(1) <= 1,50,000 (old regime) --
    if opt_out:
        trio = (_i(ded.get("Section80C")) + _i(ded.get("Section80CCC"))
                + _i(ded.get("Section80CCDEmployeeOrSE")))
        A(1, "Sum of deductions u/s 80C, 80CCC & 80CCD(1) exceeds Rs. 1,50,000.",
          trio > C.CAP_80C_CCC_CCD1 + 1)

    # -- Rule 2/3: 80CCD(1) percentage cap --
    if opt_out:
        if employer_cat in C.PENSIONER_CATEGORIES or employer_cat == "NA":
            A(2, "Deduction u/s 80CCD(1) exceeds 20% of Gross Total Income.",
              _i(ded.get("Section80CCDEmployeeOrSE")) > int(gti * 0.20) + 1)
        else:
            A(3, "Deduction u/s 80CCD(1) exceeds 10% of salary u/s 17(1).",
              _i(ded.get("Section80CCDEmployeeOrSE")) > int(_i(inc.get("Salary")) * 0.10) + 1)

    # -- Rule 4/120/216: 80CCD(2) cap --
    sal171 = _i(inc.get("Salary"))
    ccd2 = _i(ded.get("Section80CCDEmployer"))
    if opt_out:
        pct = C.CAP_80CCD2_PCT_CG_SG_OLD if employer_cat in ("CGOV", "SGOV") \
            else C.CAP_80CCD2_PCT_OTHERS_OLD
        A(4 if pct == 0.10 else 120,
          f"Deduction u/s 80CCD(2) exceeds {int(pct * 100)}% of salary.",
          ccd2 > int(sal171 * pct) + 1)
    else:
        A(216, "Deduction u/s 80CCD(2) exceeds 14% of salary under new regime.",
          ccd2 > int(sal171 * C.CAP_80CCD2_PCT_NEW) + 1)
    A(116, "80CCD(2) cannot be claimed by pensioners / NA employer category.",
      ccd2 > 0 and (employer_cat in C.PENSIONER_CATEGORIES or employer_cat == "NA"))

    # -- Rule 11-16: 80TTA / 80TTB --
    tta = _i(ded.get("Section80TTA"))
    ttb = _i(ded.get("Section80TTB"))
    if opt_out:
        A(11, "Deduction u/s 80TTA exceeds Rs. 10,000.", tta > C.CAP_80TTA)
        A(13, "80TTA cannot be claimed by a senior citizen.", tta > 0 and senior)
        A(14, "Deduction u/s 80TTB exceeds Rs. 50,000.", ttb > C.CAP_80TTB)
        A(15, "Assessee below 60 years cannot claim 80TTB.", ttb > 0 and not senior)
    savings_interest = sum(_i(r.get("OthSrcOthAmount"))
                           for r in _dig(inc, "OthersInc", "OthersIncDtlsOthSrc", default=[])
                           if r.get("OthSrcNatureDesc") == "SAV")
    A(12, "80TTA is restricted to savings account interest income.",
      tta > savings_interest + 1)

    # -- Rule 17/18: VIA totals --
    via_fields = ["Section80C", "Section80CCC", "Section80CCDEmployeeOrSE",
                  "Section80CCD1B", "Section80CCDEmployer", "Section80D", "Section80DD",
                  "Section80DDB", "Section80E", "Section80EE", "Section80EEA",
                  "Section80EEB", "Section80G", "Section80GG", "Section80GGA",
                  "Section80GGC", "Section80U", "Section80TTA", "Section80TTB",
                  "AnyOthSec80CCH"]
    via_sum = sum(_i(ded.get(f)) for f in via_fields)
    A(17, "Total of Chapter VI-A deductions does not match the sum of individual "
          "deductions (restricted to GTI).",
      not D_close(ded.get("TotalChapVIADeductions"), min(via_sum, gti), tol=10))
    A(18, "Chapter VI-A deductions exceed Gross Total Income.",
      _i(ded.get("TotalChapVIADeductions")) > gti + 10)

    # -- Rule 22 / 160 / 174: GTI composition --
    A(22 if opt_out else (160 if hp < 0 else 174),
      "Gross Total Income is not equal to the total of Salary, House Property & "
      "Other Sources income.",
      not D_close(gti, sal + hp + os_inc))
    if not opt_out:
        A(160, "House property loss has been set off under the new regime.", hp < 0)

    # -- Rule 24: Total income --
    A(24, "Total income should be GTI minus deductions or zero.",
      abs(ti - T_round10(max(0, gti_inc - _i(ded.get("TotalChapVIADeductions"))))) > 10)

    # -- Rule 25-28: tax computation identities --
    A(25, "Tax after rebate should equal tax payable minus rebate u/s 87A.",
      not D_close(tc.get("TaxPayableOnRebate"),
                  _i(tc.get("TotalTaxPayable")) - _i(tc.get("Rebate87A")), tol=2))
    A(26, "Total tax and cess should equal tax after rebate plus H&E cess.",
      not D_close(tc.get("GrossTaxLiability"),
                  _i(tc.get("TaxPayableOnRebate")) + _i(tc.get("EducationCess")), tol=2))
    ip = tc.get("IntrstPay", {})
    fee_sum = (_i(ip.get("IntrstPayUs234A")) + _i(ip.get("IntrstPayUs234B"))
               + _i(ip.get("IntrstPayUs234C")) + _i(ip.get("LateFilingFee234F"))
               + _i(ip.get("FeeFurnish234I")))
    A(28, "Total interest/fee payable should equal the sum of 234A+234B+234C+234F+234I.",
      not D_close(tc.get("TotalIntrstPay"), fee_sum, tol=2))
    A(27, "Total tax, fee & interest should equal (Gross tax liability - relief 89) "
          "plus interest and fee.",
      abs(_i(tc.get("TotTaxPlusIntrstPay"))
          - T_round10(_i(tc.get("GrossTaxLiability")) - _i(tc.get("Section89")) + fee_sum)) > 25)

    # -- Rule 23 / 191 / 192 / 148: rebate 87A --
    rebate = _i(tc.get("Rebate87A"))
    if opt_out:
        A(23, "Rebate u/s 87A claimed while total income exceeds Rs. 5,00,000.",
          rebate > 0 and ti > C.REBATE_87A_OLD_INCOME_LIMIT)
        A(192, "Rebate u/s 87A cannot exceed Rs. 12,500 under the old regime.",
          rebate > C.REBATE_87A_OLD_MAX)
    else:
        A(191, "Rebate u/s 87A claimed while total income (excl. LTCG) exceeds "
               "Rs. 12,70,590.",
          rebate > 0 and (ti - ltcg) > C.REBATE_87A_NEW_MR_CEILING)
        A("191a", "Rebate u/s 87A cannot exceed Rs. 60,000.",
          rebate > C.REBATE_87A_NEW_MAX)

    # -- Rule 29 / 30 / 184: exempt income --
    ex = inc.get("ExemptIncAgriOthUs10", {})
    ex_rows = ex.get("ExemptIncAgriOthUs10Dtls", []) or []
    agri = sum(_i(r.get("OthAmount")) for r in ex_rows if r.get("SubCategory") == "10(1)")
    A(29, "Agricultural income shown as exempt exceeds Rs. 5,000.", agri > C.MAX_AGRICULTURE_INCOME)
    A(30, "Exempt income total should equal the sum of individual amounts.",
      not D_close(ex.get("ExemptIncAgriOthUs10Total"),
                  sum(_i(r.get("OthAmount")) for r in ex_rows)))
    subs = [r.get("SubCategory") for r in ex_rows if r.get("SubCategory")]
    A(184, "A nature of exempt income has been selected more than once.",
      len(subs) != len(set(subs)))

    # -- Rule 43-49 / 240 / 253 / 295-297 / 336: house property --
    props = inc.get("PropertyDetails", []) or []
    for p in props:
        rd = p.get("Rentdetails", {})
        occ = p.get("ifLetOut")
        A(43, "Standard deduction on house property should equal 30% of annual value.",
          not D_close(rd.get("ThirtyPercentOfBalance"),
                      round(_i(rd.get("AnnualOfPropOwned")) * C.HP_STD_DEDUCTION_RATE), tol=2))
        A(46, "Annual value should be gross rent minus unrealised rent and taxes.",
          not D_close(rd.get("BalanceALV"),
                      _i(rd.get("AnnualLetableValue")) - _i(rd.get("TotalUnrealizedAndTax"))))
        A(47, "Income from house property should equal (iii - iv - v + vi).",
          not D_close(rd.get("IncomeOfHP"),
                      _i(rd.get("AnnualOfPropOwned")) - _i(rd.get("TotalDeduct"))
                      + _i(rd.get("ArrearsUnrealizedRentRcvd")), tol=2))
        A("47a", "Total deductions should equal 30% of balance plus interest on borrowed capital.",
          not D_close(rd.get("TotalDeduct"),
                      _i(rd.get("ThirtyPercentOfBalance")) + _i(rd.get("IntOnBorwCap"))))
        if occ == "S":
            if opt_out:
                A(48, "Interest on borrowed capital for self-occupied property exceeds "
                      "Rs. 2,00,000 (old regime).",
                  _i(rd.get("IntOnBorwCap")) > C.HP_SELF_OCCUPIED_INTEREST_CAP_OLD)
            else:
                A(162, "Interest on borrowed capital for self-occupied property is not "
                       "allowed in the new regime.", _i(rd.get("IntOnBorwCap")) > 0)
            A(49, "Municipal tax is not allowed for a self-occupied property.",
              _i(rd.get("LocalTaxes")) > 0)
        if occ in ("L", "D"):
            A(45, "Gross rent should be more than zero for a let-out / deemed let-out property.",
              _i(rd.get("AnnualLetableValue")) <= 0)
        A(336, "Unrealised rent exceeds gross rent.",
          _i(rd.get("RentNotRealized")) > _i(rd.get("AnnualLetableValue")))
        s24 = rd.get("Section24B")
        if s24 is not None and opt_out is False and occ == "S":
            A(253, "Interest u/s 24(b) cannot be claimed for self-occupied property "
                   "under the new regime.", _i(s24.get("TotalInterestUs24B")) > 0)
        if s24 is None and _i(rd.get("IntOnBorwCap")) > 0:
            A(220, "Details of the bank from which the loan is taken must be provided "
                   "in schedule 24(b) for claiming interest on borrowed capital.")
        if s24 is not None:
            A(240, "Interest on borrowed capital differs from schedule 24(b) total.",
              not D_close(rd.get("IntOnBorwCap"),
                          round(_i(s24.get("TotalInterestUs24B"))
                                * _share(p)), tol=2))

    # -- Rule 50-56: other sources --
    os_rows = _dig(inc, "OthersInc", "OthersIncDtlsOthSrc", default=[]) or []
    A(52, "Income from other sources should equal the sum of individual amounts "
          "less deduction u/s 57(iia).",
      not D_close(os_inc, sum(_i(r.get("OthSrcOthAmount")) for r in os_rows)
                  - _i(inc.get("DeductionUs57iia"))))
    for nature, ruleno in (("SAV", 50), ("IFD", 51), ("TAX", 55), ("FAP", 56)):
        count = sum(1 for r in os_rows if r.get("OthSrcNatureDesc") == nature)
        A(ruleno, f"'{C.OTHER_SOURCE_NATURES[nature]}' selected more than once.", count > 1)
    fap = sum(_i(r.get("OthSrcOthAmount")) for r in os_rows
              if r.get("OthSrcNatureDesc") == "FAP")
    cap_57 = C.FAMILY_PENSION_DED_CAP_OLD if opt_out else C.FAMILY_PENSION_DED_CAP_NEW
    A(53 if opt_out else 214,
      f"Deduction u/s 57(iia) exceeds lower of 1/3rd of family pension or Rs. {cap_57}.",
      _i(inc.get("DeductionUs57iia")) > min(fap // 3, cap_57) + 1)

    # -- Rule 57-63, 112, 163, 168, 210, 215, 264: salary --
    A(59, "Gross salary should equal salary 17(1) + perquisites 17(2) + profits 17(3).",
      not D_close(inc.get("GrossSalary"),
                  sal171 + _i(inc.get("PerquisitesValue")) + _i(inc.get("ProfitsInSalary"))))
    total_allw = _i(_dig(inc, "AllwncExemptUs10", "TotalAllwncExemptUs10"))
    A(60, "Net salary should equal gross salary minus exempt allowances u/s 10.",
      not D_close(inc.get("NetSalary"), _i(inc.get("GrossSalary")) - total_allw))
    d16 = (_i(inc.get("DeductionUs16ia")) + _i(inc.get("EntertainmentAlw16ii"))
           + _i(inc.get("ProfessionalTaxUs16iii")))
    A(61, "Deductions u/s 16 should equal 16(ia) + 16(ii) + 16(iii).",
      not D_close(inc.get("DeductionUs16"), d16))
    A(62, "Income from salary should equal net salary minus deductions u/s 16.",
      not D_close(sal, _i(inc.get("NetSalary")) - _i(inc.get("DeductionUs16"))))
    A(63, "Exempt allowances u/s 10 cannot exceed gross salary.",
      total_allw > _i(inc.get("GrossSalary")))
    if opt_out:
        A(112, "Standard deduction u/s 16(ia) exceeds Rs. 50,000 (old regime).",
          _i(inc.get("DeductionUs16ia")) > C.STD_DEDUCTION_16IA_OLD)
        if _i(inc.get("EntertainmentAlw16ii")) > 0 and employer_cat not in ("CGOV", "SGOV", "PSU"):
            A(58, "Entertainment allowance u/s 16(ii) is allowed only to CG/SG/PSU employees.")
    else:
        A(215, "Standard deduction u/s 16(ia) exceeds Rs. 75,000 (new regime).",
          _i(inc.get("DeductionUs16ia")) > C.STD_DEDUCTION_16IA_NEW)
        A(163, "Entertainment allowance u/s 16(ii) is not allowed in the new regime.",
          _i(inc.get("EntertainmentAlw16ii")) > 0)
        A(168, "Professional tax u/s 16(iii) is not allowed in the new regime.",
          _i(inc.get("ProfessionalTaxUs16iii")) > 0)
    A(57, "Entertainment allowance u/s 16(ii) exceeds Rs. 5,000 or 1/5th of salary.",
      _i(inc.get("EntertainmentAlw16ii")) > min(C.ENTERTAINMENT_ALLOW_16II_MAX, sal171 // 5) + 1)
    A(210, "Nature of employment cannot be 'Not Applicable' when salary details are given.",
      _i(inc.get("GrossSalary")) > 0 and employer_cat == "NA")

    # -- exempt allowance regime restrictions (rules 149/150/161/164-167/301) --
    allw_rows = _dig(inc, "AllwncExemptUs10", "AllwncExemptUs10Dtls", default=[]) or []
    for row in allw_rows:
        sec = row.get("SalNatureDesc")
        amt = _i(row.get("SalOthAmount"))
        if not opt_out:
            if sec in ("10(5)", "10(13A)", "10(14)(i)", "10(14)(ii)"):
                A(149, f"Exempt allowance {sec} is more than zero under the new regime.", amt > 0)
            if sec == "10(17)":
                A(161, "Allowance u/s 10(17) (MP/MLA/MLC) is more than zero under new regime.", amt > 0)
            if sec == "10(14)(ii)(115BAC)":
                A(148, "Transport allowance u/s 10(14)(ii)(115BAC) exceeds Rs. 38,400.", amt > 38400)
            if sec == "EIC":
                A(301, "Exempt income of Judge cannot be more than 0 under the new regime.", amt > 0)
        else:
            if sec in ("10(14)(i)(115BAC)", "10(14)(ii)(115BAC)"):
                A(150, f"Exempt allowance {sec} is not available under the old regime.", amt > 0)
    secs = [r.get("SalNatureDesc") for r in allw_rows]
    A(213, "An exempt allowance section is disclosed more than once.",
      len(secs) != len(set(secs)))

    # -- HRA schedule (rules 261-263, 265, 266, 269) --
    hra = itr1.get("ScheduleEA10_13A")
    hra_amt = sum(_i(r.get("SalOthAmount")) for r in allw_rows
                  if r.get("SalNatureDesc") == "10(13A)")
    if hra is not None:
        a = _i(hra.get("ActlHRARecv"))
        b_ = _i(hra.get("ActlRentPaid10Per"))
        c_ = _i(hra.get("Sal40Or50Per"))
        A(263, "HRA exemption should be the least of actual HRA, rent minus 10% of "
               "salary, and 40/50% of salary.",
          _i(hra.get("EligbleExmpAllwncUs13A")) > min(a, b_, c_) + 1)
        A(269, "Exempt allowance u/s 10(13A) in salary does not match schedule 10(13A).",
          not D_close(hra_amt, hra.get("EligbleExmpAllwncUs13A")))
    elif hra_amt > 0:
        A(265, "Schedule 10(13A) must be filled for claiming HRA exemption.")

    # -- Rule 121-124, 169-172, 227-232, 252: loan-linked deductions --
    A(121, "80EE exceeds Rs. 50,000.", _i(ded.get("Section80EE")) > C.CAP_80EE)
    A(122, "80EEA exceeds Rs. 1,50,000.", _i(ded.get("Section80EEA")) > C.CAP_80EEA)
    A(124, "80EEB exceeds Rs. 1,50,000.", _i(ded.get("Section80EEB")) > C.CAP_80EEB)
    A(123, "Only one of 80EE / 80EEA is allowed.",
      _i(ded.get("Section80EE")) > 0 and _i(ded.get("Section80EEA")) > 0)
    eea = itr1.get("Schedule80EEA")
    if eea is not None:
        A(229, "80EEA allowed only for stamp-duty value up to Rs. 45 lakh.",
          _i(eea.get("PropStmpDtyVal")) > C.CAP_80EEA_STAMP_DUTY_VALUE)
        for ln in eea.get("Schedule80EEADtls", []) or []:
            A(230, "80EEA loan sanction date must be between 01-04-2019 and 31-03-2022.",
              not ("2019-04-01" <= str(ln.get("DateofLoan")) <= "2022-03-31"))
    ee = itr1.get("Schedule80EE")
    if ee is not None:
        for ln in ee.get("Schedule80EEDtls", []) or []:
            A(252, "80EE loan sanction date must be between 01-04-2016 and 31-03-2017.",
              not ("2016-04-01" <= str(ln.get("DateofLoan")) <= "2017-03-31"))
            A(227, "80EE admissible only for loans up to Rs. 35 lakh.",
              _i(ln.get("TotalLoanAmt")) > C.CAP_80EE_LOAN_AMOUNT)
    eeb = itr1.get("Schedule80EEB")
    if eeb is not None:
        for ln in eeb.get("Schedule80EEBDtls", []) or []:
            A(232, "80EEB loan sanction date must be between 01-04-2019 and 31-03-2023.",
              not ("2019-04-01" <= str(ln.get("DateofLoan")) <= "2023-03-31"))

    # -- Rule 115 / 169: 80CCD(1B) --
    A(115, "80CCD(1B) exceeds Rs. 50,000.", _i(ded.get("Section80CCD1B")) > C.CAP_80CCD1B)
    A(226, "PRAN is required to claim deduction u/s 80CCD(1)/80CCD(1B).",
      (_i(usr.get("Section80CCDEmployeeOrSE")) > 0 or _i(usr.get("Section80CCD1B")) > 0)
      and not usr.get("PRANDtls"))

    # -- Rule 186/187: 80CCH --
    cch = _i(ded.get("AnyOthSec80CCH"))
    A(186, "Deduction u/s 80CCH exceeds 46.2% of salary u/s 17(1).",
      cch > int(sal171 * C.CAP_80CCH_PCT_OF_SALARY) + 1)
    A(187, "80CCH can be claimed only by Central Government employees (Agnipath).",
      cch > 0 and employer_cat != "CGOV")

    # -- Rule 200-209: 80DD / 80U --
    dd = itr1.get("Schedule80DD")
    if dd is not None and opt_out:
        expected = C.CAP_80DD_SEVERE if dd.get("NatureOfDisability") == "2" else C.CAP_80DD
        A(203, "80DD amount should be Rs. 75,000 / Rs. 1,25,000 as per disability.",
          not D_close(dd.get("DeductionAmount"), min(expected, max(gti, 0))))
    uu = itr1.get("Schedule80U")
    if uu is not None and opt_out:
        expected = C.CAP_80U_SEVERE if uu.get("NatureOfDisability") == "2" else C.CAP_80U
        A(200, "80U amount should be Rs. 75,000 / Rs. 1,25,000 as per disability.",
          not D_close(uu.get("DeductionAmount"), min(expected, max(gti, 0))))
    if not opt_out:
        for rule_no, f in ((154, "Section80DD"), (155, "Section80DDB"),
                           (159, "Section80U")):
            A(rule_no, f"{f} cannot be claimed under the new regime.", _i(ded.get(f)) > 0)

    # -- Rule 146 / 153-175: new regime restrictions on VIA --
    if not opt_out:
        for f in via_fields:
            if f in ("Section80CCDEmployer", "AnyOthSec80CCH"):
                continue
            A(146, f"Deduction at {f} cannot be more than zero under the new regime.",
              _i(ded.get(f)) > 0)
        for rule_no, sched in ((156, "Schedule80G"), (173, "Schedule80D"),
                               (175, "Schedule80GGA")):
            if sched in itr1:
                node = itr1[sched]
                nonzero = any(isinstance(v, int) and v > 0 for v in node.values())
                A(rule_no, f"{sched} is more than zero / provided under the new regime.",
                  nonzero)
        A(255, "Details in 80C/10(13A)/80E/80EE/80EEA/80EEB schedules are not allowed "
               "under the new regime.",
          any(k in itr1 for k in ("Schedule80C", "ScheduleEA10_13A", "Schedule80E",
                                  "Schedule80EE", "Schedule80EEA", "Schedule80EEB")))

    # -- Rule 5/7: 80DDB caps --
    if opt_out:
        ddb_type = _dig(usr, "Section80DDBUsrType", default="1")
        cap = C.CAP_80DDB_SR if ddb_type == "2" else C.CAP_80DDB_NORMAL
        A(5 if ddb_type == "2" else 7,
          f"80DDB exceeds Rs. {cap:,} for the chosen category.",
          _i(ded.get("Section80DDB")) > cap)
        A(6, "Eligible category not provided for 80DDB claim.",
          _i(usr.get("Section80DDB")) > 0 and not _dig(usr, "Section80DDBUsrType", default=None))
        A(239, "Specified disease is required for an 80DDB claim.",
          _i(usr.get("Section80DDB")) > 0 and not _dig(usr, "NameOfSpecDisease80DDB", default=None))

    # -- Rule 8-9 / 88 / 139: 80G --
    g80 = itr1.get("Schedule80G")
    if opt_out and _i(ded.get("Section80G")) > 0 and g80 is None:
        A(8, "80G claimed but details not provided in Schedule 80G.")
    if g80 is not None:
        A(10, "80G claimed in VIA exceeds the eligible amount in Schedule 80G.",
          _i(ded.get("Section80G")) > _i(g80.get("TotalEligibleDonationsUs80G")) + 1)
        A(139, "Eligible donations cannot exceed total donations.",
          _i(g80.get("TotalEligibleDonationsUs80G")) > _i(g80.get("TotalDonationsUs80G")) + 1)
        for blk in ("Don100Percent", "Don50PercentNoApprReqd",
                    "Don100PercentApprReqd", "Don50PercentApprReqd"):
            for row in (g80.get(blk) or {}).get("DoneeWithPan", []) or []:
                A(88, "Cash donation above Rs. 2,000 to a donee is included in eligible "
                      "donation (80G).",
                  _i(row.get("DonationAmtCash")) > C.CASH_DONATION_LIMIT
                  and _i(row.get("EligibleDonationAmt")) > _i(row.get("DonationAmtOtherMode")))
                A(78, "Donee PAN is the same as the assessee PAN.",
                  row.get("DoneePAN") == _dig(pi, "PAN", default=""))

    # -- Rule 89-94 / 118 / 143 / 144: 80GGA --
    gga = itr1.get("Schedule80GGA")
    if gga is not None:
        A(93, "80GGA in VIA exceeds eligible amount in Schedule 80GGA.",
          _i(ded.get("Section80GGA")) > _i(gga.get("TotalEligibleDonationAmt80GGA")) + 1)
        A(92, "80GGA eligible amount exceeds total donations.",
          _i(gga.get("TotalEligibleDonationAmt80GGA")) > _i(gga.get("TotalDonationsUs80GGA")) + 1)
        for row in gga.get("DonationDtlsSciRsrchRuralDev", []) or []:
            A(143, "80GGA not allowed for cash donation above Rs. 2,000.",
              _i(row.get("DonationAmtCash")) > C.CASH_DONATION_LIMIT
              and _i(row.get("EligibleDonationAmt")) > _i(row.get("DonationAmtOtherMode")))

    # -- Rule 193-199 / 329-330: 80GGC --
    ggc = itr1.get("Schedule80GGC")
    if opt_out and _i(ded.get("Section80GGC")) > 0 and ggc is None:
        A(193, "80GGC claimed but details not provided in Schedule 80GGC.")
    if ggc is not None:
        A(194, "80GGC eligible amount must equal contribution in other mode.",
          not D_close(ggc.get("TotalEligibleDonationAmt80GGC"),
                      ggc.get("TotalDonationAmtOtherMode80GGC")))
        for row in ggc.get("Schedule80GGCDetails", []) or []:
            A(329, "Name and PAN of the political party are required for 80GGC.",
              not row.get("PoliticalPartyName") or not row.get("PoliticalPartyPAN"))
            A(330, "Either cash or other-mode donation must be entered in each 80GGC row.",
              _i(row.get("DonationAmtCash")) == 0 and _i(row.get("DonationAmtOtherMode")) == 0)
            A(198, "Date of contribution is mandatory for 80GGC.", not row.get("DonationDate"))

    # -- Rule 220/224/233/241: supporting schedules mandatory --
    c80_rows = _dig(itr1, "Schedule80C", "Schedule80CDtls", default=None)
    if opt_out and _i(usr.get("Section80C")) > 0 and not c80_rows:
        A(224, "Identification number of supporting documents is required in schedule "
               "80C to claim the deduction.")
    if c80_rows:
        A(241, "Deduction u/s 80C should equal the total of payments in schedule 80C.",
          not D_close(_dig(itr1, "Schedule80C", "TotalAmt"), 
                      _i(usr.get("Section80C"))))
        A(247, "Schedule 80C: sum of individual rows must equal the total.",
          not D_close(_dig(itr1, "Schedule80C", "TotalAmt"),
                      sum(_i(r.get("Amount")) for r in c80_rows)))
    if opt_out and _i(usr.get("Section80GG")) > 0:
        A(233, "Acknowledgement of Form 10BA is required to claim deduction u/s 80GG.",
          not _dig(usr, "Form10BAAckNum", default=None))

    # -- Rule 127-137: 80D --
    d80 = _dig(itr1, "Schedule80D", "Sec80DSelfFamSrCtznHealth", default=None)
    if d80 is not None:
        A(127, "80D 'Self and Family' exceeds Rs. 25,000.",
          _i(d80.get("SelfAndFamily")) > C.CAP_80D_SELF_FAMILY)
        A(130, "80D 'Self and Family (Senior Citizen)' exceeds Rs. 50,000.",
          _i(d80.get("SelfAndFamilySeniorCitizen")) > C.CAP_80D_SELF_FAMILY_SR)
        A(132, "80D 'Parents' exceeds Rs. 25,000.", _i(d80.get("Parents")) > C.CAP_80D_PARENTS)
        A(134, "80D 'Parents (Senior Citizen)' exceeds Rs. 50,000.",
          _i(d80.get("ParentsSeniorCitizen")) > C.CAP_80D_PARENTS_SR)
        prev = (_i(d80.get("PrevHlthChckUpSlfFam")) + _i(d80.get("PrevHlthChckUpSlfFamSrCtzn"))
                + _i(d80.get("PrevHlthChckUpParents")) + _i(d80.get("PrevHlthChckUpParentsSrCtzn")))
        A(129, "Preventive health check-up exceeds Rs. 5,000 combined.", prev > C.CAP_80D_PREV_CHECKUP)
        A(136, "80D eligible amount exceeds Rs. 1,00,000.",
          _i(d80.get("EligibleAmountOfDedn")) > C.CAP_80D_TOTAL)
        A(178, "80D 1a claimed while the senior-citizen flag is not 'No'.",
          _i(d80.get("SelfAndFamily")) > 0 and d80.get("SeniorCitizenFlag") != "N")
        A(179, "80D 1b claimed while the senior-citizen flag is not 'Yes'.",
          _i(d80.get("SelfAndFamilySeniorCitizen")) > 0 and d80.get("SeniorCitizenFlag") != "Y")
        A(180, "80D 2a claimed while the parents flag is not 'No'.",
          _i(d80.get("Parents")) > 0 and d80.get("ParentsSeniorCitizenFlag") != "N")
        A(181, "80D 2b claimed while the parents flag is not 'Yes'.",
          _i(d80.get("ParentsSeniorCitizen")) > 0 and d80.get("ParentsSeniorCitizenFlag") != "Y")
        # rules 234-238 / 256-259: insurer name & policy for each premium claimed
        ins_bands = (
            ("HealthInsPremSlfFam", "Sec80DSelfFamHIDtls", 256),
            ("HlthInsPremSlfFamSrCtzn", "Sec80DSelfFamSrCtznHIDtls", 257),
            ("HlthInsPremParents", "Sec80DParentsHIDtls", 258),
            ("HlthInsPremParentsSrCtzn", "Sec80DParentsSrCtznHIDtls", 259),
        )
        for prem_key, dtl_key, rule_no in ins_bands:
            prem = _i(d80.get(prem_key))
            if prem <= 0:
                continue
            rows = _dig(d80, dtl_key, "Sch80DInsDtls", default=[]) or []
            A(rule_no, "Insurer name and policy number are required in schedule 80D "
                       "for the health-insurance premium claimed.", not rows)
            if rows:
                tot = sum(_i(r.get("HealthInsAmt")) for r in rows)
                A("234-237", "Schedule 80D: breakup of premium must equal the premium "
                             "claimed in the corresponding row.", not D_close(tot, prem))

    # -- Rule 95-111: taxes paid schedules --
    tds1 = itr1.get("TDSonSalaries", {})
    tds2 = itr1.get("TDSonOthThanSals", {})
    tds3 = itr1.get("ScheduleTDS3Dtls", {})
    tcs = itr1.get("ScheduleTCS", {})
    it = itr1.get("TaxPayments", {})
    A(100, "TDS1 total is not the sum of individual rows.",
      not D_close(tds1.get("TotalTDSonSalaries"),
                  sum(_i(r.get("TotalTDSSal")) for r in tds1.get("TDSonSalary", []) or [])))
    A(101, "TDS2 claimed total is not the sum of individual rows.",
      not D_close(tds2.get("TotalTDSonOthThanSals"),
                  sum(_i(r.get("ClaimOutOfTotTDSOnAmtPaid")) for r in tds2.get("TDSonOthThanSal", []) or [])))
    for r in tds2.get("TDSonOthThanSal", []) or []:
        A(98, "TDS claimed this year exceeds tax deducted (TDS2).",
          _i(r.get("ClaimOutOfTotTDSOnAmtPaid")) > _i(r.get("TotTDSOnAmtPaid")) + 1)
    A(102, "TDS3 claimed total is not the sum of individual rows.",
      not D_close(tds3.get("TotalTDS3Details"),
                  sum(_i(r.get("TDSClaimed")) for r in tds3.get("TDS3Details", []) or [])))
    A(97, "TCS claimed total is not the sum of individual rows.",
      not D_close(tcs.get("TotalSchTCS"),
                  sum(_i(r.get("AmtTCSClaimedThisYear")) for r in tcs.get("TCS", []) or [])))
    for r in tcs.get("TCS", []) or []:
        A(96, "TCS claimed this year exceeds tax collected.",
          _i(r.get("AmtTCSClaimedThisYear")) > _i(r.get("TotalTCS")) + 1)
    A(95, "Schedule IT total tax paid is not the sum of individual values.",
      not D_close(it.get("TotalTaxPayments"),
                  sum(_i(r.get("Amt")) for r in it.get("TaxPayment", []) or [])))
    tds_total = (_i(tds1.get("TotalTDSonSalaries")) + _i(tds2.get("TotalTDSonOthThanSals"))
                 + _i(tds3.get("TotalTDS3Details")))
    A(108, "Total TDS claimed is not the sum of TDS1, TDS2 and TDS3.",
      not D_close(paid.get("TDS"), tds_total))
    A(109, "Total TCS claimed is not equal to the sum in the TCS schedule.",
      not D_close(paid.get("TCS"), tcs.get("TotalSchTCS")))
    A(104, "Total taxes paid should equal TDS + TCS + advance tax + self-assessment tax.",
      not D_close(paid.get("TotalTaxesPaid"),
                  _i(paid.get("TDS")) + _i(paid.get("TCS")) + _i(paid.get("AdvanceTax"))
                  + _i(paid.get("SelfAssessmentTax"))))

    # advance vs self-assessment split by deposit date (rules 110/111)
    adv = sum(_i(r.get("Amt")) for r in it.get("TaxPayment", []) or []
              if "2025-04-01" <= str(r.get("DateDep")) <= "2026-03-31")
    sat = sum(_i(r.get("Amt")) for r in it.get("TaxPayment", []) or []
              if str(r.get("DateDep")) > "2026-03-31")
    A(110, "Advance tax is not equal to challans deposited within the previous year.",
      not D_close(paid.get("AdvanceTax"), adv))
    A(111, "Self-assessment tax is not equal to challans deposited after 31-03-2026.",
      not D_close(paid.get("SelfAssessmentTax"), sat))

    # -- Rule 105/106: refund / balance --
    tot_tax_int = _i(tc.get("TotTaxPlusIntrstPay"))
    A(105, "Refund claimed should equal 'Total taxes paid' minus 'Total tax and "
           "interest payable'.",
      not D_close(refund.get("RefundDue"), max(0, _i(paid.get("TotalTaxesPaid")) - tot_tax_int), tol=10))
    A(106, "Tax payable should equal 'Total tax and interest payable' minus 'Total "
           "taxes paid'.",
      not D_close(_dig(itr1, "TaxPaid", "BalTaxPayable"),
                  max(0, tot_tax_int - _i(paid.get("TotalTaxesPaid"))), tol=10))

    # -- Rule 117 / 217 / 218 / 292: amount ceilings --
    A(117, "Total income excluding LTCG exceeds Rs. 50 lakh; ITR-1 cannot be used.",
      ti - ltcg > C.MAX_INCOME_EXCL_LTCG)
    A(217, "LTCG u/s 112A exceeds Rs. 1,25,000; ITR-1 cannot be used.",
      ltcg > C.MAX_LTCG_112A)
    A(218, "LTCG u/s 112A should equal sale consideration minus cost of acquisition.",
      not D_close(ltcg, _i(_dig(doc, "ITR", "ITR1", "LTCG112A", "TotSaleCnsdrn"))
                  - _i(_dig(doc, "ITR", "ITR1", "LTCG112A", "TotCstAcqisn"))))
    A(292, "LTCG u/s 112A should equal (GTI incl. LTCG) minus (GTI excl. LTCG).",
      not D_close(ltcg, gti_inc - gti))

    # -- Rule 125 / D1: relief 89 --
    if _i(tc.get("Section89")) > 0:
        sal_parts = sal171 + _i(inc.get("PerquisitesValue")) + _i(inc.get("ProfitsInSalary")) + fap
        A(125, "Relief u/s 89 requires salary or family pension income.", sal_parts <= 0)
        D(1, "Relief u/s 89(1) claimed - ensure Form 10E has been furnished, otherwise "
             "the claim may not be entertained.")

    # -- Category B: TDS sections not permitted in ITR-1 (rules B3-B8) --
    special = {"94B", "4BB", "94BA", "4IA", "4IC", "4LA", "94S", "94E", "4LB", "4LC1",
               "4LC2", "4LC3", "4BA1", "4BA2", "LBA1", "LBA2", "LBA3", "195", "96A",
               "96B", "96C", "96D", "96DA", "94Q", "94C", "94R"}
    for sched_key, rows_key in (("TDSonOthThanSals", "TDSonOthThanSal"),
                                ("ScheduleTDS3Dtls", "TDS3Details")):
        for r in itr1.get(sched_key, {}).get(rows_key, []) or []:
            sec = r.get("TDSSection")
            if sec in special:
                B("B3-B8", f"TDS section {sec} indicates income for which ITR-1 is not "
                           f"applicable.")
            if sched_key != "TDSonSalaries" and sec in ("92A", "92B", "92C"):
                A(260, "Section 192 (salary TDS) selected under TDS schedule for other "
                       "than salary.")
    B("B9", "TDS deducted on salary exceeds total gross salary.",
      _i(tds1.get("TotalTDSonSalaries")) > _i(inc.get("GrossSalary")))

    # -- Rule 151 / 189 / 190: regime vs filing section --
    if opt_out and _i(fs.get("ReturnFileSec")) == 12:
        A(151, "Old tax regime cannot be selected when filing after the due date "
               "u/s 139(1).")
    if opt_out and _i(fs.get("ReturnFileSec")) == 17:
        A(189, "Old regime cannot be selected in a revised return when the original "
               "was filed u/s 139(4).")

    # -- Rule 20/21: disclosure consistency --
    A(21, "Income details / tax computation are required when taxes paid are disclosed.",
      _i(paid.get("TotalTaxesPaid")) > 0 and _i(inc.get("GrossTotIncomeIncLTCG112A")) == 0
      and _i(tc.get("TotTaxPlusIntrstPay")) == 0)

    # -- Rule 234-I fee revised return (324/328) --
    if _i(fs.get("ReturnFileSec")) == 17:
        expected = C.FEE_234I_SMALL if ti <= 500_000 else C.FEE_234I_FULL
        A("324/328", "Fee u/s 234-I for a revised return filed after 31-12-2026 must be "
                     f"Rs. {expected:,}.",
          str(_dig(doc, "ITR", "ITR1", "CreationInfo", "JSONCreationDate", default="")) >
          C.REVISED_RETURN_FEE_CUTOFF
          and _i(ip.get("FeeFurnish234I")) != expected)

    # -- verification --
    if _dig(ver, "Capacity") == "R" and _dig(fs, "AsseseeRepFlg", default="N") != "Y":
        A(294, "Return signed in representative capacity but representative details "
               "are not provided.")

    return issues


def _share(prop: dict) -> float:
    """Assessee's fractional share of a co-owned property."""
    try:
        return float(prop.get("AsseseeShareProperty", 100.0)) / 100.0
    except (TypeError, ValueError):
        return 1.0


def T_round10(v):
    return int(round(float(v) / 10.0)) * 10


# ---------------------------------------------------------------------------
# Layer 2 (multi-form) — generic CBDT rules for ITR-2 / ITR-4
# ---------------------------------------------------------------------------

def detect_form_in_doc(doc: dict) -> str | None:
    """Read the form name out of a built JSON (``ITR.{ITR1,ITR2,ITR4}``)."""
    itr = doc.get("ITR") or {}
    for key in itr:
        k = str(key).upper()
        if k in ("ITR1", "ITR2", "ITR3", "ITR4"):
            return k
    return None


def validate_rules_generic(doc: dict, form: str) -> list[Issue]:
    """Cross-form CBDT checks (identity, verification, refund banking,
    deduction ceilings, tax-paid arithmetic and the balance equation) for the
    forms whose rule set is not yet transcribed rule-by-rule.  Everything here
    comes from the CBDT validation-rule PDFs that apply identically across
    forms."""
    issues: list[Issue] = []

    def A(rule, msg, cond=True):
        if cond:
            issues.append(Issue("A", f"Rule {rule}", msg))

    def B(rule, msg, cond=True):
        if cond:
            issues.append(Issue("B", f"Rule {rule}", msg))

    node = _dig(doc, "ITR", form, default={})
    if form == "ITR2":
        pi = _dig(node, "PartA_GEN1", "PersonalInfo", default={})
        fs = _dig(node, "PartA_GEN1", "FilingStatus", default={})
        ti_root = node.get("PartB-TI", {}) or {}
        tti = node.get("PartB_TTI", {}) or {}
        comp = tti.get("ComputationOfTaxLiability", {}) or {}
        paid = _dig(tti, "TaxPaid", "TaxesPaid", default={})
        bal_pay = _i(_dig(tti, "TaxPaid", "BalTaxPayable"))
        refund = tti.get("Refund", {}) or {}
        ver = node.get("Verification", {}) or {}
        usr_via = _dig(node, "ScheduleVIA", "UsrDeductUndChapVIA", default={})
        gti = _i(ti_root.get("GrossTotalIncome"))
        ti = _i(ti_root.get("TotalIncome"))
        ded_total = _i(ti_root.get("DeductionsUnderScheduleVIA"))
        liability = _i(comp.get("AggregateTaxInterestLiability"))
        rebate = _i(comp.get("Rebate87A"))
    else:  # ITR-4
        pi = node.get("PersonalInfo", {}) or {}
        fs = node.get("FilingStatus", {}) or {}
        inc = node.get("IncomeDeductions", {}) or {}
        comp = node.get("TaxComputation", {}) or {}
        paid = _dig(node, "TaxPaid", "TaxesPaid", default={})
        bal_pay = _i(_dig(node, "TaxPaid", "BalTaxPayable"))
        refund = node.get("Refund", {}) or {}
        ver = node.get("Verification", {}) or {}
        usr_via = inc.get("UsrDeductUndChapVIA", {}) or {}
        gti = _i(inc.get("GrossTotIncome"))
        ti = _i(inc.get("TotalIncome"))
        ded_total = _i(_dig(inc, "DeductUndChapVIA", "TotalChapVIADeductions"))
        liability = _i(comp.get("TotTaxPlusIntrstPay"))
        rebate = _i(comp.get("Rebate87A"))

    # -- identity (category A: return is invalid if wrong) --
    pan = str(pi.get("PAN", "") or "")
    A("MPAN", "PAN must match the pattern AAAAA9999A.",
      not re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", pan))
    dob = str(pi.get("DOB", "") or "")
    A("MDOB", "Date of birth must be in YYYY-MM-DD format.",
      not re.match(r"^\d{4}-\d{2}-\d{2}$", dob))
    A("MNAME", "Assessee's first name is mandatory.",
      not str(_dig(pi, "AssesseeName", "FirstName", default="")).strip())
    addr = pi.get("Address", {}) or {}
    A("MPIN", "PIN code must be a 6-digit number.",
      not re.match(r"^\d{6}$", str(addr.get("PinCode", "") or "")))
    A("MMOB", "Mobile number must be a 10-digit number.",
      not re.match(r"^\d{10}$", str(addr.get("MobileNo", "") or "")))
    A("MEMAIL", "Email address is mandatory.",
      "@" not in str(addr.get("EmailAddress", "") or ""))

    # -- verification --
    decl = ver.get("Declaration", {}) or {}
    A("MVER", "Verification: declarant name is mandatory.",
      not str(decl.get("AssesseeVerName", "") or "").strip())
    A("MPLACE", "Verification: place of filing is mandatory.",
      not str(ver.get("Place", "") or "").strip())
    if str(ver.get("Capacity", "S")) == "S":
        A("VPAN", "Verification PAN must match the assessee's PAN (self capacity).",
          str(decl.get("AssesseeVerPAN", "") or "") != pan)

    # -- refund banking --
    refund_due = _i(refund.get("RefundDue"))
    banks = _dig(refund, "BankAccountDtls", "AddtnlBankDetails", default=[]) or []
    if refund_due > 0:
        A("MBANK", "Refund is due but no validated bank account is given.",
          not banks)
        if banks:
            b0 = banks[0] or {}
            A("MIFSC", f"IFSC '{b0.get('IFSCCode', '')}' is invalid.",
              not re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", str(b0.get("IFSCCode", "") or "")))

    # -- deduction ceilings (sec 80CCE + 80TTA/80TTB) --
    cce = sum(_i(usr_via.get(k)) for k in
              ("Section80C", "Section80CCC", "Section80CCDEmployeeOrSE"))
    A("80CCE", f"80C + 80CCC + 80CCD(1) = Rs. {cce:,} exceeds the Rs. 1,50,000 "
               "80CCE ceiling.", cce > 150000)
    A("80TTA", "Deduction u/s 80TTA cannot exceed Rs. 10,000.",
      _i(usr_via.get("Section80TTA")) > 10000)
    senior = dob <= C.SENIOR_CITIZEN_DOB_ON_OR_BEFORE
    if _i(usr_via.get("Section80TTB")) > 0:
        A("80TTBS", "80TTB (Rs. 50,000) is only for resident senior citizens; "
                    "others must use 80TTA.", not senior)
        A("80TTB", "Deduction u/s 80TTB cannot exceed Rs. 50,000.",
          _i(usr_via.get("Section80TTB")) > 50000)
    A("VIAGTI", f"Chapter VI-A deductions (Rs. {ded_total:,}) exceed the gross "
                f"total income (Rs. {gti:,}).", ded_total > gti)
    B("TIGEQ", "Total income cannot be negative in a normal return.", ti < 0)

    # -- rebate 87A guard rails --
    old_regime = str(_dig(fs, "OptOutNewTaxRegime", default="N")) == "Y"
    if old_regime:
        A("87A", "Rebate u/s 87A (old regime) needs total income ≤ Rs. 5,00,000 "
                 "and is capped at Rs. 12,500.",
          rebate > 0 and (ti > 500000 or rebate > 12500))
    else:
        A("87A", "Rebate u/s 87A (new regime) needs total income ≤ Rs. 12,00,000 "
                 "and is capped at Rs. 60,000.",
          rebate > 0 and (ti > 1200000 or rebate > 60000))

    # -- tax-paid arithmetic + balance equation --
    total_paid = _i(paid.get("TotalTaxesPaid"))
    parts = sum(_i(paid.get(k)) for k in ("AdvanceTax", "TDS", "TCS", "SelfAssessmentTax"))
    A("TPAID", f"Taxes paid break-up (Rs. {parts:,}) does not tally with "
               f"TotalTaxesPaid (Rs. {total_paid:,}).",
      "TotalTaxesPaid" in paid and parts != total_paid)
    expected = liability - total_paid
    A("BAL", f"Balance equation broken: liability Rs. {liability:,} − taxes paid "
             f"Rs. {total_paid:,} ≠ BalTaxPayable Rs. {bal_pay:,} − RefundDue "
             f"Rs. {refund_due:,}.",
      bal_pay - refund_due != expected)
    A("BALSIGN", "RefundDue must be zero when tax is still payable "
                 "(and vice-versa).", bal_pay > 0 and refund_due > 0)
    return issues


# ---------------------------------------------------------------------------
# Combined entry point
# ---------------------------------------------------------------------------

def validate_all(doc: dict, *, form: str | None = None, include_schema: bool = True,
                 schema: dict | None = None, rules: bool = True) -> list[Issue]:
    """Validate a built return JSON.

    ``form`` selects the official schema ("ITR1" default; "ITR2"/"ITR4" routed
    through :mod:`itr1.schemakit`) and the matching rule set.  When omitted the
    form is read from the document itself (``ITR.<form>``), falling back to
    ITR-1.
    """
    form = (form or detect_form_in_doc(doc) or "ITR1").upper()
    issues: list[Issue] = []
    if include_schema:
        if form == "ITR1":
            issues.extend(validate_schema(doc, schema))
        else:
            from .schemakit import validate_with_schema
            for msg in validate_with_schema(form, doc):
                issues.append(Issue("SCHEMA", "schema", msg))
    if rules:
        if form == "ITR1":
            issues.extend(validate_rules(doc))
        elif form in ("ITR2", "ITR4"):
            issues.extend(validate_rules_generic(doc, form))
        else:
            issues.append(Issue("D", "Rule FORM",
                                f"Rule checks for {form} are not implemented yet."))
    return issues


def summarize(issues: list[Issue]) -> dict:
    counts = {"SCHEMA": 0, "A": 0, "B": 0, "D": 0}
    for i in issues:
        counts[i.category] = counts.get(i.category, 0) + 1
    counts["ok"] = counts["SCHEMA"] == 0 and counts["A"] == 0
    return counts
