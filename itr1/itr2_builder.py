"""ITR-2 JSON builder (AY 2026-27, official schema V1.1).

Builds a schema-valid ``{"ITR": {"ITR2": ...}}`` document from the same
friendly data model as ITR-1, reusing the shared income-head arithmetic
(salary / house property / other sources / Chapter VI-A / tax payments) and
adding the schedules that make a return ITR-2:

* ScheduleS — per-employer salary with nature-wise breakup,
* ScheduleHP — up to two house properties,
* ScheduleCGFor23 + Schedule112A — every capital-gains class (111A / 112A /
  other STCG / land & building with the indexation choice / 54-54F
  exemptions / loss set-off notes),
* ScheduleOS — other-sources gross columns,
* ScheduleCYLA / ScheduleBFLA — set-off statements (zero-loss here),
* ScheduleVIA + deduction schedules, ScheduleEI, ScheduleSI,
* PartB-TI / PartB_TTI, bank details, tax-payment schedules.

The builder validates against the official JSON schema via
``schemakit.validate_with_schema`` in the test-suite.
"""

from __future__ import annotations

import copy

from . import (SW_CREATED_BY, UTILITY_SW_VERSION_BY_FORM,
               UTILITY_FORM_DESCRIPTION)
from . import constants as C
from . import tax_engine as T
from . import salaryx as SX
from . import capgains as CG
from . import special_rates as SR
from .builder import (_build_salary, _build_house_property, _build_other_sources,
                      _build_exempt_income, _build_vi_a, _build_tax_paid,
                      _g, _s, TODAY)
from .errors import field_scope, reset_scopes
from .schemakit import load_schema

D2 = load_schema("ITR2")["definitions"]


def _resolve(node):
    while isinstance(node, dict) and "$ref" in node:
        node = D2[node["$ref"].split("/")[-1]]
    return node if isinstance(node, dict) else {}


_ITR2_PROPS = set((_resolve(D2["ITR"]["properties"]["ITR2"]).get("properties") or {}))

# salary nature-of-income codes for ScheduleS nature-wise rows
NAT_SAL = {"BASIC": "1", "DA": "2", "CONVEY": "3", "HRA": "4", "LTA": "5",
           "CEA": "6", "BONUS": "7", "COMMISSION": "7", "SPECIAL": "7",
           "CCA": "7", "OT": "7", "MEDICAL": "7", "PHONE": "7",
           "HOSTEL": "7", "UNIFORM": "7", "CARMNT": "7", "ARREARS": "15",
           "OTHR": "OTH"}
TERMINAL_CODES = (("Gratuity", "13"), ("Leave encashment", "16"),
                  ("Commuted pension", "12"), ("VRS", "7"),
                  ("Retrenchment", "7"))


def _creation_block(creation_info) -> dict:
    ci = dict(creation_info or {})
    return {
        "SWVersionNo": str(ci.get("sw_version", UTILITY_SW_VERSION_BY_FORM["ITR2"])),
        "SWCreatedBy": ci.get("sw_created_by", SW_CREATED_BY),
        "JSONCreatedBy": ci.get("json_created_by", ci.get("sw_created_by", SW_CREATED_BY)),
        "JSONCreationDate": str(ci.get("json_creation_date", ci.get("creation_date", TODAY()))),
        "IntermediaryCity": str(ci.get("intermediary_city", "Delhi")),
        "Digest": str(ci.get("digest", "-")),
    }


# ---------------------------------------------------------------------------
# PartA_GEN1
# ---------------------------------------------------------------------------

def _personal_info(data) -> dict:
    personal = data.get("personal", {}) or {}
    addr_in = personal.get("address", {}) or {}
    addr = {
        "ResidenceNo": _s(addr_in, "residence_no"),
        "ResidenceName": _s(addr_in, "residence_name"),
        "RoadOrStreet": _s(addr_in, "road_or_street"),
        "LocalityOrArea": _s(addr_in, "locality_or_area"),
        "CityOrTownOrDistrict": _s(addr_in, "city_or_town_or_district"),
        "StateCode": _s(addr_in, "state_code"),
        "CountryCode": _s(addr_in, "country_code", "91"),
        "CountryCodeMobile": _g(addr_in, "country_code_mobile", 91),
        "MobileNo": _g(addr_in, "mobile_no"),
        "EmailAddress": _s(addr_in, "email_address"),
    }
    if _g(addr_in, "mobile_no_sec"):
        addr["CountryCodeMobileNoSec"] = _g(
            addr_in, "country_code_mobile_sec", 91)
        addr["MobileNoSec"] = _g(addr_in, "mobile_no_sec")
    if _s(addr_in, "email_address_sec"):
        addr["EmailAddressSec"] = _s(addr_in, "email_address_sec")
    pin = _g(addr_in, "pin_code")
    if pin:
        addr["PinCode"] = pin
    pi = {
        "AssesseeName": {
            "FirstName": _s(personal, "first_name"),
            "MiddleName": _s(personal, "middle_name"),
            "SurNameOrOrgName": _s(personal, "last_name") or _s(personal, "first_name"),
        },
        "PAN": _s(personal, "pan"),
        "Address": addr,
        "DOB": _s(personal, "dob") or "1990-01-01",
        "Status": _s(personal, "status") or "I",
    }
    # secondary address (utility: 'Is the secondary address same as primary?')
    secondary_same = personal.get("secondary_address_same", True)
    pi["SecondaryAdd"] = "Y" if secondary_same else "N"
    if not secondary_same:
        alt_in = personal.get("alternate_address", {}) or {}
        alt = {
            "ResidenceNo": _s(alt_in, "residence_no"),
            "ResidenceName": _s(alt_in, "residence_name"),
            "RoadOrStreet": _s(alt_in, "road_or_street"),
            "LocalityOrArea": _s(alt_in, "locality_or_area"),
            "CityOrTownOrDistrict": _s(alt_in, "city_or_town_or_district"),
            "StateCode": _s(alt_in, "state_code"),
            "CountryCode": _s(alt_in, "country_code", "91"),
        }
        alt = {k: v for k, v in alt.items() if v not in ("", None)}
        alt["PinCode"] = _g(alt_in, "pin_code")
        pi["AlternateAddress"] = alt
    if not pi["AssesseeName"]["MiddleName"]:
        pi["AssesseeName"].pop("MiddleName")
    if _s(personal, "aadhaar"):
        pi["AadhaarCardNo"] = _s(personal, "aadhaar")
    return pi


def _residential_status(data) -> str:
    """Normalised residential status (utility dropdown
    ``sheet1.ResidentialStatus1``): RES / NRI / NOR."""
    filing = data.get("filing", {}) or {}
    return (_s((data.get("personal") or {}), "residential_status")
            or _s(filing, "residential_status") or "RES").upper()[:3]


def _filing_status(data) -> dict:
    filing = data.get("filing", {}) or {}
    return_section = _g(filing, "return_section", 11)
    fs = {
        "ReturnFileSec": return_section,
        "OptOutNewTaxRegime": "Y" if filing.get("opt_out_new_regime") else "N",
        "SeventhProvisio139": "Y" if filing.get("seventh_provisio") else "N",
        # utility dropdown sheet1.ResidentialStatus1: RES / NRI / NOR
        "ResidentialStatus": _residential_status(data),
        "HeldUnlistedEqShrPrYrFlg": "N",
        "FiiFpiFlag": "N",
        "ItrFilingDueDate": C.DUE_DATE_FILING,
    }
    if filing.get("seventh_provisio"):
        fs["DepAmtAggAmtExcd1CrPrYrFlg"] = "Y" if filing.get("deposit_exceeds_1cr") else "N"
        fs["IncrExpAggAmt2LkTrvFrgnCntryFlg"] = "Y" if filing.get("foreign_travel_exceeds_2l") else "N"
        fs["IncrExpAggAmt1LkElctrctyPrYrFlg"] = "Y" if filing.get("electricity_exceeds_1l") else "N"
    if return_section == 17:
        fs["ReceiptNo"] = _s(filing, "original_receipt_no")
        fs["OrigRetFiledDate"] = _s(filing, "original_filing_date")
    if return_section in (13, 14, 16, 18):
        fs["NoticeNo"] = _s(filing, "notice_no")
        fs["NoticeDate"] = _s(filing, "notice_date")
    elig = data.get("eligibility") or {}
    if elig.get("director_unlisted"):
        fs["CompDirectorPrvYrFlg"] = "Y"
        fs["HeldUnlistedEqShrPrYrFlg"] = "Y" if elig.get("unlisted_equity") else "N"
    return fs


# ---------------------------------------------------------------------------
# ScheduleS  (salary)
# ---------------------------------------------------------------------------

_HRA_KEYS = ("Placeofwork", "ActlHRARecv", "ActlRentPaid", "DtlsSalUsSec171",
             "ActlRentPaid10Per", "Sal40Or50Per", "EligbleExmpAllwncUs13A")


def _schedule_s(data, salary_block, hra_block, exp) -> dict:
    personal = data.get("personal", {}) or {}
    sal_in = data.get("salary", {}) or {}
    employer = (exp or {}).get("employer") or {}

    nature_rows = []
    if exp is not None:
        for comp in exp["component_breakdown"]:
            code = NAT_SAL.get(comp["code"], "OTH")
            row = {"NatureDesc": code, "OthAmount": comp["amount"]}
            if code == "OTH":
                row["OthNatOfInc"] = comp["label"]
            nature_rows.append(row)
        for tw in exp["terminal_workings"]:
            if not tw["taxable"]:
                continue
            name = tw["name"].split(" —")[0]
            code = next((c for k, c in TERMINAL_CODES if name.startswith(k)), "7")
            nature_rows.append({"NatureDesc": code, "OthAmount": tw["taxable"]})

    salarys = {
        "GrossSalary": salary_block["GrossSalary"],
        "Salary": salary_block["Salary"],
        "ValueOfPerquisites": salary_block["PerquisitesValue"],
        "ProfitsinLieuOfSalary": salary_block["ProfitsInSalary"],
        "IncomeNotified89A": 0,
        "IncomeNotifiedOther89A": 0,
    }
    if nature_rows:
        salarys["NatureOfSalary"] = {"OthersIncDtls": nature_rows}
    if exp is not None and exp.get("perquisites_rows"):
        salarys["NatureOfPerquisites"] = {"OthersIncDtls": [
            {"NatureDesc": "OTH", "OthNatOfInc": r["label"], "OthAmount": r["amount"]}
            for r in exp["perquisites_rows"]]}
    if exp is not None and exp.get("profits_rows"):
        salarys["NatureOfProfitInLieuOfSalary"] = [
            {"NatureDesc": "OTH", "OthNatOfInc": r["label"], "OthAmount": r["amount"]}
            for r in exp["profits_rows"]]

    nature_emp = (personal.get("employer_category") or employer.get("category") or "NA")
    if nature_emp == "NA":
        nature_emp = "OTH"
    sal_row = {
        "NameOfEmployer": employer.get("name") or _s(sal_in, "employer_name") or "Employer",
        "NatureOfEmployment": nature_emp,
        "AddressDetail": {
            "AddrDetail": employer.get("address") or _s(sal_in, "employer_address") or "-",
            "CityOrTownOrDistrict": employer.get("city") or
            _s((personal.get("address") or {}), "city_or_town_or_district") or "-",
            "StateCode": employer.get("state_code") or
            _s((personal.get("address") or {}), "state_code") or "01",
        },
        "Salarys": salarys,
    }
    tan = employer.get("tan") or _s(sal_in, "employer_tan")
    if tan:
        sal_row["TANofEmployer"] = tan
    epin = employer.get("pin_code")
    if epin:
        sal_row["AddressDetail"]["PinCode"] = int(epin)

    ex = salary_block["AllwncExemptUs10"]
    sched = {
        "Salaries": [sal_row],
        "TotalGrossSalary": salary_block["GrossSalary"],
        "AllwncExtentExemptUs10": ex.get("TotalAllwncExemptUs10", 0),
        "NetSalary": salary_block["NetSalary"],
        "DeductionUS16": salary_block["DeductionUs16"],
        "DeductionUnderSection16ia": salary_block["DeductionUs16ia"],
        "EntertainmntalwncUs16ii": salary_block["EntertainmentAlw16ii"],
        "ProfessionalTaxUs16iii": salary_block["ProfessionalTaxUs16iii"],
        "TotIncUnderHeadSalaries": salary_block["IncomeFromSal"],
    }
    if ex.get("AllwncExemptUs10Dtls"):
        sched["AllwncExemptUs10"] = {
            "AllwncExemptUs10Dtls": ex["AllwncExemptUs10Dtls"],
        }
    if hra_block:
        sched["Section10_13A"] = {k: hra_block[k] for k in _HRA_KEYS if k in hra_block}
    return sched


# ---------------------------------------------------------------------------
# ScheduleCGFor23 + Schedule112A
# ---------------------------------------------------------------------------

def _deduct48(acq_cost, improve_cost, exp_trans):
    return {"AquisitCost": acq_cost, "ImproveCost": improve_cost,
            "ExpOnTrans": exp_trans,
            "TotalDedn": acq_cost + improve_cost + exp_trans}


def _eq94(full_consideration, acq_cost, improve_cost, exp_trans, balance,
          capgains, is_54f=None):
    """EquityOrUnitSec94Type / EquityOrUnitSec54Type row."""
    row = {
        "FullValueConsdRecvUnqshr": 0,
        "FairMrktValueUnqshr": 0,
        "FullValueConsdSec50CA": 0,
        "FullValueConsdOthUnqshr": 0,
        "FullConsideration": full_consideration,
        "DeductSec48": _deduct48(acq_cost, improve_cost, exp_trans),
        "BalanceCG": balance,
        "CapgainonAssets": capgains,
    }
    if is_54f is None:
        row["LossSec94of7Or94of8"] = 0
    else:
        row["DeductionUs54F"] = is_54f
    return row


def _q_daterange(qmap) -> dict:
    qmap = qmap or {}
    return {"DateRange": {
        "Upto15Of6": int(qmap.get("Q1", 0)),
        "Upto15Of9": int(qmap.get("Q2", 0)),
        "Up16Of9To15Of12": int(qmap.get("Q3", 0)),
        "Up16Of12To15Of3": int(qmap.get("Q4", 0)),
        "Up16Of3To31Of3": int(qmap.get("Q5", 0)),
    }}


def _zero_daterange_obj() -> dict:
    return _q_daterange(None)


def _schedule_112a_rows(eq_lt_rows) -> list[dict]:
    dtls = []
    for r in eq_lt_rows:
        qty = int(float(r.get("qty") or 100) or 100)
        qty = max(qty, 1)
        before = bool(r["buy_date"] and r["buy_date"] <= CG.GRANDFATHER_DATE)
        fmv = r["fmv"] if r["fmv"] else r["cost"]
        dtls.append({
            "ShareOnOrBefore": "BE" if before else "AE",
            "ISINCode": r["isin"][:12] if r["isin"] else "INNOTREQUIRD",
            "ShareUnitName": r["scrip"],
            "NumSharesUnits": qty,
            "SalePricePerShareUnit": round(r["sale"] / qty, 4),
            "TotSaleValue": r["sale"],
            "CostAcqWithoutIndx": r["cost"],
            "AcquisitionCost": r["coa"],
            "LTCGBeforelowerB1B2": r["cost"],
            "FairMktValuePerShareunit": round(fmv / qty, 4),
            "TotFairMktValueCapAst": fmv,
            "ExpExclCnctTransfer": r["expenses"],
            "TotalDeductions": r["coa"] + r["expenses"],
            "Balance": r["gain"],
        })
    return dtls


def _curr_yr_losses(cg) -> dict:
    from .schemakit import init_required
    node = init_required(D2, "CurrYrLosses")
    class_income = {"InStcg20Per": cg["stcg_111a"], "InStcg30Per": 0,
                    "InStcgAppRate": cg["stcg_apprate"], "InStcgDTAARate": 0,
                    "InLtcg12_5Per": cg["ltcg_112a_taxable"] + cg["ltcg_other_taxable"]
                    + cg["exempt_112a_used"],
                    "InLtcgDTAARate": 0}
    for key, income in class_income.items():
        sub = node.get(key)
        if isinstance(sub, dict):
            for fld in ("CurrYearIncome", "CurrYrCapGain"):
                if fld in sub:
                    sub[fld] = income
    return node


def _land_dtls_lt(rows, cg) -> list[dict]:
    out = []
    for r in rows:
        base_cost = r["cost"] + r["improvement"]
        eff_cost = int(r.get("indexed_cost", 0)) or base_cost
        ex_rows = [{"ExemptionSecCode": e["section"] if e["section"] != "54EE" else "54EC",
                    "ExemptionAmount": e["applied"]}
                   for e in cg["exemptions_applied"] if e["section"] in ("54", "54EC", "54EE")]
        ex_total = sum(e["ExemptionAmount"] for e in ex_rows)
        buy_fy = f"{r['buy_date'][:4]}-{int(r['buy_date'][2:4]) + 1:02d}" \
            if r.get("buy_date") else "2019-20"
        buy_year = int(r["buy_date"][:4]) if r.get("buy_date") else 2025
        improve_dtls = ([{"slno": 1, "ImproveCost": r["improvement"],
                          "ImproveDate": buy_fy,
                          "CostOfImpIndex": int(CG._indexed_cost(r["improvement"], buy_year))
                          if r.get("indexed_cost") else 0}]
                        if r["improvement"] else [])
        out.append({
            "FullConsideration": r["sale"],
            "DateofPurchase": r["buy_date"] or "2024-01-01",
            "DateofSale": r["sale_date"] or "2025-06-01",
            "PropertyValuation": r["sale"],
            "FullConsideration50C": r["sale"],
            "AquisitCost": r["cost"],
            "AquisitCostIndex": int(r.get("indexed_cost", 0)),
            "CostOfImprovements": {"CostOfImprovementsDtls": improve_dtls,
                                   "TotalImprovecost": r["improvement"],
                                   "TotalindexImprovecost": 0},
            "ExpOnTrans": r["expenses"],
            "TotalDedn": eff_cost + r["expenses"],
            "TotalDednForEiB": base_cost + r["expenses"],
            "Balance": r["sale"] - eff_cost - r["expenses"],
            "BalanceForEiB": r["sale"] - base_cost - r["expenses"],
            "ExemptionOrDednUs54": {"ExemptionOrDednUs54Dtls": ex_rows,
                                    "ExemptionGrandTotal": ex_total},
            "LTCGonImmvblPrprty": r["gain"],
            "LTCGonImmvblPrprtyBE": (r.get("gain_wo_index", r["gain"])
                                     if r.get("type") == "ltcg_indexed20" else r["gain"]),
            "TaxSec1121aiiB": round(max(0, r.get("gain_wo_index", r["gain"]))
                                    * C.LTCG_GENERAL_125_RATE),
            "TaxSec1121a": (round(max(0, r.get("gain_with_index", r["gain"]))
                                  * C.LTCG_INDEXED_20_RATE)
                            if r.get("type") == "ltcg_indexed20" else 0),
            "ExcessAmtSec1121a": 0,
        })
    return out


def _schedule_cg(cg) -> dict:
    st_eq = cg["rows"]["equity_st"]
    other_st = cg["rows"]["other_st"]
    land = cg["rows"]["land_building"]
    other_lt = cg["rows"]["other_lt"]

    st_land = [r for r in land if r.get("type") == "stcg_apprate"]
    lt_land = [r for r in land if r.get("type") != "stcg_apprate"]

    # ---- short-term ----
    st23 = {
        "NRITransacSec48Dtl": {"NRItaxSTTPaid": 0, "NRItaxSTTNotPaid": 0},
        "NRISecur115AD": _eq94(0, 0, 0, 0, 0, 0),
        "SaleOnOtherAssets": _eq94(
            sum(r["sale"] for r in other_st),
            sum(r["cost"] for r in other_st),
            sum(r["improvement"] for r in other_st),
            sum(r["expenses"] for r in other_st),
            sum(r["sale"] - r["cost"] - r["improvement"] - r["expenses"]
                for r in other_st),
            cg["stcg_apprate"]),
        "TotalAmtDeemedStcg": 0,
        "PassThrIncNatureSTCG": 0,
        "TotalAmtNotTaxUsDTAAStcg": 0,
        "TotalAmtTaxUsDTAAStcg": 0,
        "TotalSTCG": cg["stcg_111a"] + cg["stcg_apprate"],
    }
    if st_eq and cg["stcg_111a"]:
        st23["EquityMFonSTT"] = [{
            "MFSectionCode": "1A",
            "EquityMFonSTTDtls": {
                "FullConsideration": sum(r["sale"] for r in st_eq),
                "DeductSec48": _deduct48(sum(r["cost"] for r in st_eq), 0,
                                         sum(r["expenses"] for r in st_eq)),
                "BalanceCG": sum(r["sale"] - r["cost"] - r["expenses"] for r in st_eq),
                "LossSec94of7Or94of8": 0,
                "CapgainonAssets": cg["stcg_111a"],
            },
        }]
    if st_land:
        st23["SaleofLandBuild"] = {
            "SaleofLandBuildDtls": [{
                "FullConsideration50C": r["sale"],
                "AquisitCost": r["cost"],
                "ImproveCost": r["improvement"],
                "ExpOnTrans": r["expenses"],
                "TotalDedn": r["cost"] + r["improvement"] + r["expenses"],
                "Balance": r["sale"] - r["cost"] - r["improvement"] - r["expenses"],
                "DeductionUs54B": 0,
                "STCGonImmvblPrprty": r["gain"],
            } for r in st_land],
        }

    # ---- long-term ----
    ded_54f = sum(e["applied"] for e in cg["exemptions_applied"] if e["section"] == "54F")
    ltcg_112a_gross = cg["ltcg_112a_gross"]
    lt23 = {
        "SaleOfEquityShareUs112A": {
            "BalanceCG": ltcg_112a_gross + ded_54f,
            "DeductionUs54F": ded_54f,
            "CapgainonAssets": ltcg_112a_gross,
        },
        "NRISaleOfEquityShareUs112A": {"BalanceCG": 0, "DeductionUs54F": 0,
                                       "CapgainonAssets": 0},
        "NRISaleofForeignAsset": {"SaleonSpecAsset": 0, "DednSpecAssetus115": 0,
                                  "BalonSpeciAsset": 0},
        "SaleofAssetNADtls": {"SaleofAssetNA": _eq94(
            sum(r["sale"] for r in other_lt),
            sum(r["cost"] for r in other_lt),
            sum(r["improvement"] for r in other_lt),
            sum(r["expenses"] for r in other_lt),
            sum(r["sale"] - r["cost"] - r["improvement"] - r["expenses"]
                for r in other_lt),
            sum(r["gain"] for r in other_lt), 0)},
        "TotalAmtDeemedLtcg": 0,
        "PassThrIncNatureLTCG": 0,
        "TotalAmtNotTaxUsDTAALtcg": 0,
        "TotalAmtTaxUsDTAALtcg": 0,
        "TotalLTCG": cg["ltcg_112a_gross"] + cg["ltcg_other_taxable"],
    }
    if lt_land:
        lt23["SaleofLandBuild"] = {
            "SaleofLandBuildDtls": _land_dtls_lt(lt_land, cg),
            "TotalLTCGImmblPrprty": sum(r["gain"] for r in lt_land),
            "TotalExcessTax": 0,
        }

    exemptions_total = sum(e["applied"] for e in cg["exemptions_applied"])
    cgfor23 = {
        "ShortTermCapGainFor23": st23,
        "LongTermCapGain23": lt23,
        "SumOfCGIncm": (cg["stcg_111a"] + cg["stcg_apprate"]
                        + cg["ltcg_112a_gross"] + cg["ltcg_other_taxable"]),
        "IncmFromVDATrnsf": 0,
        "TotScheduleCGFor23": cg["total_cg_incl_exempt112a"],
        "CurrYrLosses": _curr_yr_losses(cg),
        "AccruOrRecOfCG": {
            "ShortTermUnder20Per": _q_daterange(cg["quarterly"]["111a"]),
            "ShortTermUnder30Per": _q_daterange(None),
            "ShortTermUnderAppRate": _q_daterange(cg["quarterly"]["apprate_st"]),
            "ShortTermUnderDTAARate": _q_daterange(None),
            "LongTermUnder12_5Per": _q_daterange(
                {q: cg["quarterly"]["112a"].get(q, 0) + cg["quarterly"]["ltcg_rate"].get(q, 0)
                 for q in ("Q1", "Q2", "Q3", "Q4", "Q5")}),
            "LongTermUnderDTAARate": _q_daterange(None),
        },
    }
    if exemptions_total:
        by_sec: dict[str, int] = {}
        for e in cg["exemptions_applied"]:
            by_sec[e["section"]] = by_sec.get(e["section"], 0) + e["applied"]
        first_sale = next((r["sale_date"] for grp in cg["rows"].values()
                           for r in grp if r.get("sale_date")), "2025-07-01")
        dtls = {"TotDeductClaim": exemptions_total}
        if by_sec.get("54"):
            dtls["DeducClaimDtlsUs54"] = [{
                "DateofTransfer": first_sale,
                "CostofNewResHouse": by_sec["54"],
                "AmtDeducted": by_sec["54"]}]
        if by_sec.get("54EC") or by_sec.get("54EE"):
            amt = by_sec.get("54EC", 0) + by_sec.get("54EE", 0)
            dtls["DeducClaimDtlsUs54EC"] = [{
                "DateofTransfer": first_sale, "AmtInvested": amt,
                "AmtDeducted": amt}]
        if by_sec.get("54F"):
            dtls["DeducClaimDtlsUs54F"] = [{
                "DateofTransfer": first_sale,
                "CostofNewResHouse": by_sec["54F"],
                "AmtDeducted": by_sec["54F"]}]
        cgfor23["DeducClaimInfo"] = dtls
    return cgfor23


def _schedule_112a(cg) -> dict:
    eq_lt = cg["rows"]["equity_lt"]
    ded_54f = sum(e["applied"] for e in cg["exemptions_applied"] if e["section"] == "54F")
    return {
        "Schedule112ADtls": _schedule_112a_rows(eq_lt),
        "SaleValue112A": sum(r["sale"] for r in eq_lt),
        "CostAcqWithoutIndx112A": sum(r["cost"] for r in eq_lt),
        "AcquisitionCost112A": sum(r["coa"] for r in eq_lt),
        "LTCGBeforelowerB1B2112A": sum(r["cost"] for r in eq_lt),
        "FairMktValueCapAst112A": sum((r["fmv"] if r["fmv"] else r["cost"]) for r in eq_lt),
        "ExpExclCnctTransfer112A": sum(r["expenses"] for r in eq_lt),
        "Deductions112A": sum(r["coa"] + r["expenses"] for r in eq_lt),
        "Balance112A": sum(r["gain"] for r in eq_lt),
        "TotalBalance112A": cg["ltcg_112a_gross"] + ded_54f,
    }


# ---------------------------------------------------------------------------
# ScheduleOS / ScheduleEI / ScheduleSI
# ---------------------------------------------------------------------------

def _schedule_os(data, oth_total, ded_57iia, os_rows=None) -> dict:
    os_rows = os_rows or []
    os_special_total = sum(r["amount"] for r in os_rows)
    rows = data.get("other_sources") or []
    gross = {"DividendGross": 0, "IntrstFrmSavingBank": 0, "IntrstFrmTermDeposit": 0,
             "IntrstFrmIncmTaxRefund": 0, "IntrstFrmOthers": 0, "FamilyPension": 0,
             "OthersGross": 0}
    for r in rows:
        with field_scope("Income from other sources", "other_sources"):
            amt = _g(r, "amount")
        if amt <= 0:
            continue
        nat = _s(r, "nature", "OTH")
        if nat == "SAV":
            gross["IntrstFrmSavingBank"] += amt
        elif nat == "IFD":
            gross["IntrstFrmTermDeposit"] += amt
        elif nat == "TAX":
            gross["IntrstFrmIncmTaxRefund"] += amt
        elif nat == "DIV":
            gross["DividendGross"] += amt
        elif nat == "FAP":
            gross["FamilyPension"] += amt
        elif nat == "INT":
            gross["IntrstFrmOthers"] += amt
        else:
            gross["OthersGross"] += amt
    interest_gross = (gross["IntrstFrmSavingBank"] + gross["IntrstFrmTermDeposit"]
                      + gross["IntrstFrmIncmTaxRefund"] + gross["IntrstFrmOthers"])
    gross_total = sum(gross.values())

    item = {
        "GrossIncChrgblTaxAtAppRate": gross_total,
        "DividendGross": gross["DividendGross"],
        "InterestGross": interest_gross,
        "IntrstFrmSavingBank": gross["IntrstFrmSavingBank"],
        "IntrstFrmTermDeposit": gross["IntrstFrmTermDeposit"],
        "IntrstFrmIncmTaxRefund": gross["IntrstFrmIncmTaxRefund"],
        "NatofPassThrghIncome": 0,
        "IntrstFrmOthers": gross["IntrstFrmOthers"],
        "RentFromMachPlantBldgs": 0,
        "Tot562x": 0, "Aggrtvaluewithoutcons562x": 0, "Immovpropwithoutcons562x": 0,
        "Immovpropinadeqcons562x": 0, "Anyotherpropwithoutcons562x": 0,
        "Anyotherpropinadeqcons562x": 0,
        "FamilyPension": gross["FamilyPension"],
        "AnyOtherIncome": gross["OthersGross"],
        "IncChargeableSpecialRates": os_special_total,
        "LtryPzzlChrgblUs115BB": 0, "IncChrgblUs115BBJ": 0, "IncChrgblUs115BBE": 0,
        "CashCreditsUs68": 0, "UnExplndInvstmntsUs69": 0, "UnExplndMoneyUs69A": 0,
        "UnDsclsdInvstmntsUs69B": 0, "UnExplndExpndtrUs69C": 0,
        "AmtBrwdRepaidOnHundiUs69D": 0,
        "OthersGross": gross["OthersGross"],
        "PassThrIncOSChrgblSplRate": 0,
        "Deductions": {"Expenses": 0, "UsrIntExp57": 0, "IntExp57": 0,
                       "DeductionUs57iia": ded_57iia, "Depreciation": 0,
                       "TotDeductions": ded_57iia},
        # balance after the special-rate buckets are taken out (they are
        # shown on their own nodes and taxed at special rates)
        "BalanceNoRaceHorse": oth_total - os_special_total,
        "IncomeNotified89AOS": 0,
        "TaxAccumulatedBalRecPF": {"TaxAccmltdBalRecPFDtls": [],
                                   "TotalIncomeBenefit": 0, "TotalTaxBenefit": 0},
    }
    zero_q = _zero_daterange_obj()
    # OS special-rate windows (Schedule SI family) — dedicated schema nodes;
    # with only an annual figure available the amount is reported in Q5
    # (16-Mar to 31-Mar window), as noted in the computation report.
    for r in os_rows:
        node = SR.OS_SPECIAL_INPUTS[r["key"]][1]
        item[node] = r["amount"]
        qnode = SR.OS_SPECIAL_INPUTS[r["key"]][2]
        if qnode:
            item.setdefault("_q", {})[qnode] = _q_daterange({"Q5": r["amount"]})
    q_extra = item.pop("_q", {})
    out = {
        "IncOthThanOwnRaceHorse": item,
        "TotOthSrcNoRaceHorse": oth_total,
        "IncFromOwnHorse": {"Receipts": 0, "DeductSec57": 0, "AmtNotDeductibleUs58": 0,
                            "ProfitChargTaxUs59": 0, "BalanceOwnRaceHorse": 0},
        "IncFrmLottery": zero_q, "IncFrmOnGames": _zero_daterange_obj(),
        "DividendIncUs115BBDA": _zero_daterange_obj(),
        "DividendIncUs115BBDAaiii": _zero_daterange_obj(),
        "DividendIncUs115A1ai": _zero_daterange_obj(),
        "DividendIncUs115A1aA": _zero_daterange_obj(),
        "DividendIncUs115AC": _zero_daterange_obj(),
        "DividendIncUs115ACA": _zero_daterange_obj(),
        "DividendIncUs115AD1i": _zero_daterange_obj(),
        "DividendDTAA": _zero_daterange_obj(),
        "NOT89A": _zero_daterange_obj(),
        "IncChargeable": oth_total,
    }
    out.update(q_extra)
    return out


def _schedule_ei(data) -> dict:
    tot_agri = 0
    other_rows = []
    tot_other = 0
    for r in (data.get("exempt_income") or []):
        amt = _g(r, "amount")
        if amt <= 0:
            continue
        cat = _s(r, "category", "OTH")
        if cat == "AGRI":
            tot_agri += amt
        else:
            tot_other += amt
            row = {"Category": "OTH", "SubCategory": _s(r, "sub_category") or "OTH",
                   "Description": _s(r, "description") or "-", "OthAmount": amt}
            other_rows.append(row)
    return {
        "InterestInc": 0,
        "GrossAgriRecpt": tot_agri, "ExpIncAgri": 0, "UnabAgriLossPrev8": 0,
        "NetAgriIncOrOthrIncRule7": tot_agri,
        "ExcNetAgriInc": {"ExcNetAgriIncDtls": []},
        "OthersInc": {"OthersIncDtls": other_rows},
        "Others": tot_other,
        "IncNotChrgblAsPerDTAA": {"IncNotChrgblAsPerDTAADtls": []},
        "IncNotChrgblToTax": tot_agri + tot_other,
        "PassThrIncNotChrgblTax": 0,
        "TotalExemptInc": tot_agri + tot_other,
    }


def _schedule_si(cg, tax, os_rows=None) -> dict:
    rows = []
    if cg["stcg_111a"]:
        rows.append({"SecCode": "1A", "SplRatePercent": 20,
                     "SplRateInc": cg["stcg_111a"],
                     "SplRateIncTax": round(cg["stcg_111a"] * C.STCG_111A_RATE)})
    if cg["ltcg_112a_taxable"]:
        rows.append({"SecCode": "2A", "SplRatePercent": 12.5,
                     "SplRateInc": cg["ltcg_112a_taxable"],
                     "SplRateIncTax": round(cg["ltcg_112a_taxable"] * C.LTCG_112A_RATE)})
    if cg["ltcg_other_125"]:
        rows.append({"SecCode": "21", "SplRatePercent": 12.5,
                     "SplRateInc": cg["ltcg_other_125"],
                     "SplRateIncTax": round(cg["ltcg_other_125"] * C.LTCG_GENERAL_125_RATE)})
    if cg["ltcg_indexed_20"]:
        rows.append({"SecCode": "21", "SplRatePercent": 20,
                     "SplRateInc": cg["ltcg_indexed_20"],
                     "SplRateIncTax": round(cg["ltcg_indexed_20"] * C.LTCG_INDEXED_20_RATE)})
    for r in (os_rows or []):
        rows.append({"SecCode": r["code"], "SplRatePercent": r["rate"],
                     "SplRateInc": r["amount"],
                     "SplRateIncTax": round(r["tax"])})
    total_inc = sum(r["SplRateInc"] for r in rows)
    total_tax = sum(r["SplRateIncTax"] for r in rows)
    out = {"TotSplRateInc": total_inc, "TotSplRateIncTax": total_tax}
    if rows:
        out["SplCodeRateTax"] = rows
    return out


# ---------------------------------------------------------------------------
# ScheduleCYLA / ScheduleBFLA  (no losses in this implementation — clean zeros)
# ---------------------------------------------------------------------------

def _fill_cyla(node, income):
    inner = node.get("IncCYLA") or node.get("IncBFLA") or {}
    for k in list(inner):
        if k in ("IncOfCurYrUnderThatHead", "IncOfCurYrUndHeadFromCYLA",
                 "IncOfCurYrAfterSetOff", "IncOfCurYrAfterSetOffLosses",
                 "IncOfCurYrAftSetOffBFLosses", "IncOfCurYrAfterSetOffBFLosses"):
            inner[k] = income
    return node


def _cyla_bfla(ti, cg):
    from .schemakit import init_required
    cyla = init_required(D2, "ScheduleCYLA")
    bfla = init_required(D2, "ScheduleBFLA")

    class_income = {
        "Salary": ti["salary"], "HP": ti["hp"],
        "STCG20Per": cg["stcg_111a"], "STCG30Per": 0,
        "STCGAppRate": cg["stcg_apprate"], "STCGDTAARate": 0,
        "LTCG12_5Per": cg["ltcg_112a_taxable"] + cg["ltcg_other_taxable"]
        + cg["exempt_112a_used"],
        "LTCGDTAARate": 0,
        "OthSrcExclRaceHorse": ti["os"], "OthSrcRaceHorse": 0, "IncOSDTAA": 0,
    }
    for key, income in class_income.items():
        if key in cyla and isinstance(cyla[key], dict):
            _fill_cyla(cyla[key], income)
        if key in bfla and isinstance(bfla[key], dict):
            _fill_cyla(bfla[key], income)
    bfla["IncomeOfCurrYrAftCYLABFLA"] = ti["gross"] + cg["total_cg_incl_exempt112a"]
    return cyla, bfla


# ---------------------------------------------------------------------------
# TDS2 / TDS3 / TCS in the ITR-2 shapes
# ---------------------------------------------------------------------------

def _tax_credit_dtls(deducted, claimed):
    return {
        "TaxDeductedOwnHands": deducted,
        "TaxDeductedIncome": deducted,
        "TaxDeductedTDS": deducted,
        "TaxClaimedOwnHands": claimed,
        "TaxClaimedIncome": claimed,
        "TaxClaimedTDS": claimed,
    }


def _bf_year(r, key):
    """Brought-forward credit year: only genuine earlier years are emitted
    (the schema's DeductedYr enum intentionally stops before the current PY)."""
    y = str(r.get(key) or "")
    return int(y) if y.isdigit() and int(y) < 2025 else None


def _itr2_tds2(data):
    tp = data.get("tax_paid") or {}
    rows = []
    total = 0
    for r in (tp.get("tds_others") or []):
        with field_scope("TDS (other than salary)", "tax_paid.tds_others"):
            claimed = min(_g(r, "tds_claimed"), _g(r, "tds_deducted"))
        if claimed <= 0:
            continue
        row = {
            "TDSCreditName": "S",
            "TANOfDeductor": _s(r, "tan"),
            "TDSSection": _s(r, "section", "94A"),
            "TaxDeductCreditDtls": _tax_credit_dtls(_g(r, "tds_deducted"), claimed),
            "GrossAmount": _g(r, "amount_paid"),
            "HeadOfIncome": _s(r, "head_of_income") or "OS",
            "AmtCarriedFwd": 0,
        }
        yr = _bf_year(r, "deduction_year")
        if yr:
            row["DeductedYr"] = yr
        rows.append(row)
        total += claimed
    return {"TDSOthThanSalaryDtls": rows, "TotalTDSonOthThanSals": total}


def _itr2_tds3(data):
    tp = data.get("tax_paid") or {}
    rows = []
    total = 0
    for r in (tp.get("tds3") or []):
        with field_scope("TDS on rent/immovable property", "tax_paid.tds3"):
            claimed = min(_g(r, "tds_claimed"), _g(r, "tds_deducted"))
        if claimed <= 0:
            continue
        row = {
            "TDSCreditName": "S",
            "PANOfBuyerTenant": _s(r, "tenant_pan"),
            "TDSSection": _s(r, "section", "4-IB"),
            "TaxDeductCreditDtls": _tax_credit_dtls(_g(r, "tds_deducted"), claimed),
            "GrossAmount": _g(r, "gross_receipt"),
            "HeadOfIncome": _s(r, "head_of_income") or "HP",
            "AmtCarriedFwd": 0,
        }
        yr = _bf_year(r, "deduction_year")
        if yr:
            row["DeductedYr"] = yr
        rows.append(row)
        total += claimed
    return {"TDS3onOthThanSalDtls": rows, "TotalTDS3OnOthThanSal": total}


def _itr2_tcs(data):
    tp = data.get("tax_paid") or {}
    rows = []
    total = 0
    for r in (tp.get("tcs") or []):
        with field_scope("TCS", "tax_paid.tcs"):
            claimed = min(_g(r, "tcs_claimed"), _g(r, "tcs_collected"))
        if claimed <= 0:
            continue
        row = {
            "TCSCreditOwner": "1",
            "EmployerOrDeductorOrCollectTAN": _s(r, "tan"),
            "TCSCurrFYDtls": {"TCSAmtCollOwnHand": _g(r, "tcs_collected"),
                              "TCSAmtCollSpouseOrOthrHand": 0},
            "TCSClaimedThisYearDtls": {"TCSAmtCollOwnHand": claimed,
                                       "TCSAmtCollSpouseOrOthrHand": 0},
            "AmtCarriedFwd": 0,
        }
        yr = _bf_year(r, "collection_year")
        if yr:
            row["DeductedYr"] = yr
        rows.append(row)
        total += claimed
    return {"TCS": rows, "TotalSchTCS": total}


# ---------------------------------------------------------------------------
# PartB_TTI inner blocks
# ---------------------------------------------------------------------------

def _partb_tti_tax(tax):
    working = tax["Working"]
    spl_tax = max(0, tax["TotalTaxPayable"] - tax["Rebate87A"]
                  - working["slab_tax"] - tax["SurchargeComputed"])
    return {
        "TaxPayableOnTI": {
            "TaxAtNormalRatesOnAggrInc": working["slab_tax"],
            "TaxAtSpecialRates": spl_tax,
            "RebateOnAgriInc": 0,
            "TaxPayableOnTotInc": working["slab_tax"] + spl_tax,
        },
        "Rebate87A": tax["Rebate87A"],
        "TaxPayableOnRebate": tax["TaxPayableOnRebate"],
        "Surcharge25ofSI": tax["SurchargeComputed"],
        "SurchargeOnAboveCrore": 0,
        "Surcharge25ofSIBeforeMarginal": tax["SurchargeComputed"],
        "SurchargeOnAboveCroreBeforeMarginal": 0,
        "TotalSurcharge": tax["SurchargeComputed"],
        "EducationCess": tax["EducationCess"],
        "GrossTaxLiability": tax["GrossTaxLiability"],
        "GrossTaxPayable": tax["GrossTaxLiability"],
        "GrossTaxPay": {"TaxInc17": tax["GrossTaxLiability"], "TaxDeferred17": 0,
                        "TaxDeferredPayableCY": 0},
        "CreditUS115JD": 0,
        "TaxPayAfterCreditUs115JD": tax["GrossTaxLiability"],
        "TaxRelief": {"Section89": tax["Section89"], "Section90": 0, "Section91": 0,
                      "TotTaxRelief": tax["Section89"]},
        "NetTaxLiability": tax["NetTaxLiability"],
        "IntrstPay": {
            "IntrstPayUs234A": tax["IntrstPay"]["IntrstPayUs234A"],
            "IntrstPayUs234B": tax["IntrstPay"]["IntrstPayUs234B"],
            "IntrstPayUs234C": tax["IntrstPay"]["IntrstPayUs234C"],
            "LateFilingFee234F": tax["IntrstPay"]["LateFilingFee234F"],
            "TotalIntrstPay": tax["TotalIntrstPay"],
        },
        "AggregateTaxInterestLiability": tax["TotTaxPlusIntrstPay"],
    }


# ---------------------------------------------------------------------------
# public entry
# ---------------------------------------------------------------------------

def build_itr2(data: dict, *, creation_info: dict | None = None) -> dict:
    """Build the official ITR-2 JSON plus the extended computation report."""
    reset_scopes()
    data = copy.deepcopy(data or {})
    warnings: list[str] = []

    personal = data.get("personal", {}) or {}
    filing = data.get("filing", {}) or {}
    opt_out = bool(filing.get("opt_out_new_regime"))
    dob = _s(personal, "dob") or "1990-01-01"

    salary_block, hra_block = _build_salary(data, opt_out, warnings)
    exp = data.get("_salary_expansion")
    hp_blocks, hp_total = _build_house_property(data.get("house_properties"),
                                                opt_out, warnings)
    for blk in hp_blocks:  # ITR-2 makes the share % mandatory on every row
        blk.setdefault("AsseseeShareProperty", 100)
    others_block, oth_src_total, ded_57iia = _build_other_sources(
        data.get("other_sources"), opt_out, warnings)
    # OS-family special-rate incomes (Schedule SI: 115BB / 115BBJ / 115BBE)
    os_rows = SR.os_special_rows(data)
    os_total, os_tax = SR.os_special_totals(os_rows)
    if os_total:
        warnings.append(
            f"Special-rate other-source income: Rs. {os_total:,} taxed at "
            f"special rates (tax Rs. {os_tax:,}); see Schedule SI.")
        oth_src_total += os_total
    _build_exempt_income(data.get("exempt_income"), warnings)

    # ---- capital gains ----
    cg = CG.compute_capital_gains(data)
    warnings.extend(cg["setoff_notes"])
    l_in = data.get("ltcg112a") or {}
    legacy_112a = max(0, _g(l_in, "sale_consideration") - _g(l_in, "cost_of_acquisition"))
    if legacy_112a and not cg["has_any_cg"]:
        cg["ltcg_112a_gross"] = min(legacy_112a, CG.EXEMPT_112A) + max(0, legacy_112a - CG.EXEMPT_112A)
        cg["ltcg_112a_taxable"] = max(0, legacy_112a - CG.EXEMPT_112A)
        cg["exempt_112a_used"] = min(legacy_112a, CG.EXEMPT_112A)
        cg["total_cg"] += cg["ltcg_112a_taxable"]
        cg["total_cg_incl_exempt112a"] += legacy_112a

    salary_171 = salary_block["Salary"]
    with field_scope("Income from other sources", "other_sources"):
        interest_income = sum(_g(r, "amount") for r in (data.get("other_sources") or [])
                              if _s(r, "nature") in ("SAV", "IFD", "TAX"))
        savings_interest = sum(_g(r, "amount") for r in (data.get("other_sources") or [])
                               if _s(r, "nature") == "SAV")

    gti_excl_cg = salary_block["IncomeFromSal"] + hp_total + oth_src_total
    gti_incl_cg = gti_excl_cg + cg["total_cg_incl_exempt112a"]

    usr_via, elig_via, ded_scheds, via_total = _build_vi_a(
        data, opt_out, gti_excl_cg, salary_171,
        _s(personal, "employer_category", "NA"), dob,
        interest_income, savings_interest, ded_57iia > 0, warnings)

    total_income = T.round_off_income(max(0, gti_incl_cg - via_total))

    special = {"stcg_111a": cg["stcg_111a"],
               "ltcg_112a": cg["ltcg_112a_gross"],
               "ltcg_125": cg["ltcg_other_125"],
               "ltcg_idx20": cg["ltcg_indexed_20"],
               "os_special": os_rows}
    sched_tds1, _sched_tds2_itr1, _sched_tds3_itr1, _sched_tcs_itr1, sched_it, \
        tds_total, tcs_total, _challans = _build_tax_paid(data.get("tax_paid"))
    filing_date = _s(filing, "filing_date") or str(TODAY())
    return_section = _g(filing, "return_section", 11)
    with field_scope("Relief u/s 89", "relief_89"):
        section89 = _g(data, "relief_89")
    if exp and exp.get("arrears_rows") and not section89:
        dob_year = int(dob[:4]) if dob[:4].isdigit() else None
        w10e = SX.relief_89_working(total_income, exp["arrears_rows"],
                                    regime="old" if opt_out else "new",
                                    dob_year=dob_year)
        data["_form10e"] = w10e
        section89 = w10e["relief89"]
        if section89 > 0:
            warnings.append(f"Relief u/s 89 (Form 10E): Rs. {section89:,} computed "
                            "automatically from arrear details.")
    tax = T.compute_tax_v2(total_income=total_income, special=special,
                           opt_out_new_regime=opt_out, dob=dob, section89=section89,
                           filing_date=filing_date, return_section=return_section,
                           tds=tds_total, tcs=tcs_total,
                           advance_tax_challans=(data.get("tax_paid") or {}).get("challans"),
                           interest_override=data.get("interest_override"),
                           residential_status=_residential_status(data),
                           status=_s((data.get("personal") or {}), "status") or "I")

    # ------------------------------------------------------------------ doc
    itr2: dict = {}
    itr2["CreationInfo"] = _creation_block(creation_info)
    itr2["Form_ITR2"] = {
        "FormName": "ITR-2",
        "Description": UTILITY_FORM_DESCRIPTION["ITR2"],
        "AssessmentYear": "2026", "SchemaVer": "Ver1.0", "FormVer": "Ver1.0",
    }
    itr2["PartA_GEN1"] = {"PersonalInfo": _personal_info(data),
                          "FilingStatus": _filing_status(data)}

    itr2["ScheduleS"] = _schedule_s(data, salary_block, hra_block, exp)
    itr2["ScheduleHP"] = {"PropertyDetails": hp_blocks,
                          "PassThroghIncome": 0,
                          "TotalIncomeChargeableUnHP": hp_total}
    if cg["has_any_cg"] or legacy_112a:
        itr2["ScheduleCGFor23"] = _schedule_cg(cg)
        if cg["rows"]["equity_lt"] or legacy_112a:
            itr2["Schedule112A"] = _schedule_112a(cg)
    itr2["ScheduleOS"] = _schedule_os(data, oth_src_total, ded_57iia, os_rows)

    ti_map = {"salary": salary_block["IncomeFromSal"], "hp": hp_total,
              "os": oth_src_total, "gross": gti_excl_cg}
    cyla, bfla = _cyla_bfla(ti_map, cg)
    itr2["ScheduleCYLA"] = cyla
    itr2["ScheduleBFLA"] = bfla

    # Chapter VI-A (filter the ITR-1-shaped blocks to ITR-2's property sets)
    usr_props = set(D2["UsrDeductUndChapVIA"].get("properties", {}))
    elig_props = set(D2["DeductUndChapVIA"].get("properties", {}))
    itr2["ScheduleVIA"] = {
        "UsrDeductUndChapVIA": {k: v for k, v in usr_via.items() if k in usr_props},
        "DeductUndChapVIA": {k: v for k, v in elig_via.items() if k in elig_props},
    }
    for key, sched in ded_scheds.items():
        if key in _ITR2_PROPS:
            itr2[key] = sched
    if "Schedule80D" in itr2:  # ITR-2 makes every 80D sub-field mandatory
        blk = itr2["Schedule80D"].get("Sec80DSelfFamSrCtznHealth")
        if isinstance(blk, dict):
            for fld in ("SelfAndFamily", "SelfAndFamilySeniorCitizen",
                        "Parents", "ParentsSeniorCitizen", "EligibleAmountOfDedn"):
                blk.setdefault(fld, 0)
    itr2["ScheduleEI"] = _schedule_ei(data)
    itr2["ScheduleSI"] = _schedule_si(cg, tax, os_rows)

    short_total = cg["stcg_111a"] + cg["stcg_apprate"]
    long_total = cg["ltcg_112a_taxable"] + cg["ltcg_other_taxable"]
    special_income_total = cg["stcg_111a"] + cg["ltcg_112a_taxable"] + cg["ltcg_other_taxable"]
    itr2["PartB-TI"] = {
        "Salaries": salary_block["IncomeFromSal"],
        "IncomeFromHP": hp_total,
        "CapGain": {
            "ShortTerm": {"ShortTerm20Per": cg["stcg_111a"], "ShortTerm30Per": 0,
                          "ShortTermAppRate": cg["stcg_apprate"],
                          "ShortTermSplRateDTAA": 0, "TotalShortTerm": short_total},
            "LongTerm": {"LongTerm12_5Per": long_total + cg["exempt_112a_used"],
                         "LongTermSplRateDTAA": 0,
                         "TotalLongTerm": long_total + cg["exempt_112a_used"]},
            "ShortTermLongTermTotal": short_total + long_total + cg["exempt_112a_used"],
            "CapGains30Per115BBH": 0,
            "TotalCapGains": short_total + long_total + cg["exempt_112a_used"],
        },
        "IncFromOS": {"OtherSrcThanOwnRaceHorse": oth_src_total - os_total,
                      "IncChargblSplRate": os_total,
                      "FromOwnRaceHorse": 0, "TotIncFromOS": oth_src_total},
        "TotalTI": gti_excl_cg + short_total + long_total + cg["exempt_112a_used"],
        "CurrentYearLoss": 0,
        "BalanceAfterSetoffLosses": gti_excl_cg + short_total + long_total
        + cg["exempt_112a_used"],
        "BroughtFwdLossesSetoff": 0,
        "GrossTotalIncome": gti_incl_cg,
        "IncChargeTaxSplRate111A112": special_income_total,
        "DeductionsUnderScheduleVIA": via_total,
        "TotalIncome": total_income,
        "IncChargeableTaxSplRates": special_income_total,
        "NetAgricultureIncomeOrOtherIncomeForRate": 0,
        "AggregateIncome": total_income,
        "LossesOfCurrentYearCarriedFwd": cg["cg_loss_carry_st"] + cg["cg_loss_carry_lt"],
        "DeemedIncomeUs115JC": 0,
    }

    itr2["PartB_TTI"] = {
        "TaxPayDeemedTotIncUs115JC": 0,
        "Surcharge": 0,
        "HealthEduCess": 0,
        "TotalTaxPayablDeemedTotInc": 0,
        "ComputationOfTaxLiability": _partb_tti_tax(tax),
        "TaxPaid": {
            "TaxesPaid": {
                "AdvanceTax": tax["TaxesPaid"]["AdvanceTax"],
                "TDS": tax["TaxesPaid"]["TDS"],
                "TCS": tax["TaxesPaid"]["TCS"],
                "SelfAssessmentTax": tax["TaxesPaid"]["SelfAssessmentTax"],
                "TotalTaxesPaid": tax["TaxesPaid"]["TotalTaxesPaid"],
            },
            "BalTaxPayable": tax["BalTaxPayable"],
        },
        "Refund": {"RefundDue": tax["RefundDue"],
                   "BankAccountDtls": {"BankDtlsFlag": "Y", "AddtnlBankDetails": []}},
        "AssetOutIndiaFlag": "NO",
    }
    bank_rows = []
    for b in (data.get("bank_accounts") or []):
        row = {
            "IFSCCode": _s(b, "ifsc"), "BankName": _s(b, "bank_name"),
            "BankAccountNo": _s(b, "account_no"),
            "AccountType": _s(b, "account_type", "SB"),
            "UseForRefund": "true" if b.get("use_for_refund") else "false",
        }
        if bank_rows or row["UseForRefund"] == "true":
            bank_rows.append(row)
    itr2["PartB_TTI"]["Refund"]["BankAccountDtls"]["AddtnlBankDetails"] = bank_rows
    if not bank_rows:
        warnings.append("At least one bank account is required for the return.")

    itr2["ScheduleIT"] = sched_it
    sched_tds2 = _itr2_tds2(data)
    sched_tds3 = _itr2_tds3(data)
    sched_tcs = _itr2_tcs(data)
    if tds_total or (data.get("tax_paid") or {}).get("tds_on_salary"):
        itr2["ScheduleTDS1"] = sched_tds1
    if sched_tds2["TDSOthThanSalaryDtls"]:
        itr2["ScheduleTDS2"] = sched_tds2
    if sched_tds3["TDS3onOthThanSalDtls"]:
        itr2["ScheduleTDS3"] = sched_tds3
    if sched_tcs["TCS"]:
        itr2["ScheduleTCS"] = sched_tcs

    ver = data.get("verification", {}) or {}
    itr2["Verification"] = {
        "Declaration": {
            "AssesseeVerName": _s(ver, "name"),
            "FatherName": _s(ver, "father_name") or _s(personal, "father_name"),
            "AssesseeVerPAN": _s(ver, "pan") or _s(personal, "pan"),
        },
        "Capacity": _s(ver, "capacity", "S"),
        "Place": _s(ver, "place"),
        "Date": _s(ver, "date") or str(TODAY()),
    }

    doc = {"ITR": {"ITR2": itr2}}
    report = {
        "form": "ITR2", "regime": "old" if opt_out else "new",
        "salary_income": salary_block["IncomeFromSal"], "hp_income": hp_total,
        "other_sources_income": oth_src_total,
        "gross_total_income_excl_cg": gti_excl_cg,
        "gross_total_income_incl_cg": gti_incl_cg,
        "capital_gains": cg, "total_income": total_income,
        "deductions_eligible": elig_via, "tax": tax, "warnings": warnings,
        "salary_expansion": exp, "form10e": data.get("_form10e"),
    }
    return {"json": doc, "report": report}
