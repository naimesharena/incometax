"""ITR-2 builder: schema validation + arithmetic on the official AY 26-27 schema."""

import copy
import json
import os

import pytest

from itr1.itr2_builder import build_itr2
from itr1.schemakit import validate_with_schema

SAMPLE = os.path.join(os.path.dirname(__file__), "..", "samples", "sample_itr2.json")


def load_sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def test_rich_sample_is_schema_valid():
    out = build_itr2(load_sample())
    errs = validate_with_schema("ITR2", out["json"])
    assert errs == [], "\n".join(errs)


def test_rich_sample_arithmetic():
    out = build_itr2(load_sample())
    rep = out["report"]
    cg = rep["capital_gains"]
    # equity STCG: 190000 - 150000 - 500 = 39500 @ 20%
    assert cg["stcg_111a"] == 39500
    # gold STCG: 250000 - 200000 - 1000 = 49000 at slab rate
    assert cg["stcg_apprate"] == 49000
    # 112A: TCS (grandfathered COA=155000 → 245000-1000=244000) + HDFC (330000-280000-800=49200)
    #        = 293200 gross; exemption 125000 → taxable 168200
    assert cg["ltcg_112a_gross"] == 293200
    assert cg["exempt_112a_used"] == 125000
    assert cg["ltcg_112a_taxable"] == 168200
    # land: bought 2019 (CII 289), sold 2025 → 12.5% w/o indexation vs 20% indexed
    #   w/o: 52,00,000 - 25,00,000 - 1,00,000 - 52,000 = 25,48,000 → tax 318,500
    #   idx: cost 26,00,000×376/289 = 33,82,006 → gain 17,65,994 → tax 353,199 → 12.5% chosen
    #   then 54EC ₹5,00,000 applied → 20,48,000
    assert cg["ltcg_other_125"] == 2048000
    assert cg["ltcg_indexed_20"] == 0
    assert cg["exemptions_applied"][0]["applied"] == 500000
    # PartB-TI coherence
    itr2 = out["json"]["ITR"]["ITR2"]
    pbti = itr2["PartB-TI"]
    assert pbti["TotalIncome"] == rep["total_income"]
    assert pbti["GrossTotalIncome"] == rep["gross_total_income_incl_cg"]
    # SI rows cover the special-rate buckets
    si = itr2["ScheduleSI"]
    codes = {r["SecCode"] for r in si["SplCodeRateTax"]}
    assert codes == {"1A", "2A", "21"}
    assert si["TotSplRateInc"] == (39500 + 168200 + 2048000)


def test_legacy_112a_only_sample():
    data = {
        "personal": {"first_name": "Beena", "last_name": "Patel", "pan": "AAAPB1234C",
                     "dob": "1992-03-21", "father_name": "K Patel",
                     "address": {"residence_no": "1", "locality_or_area": "Varachha",
                                 "city_or_town_or_district": "Surat", "state_code": "06",
                                 "country_code": "91", "pin_code": 395006,
                                 "mobile_no": 9999999999, "email_address": "b@example.com"}},
        "filing": {"opt_out_new_regime": False, "return_section": 11},
        "salary": {"salary_17_1": 900000},
        "other_sources": [{"nature": "SAV", "amount": 8000}],
        "ltcg112a": {"sale_consideration": 200000, "cost_of_acquisition": 110000},
        "tax_paid": {"tds_on_salary": [{"tan": "SRTA01234B", "employer_name": "X Ltd",
                                        "salary_income": 900000, "tds": 25000}]},
        "bank_accounts": [{"ifsc": "SBIN0001234", "bank_name": "SBI",
                           "account_no": "12345678901", "account_type": "SB",
                           "use_for_refund": True}],
        "verification": {"name": "Beena Patel", "father_name": "K Patel",
                         "pan": "AAAPB1234C", "place": "Surat"},
    }
    out = build_itr2(data)
    assert validate_with_schema("ITR2", out["json"]) == []
    # ₹90,000 LTCG fully inside the ₹1.25L exemption → no tax on it
    assert out["report"]["capital_gains"]["ltcg_112a_taxable"] == 0
    assert out["report"]["tax"]["Working"]["ltcg_tax"] == 0


def test_cg_only_retired_investor():
    data = {
        "personal": {"first_name": "Chirag", "last_name": "Desai", "pan": "BBBPD5678D",
                     "dob": "1958-11-02", "father_name": "M Desai",
                     "address": {"residence_no": "9", "locality_or_area": "Athwa",
                                 "city_or_town_or_district": "Surat", "state_code": "06",
                                 "country_code": "91", "pin_code": 395001,
                                 "mobile_no": 9876500000, "email_address": "c@example.com"}},
        "filing": {"opt_out_new_regime": True, "return_section": 11},
        "capital_gains": {"equity_lt": [
            {"scrip": "RELIANCE", "isin": "INE002A01018", "buy_date": "2017-05-01",
             "sale_date": "2025-08-20", "qty": 100, "cost_of_acquisition": 90000,
             "fmv_31_01_2018": 95000, "sale_consideration": 320000, "expenses": 0}]},
        "other_sources": [{"nature": "IFD", "amount": 240000}],
        "bank_accounts": [{"ifsc": "ICIC0001234", "bank_name": "ICICI",
                           "account_no": "9990012", "account_type": "SB",
                           "use_for_refund": True}],
        "verification": {"name": "Chirag Desai", "father_name": "M Desai",
                         "pan": "BBBPD5678D", "place": "Surat"},
    }
    out = build_itr2(data)
    assert validate_with_schema("ITR2", out["json"]) == []
    # senior citizen, old regime: 240000 interest + 100000 taxable 112A...
    # 112A gross 225000 (grandfathered 95000) → taxable 100000
    assert out["report"]["capital_gains"]["ltcg_112a_taxable"] == 100000


def test_112a_rows_and_grandfathering():
    out = build_itr2(load_sample())
    rows = out["json"]["ITR"]["ITR2"]["Schedule112A"]["Schedule112ADtls"]
    tcs_row = next(r for r in rows if r["ShareUnitName"] == "TCS")
    assert tcs_row["ShareOnOrBefore"] == "BE"
    assert tcs_row["AcquisitionCost"] == 155000  # grandfathered
    hdfc_row = next(r for r in rows if r["ShareUnitName"] == "HDFCBANK")
    assert hdfc_row["ShareOnOrBefore"] == "AE"
    assert hdfc_row["AcquisitionCost"] == 280000


def test_terminal_salary_nature_rows():
    out = build_itr2(load_sample())
    sal = out["json"]["ITR"]["ITR2"]["ScheduleS"]["Salaries"][0]
    natures = {r["NatureDesc"] for r in sal["Salarys"]["NatureOfSalary"]["OthersIncDtls"]}
    assert "1" in natures        # basic
    assert "13" in natures       # gratuity
    assert "16" in natures       # leave encashment
    assert "15" in natures       # arrears
    assert sal["NatureOfEmployment"] == "OTH"


def test_tax_engine_special_buckets_match_si():
    out = build_itr2(load_sample())
    w = out["report"]["tax"]["Working"]
    assert w["tax_111a"] == round(39500 * 0.20)
    assert w["ltcg_tax"] == round(168200 * 0.125)
    cg_parts = out["json"]["ITR"]["ITR2"]["ScheduleSI"]["SplCodeRateTax"]
    si_tax = {r["SecCode"]: r["SplRateIncTax"] for r in cg_parts}
    assert si_tax["1A"] == w["tax_111a"]
    assert si_tax["2A"] == w["ltcg_tax"]
