"""Tests for the ITR-1 JSON builder against the official schema."""

import copy
import json
import os

import pytest

from itr1.builder import build_return
from itr1.validator import validate_all, validate_schema

SAMPLE = os.path.join(os.path.dirname(__file__), os.pardir,
                      "samples", "sample_taxpayer.json")


@pytest.fixture()
def sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def test_sample_builds_and_validates_new_regime(sample):
    res = build_return(sample)
    doc = res["json"]
    schema_issues = validate_schema(doc)
    assert schema_issues == [], f"schema issues: {[i.as_dict() for i in schema_issues]}"
    issues = validate_all(doc)
    blocking = [i for i in issues if i.category in ("A", "SCHEMA")]
    assert blocking == [], f"blocking issues: {[i.as_dict() for i in blocking]}"


def test_sample_builds_and_validates_old_regime(sample):
    sample["filing"]["opt_out_new_regime"] = True
    res = build_return(sample)
    doc = res["json"]
    assert res["json"]["ITR"]["ITR1"]["FilingStatus"]["OptOutNewTaxRegime"] == "Y"
    schema_issues = validate_schema(doc)
    assert schema_issues == [], f"schema issues: {[i.as_dict() for i in schema_issues]}"
    issues = validate_all(doc)
    blocking = [i for i in issues if i.category in ("A", "SCHEMA")]
    assert blocking == [], f"blocking issues: {[i.as_dict() for i in blocking]}"


def test_sample_expected_numbers_new_regime(sample):
    rep = build_return(sample)["report"]
    # 960000 salary - 75000 SD + 23500 other = 908500 GTI (exempt 10(14)(i) removed in new regime)
    assert rep["gross_total_income_excl_ltcg"] == 908_500
    assert rep["gross_total_income_incl_ltcg"] == 1_008_500
    # only 80CCD(2)=96,000 survives in the new regime
    assert rep["deductions_eligible"]["TotalChapVIADeductions"] == 96_000
    assert rep["total_income"] == 912_500
    # taxable excl LTCG = 812,500 -> 5% of 4,00,000 + 10% of 12,500 = 21,250,
    # fully rebated (income excl LTCG <= 12 lakh)
    assert rep["tax"]["TotalTaxPayable"] == 21_250
    assert rep["tax"]["Rebate87A"] == 21_250
    assert rep["tax"]["TotTaxPlusIntrstPay"] == 0
    assert rep["tax"]["RefundDue"] == 36_500


def test_sample_expected_numbers_old_regime(sample):
    sample["filing"]["opt_out_new_regime"] = True
    rep = build_return(sample)["report"]
    # HRA eligible = min(192000, 180000-48000, 192000) = 132,000; + 20,000 (10(14)(i))
    # salary income = 960000 - 152000 exempt - 50000 SD - 2400 professional tax
    assert rep["salary_income"] == 755_600
    assert rep["gross_total_income_excl_ltcg"] == 755_600 + 23_500
    # deductions: 80C 150k + 80D 20k + 80CCD2 96k = 266,000
    assert rep["deductions_eligible"]["TotalChapVIADeductions"] == 266_000
    assert rep["total_income"] == round(879_100 - 266_000, -1)
    # taxable excl LTCG = 513,100 -> 12,500 + 20% of 13,100 = 15,120
    assert rep["tax"]["TotalTaxPayable"] == 15_120


def test_root_structure_and_versions(sample):
    doc = build_return(sample)["json"]
    itr1 = doc["ITR"]["ITR1"]
    form = itr1["Form_ITR1"]
    assert form["FormName"] == "ITR-1"
    assert form["AssessmentYear"] == "2026"
    assert form["SchemaVer"] == "Ver1.0"
    assert form["FormVer"] == "Ver1.0"
    assert re.match(r"SW\d{8}", itr1["CreationInfo"]["SWCreatedBy"])
    for key in ("PersonalInfo", "FilingStatus", "ITR1_IncomeDeductions",
                "ITR1_TaxComputation", "TaxPaid", "Refund", "Verification"):
        assert key in itr1


def test_hra_schedule_present_old_regime(sample):
    sample["filing"]["opt_out_new_regime"] = True
    doc = build_return(sample)["json"]
    hra = doc["ITR"]["ITR1"]["ScheduleEA10_13A"]
    assert hra["EligbleExmpAllwncUs13A"] == 132_000
    assert hra["ActlRentPaid10Per"] == 180_000 - 48_000


def test_no_hra_in_new_regime(sample):
    doc = build_return(sample)["json"]
    assert "ScheduleEA10_13A" not in doc["ITR"]["ITR1"]
    allws = doc["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["AllwncExemptUs10"]
    assert allws["TotalAllwncExemptUs10"] == 0


def test_80c_combined_cap(sample):
    sample["filing"]["opt_out_new_regime"] = True
    sample["deductions"]["80c"] = {"amount": 150000}
    sample["deductions"]["80ccc"] = {"amount": 40000,
                                     "funds": [{"type": "PRAN", "name": "LIC", "amount": 40000}]}
    sample["deductions"]["80ccd1"] = 60000
    res = build_return(sample)
    via = res["json"]["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["DeductUndChapVIA"]
    assert via["Section80C"] + via["Section80CCC"] + via["Section80CCDEmployeeOrSE"] == 150_000


def test_house_property_setoff_rules(sample):
    sample["house_properties"] = [{
        "occupancy": "S", "address_line": "A-1", "city": "Mumbai", "state_code": "19",
        "pin_code": 400001, "interest_on_borrowed_capital": 250000,
    }]
    # old regime: capped at 2L -> HP loss -2,00,000 reduces GTI
    sample["filing"]["opt_out_new_regime"] = True
    rep_old = build_return(sample)["report"]
    assert rep_old["hp_income"] == -200_000
    # new regime: SOP interest not allowed; no loss set-off
    sample["filing"]["opt_out_new_regime"] = False
    rep_new = build_return(sample)["report"]
    assert rep_new["hp_income"] == 0


def test_let_out_property_math(sample):
    sample["house_properties"] = [{
        "occupancy": "L", "address_line": "A-1", "city": "Pune", "state_code": "19",
        "pin_code": 411001, "annual_lettable_value": 300000, "municipal_taxes": 10000,
        "interest_on_borrowed_capital": 150000,
    }]
    rep = build_return(sample)["report"]
    # NAV 290,000 - 30% (87,000) - 150,000 = 53,000
    assert rep["hp_income"] == 53_000
    rd = build_return(sample)["json"]["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["PropertyDetails"][0]["Rentdetails"]
    assert rd["ThirtyPercentOfBalance"] == 87_000


def test_ltcg_capped_at_125000(sample):
    sample["ltcg112a"] = {"sale_consideration": 500000, "cost_of_acquisition": 200000}
    rep = build_return(sample)
    assert rep["report"]["ltcg112a"] == 125_000
    assert any("ITR-2" in w for w in rep["report"]["warnings"])


def test_tds_totals_flow_into_taxes_paid(sample):
    sample["tax_paid"]["challans"] = [
        {"bsr_code": "0012345", "date": "2026-03-10", "challan_no": 1234, "amount": 5000},
        {"bsr_code": "0012345", "date": "2026-07-10", "challan_no": 1234, "amount": 8000},
    ]
    rep = build_return(sample)["report"]
    tp = rep["tax"]["TaxesPaid"]
    assert tp["TDS"] == 36_500
    assert tp["AdvanceTax"] == 5_000
    assert tp["SelfAssessmentTax"] == 8_000
    assert tp["TotalTaxesPaid"] == 36_500 + 13_000


def test_senior_citizen_old_regime(sample):
    sample["personal"]["dob"] = "1960-02-10"
    sample["filing"]["opt_out_new_regime"] = True
    sample["deductions"]["80ttb"] = 50000
    rep = build_return(sample)["report"]
    assert rep["deductions_eligible"]["Section80TTB"] == min(50_000, 23_500)  # interest income
    assert rep["deductions_eligible"]["Section80TTA"] == 0


def test_80tta_not_for_senior(sample):
    sample["personal"]["dob"] = "1960-02-10"
    sample["filing"]["opt_out_new_regime"] = True
    sample["deductions"]["80tta"] = 8000
    rep = build_return(sample)["report"]
    assert rep["deductions_eligible"]["Section80TTA"] == 0


import re  # noqa: E402
