"""Tests for the CBDT validation-rules engine."""

import copy
import json
import os

import pytest

from itr1.builder import build_return
from itr1.validator import validate_all, validate_rules, summarize

SAMPLE = os.path.join(os.path.dirname(__file__), os.pardir,
                      "samples", "sample_taxpayer.json")


@pytest.fixture()
def sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def _blocking(doc):
    return [i for i in validate_rules(doc) if i.category in ("A", "SCHEMA")]


def test_clean_return_has_no_blocking_issues(sample):
    doc = build_return(sample)["json"]
    assert _blocking(doc) == []


def test_detects_violation_of_rule_1(sample):
    sample["filing"]["opt_out_new_regime"] = True
    doc = build_return(sample)["json"]
    via = doc["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["DeductUndChapVIA"]
    via["Section80C"] = 150_000
    via["Section80CCC"] = 100_000          # now 80C+80CCC > 1.5L -> rule 1
    issues = validate_rules(doc)
    assert any(i.rule == "Rule 1" for i in issues)


def test_detects_wrong_refund_rule_105(sample):
    doc = build_return(sample)["json"]
    doc["ITR"]["ITR1"]["Refund"]["RefundDue"] += 1000   # tamper
    issues = validate_rules(doc)
    assert any("Rule 105" in i.rule for i in issues)


def test_detects_new_regime_deduction_rule_146(sample):
    doc = build_return(sample)["json"]      # new regime
    via = doc["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["DeductUndChapVIA"]
    via["Section80C"] = 50_000
    issues = validate_rules(doc)
    assert any(i.rule == "Rule 146" for i in issues)


def test_detects_87a_over_claim_rule_191(sample):
    sample["salary"]["salary_17_1"] = 2_200_000   # pushes income beyond rebate limit
    doc = build_return(sample)["json"]
    tc = doc["ITR"]["ITR1"]["ITR1_TaxComputation"]
    tc["Rebate87A"] = 60_000                # illegal at this income
    issues = validate_rules(doc)
    assert any("Rule 191" in i.rule or i.rule == "Rule 191a" for i in issues)


def test_detects_tds_mismatch_rule_108(sample):
    doc = build_return(sample)["json"]
    doc["ITR"]["ITR1"]["TaxPaid"]["TaxesPaid"]["TDS"] += 500
    issues = validate_rules(doc)
    assert any(i.rule == "Rule 108" for i in issues)


def test_detects_bad_gti_rule_22(sample):
    doc = build_return(sample)["json"]
    doc["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["GrossTotIncome"] += 10_000
    issues = validate_rules(doc)
    assert any(i.rule.startswith("Rule 22") or "160" in i.rule or "174" in i.rule
               for i in issues)


def test_house_property_math_checked(sample):
    sample["house_properties"] = [{
        "occupancy": "L", "address_line": "A-1", "city": "Pune", "state_code": "19",
        "pin_code": 411001, "annual_lettable_value": 300000, "municipal_taxes": 10000,
        "interest_on_borrowed_capital": 150000,
    }]
    doc = build_return(sample)["json"]
    rd = doc["ITR"]["ITR1"]["ITR1_IncomeDeductions"]["PropertyDetails"][0]["Rentdetails"]
    rd["ThirtyPercentOfBalance"] = 50_000   # should be 87,000 -> rule 43
    issues = validate_rules(doc)
    assert any(i.rule == "Rule 43" for i in issues)


def test_category_b_tds_section_flag(sample):
    sample["tax_paid"]["tds_others"] = [{
        "tan": "MUMS67890F", "deductor_name": "Casino", "section": "94B",
        "amount_paid": 50000, "deduction_year": "2025", "tds_deducted": 15000,
        "tds_claimed": 15000,
    }]
    doc = build_return(sample)["json"]
    issues = validate_rules(doc)
    assert any(i.category == "B" and "B3" in i.rule for i in issues)


def test_category_d_form_10e_notice(sample):
    sample["relief_89"] = 5000
    doc = build_return(sample)["json"]
    issues = validate_rules(doc)
    assert any(i.category == "D" for i in issues)


def test_schema_validation_catches_bad_pan(sample):
    doc = build_return(sample)["json"]
    doc["ITR"]["ITR1"]["PersonalInfo"]["PAN"] = "XX"
    issues = validate_all(doc)
    assert any(i.category == "SCHEMA" for i in issues)


def test_summary_counts(sample):
    doc = build_return(sample)["json"]
    rep = summarize(validate_all(doc))
    assert rep["ok"] is True
