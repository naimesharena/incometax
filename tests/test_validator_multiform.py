"""Validator v2: form-aware dispatch (ITR-2 / ITR-4 schema + generic rules)."""

import json
import os

import pytest

from itr1.forms import build_for_form
from itr1.validator import (Issue, detect_form_in_doc, summarize, validate_all,
                            validate_rules_generic)

HERE = os.path.dirname(__file__)


def _doc(sample_name, form):
    with open(os.path.join(HERE, os.pardir, "samples", sample_name),
              encoding="utf-8") as fh:
        data = json.load(fh)
    return build_for_form(data, form)["json"]


@pytest.fixture(scope="module")
def doc2():
    return _doc("sample_itr2.json", "ITR2")


@pytest.fixture(scope="module")
def doc4():
    return _doc("sample_itr4.json", "ITR4")


def test_clean_builds_have_zero_issues(doc2, doc4):
    for form, doc in (("ITR2", doc2), ("ITR4", doc4)):
        issues = validate_all(doc, form=form)
        assert issues == [], [str(i) for i in issues[:5]]


def test_form_autodetected_from_document(doc2, doc4):
    assert detect_form_in_doc(doc2) == "ITR2"
    assert detect_form_in_doc(doc4) == "ITR4"
    assert detect_form_in_doc({}) is None
    assert summarize(validate_all(doc4))["ok"] is True     # no form= passed


def test_schema_layer_routes_to_right_form(doc2, doc4):
    doc2["ITR"]["ITR2"]["PartA_GEN1"]["PersonalInfo"]["PAN"] = "XX"
    issues = validate_all(doc2, form="ITR2", rules=False)
    assert any(i.category == "SCHEMA" for i in issues)
    doc4["ITR"]["ITR4"]["PersonalInfo"]["PAN"] = "XX"
    issues = validate_all(doc4, form="ITR4", rules=False)
    assert any(i.category == "SCHEMA" for i in issues)


def _rules(issues):
    return {i.rule for i in issues}


def test_generic_rules_fire_on_bad_identity(doc2):
    node = doc2["ITR"]["ITR2"]["PartA_GEN1"]["PersonalInfo"]
    node["PAN"] = "bad"
    node["Address"]["MobileNo"] = 123
    node["Address"]["EmailAddress"] = ""
    issues = validate_rules_generic(doc2, "ITR2")
    assert {"Rule MPAN", "Rule MMOB", "Rule MEMAIL"} <= _rules(issues)
    assert any(i.category == "A" for i in issues)


def test_generic_rules_refund_needs_bank(doc4):
    doc4["ITR"]["ITR4"]["Refund"]["RefundDue"] = 5000
    doc4["ITR"]["ITR4"]["Refund"]["BankAccountDtls"]["AddtnlBankDetails"] = []
    issues = validate_rules_generic(doc4, "ITR4")
    assert "Rule MBANK" in _rules(issues)


def test_generic_rules_80cce_ceiling(doc2):
    via = doc2["ITR"]["ITR2"]["ScheduleVIA"]["UsrDeductUndChapVIA"]
    via["Section80C"] = 150000
    via["Section80CCC"] = 25000
    issues = validate_rules_generic(doc2, "ITR2")
    assert "Rule 80CCE" in _rules(issues)


def test_generic_rules_rebate_87a_regimes(doc2, doc4):
    # new regime: TI > 12L with rebate → violation (sample_itr2 TI is ~40L)
    comp2 = doc2["ITR"]["ITR2"]["PartB_TTI"]["ComputationOfTaxLiability"]
    comp2["Rebate87A"] = 60000
    issues = validate_rules_generic(doc2, "ITR2")
    assert "Rule 87A" in _rules(issues)
    # old regime → cap 12,500 (sample_itr4 old-regime TI is ~10L → any rebate is wrong)
    fs4 = doc4["ITR"]["ITR4"]
    fs4["FilingStatus"]["OptOutNewTaxRegime"] = "Y"
    fs4["TaxComputation"]["Rebate87A"] = 20000
    issues = validate_rules_generic(doc4, "ITR4")
    assert "Rule 87A" in _rules(issues)


def test_generic_rules_tax_paid_and_balance(doc4):
    paid = doc4["ITR"]["ITR4"]["TaxPaid"]["TaxesPaid"]
    paid["TDS"] = paid["TDS"] + 1000           # break the break-up tally
    issues = validate_rules_generic(doc4, "ITR4")
    assert "Rule TPAID" in _rules(issues)
    assert "Rule BAL" in _rules(issues)


def test_generic_rules_verification_mismatch(doc2):
    ver = doc2["ITR"]["ITR2"]["Verification"]
    ver["Declaration"]["AssesseeVerPAN"] = "AAAAA1111A"
    ver["Place"] = ""
    issues = validate_rules_generic(doc2, "ITR2")
    assert "Rule VPAN" in _rules(issues)
    assert "Rule MPLACE" in _rules(issues)


def test_unknown_form_flags_not_implemented():
    issues = validate_all({"ITR": {"ITR3": {}}}, include_schema=False)
    assert any(i.category == "D" and "ITR3" in i.message for i in issues)
