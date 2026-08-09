"""Tests for the detailed computation report + PDF/DOCX exports."""

import json
import os

import pytest

from itr1.report import build_computation
from itr1.compfile import computation_docx, computation_pdf

SAMPLE = os.path.join(os.path.dirname(__file__), os.pardir,
                      "samples", "sample_taxpayer.json")


@pytest.fixture()
def sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def test_computation_structure_both_regimes(sample):
    comp = build_computation(sample)
    assert comp["meta"]["pan"] == "BQKPS1234F"
    assert comp["meta"]["assessment_year"] == "2026-27"
    for key in ("new", "old"):
        r = comp[key]
        for field in ("salary", "via_rows", "slab_rows", "tax_payable", "rebate",
                      "interest", "taxes_paid", "tot_tax_plus_interest"):
            assert field in r
    assert comp["comparison"]["recommended"] == "new"
    assert comp["comparison"]["new_total"] == 0
    assert comp["comparison"]["saving"] == comp["old"]["tot_tax_plus_interest"]


def test_computation_numbers_new_regime(sample):
    r = build_computation(sample)["new"]
    assert r["total_income"] == 912_500
    assert r["slab_tax"] == 21_250
    assert r["rebate"] == 21_250
    assert r["taxes_paid"]["total"] == 36_500
    assert r["refund_due"] == 36_500
    assert "12,00,000" in r["rebate_note"]


def test_computation_numbers_old_regime(sample):
    r = build_computation(sample)["old"]
    assert r["total_income"] == 613_100
    assert r["slab_tax"] == 15_120
    via = {row["section"]: row["eligible"] for row in r["via_rows"]}
    assert via["80C"] == 150_000
    assert via["80D"] == 20_000
    assert via["80CCD(2)"] == 96_000


def test_computation_works_with_minimal_input():
    # the original crash report: salary filled, DOB still empty
    comp = build_computation({"salary": {"salary_17_1": 900000}})
    assert comp["new"]["total_income"] == 825_000
    assert comp["new"]["slab_tax"] == 22_500


def test_interest_detail_present(sample):
    sample["filing"]["filing_date"] = "2026-09-30"
    sample["filing"]["return_section"] = 12
    comp = build_computation(sample)
    new = comp["new"]
    labels = [i["label"] for i in new["interest"]]
    assert any("234F" in l for l in labels)
    fee = next(i for i in new["interest"] if "234F" in i["label"])
    assert fee["amount"] == 5_000            # total income > 5 lakh
    assert "due date" in fee["note"].lower() or "after" in fee["note"].lower()


def test_pdf_generation(sample):
    payload = computation_pdf(build_computation(sample))
    assert payload.startswith(b"%PDF")
    assert len(payload) > 5_000


def test_docx_generation(sample):
    payload = computation_docx(build_computation(sample))
    assert payload.startswith(b"PK")         # docx is a zip container
    assert len(payload) > 5_000


def test_pdf_and_docx_with_complex_return():
    data = {
        "personal": {"first_name": "Meera", "last_name": "Iyer", "pan": "AXZPK5678Q",
                     "dob": "1958-03-20", "employer_category": "PESG",
                     "address": {"residence_no": "7", "locality_or_area": "HSR",
                                 "city_or_town_or_district": "Bengaluru",
                                 "state_code": "14", "pin_code": 560102}},
        "filing": {"return_section": 11, "opt_out_new_regime": False,
                   "filing_date": "2026-07-15"},
        "salary": {"salary_17_1": 1400000},
        "house_properties": [{
            "occupancy": "L", "address_line": "9th Main", "city": "Blr",
            "state_code": "14", "pin_code": 560102, "annual_lettable_value": 360000,
            "municipal_taxes": 12000, "interest_on_borrowed_capital": 180000,
            "loans": [{"lender_type": "B", "lender_name": "HDFC", "loan_account_no": "X1",
                       "loan_date": "2020-06-15", "loan_amount": 4500000,
                       "outstanding": 3200000, "interest": 180000}]}],
        "other_sources": [{"nature": "SAV", "amount": 22000},
                          {"nature": "FAP", "amount": 180000}],
        "deductions": {"80ccd2": 196000},
        "tax_paid": {"challans": [
            {"bsr_code": "0012345", "date": "2025-09-10", "challan_no": 1, "amount": 15000}]},
        "bank_accounts": [{"ifsc": "HDFC0000456", "bank_name": "HDFC",
                            "account_no": "50100223344", "account_type": "SB",
                            "use_for_refund": True}],
        "verification": {"name": "Meera Iyer", "father_name": "R Iyer", "place": "Bengaluru"},
    }
    comp = build_computation(data)
    # 360,000 - 12,000 taxes = 348,000; - 30% (104,400) - 180,000 interest = 63,600
    assert comp["new"]["house_properties"][0]["income"] == 63_600
    assert computation_pdf(comp).startswith(b"%PDF")
    assert computation_docx(comp).startswith(b"PK")
    # family pension 57(iia) deduction visible in other sources rows
    labels = " ".join(str(x.get("label", "")) + str(x.get("note", ""))
                      for x in comp["new"]["other_sources"]["rows"])
    assert "57(iia)" in labels
