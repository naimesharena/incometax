"""Residential-status behaviour — verified against the official artefacts.

Evidence base (all extracted under extracted/xlsm/ + the validation PDFs):

* ITR-1  EfilingCommon.bas: "ITR 1 is for individuals being a resident
  (other than not ordinarily resident) … Please file another ITR" → the
  ITR-1 builder hard-blocks NRI *and* RNOR filers.
* ITR-4  Sheet1 Worksheet_Change (`Lock44AD`): when ResidentialStatus1
  starts with "NR" (NRI only — "NOR" does not match) the utility locks and
  CLEARS the 44AD (BP_E1a/E1b/E2a/E2b) and 44ADA (BP_E3_GR, BP_E4_PI44ADA)
  cells → the ITR-4 builder raises for NRI + 44AD/44ADA rows.
* ITR-4  validation rule 212: "HUF is not eligible to claim presumptive
  income u/s 44ADA."
* ITR-4  Sheet4.cls: NR goods-carriage filers only get an advisory
  ("returns for income of non-residents are required to be filed in
  ITR 3/5") → 44AE stays allowed with a warning.
* ITR-2  'Tax Calculated'!B26/C34: senior/super-senior slabs apply only
  when resStatus is "RES" or "NOR" → NRI seniors get the general slab.
* ITR-2  Part B-TI TTI!P80 (`Rebate87Aformula_new`) + portal validation
  rules 533/534: rebate u/s 87A is 0 for non-residents and non-individuals.
"""

import copy
import json
import os

import pytest

from itr1 import tax_engine as T
from itr1.builder import build_return
from itr1.errors import DataError
from itr1.forms import detect_itr_form
from itr1.itr2_builder import build_itr2
from itr1.itr4_builder import build_itr4
from itr1.schemakit import validate_with_schema

ROOT = os.path.join(os.path.dirname(__file__), "..")
SAMPLE1 = os.path.join(ROOT, "samples", "sample_taxpayer.json")
SAMPLE2 = os.path.join(ROOT, "samples", "sample_itr2.json")
SAMPLE4 = os.path.join(ROOT, "samples", "sample_itr4.json")


def load(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def itr4_sample():
    return load(SAMPLE4)


# ---------------------------------------------------------------- tax engine

def test_87a_denied_to_non_resident():
    kw = dict(total_tax_payable=20000, total_income=600000,
              income_excl_ltcg=600000, opt_out_new_regime=False)
    assert T.compute_rebate_87a(**kw) == 20000            # RES: 60000 cap, tax smaller
    assert T.compute_rebate_87a(residential_status="NOR", **kw) == 20000
    assert T.compute_rebate_87a(residential_status="NRI", **kw) == 0.0


def test_87a_denied_to_huf():
    kw = dict(total_tax_payable=20000, total_income=600000,
              income_excl_ltcg=600000, opt_out_new_regime=False)
    assert T.compute_rebate_87a(status="H", **kw) == 0.0


def test_nri_senior_gets_general_old_regime_slab():
    # born 1955 → 70 years old; old regime resident senior pays 0 up to 3L,
    # an NRI of the same age falls back to the general slab (2.5L).
    from itr1 import constants as C
    senior_res = T.normal_slab_tax(300000, True, "1955-06-01", "RES")
    senior_nor = T.normal_slab_tax(300000, True, "1955-06-01", "NOR")
    senior_nri = T.normal_slab_tax(300000, True, "1955-06-01", "NRI")
    general = T.slab_tax(300000, C.OLD_REGIME_SLABS_GENERAL)
    assert senior_res == 0.0
    assert senior_nor == 0.0                    # RNOR is resident u/s 6
    assert senior_nri == general == 2500.0      # 5% of (3L - 2.5L)


def test_tax_v2_threads_residency_into_rebate_note():
    out = T.compute_tax_v2(total_income=600000, opt_out_new_regime=False,
                           dob="1990-01-01", filing_date="2026-07-31",
                           return_section=11, residential_status="NRI")
    assert out["Rebate87A"] == 0
    assert "533" in out["Working"]["rebate_note"]
    out2 = T.compute_tax_v2(total_income=600000, opt_out_new_regime=False,
                            dob="1990-01-01", filing_date="2026-07-31",
                            return_section=11)
    assert out2["Rebate87A"] > 0


# ------------------------------------------------------------------ radar

def test_radar_nri_blocks_itr1_and_itr4():
    data = itr4_sample()
    data["personal"]["residential_status"] = "NRI"
    det = detect_itr_form(data)
    assert det["residential_status"] == "NRI"
    assert not det["eligible"]["ITR1"]["ok"]
    assert not det["eligible"]["ITR4"]["ok"]
    assert not det["eligible"]["ITR2"]["ok"]   # business income → ITR-3
    assert det["recommended"] == "ITR3"
    assert any("44AD" in r or "non-resident" in r for r in det["eligible"]["ITR4"]["reasons"])


def test_radar_nor_blocks_itr1_keeps_itr4():
    data = itr4_sample()
    data["personal"]["residential_status"] = "NOR"
    det = detect_itr_form(data)
    assert not det["eligible"]["ITR1"]["ok"]
    assert det["eligible"]["ITR4"]["ok"]       # RNOR is resident u/s 6
    assert det["recommended"] == "ITR4"        # presumptive business stays


def test_radar_nor_salary_only_goes_itr2():
    data = load(SAMPLE1)
    data["personal"]["residential_status"] = "NOR"
    det = detect_itr_form(data)
    assert det["recommended"] == "ITR2"
    assert not det["eligible"]["ITR1"]["ok"]


def test_radar_legacy_rnor_flag_maps_to_nor():
    data = itr4_sample()
    data["eligibility"] = {"rnor": True}
    det = detect_itr_form(data)
    assert det["residential_status"] == "NOR"
    assert det["eligible"]["ITR4"]["ok"]
    assert not det["eligible"]["ITR1"]["ok"]


def test_radar_nri_salary_only_goes_itr2():
    data = load(SAMPLE1)
    data["personal"]["residential_status"] = "NRI"
    det = detect_itr_form(data)
    assert det["recommended"] == "ITR2"


# ----------------------------------------------------------------- ITR-1

def test_itr1_builder_rejects_nri_and_nor():
    for status in ("NRI", "NOR"):
        data = load(SAMPLE1)
        data["personal"]["residential_status"] = status
        with pytest.raises(DataError) as ei:
            build_return(data)
        assert "ITR-1 is only for individuals who are resident" in str(ei.value)


def test_itr1_builder_res_builds_fine():
    out = build_return(load(SAMPLE1))
    assert validate_with_schema("ITR1", out["json"]) == []


# ----------------------------------------------------------------- ITR-4

def test_itr4_nri_with_44ad_blocked_like_utility_lock44ad():
    data = itr4_sample()
    data["personal"]["residential_status"] = "NRI"
    with pytest.raises(DataError) as ei:
        build_itr4(data)
    assert "44AD / 44ADA is not available to a non-resident" in str(ei.value)


def test_itr4_nri_with_44ada_alone_blocked():
    data = itr4_sample()
    data["presumptive"]["44ad"] = []
    data["presumptive"]["44ae"] = {}
    data["personal"]["residential_status"] = "NRI"
    with pytest.raises(DataError):
        build_itr4(data)


def test_itr4_nri_with_44ae_only_allowed_with_warning():
    """The utility locks only the 44AD/44ADA cells; goods carriages stay open
    (Sheet4.cls merely shows an advisory for NR TDS cases)."""
    data = itr4_sample()
    data["presumptive"]["44ad"] = []
    data["presumptive"]["44ada"] = []
    data["personal"]["residential_status"] = "NRI"
    out = build_itr4(data)
    assert any("Non-Resident" in w and "ITR-3" in w for w in out["report"]["warnings"])
    assert validate_with_schema("ITR4", out["json"]) == []
    # and 87A is 0 for the non-resident
    cmpblk = out["json"]["ITR"]["ITR4"]["TaxComputation"]
    assert cmpblk["Rebate87A"] == 0


def test_itr4_nor_with_44ad_44ada_allowed():
    data = itr4_sample()
    data["personal"]["residential_status"] = "NOR"
    out = build_itr4(data)
    assert validate_with_schema("ITR4", out["json"]) == []


def test_itr4_huf_with_44ada_blocked_rule212():
    data = itr4_sample()
    data["personal"]["status"] = "H"
    with pytest.raises(DataError) as ei:
        build_itr4(data)
    assert "44ADA" in str(ei.value) and "HUF" in str(ei.value)


def test_itr4_huf_without_44ada_builds():
    data = itr4_sample()
    data["presumptive"]["44ada"] = []
    data["personal"]["status"] = "H"
    out = build_itr4(data)
    assert validate_with_schema("ITR4", out["json"]) == []


# ----------------------------------------------------------------- ITR-2

def test_itr2_nri_schema_valid_rebate_zero():
    data = load(SAMPLE2)
    data["personal"]["residential_status"] = "NRI"
    out = build_itr2(data)
    assert validate_with_schema("ITR2", out["json"]) == []
    pi = out["json"]["ITR"]["ITR2"]["PartA_GEN1"]
    assert pi["FilingStatus"]["ResidentialStatus"] == "NRI"
    # sample_itr2 total income is below 12L → resident would get 87A; NRI must not
    assert out["report"]["tax"]["Rebate87A"] == 0


def test_itr2_res_unchanged_baseline():
    out = build_itr2(load(SAMPLE2))
    assert validate_with_schema("ITR2", out["json"]) == []
