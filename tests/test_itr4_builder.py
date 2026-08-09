"""ITR-4 builder: presumptive-income engine + schema validation."""

import copy
import json
import os

import pytest

from itr1.itr4_builder import build_itr4, compute_presumptive
from itr1.schemakit import validate_with_schema

SAMPLE = os.path.join(os.path.dirname(__file__), "..", "samples", "sample_itr4.json")


def load_sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def test_sample_is_schema_valid():
    out = build_itr4(load_sample())
    assert validate_with_schema("ITR4", out["json"]) == []


def test_presumptive_rates():
    pres = compute_presumptive(load_sample())
    # 44AD: 6% x (3,400,000 + 200,000) + 8% x 650,000 = 216,000 + 52,000
    assert pres["44ad_income"] == 268000
    # 44ADA: 50% x 620,000
    assert pres["44ada_income"] == 310000
    # 44AE: light 7,500 x 12 + heavy 1,000 x 16t x 7m = 90,000 + 112,000
    assert pres["44ae_income"] == 202000
    assert pres["total_income"] == 780000


def test_schedule_bp_shapes():
    out = build_itr4(load_sample())
    bp = out["json"]["ITR"]["ITR4"]["ScheduleBP"]
    assert bp["PersumptiveInc44AD"]["PersumptiveInc44AD6Per"] == 216000
    assert bp["PersumptiveInc44AD"]["PersumptiveInc44AD8Per"] == 52000
    assert bp["PersumptiveInc44ADA"]["TotPersumptiveInc44ADA"] == 310000
    ae = bp["GoodsDtlsUs44AE"]
    assert ae[0]["PresumptiveIncome"] == 90000
    assert ae[1]["PresumptiveIncome"] == 112000
    assert ae[1]["TonnageCapacity"] == 16
    assert bp["FinanclPartclrOfBusiness"]["TotalAssets"] == 943000
    # ScheduleIT carries the challan; roots with empty arrays are omitted
    itr4 = out["json"]["ITR"]["ITR4"]
    assert itr4["ScheduleIT"]["TotalTaxPayments"] == 20000
    assert "TDSonSalaries" not in itr4
    assert itr4["TDSonOthThanSals"]["TotalTDSonOthThanSals"] == 3000


def test_declared_below_presumptive_is_overridden():
    data = load_sample()
    data["presumptive"]["44ad"][0]["income_declared"] = 200000  # < 268,000 presumptive
    out = build_itr4(data)
    assert out["report"]["presumptive"]["44ad_income"] == 268000
    assert any("below presumptive" in w for w in out["report"]["warnings"])
    assert validate_with_schema("ITR4", out["json"]) == []


def test_turnover_cap_warning():
    data = load_sample()
    data["presumptive"]["44ad"][0]["gross_turnover"] = 32000000
    out = build_itr4(data)
    assert any("3,00,00,000" in w for w in out["report"]["warnings"])


def test_old_regime_form10iea_flag():
    data = load_sample()
    data["filing"]["opt_out_new_regime"] = True
    out = build_itr4(data)
    fs = out["json"]["ITR"]["ITR4"]["FilingStatus"]
    assert fs["Form10IEAEarlierAYOldRegime"] == "Y"
    assert fs["F10IEACurrAYOldRegime"] == "Y"
    assert any("10-IEA" in w for w in out["report"]["warnings"])
    assert validate_with_schema("ITR4", out["json"]) == []


def test_legacy_112a_box():
    data = load_sample()
    data["ltcg112a"] = {"sale_consideration": 250000, "cost_of_acquisition": 160000}
    out = build_itr4(data)
    lt = out["json"]["ITR"]["ITR4"]["LTCG112A"]
    assert lt["LongCap112A"] == 90000  # fully inside the exemption → no tax
    assert out["report"]["tax"]["Working"]["ltcg_tax"] == 0
    assert validate_with_schema("ITR4", out["json"]) == []


def test_income_ceiling_warning():
    data = load_sample()
    # declaring ₹49L (well above the 50% presumptive) pushes total income past ₹50L
    data["presumptive"]["44ada"][0]["income_declared"] = 4900000
    out = build_itr4(data)
    assert out["report"]["total_income"] > 5000000
    assert any("50,00,000" in w for w in out["report"]["warnings"])
