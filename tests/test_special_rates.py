"""Special-rate income (Schedule SI) — CG + OS special-rate buckets.

The user-facing requirement: "capital gain and other special rate charged
income … apply all logic as per macro".  The wiring here mirrors the
official utilities end-to-end:

* ``SecCode`` values verbatim from the ITR-2 schema (Schedule SI).
* OS special-rate nodes of Schedule OS (115BB / 115BBJ / 115BBE) with the
  quarterly date ranges (Q5 split) used for the 234C computation.
* Tax engine: flat rates (30% / 30% / 60%), and a separately computed FLAT
  25% surcharge on the 115BBE tax — the utility's TaxCalc sheet labels the
  rows exactly "Add surcharge on 115BBE" / "Surcharge on other than 115BBE".
"""

import copy
import json
import os

import pytest

from itr1 import special_rates as SR
from itr1 import tax_engine as T
from itr1.errors import DataError
from itr1.itr2_builder import build_itr2
from itr1.schemakit import load_schema, validate_with_schema
from itr1.validator import summarize, validate_all

ROOT = os.path.join(os.path.dirname(__file__), "..")
SAMPLE2 = os.path.join(ROOT, "samples", "sample_itr2.json")


def load_sample():
    with open(SAMPLE2, encoding="utf-8") as fh:
        return json.load(fh)


def with_special(data, **kw):
    data.setdefault("other_sources_special", {}).update(kw)
    return data


def _os_schedule(doc):
    return doc["ITR"]["ITR2"]["ScheduleOS"]["IncOthThanOwnRaceHorse"]


def _si_rows(doc):
    return doc["ITR"]["ITR2"]["ScheduleSI"]["SplCodeRateTax"]


# ------------------------------------------------------------- code tables

def test_sec_codes_cover_schema_description():
    """Every SecCode the official schema documents (in the SecCode
    description of Schedule SI) must be known to the module — parity with
    the dropdown the official utility offers."""
    import re
    schema = load_schema("ITR2")
    desc = (schema["definitions"]["ScheduleSI"]["properties"]["SplCodeRateTax"]
            ["items"]["properties"]["SecCode"]["description"])
    codes = set()
    for seg in desc.split(";"):
        m = re.match(r"\s*([0-9A-Za-z_]+)\s*[:\-]", seg)
        if m:
            codes.add(m.group(1))
    assert "5BBE" in codes and "PTI_5Ea" in codes      # sanity of the parse
    missing = codes - set(SR.SEC_CODES)
    assert not missing, f"SecCodes in schema but unknown to special_rates: {sorted(missing)}"


def test_wired_os_codes_have_dedicated_nodes():
    assert SR.OS_SPECIAL_INPUTS["lottery_115bb"][0] == "5BB"
    assert SR.OS_SPECIAL_INPUTS["online_games_115bbj"][0] == "5BBJ"
    assert SR.OS_SPECIAL_INPUTS["unexplained_115bbe"][0] == "5BBE"
    assert SR.SEC_CODES["5BB"][1] == 30.0
    assert SR.SEC_CODES["5BBJ"][1] == 30.0
    assert SR.SEC_CODES["5BBE"][1] == 60.0


def test_unknown_special_key_rejected_with_guidance():
    with pytest.raises(DataError) as ei:
        SR.os_special_rows({"other_sources_special": {"patent_115bbf": 1000}})
    msg = str(ei.value)
    assert "unknown special-rate income key" in msg
    assert "lottery_115bb" in msg       # advertises what IS supported


# ------------------------------------------------------- ITR-2 integration

def test_os_special_nodes_and_split():
    data = with_special(load_sample(), lottery_115bb=100000,
                        online_games_115bbj=50000, unexplained_115bbe=200000)
    out = build_itr2(data)
    doc = out["json"]
    assert validate_with_schema("ITR2", doc) == []

    sched = _os_schedule(doc)
    assert sched["LtryPzzlChrgblUs115BB"] == 100000
    assert sched["IncChrgblUs115BBJ"] == 50000
    assert sched["IncChrgblUs115BBE"] == 200000
    assert sched["IncChargeableSpecialRates"] == 350000
    # special buckets are carved out of the normal-rate balance
    assert sched["BalanceNoRaceHorse"] + 350000 == \
        doc["ITR"]["ITR2"]["ScheduleOS"]["TotOthSrcNoRaceHorse"]

    # quarterly date ranges: annual-only figures sit in the Q5 window
    lot = doc["ITR"]["ITR2"]["ScheduleOS"]["IncFrmLottery"]["DateRange"]
    assert lot["Up16Of3To31Of3"] == 100000
    assert lot["Upto15Of6"] == lot["Upto15Of9"] == \
        lot["Up16Of9To15Of12"] == lot["Up16Of12To15Of3"] == 0
    games = doc["ITR"]["ITR2"]["ScheduleOS"]["IncFrmOnGames"]["DateRange"]
    assert games["Up16Of3To31Of3"] == 50000

    # Part B-TI splits OS income into slab-rate vs special-rate parts
    ti = doc["ITR"]["ITR2"]["PartB-TI"]
    inc_os = ti["IncFromOS"]
    assert inc_os["IncChargblSplRate"] == 350000
    assert inc_os["TotIncFromOS"] == \
        inc_os["OtherSrcThanOwnRaceHorse"] + inc_os["IncChargblSplRate"]


def test_si_rows_carry_cg_and_os_buckets():
    data = with_special(load_sample(), lottery_115bb=100000,
                        online_games_115bbj=50000, unexplained_115bbe=200000)
    rows = _si_rows(build_itr2(data)["json"])
    by_code = {r["SecCode"]: r for r in rows}
    # CG buckets from the sample (equity 111A, 112A, land 112)
    assert by_code["1A"]["SplRatePercent"] == 20
    assert by_code["2A"]["SplRatePercent"] == 12.5
    # OS buckets, flat rates
    assert by_code["5BB"]["SplRateInc"] == 100000
    assert by_code["5BB"]["SplRateIncTax"] == 30000
    assert by_code["5BBJ"]["SplRateInc"] == 50000
    assert by_code["5BBJ"]["SplRateIncTax"] == 15000
    assert by_code["5BBE"]["SplRateInc"] == 200000
    assert by_code["5BBE"]["SplRateIncTax"] == 120000


def test_tax_engine_special_rates_and_115bbe_surcharge():
    rows = SR.os_special_rows(with_special(
        {}, lottery_115bb=100000, online_games_115bbj=50000,
        unexplained_115bbe=200000))
    assert [(r["code"], r["amount"], r["tax"]) for r in rows] == [
        ("5BB", 100000, 30000), ("5BBJ", 50000, 15000), ("5BBE", 200000, 120000)]

    total_income = 2000000 + 350000  # 20L normal + the special buckets
    tax = T.compute_tax_v2(total_income=total_income,
                           special={"os_special": rows},
                           opt_out_new_regime=False, dob="1990-01-01",
                           filing_date="2026-07-31", return_section=11)
    w = tax["Working"]
    assert w["tax_os_special"] == 165000          # 30k + 15k + 120k
    # 115BBE: flat 25% surcharge on the 115BBE tax — utility TaxCalc parity
    assert w["surcharge_115bbe"] == 30000
    # the sub-2.5L slab income means only special tax drives the surcharge
    assert tax["SurchargeComputed"] == 30000


def test_special_rate_amounts_are_not_double_counted():
    """The special buckets must be taxed at their flat rate *instead of*
    slab rate — compare against a run where the same extra income is
    normal other-sources income."""
    data = load_sample()
    base = build_itr2(data)["report"]["tax"]
    special = build_itr2(with_special(copy.deepcopy(data),
                                      lottery_115bb=100000,
                                      unexplained_115bbe=200000))["report"]["tax"]
    w = special["Working"]
    assert w["os_special"][0]["code"] == "5BB"
    # slab income is the same as the base run's slab income
    assert w["income_excl_ltcg"] == base["Working"]["income_excl_ltcg"]


def test_full_document_passes_validator():
    data = with_special(load_sample(), lottery_115bb=100000,
                        online_games_115bbj=50000, unexplained_115bbe=200000)
    out = build_itr2(data)
    issues = validate_all(out["json"], form="ITR2")
    summary = summarize(issues)
    assert summary["ok"], json.dumps([i.__dict__ for i in issues][:10], indent=1,
                                     default=str)


def test_si_schedule_absent_without_special_income():
    """Baseline behaviour is unchanged: no special-rate income → the SI rows
    list only contains the CG buckets the sample already has (no OS rows)."""
    out = build_itr2(load_sample())
    rows = _si_rows(out["json"])
    codes = [r["SecCode"] for r in rows]
    assert "5BB" not in codes and "5BBJ" not in codes and "5BBE" not in codes
