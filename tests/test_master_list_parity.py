"""Parity tests against the *contents* of the hidden master-list sheets
extracted from the official Excel utilities (``extracted/xlsm/<FORM>/sheets/``).

These guard two things at once:

* every hidden worksheet has its cell contents dumped next to the VBA, and
* the statutory dropdown masters we surface in the software (nature-of-
  business codes, bank/IFSC masters, ISIN lists) actually exist — and the
  codes we ship are genuine official codes.
"""
import gzip
import json
from pathlib import Path

import pytest

from itr1.itr4_builder import BUSINESS_CODES_44AD

ROOT = Path(__file__).resolve().parent.parent
EXTRACTED = ROOT / "extracted" / "xlsm"


def _structure(code):
    struct = EXTRACTED / code / "structure.json"
    if not struct.exists():
        pytest.skip(f"{code} extraction not present")
    return json.loads(struct.read_text())


def _dump_path(code, sheet_name):
    """Locate the dumped TSV (plain or gzipped) for a hidden sheet."""
    for sheet in _structure(code)["sheets"]:
        if sheet["name"] == sheet_name and sheet.get("dump"):
            return EXTRACTED / code / sheet["dump"]
    pytest.skip(f"{code}: sheet {sheet_name!r} not dumped")


def _read_dump(path: Path) -> str:
    if path.suffix == ".gz":
        with gzip.open(path, "rt", encoding="utf-8") as fh:
            return fh.read()
    return path.read_text(encoding="utf-8")


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR3", "ITR4"])
def test_every_hidden_sheet_is_dumped(code):
    structure = _structure(code)
    hidden = [s for s in structure["sheets"] if s["state"] != "visible"]
    assert hidden, f"{code}: expected hidden sheets"
    for sheet in hidden:
        dump = sheet.get("dump")
        assert dump, f"{code}: hidden sheet {sheet['name']!r} has no dump"
        assert (EXTRACTED / code / dump).exists(), f"{code}: {dump} missing"


def test_nature_of_business_codes_are_official():
    """Every nature-of-business code we surface in the UI (44AD/44ADA/44AE
    pick-lists) must be a genuine code from the official master lists
    hidden inside the utilities (ITR-4 'DB' sheet / ITR-3 master).

    Historical note: this is the test that caught the buggy mapping.  An
    earlier version also required every 44AD code to exist in the ITR-2
    utility's smaller DropDownValues copy — that assumption was wrong:
    ITR-2 legitimately omits the speculative/F&O trading codes
    (21009/21010/21011), 01012 and 14005 which only business filers
    (ITR-3/ITR-4) may use.  The authoritative cross-check is against the
    ITR-3 master, which is a superset of every shipped dropdown list.
    """
    import re
    from itr1.itr4_builder import (PROFESSION_CODES_44ADA,
                                   TRANSPORT_CODES_44AE)
    from itr1 import business_codes as NOB
    db = _read_dump(_dump_path("ITR4", "DB"))

    def present(code, hay):
        return re.search(rf"(?<![0-9_]){re.escape(code)}(?=-|\t|\n)", hay)

    for catalog, name in ((BUSINESS_CODES_44AD, "44AD"),
                          (PROFESSION_CODES_44ADA, "44ADA"),
                          (TRANSPORT_CODES_44AE, "44AE")):
        for code in catalog:
            assert present(code, db), \
                f"{name} code {code} missing from ITR-4 DB sheet"
    master = set(NOB.NOB_MASTER_BY_CODE)
    for catalog, name in ((NOB.NOB_44AD_BY_CODE, "44AD"),
                          (NOB.NOB_44ADA_BY_CODE, "44ADA"),
                          (NOB.NOB_44AE_BY_CODE, "44AE")):
        missing = set(catalog) - master
        assert not missing, f"{name} codes missing from the ITR-3 NOB master: {sorted(missing)}"


def test_isin_master_lists():
    """The 112A ISIN dropdown masters must be complete extracts."""
    isin2 = _read_dump(_dump_path("ITR2", "ISIN List"))
    rows2 = isin2.count("\n") - 1                       # minus header
    assert rows2 >= 85_000
    isin3 = _read_dump(_dump_path("ITR3", "ISIN List"))
    assert isin3.count("\n") - 1 >= 120_000
    # sane shape: ISINs are 12-char alphanumeric starting with a country code
    import re
    assert re.search(r"\bIN[A-Z0-9]{10}\b", isin2)


def test_ifsc_master_ITR1():
    ifsc = _read_dump(_dump_path("ITR1", "IFSC"))
    assert ifsc.count("\n") - 1 >= 40_000
    assert "HDFC0" in ifsc and "SBIN0" in ifsc


@pytest.mark.parametrize("code,min_resolved", [("ITR1", 60), ("ITR2", 150),
                                               ("ITR3", 180), ("ITR4", 80)])
def test_dropdown_lists_resolved(code, min_resolved):
    dd = EXTRACTED / code / "dropdowns.json"
    if not dd.exists():
        pytest.skip(f"{code}: dropdowns.json not present")
    lists = json.loads(dd.read_text())
    with_values = [e for e in lists if e.get("total", 0) > 0]
    assert len(with_values) >= min_resolved
    # dynamic INDIRECT()/IF() lists and bare defined names are documented,
    # not silently dropped
    import re
    dynamic = [e for e in lists if e.get("kind") == "unresolved"]
    assert all("IF(" in e["source"] or "INDIRECT" in e["source"]
               or e["source"].isdigit() or e["source"].startswith("-")
               or "$" in e["source"]
               or re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", e["source"])
               for e in dynamic)


def test_residential_status_master_values():
    """Spot-check a statutory dropdown's exact official values."""
    dd = json.loads((EXTRACTED / "ITR2" / "dropdowns.json").read_text())
    res = [e for e in dd if e.get("definedName") == "ResiStatus"]
    assert res, "ResiStatus list not resolved"
    values = res[0]["values"]
    for token in ("RES - Resident", "NRI - Non Resident",
                  "NOR - Resident but not Ordinarily Resident"):
        assert token in values, f"residential-status value {token!r} missing"
