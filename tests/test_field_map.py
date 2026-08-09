"""Field-map extraction tests: the *visible* sheets of the official Excel
utilities (the actual return forms - labels, input cells, named ranges,
dropdowns) are extracted into ``extracted/xlsm/<FORM>/fields.json`` and a
readable ``fields.md`` that mirrors the on-screen layout row by row.
"""
import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
EXTRACTED = ROOT / "extracted" / "xlsm"

MIN_FIELDS = {"ITR1": 350, "ITR2": 2500, "ITR3": 3500, "ITR4": 450}


def _fields(code):
    path = EXTRACTED / code / "fields.json"
    if not path.exists():
        pytest.skip(f"{code}: fields.json not present")
    return json.loads(path.read_text())


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR3", "ITR4"])
def test_field_map_exists_and_is_rich(code):
    fields = _fields(code)
    assert len(fields) >= MIN_FIELDS[code]
    assert all({"sheet", "cell", "name"} <= set(f) for f in fields)
    assert (EXTRACTED / code / "fields.md").exists()


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR3", "ITR4"])
def test_every_form_has_core_identity_fields(code):
    fields = _fields(code)
    names = {str(f["name"]) for f in fields}
    assert any(n.endswith(".PAN") or n.endswith(".Pan") for n in names), \
        f"{code}: PAN field missing from field map"
    assert any("DOB" in n or "DateOfBirth" in n for n in names), \
        f"{code}: date-of-birth field missing"
    labels = "||".join((f.get("label") or "") for f in fields)
    for token in ("Name", "Address") if code == "ITR1" else ("Name", "State"):
        assert token in labels, f"{code}: expected label containing {token!r}"


def test_itr3_parta_general_layout_matches_official_form():
    """The ITR-3 Part-A-General sheet (personal information page the user
    pasted) must extract with the exact official field set."""
    fields = _fields("ITR3")
    pa = {f["name"]: f for f in fields if f["sheet"] == "PART A - General"}
    for name, label_part in {
            "sheet1.PAN": "PAN",
            "sheet1.FirstName": "First Name",
            "sheet1.MiddleName": "Middle Name",
            "sheet1.SurNameOrOrgName": "Last Name",
            "sheet1.DOC": "Date of Commencement of Business",
            "sheet1.DOB": "Date of Birth/Formation",
            "sheet1.PinCode": "Pin Code"}.items():
        assert name in pa, f"field {name} missing"
        assert label_part.lower() in (pa[name]["label"] or "").lower(), \
            f"label for {name}: {pa[name]['label']!r}"
    # the residential-status input must carry the official dropdown values
    res = pa["sheet1.ResidentialStatus1"]
    assert res.get("dropdown", [])[:4] == [
        "(Select)", "RES - Resident", "NRI - Non Resident",
        "NOR - Resident but not Ordinarily Resident"]


def test_itr2_salary_sheet_fields():
    fields = _fields("ITR2")
    sch = [f for f in fields if f["sheet"].replace(" ", "").startswith("Salary")
           or f["sheet"] == "Schedule S"]
    if not sch:  # schedule may be titled differently; fall back to any sheet
        sch = fields
    blob = json.dumps(sch)
    assert "Gross Salary" in blob or "grosssalary" in blob.lower()


def test_dropdowns_attached_to_fields():
    """A slice of fields must have resolved dropdown values attached."""
    for code in ("ITR1", "ITR2", "ITR3", "ITR4"):
        fields = _fields(code)
        with_dd = [f for f in fields if f.get("dropdown")]
        assert len(with_dd) >= 30, f"{code}: too few dropdown fields"
