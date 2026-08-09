"""Parity tests against the official CBDT Excel preparation utilities.

The VBA macros and workbook structure of the four official utilities were
extracted into ``extracted/xlsm/<FORM>/`` (see ``tools/extract_xlsm.py``).
These tests pin the software to the *identity block* the utilities write into
every generated JSON (CreationInfo + Form_ITRn) so that our output mirrors
the department utility byte-for-byte on those fields, and stays schema-clean.
"""
import json
from pathlib import Path

import pytest

from itr1.builder import build_return as build_itr1
from itr1.itr2_builder import build_itr2
from itr1.itr4_builder import build_itr4
from itr1.validator import summarize, validate_all

ROOT = Path(__file__).resolve().parent.parent
EXTRACTED = ROOT / "extracted" / "xlsm"

SAMPLES = {
    "ITR1": ("samples/sample_taxpayer.json", build_itr1, "ITR1", "ITR-1"),
    "ITR2": ("samples/sample_itr2.json", build_itr2, "ITR2", "ITR-2"),
    "ITR4": ("samples/sample_itr4.json", build_itr4, "ITR4", "ITR-4"),
}

# Values seen verbatim in each utility's VBA (GenerateJson*.bas,
# Form01Header()/getSW*()/getForm*()), latest uncommented assignments.
OFFICIAL = {
    "ITR1": {
        "SWVersionNo": "R3",
        "Description": "For Indls having Income from Salary, Pension, "
                       "family pension and Interest",
    },
    "ITR2": {
        "SWVersionNo": "R4",
        "Description": "For Individuals and HUFs not having income from "
                       "profits and gains of business or profession",
    },
    "ITR4": {
        "SWVersionNo": "R2",
        # Official wording kept verbatim (incl. the utility's spelling).
        "Description": "For Individuals having Income From Presemptive Business",
    },
}
UTILITY_SW_CREATED_BY = "SW90002627"   # all four utilities
EXPECTED_SHEETS = {"ITR1": 21, "ITR2": 66, "ITR3": 66, "ITR4": 24}


def _build_doc(code):
    sample, builder, *_ = SAMPLES[code]
    data = json.loads((ROOT / sample).read_text())
    out = builder(data)
    # every builder returns {"json": <official document>, "report": ...}
    return out["json"] and out["json"] or out


def _form_node(doc, form_key):
    return doc["ITR"][form_key]


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR4"])
def test_creation_info_matches_official_utility(code):
    _, _, form_key, _ = SAMPLES[code]
    node = _form_node(_build_doc(code), form_key)
    ci = node["CreationInfo"]

    assert ci["SWVersionNo"] == OFFICIAL[code]["SWVersionNo"]
    assert ci["IntermediaryCity"] == "Delhi"
    assert ci["Digest"] == "-"                      # official pre-upload value
    import re
    assert re.fullmatch(r"SW[0-9]{8}", ci["SWCreatedBy"])
    assert re.fullmatch(r"SW[0-9]{8}", ci["JSONCreatedBy"])
    assert re.fullmatch(r"\d{4}-\d{2}-\d{2}", ci["JSONCreationDate"])


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR4"])
def test_form_identity_block_matches_official_utility(code):
    _, _, form_key, form_name = SAMPLES[code]
    node = _form_node(_build_doc(code), form_key)
    form = node[f"Form_{form_key}"]

    assert form["FormName"] == form_name
    assert form["Description"] == OFFICIAL[code]["Description"]
    assert form["AssessmentYear"] == "2026"
    assert form["SchemaVer"] == "Ver1.0"
    assert form["FormVer"] == "Ver1.0"


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR4"])
def test_document_still_schema_clean(code):
    doc = _build_doc(code)
    issues = validate_all(doc, form=code)
    assert summarize(issues)["ok"] is True


@pytest.mark.parametrize("code", ["ITR1", "ITR2", "ITR3", "ITR4"])
def test_extraction_inventory(code):
    """The extracted structure inventory must exist and agree with the
    known official workbook layout."""
    struct = EXTRACTED / code / "structure.json"
    if not struct.exists():
        pytest.skip(f"{code} extraction not present")
    data = json.loads(struct.read_text())
    assert len(data["sheets"]) == EXPECTED_SHEETS[code]
    vba_dir = EXTRACTED / code / "vba"
    assert vba_dir.is_dir()
    assert list(vba_dir.glob("*.bas")), "no VBA modules extracted"
    # The JSON-generation macro is the heart of the utility; it must be
    # present and mention the CreationInfo identity block.
    gen = [p for p in vba_dir.iterdir()
           if p.name.lower().startswith("generatejson")]
    assert gen, f"{code}: GenerateJson module missing"
    text = gen[0].read_text(encoding="utf-8", errors="replace")
    assert "CreationInfo" in text
    # the AY 2026-27 software id must appear somewhere in that utility's VBA
    # (ITR-1 keeps the getSWCreatedBy() getters in EfilingCommon.bas).
    corpus = "\n".join(p.read_text(encoding="utf-8", errors="replace")
                       for p in vba_dir.iterdir() if p.is_file())
    assert UTILITY_SW_CREATED_BY in corpus
