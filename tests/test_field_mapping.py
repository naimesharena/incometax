"""Crosswalk enforcement: official-utility fields <-> our input model <->
official JSON output (see ``itr1/fieldmap.py``).

For every crosswalk entry these tests prove three links of the chain:

1. the field EXISTS in the extracted utility field map
   (``extracted/xlsm/<FORM>/fields.json`` - cell + defined name + label);
2. our input model really ingests it (sentinel flow: inject a unique value
   at the friendly key and rebuild);
3. it lands at the documented official JSON path with that exact value.

For "exception" entries the tests pin the *decision* (so nothing is silently
unmapped - every exemption is deliberate and documented).
"""
import copy
import json
from pathlib import Path

import pytest

from itr1.builder import build_return as build_itr1
from itr1.itr2_builder import build_itr2
from itr1.itr4_builder import build_itr4
from itr1.fieldmap import FIELD_MAP, coverage

ROOT = Path(__file__).resolve().parent.parent
EXTRACTED = ROOT / "extracted" / "xlsm"

BUILDERS = {"ITR1": build_itr1, "ITR2": build_itr2, "ITR4": build_itr4}
SAMPLES = {"ITR1": "samples/sample_taxpayer.json",
           "ITR2": "samples/sample_itr2.json",
           "ITR4": "samples/sample_itr4.json"}
MIN_MAPPED = {"ITR1": 30, "ITR2": 34, "ITR4": 28}


def _sample(form):
    return json.loads((ROOT / SAMPLES[form]).read_text())


def _utility_names(form):
    path = EXTRACTED / form / "fields.json"
    if not path.exists():
        pytest.skip(f"{form}: field map extraction not present")
    fields = json.loads(path.read_text())
    return {str(f["name"]) for f in fields if f.get("name")}, fields


def _deep_set(data: dict, dotted: str, value):
    node = data
    parts = dotted.split(".")
    for p in parts[:-1]:
        node = node.setdefault(p, {})
    node[parts[-1]] = value


def _merge(data: dict, extra: dict):
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(data.get(k), dict):
            _merge(data[k], v)
        else:
            data[k] = v


def _get(node, path):
    for key in path:
        assert isinstance(node, dict) and key in node, \
            f"path broken at {key} (walking {path})"
        node = node[key]
    return node


# ---------------------------------------------------------------------------

@pytest.mark.parametrize("form", ["ITR1", "ITR2", "ITR4"])
def test_every_crosswalk_field_exists_in_utility(form):
    names, _raw = _utility_names(form)
    names_lower = {n.lower() for n in names}
    missing = [e["name"] for e in FIELD_MAP[form]
               if e["name"] not in names and e["name"].lower() not in names_lower]
    assert not missing, f"{form}: utility fields not found in extraction: {missing}"


@pytest.mark.parametrize("form", ["ITR1", "ITR2", "ITR4"])
def test_labels_match_utility(form):
    _names, fields = _utility_names(form)
    by_name = {}
    for f in fields:
        by_name.setdefault(str(f.get("name")), []).append(f.get("label") or "")
    bad = []
    for e in FIELD_MAP[form]:
        if not e["label"]:
            continue
        labels = by_name.get(e["name"]) or [
            v for n, v in by_name.items() if n.lower() == e["name"].lower()
            for v in v] or []
        if not any(e["label"].lower() in (lab or "").lower() for lab in labels):
            bad.append((e["name"], e["label"], labels[:3]))
    assert not bad, f"{form}: label mismatches: {bad[:4]}"


def _flow_cases():
    cases = []
    for form, entries in FIELD_MAP.items():
        for e in entries:
            if e.get("sent") and e.get("source") and e.get("path"):
                cases.append((form, e))
    return cases


@pytest.mark.parametrize("form,entry", _flow_cases(),
                         ids=[f"{f}:{e['name']}" for f, e in _flow_cases()])
def test_sentinel_flows_end_to_end(form, entry):
    """Inject a sentinel value at our friendly input key; the generated
    official JSON must carry it at the crosswalk's documented path."""
    data = copy.deepcopy(_sample(form))
    _merge(data, entry.get("extra_inject") or {})
    _deep_set(data, entry["source"], entry["sent"][0])
    doc = BUILDERS[form](data)["json"]
    got = _get(doc, entry["path"])
    want = entry["sent"][1]
    assert got == want, (
        f"{entry['name']}: {entry['source']}={entry['sent'][0]!r} -> "
        f"{entry['path'][-1]}={got!r}, expected {want!r}")


@pytest.mark.parametrize("form", ["ITR1", "ITR2", "ITR4"])
def test_coverage_thresholds(form):
    cov = coverage(form)
    assert cov["mapped"] >= MIN_MAPPED[form], f"{form}: {cov}"
    # every mapped (non-partial) entry must carry a sentinel - otherwise the
    # mapping is not actually provable
    unproven = [e["name"] for e in FIELD_MAP[form]
                if e["kind"] == "mapped" and not e.get("sent")]
    assert not unproven, f"{form}: mapped fields without sentinel: {unproven}"
    print(f"    [{form}] mapped={cov['mapped']} partial={cov['partial']} "
          f"exception={cov['exception']} total={cov['total']}")


def test_exceptions_all_have_reasons():
    for form, entries in FIELD_MAP.items():
        for e in entries:
            if e["kind"] == "exception":
                assert e["note"].strip(), f"{form}:{e['name']} exception w/o reason"
