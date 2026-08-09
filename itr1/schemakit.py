"""Schema-kit: minimal-object initialisation driven by the official JSON
schemas (draft-04).  ``init_required`` builds a document containing exactly
the ``required`` keys of a definition (recursively), with type-correct zero
values — a skeleton we then fill with real numbers.  Adding only real values
on top keeps the output slim while always satisfying the schema's structural
mandates.
"""

from __future__ import annotations

import json
import os

_SCHEMA_CACHE: dict[str, dict] = {}
SCHEMA_FILES = {
    "ITR1": "ITR-1_2026_Main_V1.1.json",
    "ITR2": "ITR-2_2026_Main_V1.1.json",
    "ITR4": "ITR-4_2026_Main_V1.1.json",
}
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_schema(form: str) -> dict:
    form = form.upper()
    if form not in _SCHEMA_CACHE:
        path = os.path.join(_REPO_ROOT, SCHEMA_FILES[form])
        with open(path, "r", encoding="utf-8") as fh:
            _SCHEMA_CACHE[form] = json.load(fh)
    return _SCHEMA_CACHE[form]


def _resolve(defs: dict, node: dict):
    while "$ref" in node:
        node = defs[node["$ref"].split("/")[-1]]
    return node


_ENUM_FIRST_SAFE = {
    # enums where the first value is not the sensible default
}


def _zero_value(defs: dict, node: dict):
    node = _resolve(defs, node)
    if "enum" in node and node["enum"]:
        return node["enum"][0]
    t = node.get("type")
    if t == "integer" or "integer" in (t if isinstance(t, list) else []):
        return 0
    if t == "number":
        return 0.0
    if t == "string":
        m = node.get("pattern")
        if m:
            return _pattern_string(m, node)
        ml = node.get("minLength")
        return " " * max(ml or 0, 1) if (ml or 0) > 0 else ""
    if t == "boolean":
        return False
    if t == "object" or "properties" in node:
        return init_required_node(defs, node)
    if t == "array":
        return []
    if "anyOf" in node or "oneOf" in node:
        variants = node.get("anyOf") or node.get("oneOf")
        return _zero_value(defs, variants[0])
    return 0


def _pattern_string(pattern: str, node: dict) -> str:
    """Build a small string that satisfies simple CBDT patterns."""
    if pattern in ("^[A-Z]{5}[0-9]{4}[A-Z]{1}$", r"^[A-Z]{5}[0-9]{4}[A-Z]$"):
        return "AAAAA0000A"
    if "ddmmyyyy" in pattern.lower():
        return "01011990"
    if pattern.startswith("^20"):
        # e.g. years / dates
        return "2025" if "{" not in pattern else "2025"
    if "[0-9]" in pattern or "\\d" in pattern:
        m = node.get("minLength", 1) or 1
        return "0" * m
    return ""


def init_required_node(defs: dict, node: dict) -> dict:
    node = _resolve(defs, node)
    out = {}
    for key in node.get("required", []) or []:
        prop = (node.get("properties") or {}).get(key)
        if prop is not None:
            out[key] = _zero_value(defs, prop)
    return out


def init_required(defs: dict, def_name: str) -> dict:
    return init_required_node(defs, defs[def_name])


def validate_with_schema(form: str, doc: dict):
    """jsonschema-validate ``doc`` against the official schema for ``form``.
    Returns a list of error strings (empty when valid)."""
    import jsonschema
    schema = load_schema(form)
    validator_cls = jsonschema.validators.validator_for(schema)
    validator = validator_cls(schema)
    errs = []
    for e in validator.iter_errors(doc):
        loc = "$" + "".join(f"[{p!r}]" if isinstance(p, int) else f".{p}" for p in e.absolute_path)
        errs.append(f"{loc}: {e.message}")
    return errs


def array_item(defs: dict, node: dict):
    """Return the (resolved) item-definition node of an array property."""
    node = _resolve(defs, node)
    return _resolve(defs, node.get("items", {}))
