#!/usr/bin/env python3
"""Generate ``itr1/business_codes.py`` from the official CBDT utilities.

The utilities ship their nature-of-business (NOB) masters as hidden-sheet
dropdown lists.  ``tools/extract_xlsm.py`` already resolves every dropdown to
its verbatim values (``extracted/xlsm/<FORM>/dropdowns.json``); this script
turns four of those lists into an importable Python module:

* ``NOB_MASTER``  — the complete nature-of-business master (ITR-3 utility,
  ``Nature_of_Business`` dropdown; the largest of the shipped masters).
* ``NOB_44AD``    — the presumptive u/s 44AD business dropdown (ITR-4 utility,
  BP sheet ``F13:H17`` source ``NOB44AD``; same set in the ITR-3 utility).
* ``NOB_44ADA``   — the presumptive u/s 44ADA profession dropdown (ITR-4
  utility, BP sheet ``F30:H35`` source ``NOB44ADA``).
* ``NOB_44AE``    — the goods-carriage dropdown (ITR-4 utility, BP sheet
  ``E5:E7`` source ``NOB``).

Labels are kept verbatim (trailing whitespace trimmed) and codes are the
e-filing tokens (e.g. ``16019_1`` for "Medical Profession").

Regenerate whenever the extracted artefacts are refreshed::

    python tools/extract_xlsm.py            # refresh extracted/
    python tools/build_business_codes.py    # refresh itr1/business_codes.py
"""

from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXTRACTED = ROOT / "extracted" / "xlsm"
OUT = ROOT / "itr1" / "business_codes.py"

TOKEN = re.compile(r"^(?P<code>\d{4,6}(?:_\d+)?)-(?P<label>.+)$")

# (module constant, form, dropdown source, exact expected count)
LISTS = [
    ("NOB_MASTER", "ITR3", "Nature_of_Business", 356),
    ("NOB_44AD",   "ITR4", "NOB44AD",            315),
    ("NOB_44ADA",  "ITR4", "NOB44ADA",            38),
    ("NOB_44AE",   "ITR4", "NOB",                  7),
]

DOCS = {
    "NOB_MASTER": "Complete official nature-of-business master (ITR-3 utility)",
    "NOB_44AD":   "Official dropdown for presumptive business u/s 44AD (ITR-4)",
    "NOB_44ADA":  "Official dropdown for presumptive profession u/s 44ADA (ITR-4)",
    "NOB_44AE":   "Official dropdown for goods carriages u/s 44AE (ITR-4)",
}


def _dropdown_values(form: str, source: str) -> list[str]:
    path = EXTRACTED / form / "dropdowns.json"
    entries = json.loads(path.read_text(encoding="utf-8"))
    values = None
    for entry in entries:
        if str(entry.get("source")) == source and entry.get("values"):
            # every validation bound to the same source must agree; keep the
            # first occurrence (they are identical in the official files)
            if values is None:
                values = list(entry["values"])
            elif values != list(entry["values"]):
                raise SystemExit(
                    f"{form}:{source}: inconsistent dropdown values across "
                    f"ranges ({entry.get('ranges')})")
    if values is None:
        raise SystemExit(f"{form}:{source}: dropdown not found in {path}")
    return values


def _parse(values: list[str]) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for raw in values:
        v = raw.rstrip()
        m = TOKEN.match(v)
        if not m:
            continue  # '(Select)', 'All' and other non-code rows
        pairs.append((m.group("code"), m.group("label").rstrip()))
    codes = [c for c, _ in pairs]
    if len(codes) != len(set(codes)):
        dupes = sorted({c for c in codes if codes.count(c) > 1})
        raise SystemExit(f"duplicate codes in dropdown: {dupes}")
    return pairs


def build() -> dict[str, list[tuple[str, str]]]:
    out: dict[str, list[tuple[str, str]]] = {}
    for const, form, source, expected in LISTS:
        pairs = _parse(_dropdown_values(form, source))
        if len(pairs) != expected:
            raise SystemExit(
                f"{form}:{source}: expected {expected} codes, parsed "
                f"{len(pairs)} — was the utility updated?  Re-check the "
                f"official dropdown and adjust LISTS accordingly.")
        out[const] = pairs

    # structural guarantees the rest of the code relies on
    master = {c for c, _ in out["NOB_MASTER"]}
    for const in ("NOB_44AD", "NOB_44ADA", "NOB_44AE"):
        codes = {c for c, _ in out[const]}
        missing = codes - master
        if missing:
            raise SystemExit(f"{const} codes missing from NOB_MASTER: {sorted(missing)}")
    return out


def render(lists: dict[str, list[tuple[str, str]]]) -> str:
    lines = [
        '"""Official CBDT nature-of-business code masters for FY 2025-26 (AY 2026-27).',
        "",
        "GENERATED FILE — do not edit by hand.  Produced by",
        "``tools/build_business_codes.py`` from the nature-of-business dropdown",
        "lists hidden inside the official Excel utilities (see",
        "``extracted/xlsm/<FORM>/dropdowns.json``); regenerate it whenever the",
        "extracted artefacts are refreshed.",
        "",
        "Counts pinned from the official AY 2026-27 utilities:",
    ]
    for const, form, source, _ in LISTS:
        lines.append(f"  * ``{const}``: {len(lists[const])} codes "
                     f"({form} dropdown `{source}`)")
    lines += ['"""', "", "from __future__ import annotations", ""]

    for const, _form, _source, _ in LISTS:
        pairs = lists[const]
        lines.append(f"# {DOCS[const]}")
        lines.append(f"# (code, official label) pairs, in official dropdown order.")
        lines.append(f"{const} = (")
        for code, label in pairs:
            esc = label.replace("\\", "\\\\").replace('"', '\\"')
            lines.append(f'    ("{code}", "{esc}"),')
        lines.append(")")
        lines.append("")

    lines += [
        "# Lookup dictionaries keyed by code.",
    ]
    for const, _form, _source, _ in LISTS:
        lines.append(f"{const}_BY_CODE = dict({const})")
    lines += [
        "",
        "",
        "def is_valid(code: str, scope: str = \"master\") -> bool:",
        '    """True when ``code`` is an official code of the given scope',
        '    (``master`` | ``44ad`` | ``44ada`` | ``44ae``)."""',
        "    table = {",
        '        "master": NOB_MASTER_BY_CODE,',
        '        "44ad": NOB_44AD_BY_CODE,',
        '        "44ada": NOB_44ADA_BY_CODE,',
        '        "44ae": NOB_44AE_BY_CODE,',
        "    }[scope]",
        '    return str(code or "") in table',
        "",
        "",
        "def label(code: str, scope: str = \"master\") -> str:",
        '    """Official label for ``code`` (or an empty string when unknown)."""',
        "    table = {",
        '        "master": NOB_MASTER_BY_CODE,',
        '        "44ad": NOB_44AD_BY_CODE,',
        '        "44ada": NOB_44ADA_BY_CODE,',
        '        "44ae": NOB_44AE_BY_CODE,',
        "    }[scope]",
        '    return table.get(str(code or ""), "")',
        "",
        "",
        "def options(scope: str = \"master\") -> list[str]:",
        '    """Display strings ``"<code>-<label>"`` exactly as the official',
        '    utility dropdown shows them (``(Select)`` and ``All`` excluded)."""',
        "    table = {",
        '        "master": NOB_MASTER,',
        '        "44ad": NOB_44AD,',
        '        "44ada": NOB_44ADA,',
        '        "44ae": NOB_44AE,',
        "    }[scope]",
        '    return [f"{code}-{lbl}" for code, lbl in table]',
        "",
    ]
    return "\n".join(lines)


def main() -> None:
    lists = build()
    text = render(lists)
    OUT.write_text(text, encoding="utf-8")
    total = sum(len(v) for v in lists.values())
    print(f"wrote {OUT.relative_to(ROOT)} "
          f"({', '.join(f'{k}={len(v)}' for k, v in lists.items())}; {total} codes)")


if __name__ == "__main__":
    main()
