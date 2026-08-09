#!/usr/bin/env python3
"""Extract macros (VBA) plus workbook structure from the official CBDT
Excel preparation utilities (.xlsm) checked into this repository.

The official artefacts themselves are NEVER modified; everything is written
to a separate output directory (default: ``extracted/xlsm/<code>/``):

    vba/                  one text file per VBA module (.bas/.cls/.frm)
    structure.json        machine-readable workbook inventory:
                          sheets (name/visibility/dimensions/formula counts),
                          defined names, data-validation lists (dropdowns),
                          protection flags, customXml parts
    summary.md            human-readable digest of the above

Usage:
    python tools/extract_xlsm.py                      # all known utilities
    python tools/extract_xlsm.py path/to/file.xlsm    # one utility

Requires: oletools (VBA decompression). Everything else is stdlib.
"""
from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

REPO = Path(__file__).resolve().parent.parent

# code -> (xlsm path relative to repo, companion CSV folder)
UTILITIES = {
    "ITR1": ("ITR1_AY_26-27_V1.2.xlsm", None),
    "ITR2": ("ITR2_AY_26-27_V1.3/ITR2_AY_26-27_V1.3.xlsm",
             "ITR2_AY_26-27_V1.3"),
    "ITR3": ("ITR3_AY_26-27_V1.2/ITR3_AY_26-27_V1.2.xlsm",
             "ITR3_AY_26-27_V1.2"),
    "ITR4": ("ITR4_AY_26-27_V1.1.xlsm", None),
}

BAS_MODULE_EXT = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}


# --------------------------------------------------------------------------
# VBA extraction
# --------------------------------------------------------------------------

def extract_vba(xlsm_path: Path, out_dir: Path) -> list[dict]:
    """Decompress every VBA module via oletools and write one file each."""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "oletools is required: pip install oletools") from exc

    modules = []
    vba_dir = out_dir / "vba"
    vba_dir.mkdir(parents=True, exist_ok=True)

    parser = VBA_Parser(str(xlsm_path))
    try:
        if not parser.detect_vba_macros():
            return modules
        for (_fname, stream_path, vba_filename, vba_code) in \
                parser.extract_macros():
            safe = re.sub(r"[^A-Za-z0-9_.-]", "_", vba_filename or "module")
            ext = Path(safe).suffix.lower()
            if ext not in (".bas", ".cls", ".frm"):
                # oletools usually keeps the extension; fall back to .bas
                safe = safe + (ext if ext else ".bas")
            target = vba_dir / safe
            text = vba_code if isinstance(vba_code, str) else \
                vba_code.decode("utf-8", "replace")
            # strip the p-code dir stream noise oletools may prepend
            target.write_text(text, encoding="utf-8", errors="replace")
            modules.append({
                "module": vba_filename,
                "stream": stream_path,
                "file": f"vba/{safe}",
                "lines": text.count("\n") + 1,
                "chars": len(text),
            })
    finally:
        parser.close()
    modules.sort(key=lambda m: (m["module"] or "").lower())
    return modules


# --------------------------------------------------------------------------
# Workbook structure (parsed straight from the OOXML parts; no openpyxl load,
# which is very slow on these 10 MB+ utilities)
# --------------------------------------------------------------------------

def _rels_map(zf: zipfile.ZipFile) -> dict[str, str]:
    rels = {}
    data = zf.read("xl/_rels/workbook.xml.rels")
    root = ET.fromstring(data)
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = "xl/" + target.lstrip("/") \
                if not target.startswith("xl/") else target
    return rels


def extract_structure(zf: zipfile.ZipFile) -> dict:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = _rels_map(zf)

    defined_names = []
    dn_node = wb_root.find(f"{NS_MAIN}definedNames")
    if dn_node is not None:
        for dn in dn_node.findall(f"{NS_MAIN}definedName"):
            defined_names.append({
                "name": dn.get("name"),
                "localSheetId": dn.get("localSheetId"),
                "hidden": dn.get("hidden") in ("1", "true"),
                "refersTo": (dn.text or "").strip(),
            })

    sheets = []
    for sh in wb_root.find(f"{NS_MAIN}sheets") or []:
        rid = sh.get(f"{NS_REL}id")
        sheets.append({
            "name": sh.get("name"),
            "sheetId": sh.get("sheetId"),
            "state": sh.get("state", "visible"),
            "part": rels.get(rid),
        })

    for sh in sheets:
        part = sh.get("part")
        if not part or part not in zf.namelist():
            continue
        raw = zf.read(part)
        try:
            sroot = ET.fromstring(raw)
        except ET.ParseError:
            continue
        dim = sroot.find(f"{NS_MAIN}dimension")
        sh["dimensions"] = dim.get("ref") if dim is not None else None
        sh["protection"] = sroot.find(f"{NS_MAIN}sheetProtection") is not None
        sh["formulas"] = len(sroot.findall(f".//{NS_MAIN}f"))
        merges = sroot.find(f"{NS_MAIN}mergeCells")
        sh["mergedCells"] = int(merges.get("count", 0)) if merges is not None else 0

        # data validations (dropdown lists)
        dvals = []
        for dv in sroot.findall(
                f"{NS_MAIN}dataValidations/{NS_MAIN}dataValidation"):
            f1 = dv.find(f"{NS_MAIN}formula1")
            entry = {
                "ranges": dv.get("sqref"),
                "type": dv.get("type"),
                "formula1": (f1.text or "").strip() if f1 is not None else None,
            }
            dvals.append(entry)
        sh["dataValidations"] = dvals

    structure = {
        "sheets": sheets,
        "definedNames": defined_names,
        "customXml": sorted(n for n in zf.namelist()
                            if n.startswith("customXml/") and n.endswith(".xml")
                            and "Props" not in n),
    }
    # pull in the small customXml payloads verbatim (they carry utility config)
    custom = {}
    for name in structure["customXml"]:
        try:
            custom[name] = zf.read(name).decode("utf-8", "replace")
        except Exception:
            pass
    if custom:
        structure["customXmlContent"] = custom
    return structure


def companion_csvs(folder: Path | None) -> dict[str, str]:
    """Return the official CSV import/export templates that ship next to
    a utility (112A / 115AD / TDS / TCS / IT), header line included."""
    out = {}
    if not folder or not folder.is_dir():
        return out
    for csv in sorted(folder.rglob("*.csv")):
        rel = csv.relative_to(folder.parent.parent
                              if folder.parent.parent.name == ""
                              else REPO)
        text = csv.read_text(encoding="utf-8-sig", errors="replace")
        out[str(csv.relative_to(REPO))] = {
            "header": text.splitlines()[0] if text.strip() else "",
            "bytes": len(text.encode()),
        }
    return out


# --------------------------------------------------------------------------
# Cell content dumps (hidden sheets hold all the dropdown master lists)
# --------------------------------------------------------------------------

def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    strings = []
    with zf.open("xl/sharedStrings.xml") as fh:
        for _ev, elem in ET.iterparse(fh, events=("end",)):
            if elem.tag == f"{NS_MAIN}si":
                strings.append("".join(
                    t.text or "" for t in elem.iter(f"{NS_MAIN}t")))
                elem.clear()
    return strings


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


_REF_RE = re.compile(r"^([A-Za-z]{1,3})(\d{1,7})$")


def iter_sheet_cells(zf: zipfile.ZipFile, part: str,
                     shared: list[str]):
    """Yield (ref, col_letters, row, value, formula) for every stored cell,
    streaming the sheet XML so 100k-row master lists stay cheap."""
    with zf.open(part) as fh:
        for _ev, elem in ET.iterparse(fh, events=("end",)):
            if elem.tag != f"{NS_MAIN}c":
                continue
            ref = elem.get("r", "")
            ctype = elem.get("t", "")
            fnode = elem.find(f"{NS_MAIN}f")
            vnode = elem.find(f"{NS_MAIN}v")
            formula = fnode.text if fnode is not None else None
            raw = vnode.text if vnode is not None else None
            value = ""
            if ctype == "s" and raw is not None:
                idx = int(raw)
                value = shared[idx] if 0 <= idx < len(shared) else ""
            elif ctype == "inlineStr":
                value = "".join(t.text or "" for t in elem.iter(f"{NS_MAIN}t"))
            elif ctype == "b":
                value = "TRUE" if raw == "1" else "FALSE"
            elif raw is not None:
                value = raw
            m = _REF_RE.match(ref or "")
            if m:
                yield ref, m.group(1).upper(), int(m.group(2)), value, formula
            elem.clear()


def dump_hidden_sheets(zf: zipfile.ZipFile, structure: dict,
                       shared: list[str], out_dir: Path,
                       include_visible: bool = True):
    """Write the full cell contents of worksheets:
    ``sheets/NN_<name>.tsv`` (row-major grid, first column = row number) and
    ``sheets/NN_<name>.formulas.tsv`` when the sheet carries formulas.
    Hidden sheets hold the dropdown masters; visible sheets hold the actual
    form layouts (labels + input boxes).  Visible-sheet grids are also kept
    in memory for the field-map build and returned."""
    sheets_dir = out_dir / "sheets"
    dumps = []
    visible_cells: dict[str, dict[str, str]] = {}
    idx = 0
    for sh in structure["sheets"]:
        is_hidden = sh["state"] != "visible"
        if not is_hidden and not include_visible:
            continue
        idx += 1
        part = sh.get("part")
        if not part or part not in zf.namelist():
            continue
        rows: dict[int, dict[str, str]] = {}
        formulas: list[tuple[str, str, str]] = []
        max_col = 0
        n_cells = 0
        for ref, col, row, value, formula in iter_sheet_cells(zf, part, shared):
            if value != "":
                rows.setdefault(row, {})[col] = value
                max_col = max(max_col, _col_index(col))
                n_cells += 1
            if formula:
                formulas.append((ref, formula, value))
        sheets_dir.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", sh["name"]).strip("_")
        letters = []
        for i in range(1, max_col + 1):
            n, l = i, ""
            while n:
                n, r = divmod(n - 1, 26)
                l = chr(65 + r) + l
            letters.append(l)
        lines = ["row\t" + "\t".join(letters)]
        for r in sorted(rows):
            cells = rows[r]
            last = max((_col_index(c) for c in cells), default=0)
            lines.append("\t".join([str(r)] + [cells.get(letters[i], "")
                                               for i in range(last)]))
        text = "\n".join(lines) + "\n"
        # large master lists (ISIN/IFSC/…) are stored gzipped to keep the
        # repository lean; smaller sheets stay plain-tab readable.
        if len(text.encode("utf-8")) > 1_000_000:
            import gzip
            fname = f"sheets/{idx:02d}_{safe}.tsv.gz"
            with gzip.open(out_dir / fname, "wt", encoding="utf-8",
                           compresslevel=9) as fh:
                fh.write(text)
        else:
            fname = f"sheets/{idx:02d}_{safe}.tsv"
            (out_dir / fname).write_text(text, encoding="utf-8")
        rec = {"sheet": sh["name"], "state": sh["state"], "file": fname,
               "rows": len(rows), "cells": n_cells, "formulas": len(formulas)}
        if formulas:
            ff = f"sheets/{idx:02d}_{safe}.formulas.tsv"
            with (out_dir / ff).open("w", encoding="utf-8",
                                     newline="") as fh:
                fh.write("cell\tformula\tcached_value\n")
                for ref, formula, cached in formulas:
                    fh.write(f"{ref}\t{formula}\t{cached}\n")
            rec["formulasFile"] = ff
        dumps.append(rec)
        sh["dump"] = fname
        if sh["state"] == "visible":
            grid: dict[str, str] = {}
            for r, cols in rows.items():
                for c, v in cols.items():
                    grid[f"{c}{r}"] = v
            visible_cells[sh["name"]] = grid
    return dumps, visible_cells


# --------------------------------------------------------------------------
# Form field map: label -> cell -> named range -> dropdown values
# --------------------------------------------------------------------------

_SQREF_PART_RE = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?$")
_NAME_REF_RE = re.compile(
    r"^(?:'([^']+)'|([A-Za-z0-9_.&() /-]+))!\$?([A-Z]{1,3})\$?(\d{1,7})$")

_LABEL_SKIP = re.compile(r"^\(Select\)|^#REF!|^=|^\d+$")


def _col_letters(idx: int) -> str:
    n, l = idx, ""
    while n:
        n, r = divmod(n - 1, 26)
        l = chr(65 + r) + l
    return l


def _expand_sqref(sqref: str, cap: int = 500):
    """Expand a data-validation sqref like 'F10:AC10 E10' to cell refs."""
    cells = []
    for part in (sqref or "").split():
        m = _SQREF_PART_RE.match(part)
        if not m:
            continue
        c1, r1 = _col_index(m.group(1)), int(m.group(2))
        c2 = _col_index(m.group(3)) if m.group(3) else c1
        r2 = int(m.group(4)) if m.group(4) else r1
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cells.append(f"{_col_letters(c)}{r}")
                if len(cells) >= cap:
                    return cells
    return cells


def build_fields(structure: dict, dropdowns: list[dict],
                 visible_cells: dict[str, dict[str, str]],
                 out_dir: Path) -> list[dict]:
    """Build the machine-readable field map of every *visible* sheet.

    Every interactive input in the official utilities is addressable three
    ways: its cell, a (sheet-scoped) defined name (``sheet1.PAN``…), and the
    human label placed next to it.  This map joins all three plus the
    resolved dropdown values for that cell.
    """
    # dropdown values per (sheet, cell)
    dd_by_cell: dict[str, dict] = {}
    for e in dropdowns:
        if e.get("kind") not in ("range", "inline") or not e.get("total"):
            continue
        for ref in _expand_sqref(e.get("ranges") or ""):
            dd_by_cell.setdefault(f"{e['sheet']}!{ref}", e)

    fields = []
    visible = {s["name"] for s in structure["sheets"] if s["state"] == "visible"}
    for dn in structure["definedNames"]:
        m = _NAME_REF_RE.match(dn.get("refersTo") or "")
        if not m:
            continue
        sheet = (m.group(1) or m.group(2) or "").strip()
        cell = f"{m.group(3)}{m.group(4)}"
        if sheet not in visible:
            continue
        col, row = _col_index(m.group(3)), int(m.group(4))
        grid = visible_cells.get(sheet, {})

        label = ""
        for c in range(col - 1, max(col - 8, 0), -1):      # nearest left
            v = grid.get(f"{_col_letters(c)}{row}", "").strip()
            if len(v) >= 3 and not _LABEL_SKIP.match(v) and any(ch.isalpha() for ch in v):
                label = v
                break
        if not label:
            for r in range(row - 1, max(row - 4, 0), -1):  # then above
                v = grid.get(f"{m.group(3)}{r}", "").strip()
                if len(v) >= 3 and not _LABEL_SKIP.match(v) and any(ch.isalpha() for ch in v):
                    label = v
                    break

        entry = {"sheet": sheet, "cell": cell,
                 "name": dn.get("name"), "label": label[:120]}
        dd = dd_by_cell.get(f"{sheet}!{cell}")
        if dd:
            entry["dropdown"] = dd["values"][:60]
            entry["dropdownTotal"] = dd["total"]
        fields.append(entry)

    fields.sort(key=lambda f: (structure["sheets"].index(
        next(s for s in structure["sheets"] if s["name"] == f["sheet"])),
        int(re.sub(r"\D", "", f["cell"])), f["cell"]))
    (out_dir / "fields.json").write_text(
        json.dumps(fields, indent=1, ensure_ascii=False), encoding="utf-8")

    # readable digest, sheet by sheet (mirrors the on-screen form layout)
    lines = [f"# Field map — {out_dir.name}", ""]
    cur = None
    for f in fields:
        if f["sheet"] != cur:
            cur = f["sheet"]
            lines += ["", f"## {cur}", "",
                      "| Cell | Named range | Label | Dropdown values |",
                      "|------|-------------|-------|-----------------|"]
        dd = "; ".join(f.get("dropdown", [])[:8])
        if f.get("dropdownTotal", 0) > 8:
            dd += f" … ({f['dropdownTotal']})"
        lab = (f["label"] or "").replace("|", "\\|").replace("\n", " ")
        lines.append(f"| {f['cell']} | `{f['name']}` | {lab} | {dd} |")
    (out_dir / "fields.md").write_text("\n".join(lines) + "\n",
                                       encoding="utf-8")
    return fields
# --------------------------------------------------------------------------

_RANGE_RE = re.compile(
    r"^(?:'(?P<q>[^']+)'|(?P<b>[A-Za-z0-9_. ]+))!"
    r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)"
    r"(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?$")


def _cache_sheet_cells(zf, part, shared, cache, sheet_name):
    if sheet_name in cache:
        return
    cells: dict[str, str] = {}
    if part and part in zf.namelist():
        for ref, _col, _row, value, _f in iter_sheet_cells(zf, part, shared):
            if value != "":
                cells[ref] = value
    cache[sheet_name] = cells


def resolve_dropdowns(zf: zipfile.ZipFile, structure: dict,
                      shared: list[str], out_dir: Path,
                      cap: int = 400) -> list[dict]:
    """Resolve every list-type data-validation rule to its actual values —
    inline lists, defined names and hidden-sheet ranges."""
    parts = {s["name"]: s.get("part") for s in structure["sheets"]}
    names = {d["name"]: d["refersTo"] for d in structure["definedNames"]
             if d.get("name") and d.get("refersTo")}
    cache: dict[str, dict[str, str]] = {}
    resolved: list[dict] = []

    def values_from_range(sheet: str, c1: str, r1: str, c2: str, r2: str):
        _cache_sheet_cells(zf, parts.get(sheet), shared, cache, sheet)
        cells = cache.get(sheet, {})
        lo_r, hi_r = int(r1), int(r2 or r1)
        lo_c = _col_index(c1)
        hi_c = _col_index(c2 or c1)
        vals, total = [], 0
        for r in range(lo_r, hi_r + 1):
            for ci in range(lo_c, hi_c + 1):
                n, l = ci, ""
                while n:
                    n, rr = divmod(n - 1, 26)
                    l = chr(65 + rr) + l
                v = cells.get(f"{l}{r}", "")
                if v != "":
                    total += 1
                    if len(vals) < cap:
                        vals.append(v)
        return vals, total

    for sh in structure["sheets"]:
        for dv in sh.get("dataValidations", []):
            dtype = dv.get("type") or ""
            f1 = (dv.get("formula1") or "").strip()
            if not f1:
                if dtype and dtype != "list":
                    continue  # pure min/max constraint rules are not lists
                continue
            if dtype and dtype != "list":
                # whole/decimal/date constraints (formula1 is a bound like 0)
                resolved.append({"sheet": sh["name"], "ranges": dv.get("ranges"),
                                 "type": dtype, "source": f1,
                                 "kind": "constraint", "values": [], "total": 0})
                continue
            entry = {"sheet": sh["name"], "ranges": dv.get("ranges"),
                     "type": "list", "source": f1}
            f1 = f1.lstrip("=")
            if f1.startswith('"') and f1.endswith('"'):
                vals = [x.strip() for x in f1.strip('"').split(",")]
                entry.update(kind="inline", values=vals, total=len(vals))
                resolved.append(entry)
                continue
            expr = f1.lstrip("=")
            if expr in names:
                entry["definedName"] = expr
                expr = names[expr].split(",")[0].strip().lstrip("=")
            m = _RANGE_RE.match(expr)
            if m and m.group("q") in parts or (m and (m.group("b") or "")
                                             .strip() in parts):
                sheet = (m.group("q") or m.group("b") or "").strip()
                vals, total = values_from_range(
                    sheet, m.group("c1"), m.group("r1"),
                    m.group("c2"), m.group("r2"))
                entry.update(kind="range", sourceSheet=sheet,
                             values=vals, total=total)
            else:
                entry.update(kind="unresolved", values=[], total=0)
            resolved.append(entry)
    (out_dir / "dropdowns.json").write_text(
        json.dumps(resolved, indent=1, ensure_ascii=False), encoding="utf-8")
    return resolved

def write_summary(code: str, xlsm: Path, modules: list[dict],
                  structure: dict, csvs: dict, out_dir: Path) -> None:
    lines = [f"# {code} — Excel utility extraction", ""]
    lines.append(f"Source: `{xlsm.relative_to(REPO)}` "
                 f"({xlsm.stat().st_size:,} bytes)")
    lines.append("")
    sheets = structure["sheets"]
    vis = [s for s in sheets if s["state"] == "visible"]
    hid = [s for s in sheets if s["state"] != "visible"]
    lines.append(f"- Worksheets: **{len(sheets)}** "
                 f"({len(vis)} visible, {len(hid)} hidden)")
    dumps = structure.get("hiddenSheetDumps", [])
    if dumps:
        cells = sum(d["cells"] for d in dumps)
        lines.append(f"- Sheet contents (hidden + visible): **{len(dumps)} dumps** in "
                     f"`sheets/` ({cells:,} stored cell values, "
                     f"formulas extracted alongside)")
    lines.append(f"- VBA modules: **{len(modules)}**, "
                 f"{sum(m['lines'] for m in modules):,} lines of code")
    dv_total = sum(len(s["dataValidations"]) for s in sheets)
    lines.append(f"- Data-validation (dropdown) rules: **{dv_total}** "
                 f"(resolved to actual list values in `dropdowns.json`)")
    lines.append(f"- Defined names: **{len(structure['definedNames'])}**")
    lines.append("")
    lines.append("## Worksheets")
    lines.append("")
    lines.append("| # | Sheet | State | Dimensions | Formulas | Validations |")
    lines.append("|---|-------|-------|-----------|----------|-------------|")
    for i, s in enumerate(sheets, 1):
        lines.append(
            f"| {i} | {s['name']} | {s['state']} | "
            f"{s.get('dimensions') or ''} | {s.get('formulas', 0)} | "
            f"{len(s['dataValidations'])} |")
    lines.append("")
    lines.append("## VBA modules")
    lines.append("")
    lines.append("| Module | Lines | File |")
    lines.append("|--------|-------|------|")
    for m in modules:
        lines.append(f"| {m['module']} | {m['lines']} | `{m['file']}` |")
    lines.append("")
    if structure["definedNames"]:
        lines.append("## Defined names")
        lines.append("")
        for dn in structure["definedNames"][:200]:
            lines.append(f"- `{dn['name']}` → {dn['refersTo']}")
        lines.append("")
    if csvs:
        lines.append("## Companion CSV templates")
        lines.append("")
        for path, meta in csvs.items():
            lines.append(f"- `{path}` — header: `{meta['header'][:160]}`")
        lines.append("")
    if dumps:
        lines.append("## Sheet dumps (`sheets/`) — hidden and visible")
        lines.append("")
        lines.append("| Sheet | State | Rows with values | Cells | Formulas | File |")
        lines.append("|-------|-------|------------------|-------|----------|------|")
        for d in dumps:
            lines.append(f"| {d['sheet']} | {d['state']} | {d['rows']} | "
                         f"{d['cells']} | {d['formulas']} | `{d['file']}` |")
        lines.append("")
    (out_dir / "summary.md").write_text("\n".join(lines), encoding="utf-8")


def process(code: str, xlsm: Path, csv_folder: Path | None,
            out_root: Path) -> dict:
    out_dir = out_root / code
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"[{code}] {xlsm} -> {out_dir}")
    modules = extract_vba(xlsm, out_dir)
    dumps = []
    with zipfile.ZipFile(xlsm) as zf:
        structure = extract_structure(zf)
        shared = _load_shared_strings(zf)
        dumps, visible_cells = dump_hidden_sheets(zf, structure, shared,
                                                  out_dir)
        dropdowns = resolve_dropdowns(zf, structure, shared, out_dir)
        fields = build_fields(structure, dropdowns, visible_cells, out_dir)
    structure["hiddenSheetDumps"] = dumps
    structure["fieldCount"] = len(fields)
    csvs = companion_csvs(csv_folder)
    if csvs:
        structure["companionCSVs"] = csvs
    (out_dir / "structure.json").write_text(
        json.dumps(structure, indent=2, ensure_ascii=False), encoding="utf-8")
    write_summary(code, xlsm, modules, structure, csvs, out_dir)
    print(f"  sheets={len(structure['sheets'])} modules={len(modules)} "
          f"vba_lines={sum(m['lines'] for m in modules):,} "
          f"dumps={len(dumps)} "
          f"cells={sum(d['cells'] for d in dumps):,} fields={len(fields)}")
    return {"code": code, "xlsm": str(xlsm.relative_to(REPO)),
            "out": str(out_dir), "modules": modules,
            "sheet_count": len(structure["sheets"]),
            "hidden_dumps": dumps, "fields": len(fields)}


def main(argv: list[str]) -> int:
    out_root = REPO / "extracted" / "xlsm"
    results = []
    if len(argv) > 1:
        for arg in argv[1:]:
            xlsm = Path(arg)
            code = xlsm.stem
            results.append(process(code, xlsm, None, out_root))
    else:
        for code, (rel, csvdir) in UTILITIES.items():
            xlsm = REPO / rel
            if not xlsm.exists():
                print(f"[{code}] SKIP (missing {rel})")
                continue
            results.append(process(code, xlsm,
                                   REPO / csvdir if csvdir else None,
                                   out_root))
    (out_root / "index.json").write_text(
        json.dumps(results, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
