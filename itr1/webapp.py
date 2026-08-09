"""Flask web application - guided ITR-1 (Sahaj) return preparation wizard.

Run directly with::

    python run_web.py               # serves http://127.0.0.1:5000/

The UI is a single-page wizard (static/app.js) that keeps the taxpayer data
model in the browser, calls the APIs below for computation / validation, and
downloads the finished e-filing JSON.
"""

from __future__ import annotations

import copy
import json
import os

from . import ASSESSMENT_YEAR, __version__
from . import constants as C
from .builder import build_return, compute_both_regimes
from .errors import DataError
from .validator import validate_all, summarize

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, os.pardir, "web")
SAMPLE_PATH = os.path.join(BASE_DIR, os.pardir, "samples", "sample_taxpayer.json")


def create_app():
    from flask import Flask, jsonify, request, send_from_directory, Response

    app = Flask(__name__,
                static_folder=os.path.abspath(os.path.join(WEB_DIR, "static")),
                static_url_path="/static")

    def input_error(exc):
        """400 response that names the exact form field at fault."""
        return jsonify({
            "error": str(exc),
            "field": getattr(exc, "field", ""),
            "path": getattr(exc, "path", ""),
        }), 400

    # -- pages -----------------------------------------------------------
    @app.get("/")
    def index():
        return send_from_directory(os.path.join(os.path.abspath(WEB_DIR), "templates"),
                                   "index.html")

    # -- metadata --------------------------------------------------------
    @app.get("/api/meta")
    def meta():
        return jsonify({
            "version": __version__,
            "assessment_year": ASSESSMENT_YEAR,
            "due_date": C.DUE_DATE_FILING,
            "state_codes": C.STATE_CODES,
            "employer_categories": C.EMPLOYER_CATEGORIES,
            "return_sections": {str(k): v for k, v in C.RETURN_SECTIONS.items()},
            "exempt_allowance_sections": C.EXEMPT_ALLOWANCE_SECTIONS,
            "other_source_natures": C.OTHER_SOURCE_NATURES,
            "std_deduction": {"new": C.STD_DEDUCTION_16IA_NEW, "old": C.STD_DEDUCTION_16IA_OLD},
            "rebate87a": {"new_max": C.REBATE_87A_NEW_MAX,
                          "new_income_limit": C.REBATE_87A_NEW_INCOME_LIMIT,
                          "old_max": C.REBATE_87A_OLD_MAX,
                          "old_income_limit": C.REBATE_87A_OLD_INCOME_LIMIT},
        })

    @app.get("/api/sample")
    def sample():
        with open(SAMPLE_PATH, "r", encoding="utf-8") as fh:
            return jsonify(json.load(fh))

    # -- computation -----------------------------------------------------
    @app.post("/api/summary")
    def summary():
        data = request.get_json(force=True) or {}
        try:
            both = compute_both_regimes(data)
        except DataError as exc:  # which field, exactly, is wrong
            return input_error(exc)
        except Exception as exc:  # input errors -> 400 with message
            app.logger.exception("summary failed")
            return jsonify({"error": str(exc)}), 400
        out = {}
        for regime, res in both.items():
            rep = res["report"]
            tax = rep["tax"]
            out[regime] = {
                "gross_total_income": rep["gross_total_income_incl_ltcg"],
                "deductions": rep["deductions_eligible"]["TotalChapVIADeductions"],
                "total_income": rep["total_income"],
                "tax_on_total_income": tax["TotalTaxPayable"],
                "rebate_87a": tax["Rebate87A"],
                "cess": tax["EducationCess"],
                "gross_tax_liability": tax["GrossTaxLiability"],
                "interest_fee": tax["TotalIntrstPay"],
                "total_tax_interest": tax["TotTaxPlusIntrstPay"],
                "taxes_paid": tax["TaxesPaid"],
                "balance_payable": tax["BalTaxPayable"],
                "refund_due": tax["RefundDue"],
                "warnings": rep["warnings"],
                "slab_breakup": _breakup(data, regime == "old"),
            }
        chosen = bool((data.get("filing") or {}).get("opt_out_new_regime"))
        out["recommended"] = "old" if (both["old"]["report"]["tax"]["TotTaxPlusIntrstPay"]
                                       < both["new"]["report"]["tax"]["TotTaxPlusIntrstPay"]) else "new"
        out["chosen"] = "old" if chosen else "new"
        # live rule issues on the chosen regime only (schema layer skipped for
        # speed - the builder produces schema-valid documents by construction)
        from .validator import validate_rules
        chosen_doc = both["old" if chosen else "new"]["json"]
        try:
            issues = validate_rules(chosen_doc)
            out["issues"] = [i.as_dict() for i in issues]
            out["issue_summary"] = summarize(issues)
        except Exception:
            app.logger.exception("live rule validation failed")
            out["issues"] = []
            out["issue_summary"] = {"A": 0, "B": 0, "D": 0, "SCHEMA": 0, "ok": None}
        return jsonify(out)

    @app.post("/api/generate")
    def generate():
        data = request.get_json(force=True) or {}
        try:
            result = build_return(data)
            doc = result["json"]
            issues = validate_all(doc)
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("generate failed")
            return jsonify({"error": str(exc)}), 400
        return jsonify({
            "json": doc,
            "report": result["report"],
            "issues": [i.as_dict() for i in issues],
            "issue_summary": summarize(issues),
        })

    def _data_and_form():
        """Accept either the raw taxpayer object or ``{data, form}``."""
        body = request.get_json(force=True, silent=True) or {}
        if isinstance(body.get("data"), dict):
            return body["data"], (body.get("form") or "").upper() or None
        form = (body.get("form") or "").upper() or None
        body.pop("form", None)
        return body, form

    @app.post("/api/computation")
    def computation():
        """Verbose line-by-line computation report for both regimes (JSON)."""
        from .report import build_computation
        data, form = _data_and_form()
        try:
            return jsonify(build_computation(data, form=form))
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("computation failed")
            return jsonify({"error": str(exc)}), 400

    @app.post("/api/computation/pdf")
    def computation_pdf():
        from .report import build_computation
        from .compfile import computation_pdf as render_pdf
        data, form = _data_and_form()
        try:
            comp = build_computation(data, form=form)
            payload = render_pdf(comp)
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("pdf failed")
            return jsonify({"error": str(exc)}), 400
        pan = ((data.get("personal") or {}).get("pan") or "ITR").upper()
        tag = (form or comp["meta"]["form"]).split(" ")[0].replace("-", "")
        return Response(payload, mimetype="application/pdf", headers={
            "Content-Disposition":
                f"attachment; filename={tag}_Computation_AY2026-27_{pan}.pdf"})

    @app.post("/api/computation/docx")
    def computation_docx():
        from .report import build_computation
        from .compfile import computation_docx as render_docx
        data, form = _data_and_form()
        try:
            comp = build_computation(data, form=form)
            payload = render_docx(comp)
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("docx failed")
            return jsonify({"error": str(exc)}), 400
        pan = ((data.get("personal") or {}).get("pan") or "ITR").upper()
        tag = (form or comp["meta"]["form"]).split(" ")[0].replace("-", "")
        return Response(
            payload,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition":
                     f"attachment; filename={tag}_Computation_AY2026-27_{pan}.docx"})

    @app.post("/api/validate")
    def validate_upload():
        """Validate an already-built official ITR-1 JSON document."""
        doc = request.get_json(force=True) or {}
        if "ITR" not in doc:
            return jsonify({"error": "Expecting the official ITR JSON document "
                                     "(root key 'ITR')."}), 400
        issues = validate_all(doc)
        return jsonify({"issues": [i.as_dict() for i in issues],
                        "issue_summary": summarize(issues)})

    @app.post("/api/download")
    def download():
        data = request.get_json(force=True) or {}
        try:
            result = build_return(data)
            doc = result["json"]
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("download failed")
            return jsonify({"error": str(exc)}), 400
        pan = ((data.get("personal") or {}).get("pan") or "ITR1").upper()
        payload = json.dumps(doc, indent=2)
        return Response(
            payload,
            mimetype="application/json",
            headers={"Content-Disposition":
                     f"attachment; filename=ITR1_AY2026-27_{pan}.json"})

    # ==================================================================
    # V2 — multi-form suite: detect, ITR-2/4, capital gains, imports
    # ==================================================================

    from . import forms as _forms
    from . import imports as _imp

    def _builds_for(data, form):
        """Both regimes for the chosen (or detected) form."""
        return _forms.compute_both_regimes_form(data, form)

    def _rep_summary(rep):
        tax = rep["tax"]
        return {
            "gross_total_income": rep.get("gross_total_income_incl_cg",
                                          rep.get("gross_total_income_incl_ltcg", 0)),
            "deductions": rep.get("deductions_eligible", {}).get("TotalChapVIADeductions", 0),
            "total_income": rep["total_income"],
            "tax_on_total_income": tax["TotalTaxPayable"],
            "rebate_87a": tax["Rebate87A"],
            "cess": tax["EducationCess"],
            "gross_tax_liability": tax["GrossTaxLiability"],
            "interest_fee": tax["TotalIntrstPay"],
            "total_tax_interest": tax["TotTaxPlusIntrstPay"],
            "taxes_paid": tax["TaxesPaid"],
            "balance_payable": tax["BalTaxPayable"],
            "refund_due": tax["RefundDue"],
            "warnings": rep["warnings"],
        }

    @app.post("/api/forms/detect")
    def forms_detect():
        data = request.get_json(force=True) or {}
        try:
            return jsonify(_forms.detect_itr_form(data))
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("detect failed")
            return jsonify({"error": str(exc)}), 400

    @app.post("/api/summary/v2")
    def summary_v2():
        """Both-regime comparison for the selected ITR form."""
        body = request.get_json(force=True) or {}
        data = body.get("data") or body
        form = (body.get("form") or "").upper() or None
        try:
            both = _builds_for(data, form)
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("summary v2 failed")
            return jsonify({"error": str(exc)}), 400
        chosen = bool((data.get("filing") or {}).get("opt_out_new_regime"))
        out = {r: _rep_summary(both[r]["report"]) for r in ("old", "new")}
        out["form"] = both["form"]
        out["recommended"] = (
            "old" if both["old"]["report"]["tax"]["TotTaxPlusIntrstPay"]
            < both["new"]["report"]["tax"]["TotTaxPlusIntrstPay"] else "new")
        out["chosen"] = "old" if chosen else "new"
        return jsonify(out)

    def _form_generate_download(form):
        def generate():
            data = request.get_json(force=True) or {}
            try:
                result = _forms.build_for_form(data, form)
            except DataError as exc:
                return input_error(exc)
            except Exception as exc:
                app.logger.exception("generate %s failed", form)
                return jsonify({"error": str(exc)}), 400
            from .schemakit import validate_with_schema
            schema_errs = validate_with_schema(form, result["json"])
            issues = validate_all(result["json"], form=form)
            return jsonify({"json": result["json"], "report": result["report"],
                            "schema_errors": schema_errs,
                            "issues": [i.as_dict() for i in issues],
                            "issue_summary": summarize(issues)})

        def dl():
            data = request.get_json(force=True) or {}
            try:
                result = _forms.build_for_form(data, form)
            except DataError as exc:
                return input_error(exc)
            except Exception as exc:
                app.logger.exception("download %s failed", form)
                return jsonify({"error": str(exc)}), 400
            pan = ((data.get("personal") or {}).get("pan") or form).upper()
            return Response(
                json.dumps(result["json"], indent=2),
                mimetype="application/json",
                headers={"Content-Disposition":
                         f"attachment; filename={form}_AY2026-27_{pan}.json"})

        return generate, dl

    g2, d2 = _form_generate_download("ITR2")
    app.add_url_rule("/api/itr2/generate", "itr2_generate", g2, methods=["POST"])
    app.add_url_rule("/api/itr2/download", "itr2_download", d2, methods=["POST"])
    g4, d4 = _form_generate_download("ITR4")
    app.add_url_rule("/api/itr4/generate", "itr4_generate", g4, methods=["POST"])
    app.add_url_rule("/api/itr4/download", "itr4_download", d4, methods=["POST"])

    # ---- capital-gain Excel round-trip --------------------------------

    @app.post("/api/cg/xlsx")
    def cg_export_xlsx():
        from .capgains import export_xlsx
        data = request.get_json(force=True) or {}
        try:
            payload = export_xlsx(data)
        except DataError as exc:
            return input_error(exc)
        pan = ((data.get("personal") or {}).get("pan") or "CG").upper()
        return Response(payload, mimetype="application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition":
                                 f"attachment; filename=CapitalGains_{pan}.xlsx"})

    @app.post("/api/cg/import-xlsx")
    def cg_import_xlsx():
        """multipart upload of our CG workbook (or the dept CSV_112A)."""
        from .capgains import import_xlsx, import_csv_112a
        blob = b""
        filename = ""
        if request.files.get("file"):
            blob = request.files["file"].read()
            filename = request.files["file"].filename or ""
        else:
            blob = request.get_data() or b""
            filename = request.headers.get("X-Filename", "")
        if not blob:
            return jsonify({"error": "No file received.", "field": "Excel/CSV file"}), 400
        try:
            if filename.lower().endswith(".csv") or blob[:3] != b"PK\x03":
                merged = import_csv_112a(blob.decode("utf-8", "replace"))
            else:
                merged = import_xlsx(blob)
        except DataError as exc:
            return input_error(exc)
        except Exception as exc:
            app.logger.exception("cg import failed")
            return jsonify({"error": f"Could not read the file ({exc}).",
                            "field": "Excel/CSV file"}), 400
        return jsonify({"capital_gains": merged.get("capital_gains", {}),
                        "counts": {k: len(v) for k, v in
                                   (merged.get("capital_gains") or {}).items()}})

    @app.get("/api/cg/template")
    def cg_template():
        from .capgains import cg_template_json
        return Response(cg_template_json(), mimetype="application/json",
                        headers={"Content-Disposition":
                                 "attachment; filename=capital_gains_template.json"})

    # ---- imports: ITR JSON / AIS / 26AS -------------------------------

    def _posted_file_and_data():
        blob = b""
        filename = ""
        if request.files.get("file"):
            blob = request.files["file"].read()
            filename = request.files["file"].filename or ""
        raw = request.form.get("data") or request.args.get("data") or "{}"
        try:
            data = json.loads(raw) if isinstance(raw, str) else (raw or {})
        except ValueError:
            data = {}
        return blob, filename, data

    @app.post("/api/import/itr-json")
    def import_itr_json():
        blob, filename, data = _posted_file_and_data()
        if not blob and request.is_json:
            body = request.get_json(force=True) or {}
            blob = json.dumps(body.get("document") or {}).encode()
            data = body.get("data") or {}
        if not blob:
            return jsonify({"error": "Choose the previous-year ITR JSON file first.",
                            "field": "ITR JSON file"}), 400
        try:
            summary = _imp.parse_itr_json(blob)
            apply = request.form.get("apply", "1") != "0"
            out_data, notes = _imp.apply_itr_prefill(data, summary) if apply else (data, [])
        except DataError as exc:
            return input_error(exc)
        return jsonify({"summary": summary, "notes": notes,
                        "merged_data": out_data if apply else None})

    @app.post("/api/import/ais")
    def import_ais():
        blob, filename, data = _posted_file_and_data()
        if not blob:
            return jsonify({"error": "Choose the AIS file (JSON or ZIP) first.",
                            "field": "AIS file"}), 400
        try:
            ais = _imp.parse_ais(blob, filename, data)
            candidates = _imp.ais_income_candidates(ais)
            tds_pk = _ais_tax_credits(ais)
            comparison = _imp.compare_with_ais(data, ais, {})
        except DataError as exc:
            return input_error(exc)
        cats: dict[str, int] = {}
        for e in ais.get("entries", []):
            cats[e["category"]] = cats.get(e["category"], 0) + 1
        return jsonify({
            "assessment_year": ais.get("assessment_year", ""),
            "password_detected": ais.get("password_used", False),
            "entry_count": len(ais.get("entries", [])),
            "categories": cats,
            "entries": ais.get("entries", [])[:200],
            "income_candidates": candidates,
            "tax_credits": tds_pk,
            "comparison": comparison,
        })

    @app.post("/api/import/ais/apply")
    def import_ais_apply():
        body = request.get_json(force=True) or {}
        data = body.get("data") or {}
        try:
            out = data
            notes: list[str] = []
            if body.get("income_ids") is not None and body.get("candidates"):
                out, notes = _imp.apply_ais_items(out, body["candidates"],
                                                  body["income_ids"])
            if body.get("tax_credits"):
                out, tnotes = _imp.apply_26as(out, body["tax_credits"],
                                              body.get("credit_kinds"))
                notes += tnotes
        except DataError as exc:
            return input_error(exc)
        return jsonify({"merged_data": out, "notes": notes})

    @app.post("/api/import/26as")
    def import_26as():
        blob, filename, data = _posted_file_and_data()
        if not blob:
            return jsonify({"error": "Choose the Form 26AS (PDF or TXT) file first.",
                            "field": "26AS file"}), 400
        try:
            parsed = _imp.parse_26as(blob, filename, data)
            apply = request.form.get("apply", "0") == "1"
            out_data, notes = _imp.apply_26as(data, parsed) if apply else (data, [])
        except DataError as exc:
            return input_error(exc)
        return jsonify({"parsed": parsed, "password_detected": parsed.get("password_used", False),
                        "counts": {k: len(parsed.get(k) or []) for k in
                                   ("tds_salary", "tds_others", "tds_rent", "tcs", "challans")},
                        "notes": notes, "merged_data": out_data if apply else None})

    @app.post("/api/yoy")
    def yoy():
        """Year-on-year comparison: current data (any form) vs previous-year summary."""
        body = request.get_json(force=True) or {}
        data = body.get("data") or {}
        prev = body.get("prev_summary") or {}
        cur_heads_body = body.get("current_heads")
        try:
            if cur_heads_body is None:
                form = (body.get("form") or "").upper() or None
                result = _forms.build_for_form(data, form)
                rep = result["report"]
                cur_heads_body = {
                    "salary": rep.get("salary_income", 0),
                    "hp": rep.get("hp_income", 0),
                    "business": rep.get("business_income", 0),
                    "os": rep.get("other_sources_income", 0),
                    "cg": (rep.get("capital_gains") or {}).get("total_cg", 0),
                    "gti": rep.get("gross_total_income_incl_cg",
                                   rep.get("gross_total_income_incl_ltcg", 0)),
                    "total_income": rep["total_income"],
                }
            rows = _imp.year_on_year(cur_heads_body, prev)
        except DataError as exc:
            return input_error(exc)
        return jsonify({"rows": rows, "prev": {"form": prev.get("form"),
                                               "assessment_year": prev.get("assessment_year"),
                                               "losses": prev.get("losses")}})

    @app.post("/api/salary/relief89")
    def relief89():
        from . import salaryx as SX
        body = request.get_json(force=True) or {}
        data = body.get("data") or body
        exp = data.get("_salary_expansion")
        if exp is None:
            sal = data.get("salary") or {}
            exp = SX.expand_salary(data) if SX.has_rich_salary(sal) else None
        if not exp:
            return jsonify({"relief": 0, "rows": [],
                            "message": "No arrear salary rows in the salary section."})
        filing = data.get("filing") or {}
        personal = data.get("personal") or {}
        dob = str(personal.get("dob") or "1990-01-01")
        dob_year = int(dob[:4]) if dob[:4].isdigit() else None
        try:
            # compute current total income first (needed for the Form-10E deltas)
            result = _forms.build_for_form(data, None)
            total_income = result["report"]["total_income"]
            w = SX.relief_89_working(total_income, exp.get("arrears_rows") or [],
                                     regime="old" if filing.get("opt_out_new_regime") else "new",
                                     dob_year=dob_year)
        except DataError as exc:
            return input_error(exc)
        return jsonify(w)

    @app.post("/api/salary/monthly-fill")
    def monthly_fill():
        """Expand a monthly salary grid into annuals (for the double-click modal)."""
        from . import salaryx as SX
        body = request.get_json(force=True) or {}
        sal = body.get("salary") or body
        try:
            monthly = SX.aggregate_monthly(sal.get("monthly"))
        except DataError as exc:
            return input_error(exc)
        return jsonify(monthly)

    @app.get("/api/salary/catalog")
    def salary_catalog():
        from . import salaryx as SX
        from . import business_codes as NOB
        return jsonify({
            "components": {k: {"label": v[0], "section": v[1], "cap": v[2]}
                           for k, v in SX.COMPONENTS.items()},
            "perquisites": {k: v for k, v in SX.PERQUISITE_TYPES.items()},
            "terminal": ["gratuity", "leave_encashment", "commuted_pension",
                         "vrs", "retrenchment"],
            # Official nature-of-business dropdowns, extracted verbatim from
            # the CBDT utilities (itr1/business_codes.py — generated file).
            "business_codes": {
                "44ad": NOB.NOB_44AD_BY_CODE,
                "44ada": NOB.NOB_44ADA_BY_CODE,
                "44ae": NOB.NOB_44AE_BY_CODE,
            },
            "business_code_options": {
                "44ad": NOB.options("44ad"),
                "44ada": NOB.options("44ada"),
                "44ae": NOB.options("44ae"),
                "master": NOB.options("master"),
            },
        })

    return app


def _ais_tax_credits(ais: dict) -> dict:
    """Shape AIS tax-credit entries like a parsed 26AS for the apply flow."""
    import re
    out = {"tds_salary": [], "tds_others": [], "tds_rent": [], "tcs": [], "challans": []}
    tan_re = re.compile(r"[A-Z]{4}[0-9]{5}[A-Z]")
    for e in ais.get("entries", []):
        fields = e.get("fields") or {}
        source = str(fields.get("informationSource") or fields.get("deductorName") or "")
        name = source or e["description"][:60]
        tan = str(fields.get("deductorTAN") or fields.get("tanOfDeductor") or "")
        if not tan:
            m = tan_re.search(source)
            tan = m.group(0) if m else ""
        amt = int(round(e.get("amount") or 0))
        if amt <= 0:
            continue
        if e["category"] == "tds_salary":
            out["tds_salary"].append({"tan": tan, "name": name,
                                      "amount_paid": int(fields.get("grossAmount") or 0),
                                      "tax": amt, "section": "192"})
        elif e["category"] == "tds_other":
            sec = "94A" if "interest" in e["context"].lower() else "94J"
            out["tds_others"].append({"tan": tan, "name": name, "section": sec,
                                      "amount_paid": int(fields.get("grossAmount") or 0),
                                      "tax": amt})
        elif e["category"] == "tcs":
            out["tcs"].append({"tan": tan, "name": name, "tax": amt,
                               "amount_paid": int(fields.get("grossAmount") or 0)})
    return out


def _breakup(data: dict, opt_out: bool) -> list:
    from .tax_engine import slab_breakup
    from .builder import build_return
    cp = copy.deepcopy(data or {})
    cp.setdefault("filing", {})["opt_out_new_regime"] = opt_out
    rep = build_return(cp)["report"]
    ltcg = rep["ltcg112a"]
    ti_excl = max(0, rep["total_income"] - ltcg)
    dob = (cp.get("personal") or {}).get("dob", "1990-01-01")
    return slab_breakup(ti_excl, opt_out, dob)


app = create_app()

if __name__ == "__main__":  # pragma: no cover
    app.run(debug=True)
