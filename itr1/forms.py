"""ITR form eligibility radar — recommends ITR-1 / ITR-2 / ITR-3 / ITR-4 for
FY 2025-26 (AY 2026-27) from the income profile, mirroring CBDT eligibility
rules (CBDT notifications + the validation-rule PDFs).

Only generation of ITR-1, ITR-2 and ITR-4 is implemented; when ITR-3 is the
right form the recommendation tells the user exactly why (and downloads for
ITR-1/2/4 remain available for the compatible heads of income).
"""

from __future__ import annotations

import copy

from . import capgains as CG


def _num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def detect_itr_form(data: dict) -> dict:
    """Return ``{recommended, reasons, eligible}``.

    ``eligible`` maps every form to (ok, [blocking/reason strings]) so the UI
    can render a radio group with explanations per option.
    """
    elig = data.get("eligibility") or {}
    foreign = bool(elig.get("foreign_assets_income"))
    director = bool(elig.get("director_unlisted"))
    rs = (((data.get("personal") or {}).get("residential_status")
           or "").upper())
    is_nri = rs == "NRI"
    # RNOR is resident u/s 6; legacy `eligibility.rnor` flag still honoured.
    is_nor = rs == "NOR" or bool(elig.get("rnor"))

    cg = CG.compute_capital_gains(data) if (data.get("capital_gains") or {}) else None
    has_cg_beyond_112a = bool(cg and (not cg["itr1_ok"]))

    presumpt = data.get("presumptive") or {}
    has_business = any((presumpt.get(k) or []) for k in ("44ad", "44ada", "44ae"))
    manual_bp = _num(presumpt.get("business_income_manual"))
    has_business = has_business or manual_bp > 0

    hp_count = len(data.get("house_properties") or [])
    ltcg_manual = 0.0
    l_in = data.get("ltcg112a") or {}
    ltcg_manual = max(0, _num(l_in.get("sale_consideration")) - _num(l_in.get("cost_of_acquisition")))
    agri = 0.0
    for r in data.get("exempt_income") or []:
        if str(r.get("sub_category")) == "10(1)":
            agri += _num(r.get("amount"))
    losses_prior = bool((data.get("losses_cf") or {}).get("house_property") or
                        (data.get("losses_cf") or {}).get("capital_gains"))

    hp_loss = any(_num(p.get("interest_on_borrowed_capital")) + _num(p.get("municipal_taxes"))
                  > _num(p.get("annual_lettable_value")) - _num(p.get("rent_not_realized"))
                  and str(p.get("occupancy")) != "S"
                  for p in data.get("house_properties") or [])

    # ---- blocking reasons per form ----
    reasons = {"ITR1": [], "ITR2": [], "ITR3": [], "ITR4": []}

    if hp_count > 2:
        reasons["ITR1"].append("more than two house properties")
        reasons["ITR4"].append("more than two house properties")
    if ltcg_manual > 125000 and not has_cg_beyond_112a:
        reasons["ITR1"].append("LTCG u/s 112A exceeds Rs. 1.25 lakh")
        reasons["ITR4"].append("LTCG u/s 112A exceeds Rs. 1.25 lakh")
    if agri > 5000:
        reasons["ITR1"].append("agricultural income exceeds Rs. 5,000")
        reasons["ITR4"].append("agricultural income exceeds Rs. 5,000")
    if losses_prior:
        reasons["ITR1"].append("brought-forward losses to be set off")
    if has_cg_beyond_112a:
        reasons["ITR1"].append("capital gains beyond the ITR-1 scope "
                               "(STCG / LTCG above Rs. 1.25L / CG losses)")
        reasons["ITR4"].append("capital gains beyond 112A within limit")
    if foreign:
        for f in ("ITR1", "ITR4"):
            reasons[f].append("foreign assets / foreign income to report")
    if is_nri:
        # ITR-1: EfilingCommon.bas — "ITR 1 is for individuals being a
        #   resident (other than not ordinarily resident)…"
        # ITR-4: Sheet1 Worksheet_Change → Lock44AD locks & clears the
        #   44AD/44ADA cells for 'NR' (ss. 44AD(1)/44ADA(1): resident only);
        #   Sheet4.cls — "Returns for income of non-residents are required
        #   to be filed in ITR 3/5."
        reasons["ITR1"].append("non-resident — ITR-1 is only for a resident "
                               "individual (other than not ordinarily resident)")
        reasons["ITR4"].append("non-resident — presumptive sections 44AD/44ADA "
                               "require a resident assessee (the official utility "
                               "locks those cells; file ITR-3)")
    if is_nor:
        # RNOR is resident u/s 6 → 44AD/44ADA remain available in ITR-4
        # (the utility locks them only for "NR" = NRI).  ITR-1, however, is
        # ROR-only.
        reasons["ITR1"].append("resident but not ordinarily resident (RNOR) — "
                               "ITR-1 is only for an ordinarily resident individual")
    if director:
        for f in ("ITR1", "ITR2", "ITR4"):
            reasons[f].append("director in a company / unlisted equity shares held")
    if has_business:
        reasons["ITR1"].append("business / professional income")
        reasons["ITR2"].append("business / professional income")
    if manual_bp > 0:
        reasons["ITR4"].append("non-presumptive business income (books of account)")

    eligible = {f: (len(frm_reasons) == 0, frm_reasons)
                for f, frm_reasons in reasons.items()}

    # ---- recommendation ----
    if foreign or is_nri or director:
        rec = "ITR3" if has_business else "ITR2"
        why = ("foreign assets/income, non-resident status or directorship — these can only "
               "be reported in " + rec)
    elif has_business and not eligible["ITR4"][0]:
        rec, why = "ITR3", "business/professional income is not under the presumptive scheme"
    elif has_business:
        rec, why = "ITR4", "presumptive business/professional income (44AD/44ADA/44AE)"
    elif not eligible["ITR1"][0]:
        rec, why = "ITR2", "; ".join(reasons["ITR1"][:2])
    else:
        rec, why = "ITR1", "all incomes fit the ITR-1 (Sahaj) scope"

    return {"recommended": rec, "why": why,
            "eligible": {f: {"ok": ok, "reasons": frm_reasons}
                         for f, (ok, frm_reasons) in eligible.items()},
            "has_business": has_business, "has_cg_beyond_112a": has_cg_beyond_112a,
            "residential_status": rs or ("NOR" if is_nor else None)}


# ---------------------------------------------------------------------------
# build dispatcher
# ---------------------------------------------------------------------------

def supported_forms() -> list[str]:
    return ["ITR1", "ITR2", "ITR4"]


def build_for_form(data: dict, form: str | None = None, *, creation_info=None) -> dict:
    """Dispatch to the right builder.  ``form=None`` → auto-detect."""
    form = (form or "").upper() or detect_itr_form(data)["recommended"]
    if form == "ITR3":
        from .errors import DataError
        raise DataError("ITR-3 generation is not available yet (balance-sheet & P&L "
                        "schedules). File ITR-3 on the department utility, or switch "
                        "to the presumptive scheme (44AD/44ADA/44AE) and use ITR-4 "
                        "here if eligible.", field="ITR form")
    if form in ("ITR1", "ITR-1"):
        from .builder import build_return
        return build_return(data, creation_info=creation_info)
    if form in ("ITR2", "ITR-2"):
        from .itr2_builder import build_itr2
        return build_itr2(data, creation_info=creation_info)
    if form in ("ITR4", "ITR-4"):
        from .itr4_builder import build_itr4
        return build_itr4(data, creation_info=creation_info)
    from .errors import DataError
    raise DataError(f"Unknown ITR form '{form}'.", field="ITR form")


def compute_both_regimes_form(data: dict, form: str | None = None) -> dict:
    """{'old': build, 'new': build} for the selected (or recommended) form."""
    from .forms import detect_itr_form as _detect  # self-import for clarity
    form = (form or "").upper() or _detect(data)["recommended"]

    def build(opt_out):
        cp = copy.deepcopy(data or {})
        cp.setdefault("filing", {})["opt_out_new_regime"] = opt_out
        return build_for_form(cp, form)

    return {"old": build(True), "new": build(False), "form": form}
