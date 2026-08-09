"""Advanced salary engine — salary-slip components, allowances, perquisites,
terminal benefits (gratuity / leave encashment / commuted pension / VRS /
retrenchment), monthly payroll grid and arrear salary with Form 10E relief
u/s 89.

The engine *expands* a rich ``salary`` block into the plain totals the core
builder consumes (``salary_17_1``, ``perquisites_17_2``, ``profits_17_3``,
exempt-allowance rows, HRA inputs, professional tax, …) so all downstream
logic keeps working, and additionally produces human-readable workings for the
computation report / PDF.
"""

from __future__ import annotations

from . import constants as C
from . import history as H
from .errors import field_scope, parse_int

# ---------------------------------------------------------------------------
# Salary-slip component catalogue
# ---------------------------------------------------------------------------
# code → (label, exempt section or None, statutory annual cap on the exempt part)
COMPONENTS = {
    "BASIC":   ("Basic salary", None, None),
    "DA":      ("Dearness allowance (forms part of retirement benefits)", None, None),
    "HRA":     ("House rent allowance", "10(13A)", None),     # exempt via HRA schedule
    "LTA":     ("Leave travel allowance / concession", "10(5)", None),
    "CONVEY":  ("Conveyance / transport allowance", "10(14)(ii)", 19200),
    "CONVEY_DIV": ("Transport allowance (divyang employee)", "10(14)(ii)(115BAC)", 38400),
    "CEA":     ("Children education allowance (≤2 children)", "10(14)(ii)", 2400),
    "HOSTEL":  ("Hostel expenditure allowance (≤2 children)", "10(14)(ii)", 7200),
    "PHONE":   ("Telephone / mobile allowance (duty)", "10(14)(i)", None),
    "UNIFORM": ("Uniform allowance (duty)", "10(14)(i)", None),
    "CARMNT":  ("Car maintenance allowance (duty)", "10(14)(i)", None),
    "MEDICAL": ("Medical allowance (taxable since FY 2018-19)", None, None),
    "SPECIAL": ("Special allowance", None, None),
    "CCA":     ("City compensatory allowance", None, None),
    "OT":      ("Overtime allowance", None, None),
    "BONUS":   ("Bonus / performance incentive", None, None),
    "COMMISSION": ("Commission", None, None),
    "ARREARS": ("Arrears of salary (current year)", None, None),
    "OTHR":    ("Other allowance", None, None),
}

# sections never exempt under the new regime (115BAC) — mirrors builder logic
NEW_REGIME_BANNED = {"10(5)", "10(13A)", "10(14)(i)", "10(14)(ii)", "10(17)", "EIC"}

PERQUISITE_TYPES = {
    "RFA": "Rent-free / concessional accommodation",
    "CAR": "Motor car provided by employer",
    "LOAN": "Interest-free / concessional loan",
    "DOMESTIC": "Domestic servant / sweeper / gardener",
    "GAS": "Gas, electricity, water",
    "EDU": "Free / concessional education",
    "CLUB": "Club membership / credit card",
    "MOVE": "Movable asset transferred / used",
    "OTHR": "Other perquisite",
}

GRATUITY_LIMIT = 2000000          # sec 10(10) ceiling
LEAVE_ENCASH_LIMIT = 2500000      # sec 10(10AA) ceiling (raised by FA 2023)
VRS_LIMIT = 500000                # sec 10(10C)
RETRENCH_LIMIT = 500000           # sec 10(10B)

MONTHS = ["2025-04", "2025-05", "2025-06", "2025-07", "2025-08", "2025-09",
          "2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03"]
MONTH_NAMES = ["Apr 2025", "May 2025", "Jun 2025", "Jul 2025", "Aug 2025",
               "Sep 2025", "Oct 2025", "Nov 2025", "Dec 2025", "Jan 2026",
               "Feb 2026", "Mar 2026"]

RICH_SALARY_KEYS = ("components", "perquisites", "profits_detail", "terminal",
                    "monthly", "arrears", "employer")


def has_rich_salary(sal: dict) -> bool:
    """True when any of the advanced salary inputs are present/non-empty."""
    if not sal:
        return False
    if (sal.get("employer") or {}).get("name") or (sal.get("employer") or {}).get("tan"):
        return True
    for k in RICH_SALARY_KEYS:
        v = sal.get(k)
        if isinstance(v, list) and any(True for _ in v):
            return True
        if isinstance(v, dict) and any(
                (isinstance(x, dict) and any(str(y) not in ("", "0") for y in x.values()))
                or (x not in ("", None, 0, "0") and not isinstance(x, (dict, list)))
                for x in v.values()):
            return True
    return False


def _g(mapping, key, default=0, label=None):
    if mapping is None or key not in mapping:
        return default
    return parse_int(mapping.get(key), default=default,
                     field=label, key_path=str(key))


def _years_completed(years, months=0):
    """Completed years of service (6+ months rounds up — Payment of Gratuity Act)."""
    return int(years) + (1 if int(months) >= 6 else 0)


# ---------------------------------------------------------------------------
# Terminal benefits
# ---------------------------------------------------------------------------

def gratuity_exemption(d: dict) -> dict:
    """Gratuity u/s 10(10) — least-of-three for non-government employees."""
    received = _g(d, "received", label="Gratuity received")
    govt = bool(d.get("govt"))
    years = _years_completed(_g(d, "years", label="Years of service"),
                             _g(d, "months_extra", label="Extra months"))
    touching = []
    if received <= 0:
        return {"exempt": 0, "taxable": 0, "workings": touching}
    if govt:
        touching = [(f"Gratuity received from Government: Rs. {received:,}", received)]
        return {"exempt": received, "taxable": 0, "workings": touching,
                "limit": "fully exempt (Government)"}
    covered = bool(d.get("covered_act", True))
    if covered:
        salary = _g(d, "last_drawn_basic_da", label="Last drawn basic + DA")
        formula = salary * 15 * years // 26 if years else 0
        basis = f"15/26 × Rs. {salary:,} × {years} yr(s) = Rs. {formula:,}"
    else:
        salary = _g(d, "avg_10m_salary", label="Average salary of last 10 months")
        formula = salary * 15 // 30 * years if years else 0
        basis = f"15/30 × Rs. {salary:,} × {years} yr(s) = Rs. {formula:,}"
    exempt = min(received, formula, GRATUITY_LIMIT)
    touching = [
        ("(i) Gratuity actually received", received),
        (f"(ii) {basis}", formula),
        (f"(iii) Statutory ceiling u/s 10(10)", GRATUITY_LIMIT),
        (f"Exempt = least of (i), (ii), (iii)", exempt),
    ]
    return {"exempt": exempt, "taxable": received - exempt, "workings": touching}


def leave_encashment_exemption(d: dict) -> dict:
    """Leave encashment on retirement u/s 10(10AA)."""
    received = _g(d, "received", label="Leave encashment received")
    if received <= 0:
        return {"exempt": 0, "taxable": 0, "workings": []}
    govt = bool(d.get("govt"))
    if govt:
        return {"exempt": received, "taxable": 0,
                "workings": [(f"Leave encashment (Government employee): fully exempt", received)],
                "limit": "fully exempt (Government)"}
    salary = _g(d, "avg_10m_salary", label="Average salary (basic+DA) of last 10 months")
    years = _g(d, "years", label="Years of service")
    months = _g(d, "leave_months_standing", label="Earned leave standing (months)")
    months = min(months, years) if years else months            # 30 days per year of service
    leave_salary = salary * months
    ten_month_salary = salary * 10
    exempt = min(received, leave_salary, ten_month_salary, LEAVE_ENCASH_LIMIT)
    touching = [
        ("(i) Amount actually received", received),
        (f"(ii) Leave salary: {months} mo. × Rs. {salary:,} (max 30 days/yr of service)",
         leave_salary),
        (f"(iii) 10 months' average salary: 10 × Rs. {salary:,}", ten_month_salary),
        (f"(iv) Statutory ceiling u/s 10(10AA)", LEAVE_ENCASH_LIMIT),
        ("Exempt = least of (i)–(iv)", exempt),
    ]
    return {"exempt": exempt, "taxable": received - exempt, "workings": touching}


def commuted_pension_exemption(d: dict) -> dict:
    """Commuted pension u/s 10(10A)."""
    received = _g(d, "received", label="Commuted value of pension received")
    if received <= 0:
        return {"exempt": 0, "taxable": 0, "workings": []}
    govt = bool(d.get("govt"))
    if govt:
        return {"exempt": received, "taxable": 0,
                "workings": [(f"Commuted pension (Government): fully exempt", received)],
                "limit": "fully exempt (Government)"}
    if bool(d.get("received_gratuity")):
        exempt = received // 3
        note = "1/3rd of commuted pension (gratuity also received)"
    else:
        exempt = received // 2
        note = "1/2 of commuted pension (no gratuity received)"
    return {"exempt": exempt, "taxable": received - exempt,
            "workings": [(f"{note}: Rs. {received:,} → exempt Rs. {exempt:,}", exempt)]}


def vrs_exemption(d: dict) -> dict:
    """Voluntary retirement compensation u/s 10(10C) — least of three."""
    received = _g(d, "received", label="VRS compensation received")
    if received <= 0:
        return {"exempt": 0, "taxable": 0, "workings": []}
    ms = _g(d, "months_salary", label="Monthly salary (last drawn)")
    yrs = _g(d, "years_completed", label="Completed years of service")
    mrem = _g(d, "months_remaining", label="Months of service left before retirement")
    three_mo = 3 * ms * yrs
    rem = ms * mrem
    exempt = min(received, VRS_LIMIT, three_mo, rem)
    touching = [
        ("(i) Compensation actually received", received),
        (f"(ii) 3 months' salary × {yrs} yr(s) = 3 × Rs. {ms:,} × {yrs}", three_mo),
        (f"(iii) Salary × months of service remaining = Rs. {ms:,} × {mrem}", rem),
        (f"(iv) Statutory ceiling u/s 10(10C)", VRS_LIMIT),
        ("Exempt = least of (i)–(iv)", exempt),
    ]
    return {"exempt": exempt, "taxable": received - exempt, "workings": touching}


def retrenchment_exemption(d: dict) -> dict:
    """Retrenchment compensation u/s 10(10B)."""
    received = _g(d, "received", label="Retrenchment compensation received")
    if received <= 0:
        return {"exempt": 0, "taxable": 0, "workings": []}
    if bool(d.get("approved_scheme")):
        return {"exempt": received, "taxable": 0,
                "workings": [("Retrenchment compensation under approved scheme — fully exempt",
                              received)],
                "limit": "fully exempt (approved scheme)", "section": "10(10B)(ii)"}
    salary = _g(d, "avg_3m_salary", label="Average salary of last 3 months")
    years = _years_completed(_g(d, "years", label="Years of service"),
                             _g(d, "months_extra", label="Extra months (>6 counts as 1 yr)"))
    formula = salary * 15 * years // 26 if years else 0
    exempt = min(received, RETRENCH_LIMIT, formula)
    touching = [
        ("(i) Compensation actually received", received),
        (f"(ii) 15/26 × Rs. {salary:,} × {years} yr(s)", formula),
        ("(iii) Statutory ceiling u/s 10(10B)(i)", RETRENCH_LIMIT),
        ("Exempt = least of (i)–(iii)", exempt),
    ]
    return {"exempt": exempt, "taxable": received - exempt, "workings": touching,
            "section": "10(10B)(i)"}


# ---------------------------------------------------------------------------
# Monthly payroll grid
# ---------------------------------------------------------------------------

def aggregate_monthly(rows: list[dict]) -> dict:
    """Sum the 12-month payroll grid into annual figures."""
    out = {"basic": 0, "da": 0, "hra_received": 0, "rent_paid": 0,
           "other_taxable": 0, "other_exempt": 0, "professional_tax": 0, "tds": 0}
    norm = []
    for r in rows or []:
        r = dict(r or {})
        if "hra_received" not in r and "hra" in r:
            r["hra_received"] = r["hra"]
        if "pt" not in r and "professional_tax" in r:
            r["pt"] = r["professional_tax"]
        norm.append(r)
    for r in norm:
        with field_scope(f"Monthly salary {r.get('month', '?')}"):
            out["basic"] += _g(r, "basic")
            out["da"] += _g(r, "da")
            out["hra_received"] += _g(r, "hra_received")
            out["rent_paid"] += _g(r, "rent_paid")
            out["other_taxable"] += _g(r, "other_taxable")
            out["other_exempt"] += _g(r, "other_exempt")
            out["professional_tax"] += _g(r, "pt")
            out["tds"] += _g(r, "tds")
    out["months_filled"] = sum(1 for r in norm
                               if any(_g(r, k) for k in ("basic", "da", "hra_received",
                                                         "other_taxable", "other_exempt")))
    return out


# ---------------------------------------------------------------------------
# Arrears — Form 10E relief u/s 89
# ---------------------------------------------------------------------------

def relief_89_working(total_income_current: int, arrear_rows: list[dict],
                      *, regime: str, dob_year: int | None) -> dict:
    """Annexure-I (Form 10E) computation for salary arrears/advance.

    ``arrear_rows`` = [{fy: '2023-24', amount: X, income_that_fy: Y}]
    Returns relief amount, per-row workings and the two current-year tax figures.
    Tax for prior years uses statutory parameters of those years (history module).
    """
    rows = []
    total_arrears = 0
    sum_prior_deltas = 0.0
    for r in arrear_rows or []:
        fy = str(r.get("fy", "")).strip()
        amt = _g(r, "amount", label=f"Arrear amount (FY {fy})")
        inc = _g(r, "income_that_fy", label=f"Total income of FY {fy}")
        if amt <= 0:
            continue
        if fy not in H.YEARS:
            rows.append({"fy": fy, "amount": amt, "income_that_fy": inc,
                         "error": f"No tax parameters for FY {fy} (2017-18 to 2025-26 only)"})
            continue
        use_reg = "new" if regime == "new" else "old"
        # figures are total *taxable* income → no further standard deduction
        tax_no = H.tax_for_year(fy, inc, dob_year=dob_year, regime=use_reg, salaried=False)
        tax_with = H.tax_for_year(fy, inc + amt, dob_year=dob_year, regime=use_reg,
                                  salaried=False)
        delta = tax_with - tax_no
        sum_prior_deltas += delta
        total_arrears += amt
        rows.append({"fy": fy, "amount": amt, "income_that_fy": inc,
                     "tax_without": tax_no, "tax_with": tax_with, "delta": delta})
    reg = "new" if regime == "new" else "old"
    cur_no = H.tax_for_year("2025-26", total_income_current - total_arrears,
                            dob_year=dob_year, regime=reg, salaried=False)
    cur_with = H.tax_for_year("2025-26", total_income_current,
                              dob_year=dob_year, regime=reg, salaried=False)
    delta_current = cur_with - cur_no
    relief = max(0.0, delta_current - sum_prior_deltas)
    return {
        "rows": rows,
        "total_arrears": total_arrears,
        "current_tax_with_arrears": cur_with,
        "current_tax_without_arrears": cur_no,
        "current_delta": delta_current,
        "sum_prior_deltas": sum_prior_deltas,
        "relief89": int(relief + 0.5),
    }


# ---------------------------------------------------------------------------
# Master expansion
# ---------------------------------------------------------------------------

def expand_salary(data: dict) -> dict:
    """Fold every salary input (slip components, perquisites, profits, terminal
    benefits, monthly grid, arrears) into the classic totals + workings.

    Returns a dict with:
      totals:        salary_17_1 / perquisites_17_2 / profits_17_3 / professional_tax /
                     entertainment_allowance
      exempt_rows:   [{section, amount, note}]   (incl. terminal-benefit exemptions)
      hra:           {hra_received, rent_paid, basic_salary, da}  (auto from components
                     / monthly when the manual HRA block is blank)
      arrears_total, terminal_workings, component_breakdown, employer,
      monthly_totals, warnings
    """
    sal = data.get("salary", {}) or {}
    out = {"warnings": [], "component_breakdown": [], "terminal_workings": [],
           "exempt_rows": [], "terminal_taxable": 0}

    # ---- employer (used by TDS-1 sync + ITR-2 Schedule S) ----
    emp = sal.get("employer") or {}
    out["employer"] = {"name": str(emp.get("name", "") or ""), "tan": str(emp.get("tan", "") or ""),
                       "category": str(emp.get("category", "") or (data.get("personal", {}) or {})
                                     .get("employer_category", "") or "")}

    # ---- monthly grid (auto-fills annuals when components not used) ----
    monthly = aggregate_monthly(sal.get("monthly"))
    out["monthly_totals"] = monthly
    using_monthly = monthly["months_filled"] > 0

    # ---- slip components ----
    comp_total = 0
    comp_basic = comp_da = 0
    comp_exempt_rows = []
    for c in sal.get("components") or []:
        code = str(c.get("code", "OTHR"))
        label = c.get("label") or COMPONENTS.get(code, ("Other allowance",))[0]
        with field_scope(f"Salary component: {label}"):
            amt = _g(c, "amount")
        if amt <= 0:
            continue
        comp_total += amt
        section = COMPONENTS.get(code, (None, None, None))[1]
        cap = COMPONENTS.get(code, (None, None, None))[2]
        if code == "BASIC":
            comp_basic += amt
        elif code == "DA":
            comp_da += amt
        ex_amt = 0
        if section == "10(13A)":
            pass  # exemption computed by the HRA schedule, not here
        elif section:
            default_ex = amt if cap is None else min(amt, cap)
            with field_scope(f"Exempt part of {label}"):
                ex_amt = _g(c, "exempt", default_ex)
            if cap is not None and ex_amt > cap:
                ex_amt = cap
                out["warnings"].append(f"Exemption for {label} capped at Rs. {cap:,}.")
            if ex_amt > amt:
                ex_amt = amt
            if ex_amt > 0:
                comp_exempt_rows.append({"section": section, "amount": ex_amt,
                                         "note": label})
        row = {"code": code, "label": label, "amount": amt,
               "exempt": ex_amt, "section": section}
        if code == "ARREARS":
            row["arrear_fy"] = str(c.get("arrear_fy", "") or "").strip()
            row["income_that_fy"] = _g(c, "income_that_fy")
        out["component_breakdown"].append(row)

    # ---- merge with legacy/manual entry fields ----
    with field_scope("Salary income", "salary"):
        manual_171 = 0 if (using_monthly or comp_total) else _g(sal, "salary_17_1")
        manual_172 = 0 if sal.get("perquisites") else _g(sal, "perquisites_17_2")
        manual_173 = _g(sal, "profits_17_3")
        manual_pt = 0 if using_monthly else _g(sal, "professional_tax")
        manual_ent = _g(sal, "entertainment_allowance")
        std_override = _g(sal, "standard_deduction_override", None)

    monthly_171 = (monthly["basic"] + monthly["da"] + monthly["hra_received"] +
                   monthly["other_taxable"] + monthly["other_exempt"]) if using_monthly else 0

    out["salary_17_1"] = manual_171 + comp_total + monthly_171
    out["entertainment_allowance"] = manual_ent
    out["professional_tax"] = manual_pt + (monthly["professional_tax"] if using_monthly else 0)
    out["standard_deduction_override"] = std_override

    # ---- perquisites ----
    perq_rows = []
    perq_total = 0
    for p in sal.get("perquisites") or []:
        ptype = str(p.get("type", "OTHR"))
        plabel = p.get("description") or PERQUISITE_TYPES.get(ptype, "Perquisite")
        with field_scope(f"Perquisite: {plabel}"):
            amt = _g(p, "amount")
        if amt <= 0:
            continue
        perq_total += amt
        perq_rows.append({"type": ptype, "label": plabel, "amount": amt})
    out["perquisites_rows"] = perq_rows
    out["perquisites_17_2"] = perq_total + manual_172

    # ---- profits in lieu of salary (detail) ----
    profit_rows = []
    for p in sal.get("profits_detail") or []:
        label = p.get("label", "Other")
        with field_scope(f"Profits in lieu of salary: {label}"):
            amt = _g(p, "amount")
        if amt > 0:
            profit_rows.append({"label": label, "amount": amt})
    out["profits_rows"] = profit_rows
    out["profits_17_3"] = manual_173 + sum(r["amount"] for r in profit_rows)

    # ---- terminal benefits ----
    term = sal.get("terminal") or {}
    calculators = [
        ("gratuity", "Gratuity u/s 10(10)", gratuity_exemption, "10(10)"),
        ("leave_encashment", "Leave encashment u/s 10(10AA)", leave_encashment_exemption,
         "10(10AA)"),
        ("commuted_pension", "Commuted pension u/s 10(10A)", commuted_pension_exemption,
         "10(10A)"),
        ("vrs", "VRS compensation u/s 10(10C)", vrs_exemption, "10(10C)"),
        ("retrenchment", "Retrenchment compensation u/s 10(10B)", retrenchment_exemption,
         "10(10B)(i)"),
    ]
    for key, name, fn, section in calculators:
        d = term.get(key) or {}
        res = fn(d)
        if res["exempt"] <= 0 and res["taxable"] <= 0:
            continue
        sec = res.get("section", section)
        if res["exempt"] > 0:
            out["terminal_workings"].append(
                {"name": name, "section": sec, "received": res["exempt"] + res["taxable"],
                 "exempt": res["exempt"], "taxable": res["taxable"],
                 "limit": res.get("limit", ""), "lines": res["workings"]})
            out["exempt_rows"].append({"section": sec, "amount": res["exempt"],
                                       "note": name})
        if res["taxable"] > 0:
            out["terminal_taxable"] += res["taxable"]
            out["terminal_workings"].append(
                {"name": f"{name} — taxable part", "section": "", "received": 0,
                 "exempt": 0, "taxable": res["taxable"], "limit": "",
                 "lines": [(f"Taxable part added to salary u/s 17(1)", res["taxable"])]})
    out["salary_17_1"] += out["terminal_taxable"]

    # ---- arrears (they stay inside 17(1); relief-89 handles the rate shift) ----
    arrears_rows = []
    covered_fys = set()
    for r in sal.get("arrears") or []:
        fy = str(r.get("fy", "")).strip()
        with field_scope(f"Arrear salary (FY {fy})"):
            amt = _g(r, "amount")
            inc = _g(r, "income_that_fy")
        if amt > 0:
            arrears_rows.append({"fy": fy, "amount": amt, "income_that_fy": inc})
            covered_fys.add(fy)
    # payroll-slip style: an ARREARS component tagged with the earlier FY is
    # promoted into the Form-10E working (relief u/s 89 is then automatic)
    for comp in out["component_breakdown"]:
        if comp["code"] != "ARREARS":
            continue
        fy = comp.get("arrear_fy") or ""
        if not fy or fy in covered_fys:
            if not fy:
                out["warnings"].append(
                    f"Arrear salary of Rs. {comp['amount']:,}: tag the year it "
                    "belongs to (e.g. 2023-24) to get relief u/s 89 automatically.")
            continue
        covered_fys.add(fy)
        if comp.get("income_that_fy"):
            arrears_rows.append({"fy": fy, "amount": comp["amount"],
                                 "income_that_fy": comp["income_that_fy"]})
        else:
            out["warnings"].append(
                f"Arrear salary of Rs. {comp['amount']:,} for FY {fy}: also enter the "
                "total income of FY " + fy + " — relief u/s 89 is skipped until that "
                "is provided (the arrear itself stays fully taxable this year).")
    out["arrears_rows"] = arrears_rows
    out["arrears_total"] = sum(r["amount"] for r in arrears_rows)

    # ---- HRA inputs (auto from components/monthly if not explicitly set) ----
    hra = dict(sal.get("hra") or {})
    hra_rec_comp = sum(r["amount"] for r in out["component_breakdown"] if r["code"] == "HRA")
    if not _g(hra, "hra_received"):
        hra["hra_received"] = hra_rec_comp + (monthly["hra_received"] if using_monthly else 0)
    if not _g(hra, "rent_paid") and using_monthly:
        hra["rent_paid"] = monthly["rent_paid"]
    if not _g(hra, "basic_salary"):
        hra["basic_salary"] = comp_basic + (monthly["basic"] if using_monthly else 0)
    if not _g(hra, "da"):
        hra["da"] = comp_da + (monthly["da"] if using_monthly else 0)
    if comp_exempt_rows and not hra.get("_has_exempt"):
        pass
    out["hra"] = hra if any(_g(hra, k) for k in
                            ("hra_received", "rent_paid", "basic_salary", "da")) else None

    # ---- exempt rows: components + manual legacy rows + terminal ----
    legacy_rows = []
    for row in sal.get("exempt_allowances", []) or []:
        with field_scope("Exempt allowance (manual entry)"):
            amt = _g(row, "amount")
        if amt > 0:
            legacy_rows.append({"section": str(row.get("section", "")), "amount": amt,
                                "note": "manual"})
    out["exempt_rows"].extend(comp_exempt_rows)
    out["exempt_rows"].extend(legacy_rows)

    # TDS from the monthly grid (used as a convenience suggestion)
    out["tds_monthly"] = monthly["tds"] if using_monthly else 0

    # validations for the UI
    if using_monthly and monthly["months_filled"] < 12:
        out["warnings"].append(
            f"Monthly payroll: only {monthly['months_filled']} of 12 months filled — "
            "totals use the months entered.")
    return out
