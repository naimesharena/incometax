/* Open ITR Suite - ITR-1 / ITR-2 / ITR-4 wizard application (vanilla JS, no deps). */
"use strict";

/* ------------------------------------------------------------------ state */
const blankState = () => ({
  form: "auto",          // "auto" | "ITR1" | "ITR2" | "ITR4"  (ITR-3 read-only import)
  personal: {
    first_name: "", middle_name: "", last_name: "", pan: "", dob: "", aadhaar: "",
    employer_category: "OTH", status: "I", residential_status: "RES",
    address: {
      residence_no: "", residence_name: "", road_or_street: "", locality_or_area: "",
      city_or_town_or_district: "", state_code: "", pin_code: "", country_code: "91",
      mobile_no: "", email_address: "", country_code_mobile: "91", mobile_no_sec: "",
      email_address_sec: "",
    },
    secondary_address_same: true,
    alternate_address: {},
  },
  filing: {
    return_section: 11, opt_out_new_regime: false, filing_date: todayISO(),
    is_representative: false, representative: {},
    seventh_provisio: false, foreign_travel_exceeds_2l: false,
    electricity_exceeds_1l: false, deposit_exceeds_1cr: false,
    original_receipt_no: "", original_filing_date: "", notice_no: "", notice_date: "",
  },
  salary: {
    salary_17_1: "", perquisites_17_2: "", profits_17_3: "",
    exempt_allowances: [], hra: null,
    entertainment_allowance: "", professional_tax: "", standard_deduction_override: "",
    employer: {}, components: [], perquisites: [], profits_detail: [],
    terminal: {}, arrears: [], monthly: [],
  },
  house_properties: [],
  other_sources: [],
  other_sources_special: {},
  ltcg112a: { sale_consideration: "", cost_of_acquisition: "" },
  capital_gains: { equity_st: [], equity_lt: [], land_building: [], other_st: [],
                   other_lt: [], exemptions: [] },
  presumptive: { "44ad": [], "44ada": [], "44ae": [], business_income_manual: "",
                 financials: {} },
  exempt_income: [],
  deductions: {},
  tax_paid: { tds_on_salary: [], tds_others: [], tds3: [], tcs: [], challans: [] },
  relief_89: "",
  losses_cf: { stcg_cf: "", ltcg_cf: "", hp_loss_cf: "", business_loss_cf: "" },
  bank_accounts: [{ ifsc: "", bank_name: "", account_no: "", account_type: "SB", use_for_refund: true }],
  verification: { name: "", father_name: "", pan: "", place: "", capacity: "S" },
  interest_override: {},
});

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

let state = blankState();
let meta = null;
let currentStep = 0;
let lastSummary = null;

/* ------------------------------------------------------------------ utils */
const $ = (sel, root = document) => root.querySelector(sel);
const el = html => { const t = document.createElement("template"); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const inr = n => "₹ " + Number(Math.round(Number(n) || 0)).toLocaleString("en-IN");

function getPath(obj, path) {
  return path.split(".").reduce((o, k) => (o == null ? o : o[k]), obj);
}
function setPath(obj, path, value) {
  const keys = path.split(".");
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    if (cur[keys[i]] == null) cur[keys[i]] = /^\d+$/.test(keys[i + 1]) ? [] : {};
    cur = cur[keys[i]];
  }
  cur[keys[keys.length - 1]] = value;
}

/* ------------------------------------------------------------------ meta */
async function loadMeta() {
  meta = await (await fetch("/api/meta")).json();
}

const opt = (v, l, sel) => `<option value="${esc(v)}" ${String(v) === String(sel) ? "selected" : ""}>${esc(l)}</option>`;
const optsFrom = (map, sel, placeholder) =>
  (placeholder ? `<option value="">${placeholder}</option>` : "") +
  Object.entries(map).map(([v, l]) => opt(v, l, sel)).join("");

const TDS_SECTIONS = {
  "94A": "194A - Interest other than on securities", "193": "193 - Interest on securities",
  "194": "194 - Dividend", "4DA": "194DA - Life insurance policy payment",
  "4EE": "194EE - NSS deposits", "4F": "194F - Repurchase of MF/UTI units",
  "4G": "194G - Commission on lottery tickets", "4H": "194H - Commission or brokerage",
  "4-IA": "194I(a) - Rent: plant & machinery", "4-IB": "194I(b) - Rent: other",
  "4IB": "194IB - Rent by certain individuals", "94K": "194K - Mutual fund income",
  "94N": "194N - Cash payments", "94O": "194O - E-commerce", "94M": "194M - Certain payments",
  "94P": "194P - Senior citizen", "94J-A": "194J(a) - Technical services",
  "94J-B": "194J(b) - Professional services",
};
const FAMILY_PENSION_NOTE = "Family pension fetches an automatic deduction u/s 57(iia): lower of ⅓ or ₹15,000 (old) / ₹25,000 (new).";
const EXEMPT_CATEGORIES = {
  AGRI: "Agricultural & related incomes",
  GOVC: "Compensation/sums from Government entities",
  ISI: "Income from specified investments",
  SSRA: "Sums received by armed forces personnel",
  SRSC: "Sums received by senior citizens/minors",
  SRST: "Sums received by specified taxpayers",
  SRPC: "LIC/NPS/PF/SSY receipts",
  OTH: "Other incomes",
};
const EXEMPT_SUBCATEGORIES = [
  "10(1)", "10(30)", "10(31)", "10(10BB)", "10(10BC)", "10(17A)", "10(12AB)", "10(15)",
  "10(23FBB)", "10(23FD)", "10(35)", "10(35A)", "10(12C)", "10(18)", "10(19)", "10(23AA)",
  "DMD", "10(32)", "10(43)", "10(19A)", "10(26)", "10(26AAA)", "10(10D)", "10(11)",
  "10(11A)", "10(12)", "10(12A)", "10(12AA)", "10(12B)", "10(12BA)", "10(13)", "10(25)",
  "10(44)", "10(2)", "10(16)", "Incmexmptcircular", "Incmexmptnotification", "Receiptnotincme",
];
const DDB_DISEASES = { a: "Dementia", b: "Dystonia Musculorum Deformans", c: "Motor Neuron Disease", d: "Ataxia", e: "Chorea", f: "Hemiballismus", g: "Aphasia", h: "Parkinsons Disease", i: "Malignant Cancers", j: "AIDS", k: "Chronic Renal failure", l: "Hemophilia", m: "Thalassaemia", n: "Other" };

/* ------------------------------------------------------- field factories */
function money(path, label, sub = "", attrs = "") {
  return `<div class="field"><label>${esc(label)}</label>
    <input type="number" min="0" step="1" data-path="${path}" data-kind="int" value="${esc(getPath(state, path) ?? "")}" placeholder="0" ${attrs}>
    ${sub ? `<span class="sub">${esc(sub)}</span>` : ""}</div>`;
}
function moneyf(path, label, sub = "") {
  return `<div class="field"><label>${esc(label)}</label>
    <input type="number" min="0" step="0.5" data-path="${path}" data-kind="float" value="${esc(getPath(state, path) ?? "")}" placeholder="0">
    ${sub ? `<span class="sub">${esc(sub)}</span>` : ""}</div>`;
}
function txt(path, label, opts = {}) {
  const v = getPath(state, path) ?? "";
  return `<div class="field"><label>${esc(label)}${opts.req ? ' <span class="req">*</span>' : ""}</label>
    <input type="${opts.type || "text"}" data-path="${path}" value="${esc(v)}" placeholder="${esc(opts.ph || "")}" ${opts.max ? `maxlength="${opts.max}"` : ""}>
    ${opts.sub ? `<span class="sub">${esc(opts.sub)}</span>` : ""}</div>`;
}
function sel(path, label, optionsHtml, sub = "", req = false) {
  return `<div class="field"><label>${esc(label)}${req ? ' <span class="req">*</span>' : ""}</label>
    <select data-path="${path}">${optionsHtml}</select>
    ${sub ? `<span class="sub">${esc(sub)}</span>` : ""}</div>`;
}
function dateInp(path, label, sub = "", req = false) {
  return txt(path, label, { type: "date", sub, req });
}
function check(path, label) {
  const v = getPath(state, path);
  return `<label class="inline" style="gap:.4rem;font-size:.85rem"><input type="checkbox" data-path="${path}" data-kind="bool" ${v ? "checked" : ""}> ${esc(label)}</label>`;
}
function radioPair(path, label) {
  const v = getPath(state, path);
  return `<div class="field"><label>${esc(label)}</label><div class="inline">
    <label><input type="radio" name="${path}" data-path="${path}" data-kind="radio-bool" value="true" ${v === true ? "checked" : ""}> Yes</label>
    <label><input type="radio" name="${path}" data-path="${path}" data-kind="radio-bool" value="false" ${v === false ? "checked" : ""}> No</label>
  </div></div>`;
}

function section(title, inner, hint = "") {
  return `<section class="card"><h2>${esc(title)}</h2>${hint ? `<p class="hint">${hint}</p>` : ""}${inner}</section>`;
}

/* --------------------------------------------------------- repeater util */
function repeater(key, addLabel, itemHtml, options = {}) {
  const rows = (getPath(state, key) || []).map((it, i) => `
    <div class="rep-item" data-rep="${key}" data-idx="${i}">
      <div class="rep-head"><span class="rep-title">${options.itemTitle ? options.itemTitle(it, i) : `#${i + 1}`}</span>
      <button type="button" class="rep-del" data-del-rep="${key}" data-idx="${i}">Remove</button></div>
      ${itemHtml(it, i)}
    </div>`).join("");
  return `<div class="repeater" id="rep-${key}">${rows}
    <button type="button" class="btn-add" data-add-rep="${key}">+ ${esc(addLabel)}</button></div>`;
}

function addRepeaterItem(key, factory) {
  const arr = getPath(state, key) || [];
  arr.push(factory());
  setPath(state, key, arr);
  render();
}
function delRepeaterItem(key, idx) {
  const arr = getPath(state, key) || [];
  arr.splice(idx, 1);
  render();
}

/* ==================================================================== TABS */
const TABS = [
  { id: "start", title: "Start & ITR form", render: renderStart },
  { id: "personal", title: "Personal & filing", render: renderPersonal },
  { id: "salary", title: "Salary income", render: renderSalary },
  { id: "hp", title: "House property", render: renderHP },
  { id: "other", title: "Other income", render: renderOther },
  { id: "capgains", title: "Capital gains", render: renderCapGains },
  { id: "presumptive", title: "Business (44AD/AE)", render: renderPresumptive },
  { id: "deductions", title: "Deductions (VI-A)", render: renderDeductions },
  { id: "taxpaid", title: "Taxes paid", render: renderTaxPaid },
  { id: "bank", title: "Bank & verification", render: renderBank },
  { id: "imports", title: "Imports & AIS", render: renderImports },
  { id: "review", title: "Review & file", render: renderReview },
];

/* Feature catalog from the server: salary components, perquisite types,
 * business codes (44AD/44ADA/44AE). */
let CATALOG = { components: {}, perquisites: {}, business_codes: { "44ad": {}, "44ada": {}, "44ae": {} }, terminal: {} };
async function loadCatalog() {
  try {
    const c = await (await fetch("/api/salary/catalog")).json();
    if (c && c.components) CATALOG = c;
  } catch (e) { /* offline demo fallback */ }
}

/* --------------------------------------------------------------- tab: start */
const ITR_FORMS = [
  { id: "auto", title: "✨ Auto-detect (recommended)",
    note: "fill your income details — the software picks the correct form and explains why" },
  { id: "ITR1", title: "ITR-1 (Sahaj)",
    note: "resident individual · income ≤ ₹50L · salary/pension, up to 2 house properties, interest/dividend/family pension, LTCG 112A up to ₹1.25L, agricultural income ≤ ₹5,000" },
  { id: "ITR2", title: "ITR-2",
    note: "any capital gains (equity 111A/112A with grandfathering, property, others), multiple house properties, foreign income/assets, RNOR — anything except business income" },
  { id: "ITR4", title: "ITR-4 (Sugam)",
    note: "presumptive business/profession u/s 44AD / 44ADA / 44AE · total income ≤ ₹50L · turnover ≤ ₹3 crore (business) / ₹75 lakh (profession)" },
  { id: "ITR3", title: "ITR-3",
    note: "business/profession with books of account (balance sheet + P&L) — not generatable here; use the department utility", disabled: true },
];

function renderStart(host) {
  host.appendChild(el(section("Pick the ITR form for AY 2026-27", `
    <p class="hint">Guided preparation of your Indian income-tax return for previous year
    2025-26. The suite computes tax under <b>both</b> regimes, recommends the cheaper one,
    validates against the official CBDT schema &amp; validation rules of the chosen form, and
    produces the JSON you upload on the e-Filing portal.</p>
    <div class="form-cards" role="radiogroup" aria-label="ITR form">
      ${ITR_FORMS.map(f => `
        <label class="form-card ${state.form === f.id ? "selected" : ""} ${f.disabled ? "disabled" : ""}" data-form="${f.id}">
          <input type="radio" name="itr-form" value="${f.id}" ${state.form === f.id ? "checked" : ""}
                 ${f.disabled ? "disabled" : ""} data-change-form>
          <span class="fc-title">${f.title}</span>
          <span class="badge fc-elig" data-elig="${f.id}"></span>
          <span class="fc-note">${f.note}</span>
        </label>`).join("")}
    </div>
    <div id="form-detect" class="notice"><span class="muted">Analysing your income details to recommend a form…</span></div>`)));

  const facts = section("Key facts for AY 2026-27", `
    <div class="grid c2">
      <div>
        <p class="hint"><b>New regime (default)</b> — slabs: nil up to ₹4L, 5% to ₹8L, 10% to ₹12L,
        15% to ₹16L, 20% to ₹20L, 25% to ₹24L, 30% above. Standard deduction ₹75,000.
        Rebate u/s 87A up to ₹60,000 → <b>zero tax up to ₹12 lakh</b> of taxable income
        (₹12.75L for salaried).</p>
      </div>
      <div>
        <p class="hint"><b>Old regime</b> — keeps Chapter VI-A deductions (80C/80D/80G...), HRA, and
        home-loan interest on self-occupied property. Standard deduction ₹50,000; rebate ₹12,500 up to
        ₹5 lakh income. Only available if filed within the due date.</p>
      </div>
    </div>
    <p class="hint">Due date of filing (u/s 139(1)): <b>${meta?.due_date || "2026-07-31"}</b> ·
    ITR-4 (Sugam) due date convenience: 31-08-2026.</p>`);
  host.appendChild(el(facts));
  refreshFormDetect();
}

async function refreshFormDetect() {
  const box = $("#form-detect");
  if (!box) return;
  const d = await postJSON("/api/forms/detect", state);
  if (!box.isConnected) return;
  if (d.error) { box.innerHTML = `<span class="muted">${esc(d.error)}</span>`; return; }
  lastDetect = d;
  const elig = d.eligible || {};
  for (const [fid, info] of Object.entries(elig)) {
    const badge = $(`[data-elig="${fid}"]`);
    if (!badge) continue;
    badge.className = `badge fc-elig ${info.ok ? "green" : "red"}`;
    badge.textContent = info.ok ? "✓ eligible" : "✗ " + (info.reasons[0] || "not eligible");
    badge.title = (info.reasons || []).join(" • ");
  }
  const rec = d.recommended;
  if (rec === "ITR3") {
    box.innerHTML = `<b>Recommendation: ITR-3</b> — ${esc(d.why)}. This suite prepares ITR-1/2/4;
      for ITR-3 (books of account) please use the department utility.`;
    box.className = "notice warn";
  } else if (state.form === "auto") {
    box.innerHTML = `<b>Recommended: ${rec}</b> — ${esc(d.why)}. <span class="muted">Change the
      selection above only if you disagree; Review &amp; file will follow the auto-detected form.</span>`;
    box.className = "notice ok";
  } else {
    box.innerHTML = `You chose <b>${state.form}</b>. Auto-analysis suggests <b>${rec}</b> — ${esc(d.why)}.`;
    box.className = state.form === rec ? "notice ok" : "notice warn";
  }
}
let lastDetect = null;

function eligible() { return true; }

/* ============================================ mandatory-field validation
 * Every wizard step maps its *mandatory* fields (per the official ITR-1
 * schema + CBDT rules).  "Next →", step jumps forward, and Finish/Download
 * all run these checks first and name exactly which field is missing/invalid.
 * ===================================================================== */
const RE_PAN = /^[A-Z]{5}[0-9]{4}[A-Z]$/;
const RE_TAN = /^[A-Z]{4}[0-9]{5}[A-Z]$/;
const RE_IFSC = /^[A-Z]{4}0[A-Z0-9]{6}$/;
const RE_EMAIL = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
const RE_MOBILE = /^[1-9][0-9]{9}$/;
const RE_PIN = /^[1-9][0-9]{5}$/;
const RE_AADHAAR = /^[0-9]{12}$/;
const RE_DATE = /^\d{4}-\d{2}-\d{2}$/;
const RE_BSR = /^[0-9]{7}$/;
const RE_CHALLAN_NO = /^[0-9]{1,5}$/;

const blank = v => v === "" || v === null || v === undefined;
const gt0 = v => !blank(v) && Number(v) > 0;

function req(errs, path, label) {
  if (blank(getPath(state, path))) errs.push({ path, msg: `${label} is required` });
}
function fmt(errs, path, label, re, hint) {
  const v = getPath(state, path);
  if (!blank(v) && !re.test(String(v).trim()))
    errs.push({ path, msg: `${label}: '${v}' is invalid — ${hint}` });
}

/* ------------------------------------------------ per-step check functions */
function checkPersonal(errs) {
  const p = state.personal, a = p.address || {};
  if (blank(p.first_name) && blank(p.last_name))
    errs.push({ path: "personal.last_name", msg: "Name (first or last name as per PAN) is required" });
  req(errs, "personal.pan", "PAN");
  fmt(errs, "personal.pan", "PAN", RE_PAN, "format ABCDE1234F (5 letters, 4 digits, 1 letter)");
  req(errs, "personal.dob", "Date of birth");
  if (!blank(p.dob)) {
    if (!RE_DATE.test(p.dob)) errs.push({ path: "personal.dob", msg: `Date of birth '${p.dob}' is invalid — use YYYY-MM-DD` });
    else if (p.dob < "1900-01-01") errs.push({ path: "personal.dob", msg: "Date of birth cannot be before 1900" });
    else if (p.dob > "2026-04-01") errs.push({ path: "personal.dob", msg: "Date of birth cannot be after 01-04-2026 (person not born in the previous year)" });
  }
  req(errs, "personal.employer_category", "Nature of employment");
  fmt(errs, "personal.aadhaar", "Aadhaar number", RE_AADHAAR, "12 digits");
  req(errs, "personal.address.residence_no", "Flat/Door/Block no.");
  req(errs, "personal.address.locality_or_area", "Locality/Area");
  req(errs, "personal.address.city_or_town_or_district", "City/Town/District");
  req(errs, "personal.address.state_code", "State");
  req(errs, "personal.address.pin_code", "PIN code");
  fmt(errs, "personal.address.pin_code", "PIN code", RE_PIN, "6 digits, not starting with 0");
  req(errs, "personal.address.mobile_no", "Mobile no.");
  fmt(errs, "personal.address.mobile_no", "Mobile no.", RE_MOBILE, "10 digits, not starting with 0");
  req(errs, "personal.address.email_address", "Email address");
  fmt(errs, "personal.address.email_address", "Email address", RE_EMAIL, "format name@example.com");

  if (p.secondary_address_same === false) {
    req(errs, "personal.alternate_address.residence_no", "Alternate address Flat/No");
    req(errs, "personal.alternate_address.locality_or_area", "Alternate address locality");
    req(errs, "personal.alternate_address.city_or_town_or_district", "Alternate address city");
    req(errs, "personal.alternate_address.state_code", "Alternate address state");
    fmt(errs, "personal.alternate_address.pin_code", "Alternate address PIN", RE_PIN, "6 digits");
  }

  const fs = state.filing;
  req(errs, "filing.filing_date", "Date of filing");
  fmt(errs, "filing.filing_date", "Date of filing", RE_DATE, "use YYYY-MM-DD");
  if (Number(fs.return_section) === 17) {           // 139(5) revised
    req(errs, "filing.original_receipt_no", "Acknowledgement no. of original return");
    req(errs, "filing.original_filing_date", "Date of filing of original return");
  }
  if ([13, 14, 16, 18].includes(Number(fs.return_section))) {
    req(errs, "filing.notice_no", "Notice no.");
    req(errs, "filing.notice_date", "Date of notice/order");
  }
  if (fs.is_representative) {
    req(errs, "filing.representative.name", "Representative name");
    fmt(errs, "filing.representative.mobile", "Representative mobile", RE_MOBILE, "10 digits");
    fmt(errs, "filing.representative.email", "Representative email", RE_EMAIL, "format name@example.com");
  }
}

function checkSalary(errs) {
  (state.salary.exempt_allowances || []).forEach((r, i) => {
    if (gt0(r.amount) && blank(r.section))
      errs.push({ path: `salary.exempt_allowances.${i}.section`, msg: `Exempt allowance row ${i + 1}: choose the section` });
  });
  const s = state.salary || {};
  if (!blank(s.employer?.tan))
    fmt(errs, "salary.employer.tan", "Employer TAN", RE_TAN, "format AAAA00000A");
  (s.components || []).forEach((r, i) => {
    if (!gt0(r.amount)) return;
    if (blank(r.code))
      errs.push({ path: `salary.components.${i}.code`, msg: `Earning component ${i + 1}: choose the component` });
    if (r.code === "ARREARS" && blank(r.arrear_fy))
      errs.push({ path: `salary.components.${i}.arrear_fy`, msg: `Arrear component ${i + 1}: enter the year it belongs to (e.g. 2023-24) — otherwise §89 relief is skipped` });
  });
  (s.perquisites || []).forEach((r, i) => {
    if (gt0(r.amount) && blank(r.code))
      errs.push({ path: `salary.perquisites.${i}.code`, msg: `Perquisite row ${i + 1}: choose the perquisite type` });
  });
  (s.arrears || []).forEach((r, i) => {
    if (!gt0(r.amount)) return;
    req(errs, `salary.arrears.${i}.fy`, `Arrear row ${i + 1}: FY it belongs to`);
    if (blank(r.income_that_fy))
      errs.push({ path: `salary.arrears.${i}.income_that_fy`, msg: `Arrear row ${i + 1}: total income of FY ${r.fy || "?"} is needed for §89 relief` });
  });
}

function checkHP(errs) {
  (state.house_properties || []).forEach((p, i) => {
    const base = `house_properties.${i}`;
    req(errs, `${base}.occupancy`, `House property ${i + 1}: type (let out / self occupied)`);
    if ((p.occupancy === "L" || p.occupancy === "D") && !gt0(p.annual_lettable_value))
      errs.push({ path: `${base}.annual_lettable_value`, msg: `House property ${i + 1}: gross rent is required for a let-out property (enter actual rent received/receivable)` });
    fmt(errs, `${base}.pin_code`, `House property ${i + 1} PIN`, RE_PIN, "6 digits");
    if (p.co_owned) {
      const sh = Number(p.assessee_share_pct);
      if (!(sh > 0 && sh <= 100))
        errs.push({ path: `${base}.assessee_share_pct`, msg: `House property ${i + 1}: your share % must be between 1 and 100` });
      (p.co_owners || []).forEach((c, j) => {
        req(errs, `${base}.co_owners.${j}.name`, `Co-owner ${j + 1} name`);
        const cs = Number(c.share_pct);
        if (!(cs > 0 && cs <= 100))
          errs.push({ path: `${base}.co_owners.${j}.share_pct`, msg: `Co-owner ${j + 1}: share % must be between 1 and 100` });
        fmt(errs, `${base}.co_owners.${j}.pan`, `Co-owner ${j + 1} PAN`, RE_PAN, "format ABCDE1234F");
        fmt(errs, `${base}.co_owners.${j}.aadhaar`, `Co-owner ${j + 1} Aadhaar`, RE_AADHAAR, "12 digits");
      });
    }
    (p.tenants || []).forEach((t, j) => {
      fmt(errs, `${base}.tenants.${j}.pan`, `Tenant ${j + 1} PAN`, RE_PAN, "format ABCDE1234F");
      fmt(errs, `${base}.tenants.${j}.aadhaar`, `Tenant ${j + 1} Aadhaar`, RE_AADHAAR, "12 digits");
    });
    (p.loans || []).forEach((l, j) => {
      req(errs, `${base}.loans.${j}.lender_name`, `Loan ${j + 1}: lender name`);
      req(errs, `${base}.loans.${j}.loan_date`, `Loan ${j + 1}: date of loan`);
      fmt(errs, `${base}.loans.${j}.loan_date`, `Loan ${j + 1}: date of loan`, RE_DATE, "YYYY-MM-DD");
    });
  });
}

function checkOther(errs) {
  (state.other_sources || []).forEach((r, i) => {
    req(errs, `other_sources.${i}.nature`, `Other income row ${i + 1}: nature of income`);
    if (blank(r.amount))
      errs.push({ path: `other_sources.${i}.amount`, msg: `Other income row ${i + 1}: amount is required (enter 0 or remove the row)` });
    if (r.nature === "OTH") req(errs, `other_sources.${i}.description`, `Other income row ${i + 1}: description`);
  });
  (state.exempt_income || []).forEach((r, i) => {
    if (blank(r.amount))
      errs.push({ path: `exempt_income.${i}.amount`, msg: `Exempt income row ${i + 1}: amount is required (enter 0 or remove the row)` });
  });
  const l = state.ltcg112a || {};
  if (!blank(l.sale_consideration) !== !blank(l.cost_of_acquisition))
    errs.push({ path: "ltcg112a.cost_of_acquisition", msg: "LTCG 112A: fill both sale consideration and cost of acquisition (or neither)" });
}

function checkDeductions(errs) {
  const d = state.deductions || {};
  if (d["80ddb"] && gt0(d["80ddb"].amount))
    req(errs, "deductions.80ddb.disease", "80DDB: specified disease");
  ["80e", "80ee", "80eea", "80eeb"].forEach(k => {
    if (!dedEnabled(k)) return;
    (d[k].loans || []).forEach((ln, i) => {
      req(errs, `deductions.${k}.loans.${i}.lender_name`, `${k.toUpperCase()} loan ${i + 1}: lender name`);
      fmt(errs, `deductions.${k}.loans.${i}.loan_date`, `${k.toUpperCase()} loan ${i + 1}: date`, RE_DATE, "YYYY-MM-DD");
    });
  });
  const g80 = d["80g"] || {};
  ["don_100", "don_50", "don_100_ql", "don_50_ql"].forEach(bucket => {
    (g80[bucket] || []).forEach((dn, i) => {
      req(errs, `deductions.80g.${bucket}.${i}.name`, `80G donee ${i + 1}: name`);
      if (!gt0(dn.cash) && !gt0(dn.other))
        errs.push({ path: `deductions.80g.${bucket}.${i}.other`, msg: `80G donee '${dn.name || (i + 1)}': enter the donation amount (cash or other mode)` });
      fmt(errs, `deductions.80g.${bucket}.${i}.pan`, `80G donee ${i + 1} PAN`, RE_PAN, "format ABCDE1234F");
    });
  });
  const ggc = d["80ggc"] || {};
  (ggc.donations || []).forEach((dn, i) => {
    req(errs, `deductions.80ggc.donations.${i}.party_name`, `80GGC donation ${i + 1}: party name`);
    req(errs, `deductions.80ggc.donations.${i}.date`, `80GGC donation ${i + 1}: date`);
    if (!gt0(dn.other))
      errs.push({ path: `deductions.80ggc.donations.${i}.other`, msg: `80GGC donation ${i + 1}: amount is required` });
  });
}

function checkTaxPaid(errs) {
  const tp = state.tax_paid || {};
  (tp.tds_on_salary || []).forEach((r, i) => {
    if (!gt0(r.tds)) return;
    req(errs, `tax_paid.tds_on_salary.${i}.tan`, `TDS (salary) row ${i + 1}: TAN of employer`);
    fmt(errs, `tax_paid.tds_on_salary.${i}.tan`, `TDS (salary) row ${i + 1}: TAN`, RE_TAN, "format AAAA00000A");
    req(errs, `tax_paid.tds_on_salary.${i}.employer_name`, `TDS (salary) row ${i + 1}: employer name`);
  });
  (tp.tds_others || []).forEach((r, i) => {
    if (!gt0(r.tds_claimed)) return;
    req(errs, `tax_paid.tds_others.${i}.tan`, `TDS row ${i + 1}: TAN of deductor`);
    fmt(errs, `tax_paid.tds_others.${i}.tan`, `TDS row ${i + 1}: TAN`, RE_TAN, "format AAAA00000A");
    req(errs, `tax_paid.tds_others.${i}.deductor_name`, `TDS row ${i + 1}: deductor name`);
    if (Number(r.tds_claimed) > Number(r.tds_deducted))
      errs.push({ path: `tax_paid.tds_others.${i}.tds_claimed`, msg: `TDS row ${i + 1}: claimed (₹${Number(r.tds_claimed).toLocaleString("en-IN")}) cannot exceed TDS deducted (₹${Number(r.tds_deducted || 0).toLocaleString("en-IN")}) — it will be restricted` });
  });
  (tp.tds3 || []).forEach((r, i) => {
    if (!gt0(r.tds_claimed)) return;
    req(errs, `tax_paid.tds3.${i}.tenant_pan`, `TDS on rent row ${i + 1}: tenant PAN`);
    fmt(errs, `tax_paid.tds3.${i}.tenant_pan`, `TDS on rent row ${i + 1}: tenant PAN`, RE_PAN, "format ABCDE1234F");
  });
  (tp.tcs || []).forEach((r, i) => {
    if (!gt0(r.tcs_claimed)) return;
    req(errs, `tax_paid.tcs.${i}.tan`, `TCS row ${i + 1}: TAN of collector`);
    fmt(errs, `tax_paid.tcs.${i}.tan`, `TCS row ${i + 1}: TAN`, RE_TAN, "format AAAA00000A");
    if (Number(r.tcs_claimed) > Number(r.tcs_collected))
      errs.push({ path: `tax_paid.tcs.${i}.tcs_claimed`, msg: `TCS row ${i + 1}: claimed cannot exceed TCS collected — it will be restricted` });
  });
  (tp.challans || []).forEach((c, i) => {
    const anyFilled = [c.bsr_code, c.date, c.challan_no, c.amount].some(v => !blank(v) && Number(v) !== 0);
    if (!anyFilled) return;
    req(errs, `tax_paid.challans.${i}.bsr_code`, `Challan ${i + 1}: BSR code`);
    fmt(errs, `tax_paid.challans.${i}.bsr_code`, `Challan ${i + 1}: BSR code`, RE_BSR, "7 digits");
    req(errs, `tax_paid.challans.${i}.date`, `Challan ${i + 1}: date of deposit`);
    fmt(errs, `tax_paid.challans.${i}.date`, `Challan ${i + 1}: date of deposit`, RE_DATE, "YYYY-MM-DD");
    req(errs, `tax_paid.challans.${i}.challan_no`, `Challan ${i + 1}: challan serial no.`);
    fmt(errs, `tax_paid.challans.${i}.challan_no`, `Challan ${i + 1}: serial no.`, RE_CHALLAN_NO, "up to 5 digits");
    if (!gt0(c.amount)) errs.push({ path: `tax_paid.challans.${i}.amount`, msg: `Challan ${i + 1}: amount must be greater than 0` });
  });
}

function checkBank(errs) {
  if (!(state.bank_accounts || []).length)
    errs.push({ path: "bank_accounts.0.ifsc", msg: "At least one bank account is mandatory" });
  (state.bank_accounts || []).forEach((b, i) => {
    req(errs, `bank_accounts.${i}.ifsc`, `Bank account ${i + 1}: IFSC`);
    fmt(errs, `bank_accounts.${i}.ifsc`, `Bank account ${i + 1}: IFSC`, RE_IFSC, "format SBIN0000123 (4 letters, 0, 6 more)");
    req(errs, `bank_accounts.${i}.bank_name`, `Bank account ${i + 1}: bank name`);
    req(errs, `bank_accounts.${i}.account_no`, `Bank account ${i + 1}: account number`);
    fmt(errs, `bank_accounts.${i}.account_no`, `Bank account ${i + 1}: account number`, /^[A-Za-z0-9\-]{4,20}$/, "4-20 letters/digits");
  });
  if ((state.bank_accounts || []).length && !(state.bank_accounts || []).some(b => b.use_for_refund))
    errs.push({ path: "bank_accounts.0.ifsc", msg: "Nominate one bank account for refund (radio button)" });
  const v = state.verification || {};
  req(errs, "verification.name", "Verification: name of declarant");
  req(errs, "verification.father_name", "Verification: father's name");
  fmt(errs, "verification.pan", "Verification: PAN", RE_PAN, "format ABCDE1234F");
  req(errs, "verification.place", "Verification: place of filing");
}

function checkReview(errs) {
  for (const fn of [checkPersonal, checkSalary, checkHP, checkOther, checkCapGains,
                    checkPresumptive, checkDeductions, checkTaxPaid, checkBank]) fn(errs);
}

function checkCapGains(errs) {
  const cg = state.capital_gains || {};
  const dateRow = (kind, rows, nameFields) => (rows || []).forEach((r, i) => {
    const anyVal = ["sale_consideration", "cost_of_acquisition", "cost_of_improvement", "expenses"].some(k => gt0(r[k]));
    if (!anyVal) return;                          // unfinished placeholder row
    req(errs, `capital_gains.${kind}.${i}.sale_consideration`, `CG row (${kind} ${i + 1}): sale consideration`);
    req(errs, `capital_gains.${kind}.${i}.cost_of_acquisition`, `CG row (${kind} ${i + 1}): cost of acquisition`);
    if (nameFields) req(errs, `capital_gains.${kind}.${i}.${nameFields}`, `CG row (${kind} ${i + 1}): ${nameFields === "scrip" ? "scrip name" : "description"}`);
    fmt(errs, `capital_gains.${kind}.${i}.buy_date`, `CG row (${kind} ${i + 1}): buy date`, RE_DATE, "YYYY-MM-DD");
    fmt(errs, `capital_gains.${kind}.${i}.sale_date`, `CG row (${kind} ${i + 1}): sale date`, RE_DATE, "YYYY-MM-DD");
    if (!blank(r.sale_date) && (r.sale_date < "2025-04-01" || r.sale_date > "2026-03-31"))
      errs.push({ path: `capital_gains.${kind}.${i}.sale_date`, msg: `CG row (${kind} ${i + 1}): sale must fall in FY 2025-26 (2025-04-01 to 2026-03-31)` });
  });
  dateRow("equity_st", cg.equity_st, "scrip");
  dateRow("equity_lt", cg.equity_lt, "scrip");
  dateRow("land_building", cg.land_building, "description");
  dateRow("other_st", cg.other_st, "description");
  dateRow("other_lt", cg.other_lt, "description");
  (cg.exemptions || []).forEach((r, i) => {
    if (!gt0(r.amount)) return;
    req(errs, `capital_gains.exemptions.${i}.section`, `CG exemption ${i + 1}: section`);
    if (r.section === "54EC" && Number(r.amount) > 5000000)
      errs.push({ path: `capital_gains.exemptions.${i}.amount`, msg: `CG exemption ${i + 1}: 54EC is capped at ₹50,00,000 — it will be restricted` });
  });
}
function checkPresumptive(errs) {
  const p = state.presumptive || {};
  (p["44ad"] || []).forEach((r, i) => {
    if (!gt0(r.gross_turnover) && !gt0(r.income_declared)) return;
    req(errs, `presumptive.44ad.${i}.name`, `44AD row ${i + 1}: business name`);
    req(errs, `presumptive.44ad.${i}.code`, `44AD row ${i + 1}: business code`);
    req(errs, `presumptive.44ad.${i}.gross_turnover`, `44AD row ${i + 1}: gross turnover`);
    if (Number(r.gross_turnover) > 30000000)
      errs.push({ path: `presumptive.44ad.${i}.gross_turnover`, msg: `44AD row ${i + 1}: presumptive ceiling is ₹3 crore — above that ITR-3 is required` });
  });
  (p["44ada"] || []).forEach((r, i) => {
    if (!gt0(r.gross_receipts) && !gt0(r.income_declared)) return;
    req(errs, `presumptive.44ada.${i}.name`, `44ADA row ${i + 1}: profession name`);
    req(errs, `presumptive.44ada.${i}.code`, `44ADA row ${i + 1}: profession code`);
    req(errs, `presumptive.44ada.${i}.gross_receipts`, `44ADA row ${i + 1}: gross receipts`);
    if (Number(r.gross_receipts) > 7500000)
      errs.push({ path: `presumptive.44ada.${i}.gross_receipts`, msg: `44ADA row ${i + 1}: presumptive ceiling is ₹75 lakh — above that ITR-3 is required` });
  });
  (p["44ae"] || []).forEach((r, i) => {
    if (blank(r.reg_no) && !gt0(r.months)) return;
    req(errs, `presumptive.44ae.${i}.reg_no`, `44AE vehicle ${i + 1}: registration number`);
    const m = Number(r.months);
    if (!(m >= 1 && m <= 12))
      errs.push({ path: `presumptive.44ae.${i}.months`, msg: `44AE vehicle ${i + 1}: months (1–12) held is required` });
  });
}

const STEP_CHECKS = {
  start: () => [], personal: checkPersonal, salary: checkSalary, hp: checkHP,
  other: checkOther, capgains: checkCapGains, presumptive: checkPresumptive,
  deductions: checkDeductions, taxpaid: checkTaxPaid, imports: () => [],
  bank: checkBank, review: checkReview,
};

function validateStep(i) {
  const errs = [];
  (STEP_CHECKS[TABS[i].id] || (() => []))(errs);
  return errs;
}
function validateAllSteps() {
  const results = [];
  TABS.forEach((t, i) => {
    if (t.id === "review") return;                    // review aggregates anyway
    const errs = validateStep(i);
    if (errs.length) results.push({ step: i, title: t.title, errors: errs });
  });
  return results;
}

/* ------------------------------------------------- validation UI plumbing */
let hardFailed = false;   // user was blocked on this step → red styling

function markBadPaths(errors) {
  document.querySelectorAll("[data-path].bad").forEach(n => n.classList.remove("bad"));
  errors.forEach(e2 => {
    const n = document.querySelector(`[data-path="${CSS.escape(e2.path)}"]`);
    if (n) n.classList.add("bad");
  });
}

function showStepAlerts(errors, hard) {
  const box = $("#step-alerts");
  if (!box) return;
  if (!errors.length) { box.classList.add("hidden"); box.innerHTML = ""; return; }
  box.classList.remove("hidden");
  box.classList.toggle("hard", !!hard);
  const listed = [...new Set(errors.map(e2 => e2.msg))];
  box.innerHTML =
    `<b>${hard ? "⛔ Complete these mandatory fields before going to the next step:"
               : "⚠ Mandatory fields still incomplete on this step:"}</b>
     <ul>${listed.slice(0, 9).map(m => `<li>${esc(m)}</li>`).join("")}
     ${listed.length > 9 ? `<li>…and ${listed.length - 9} more</li>` : ""}</ul>`;
}

function refreshSoftValidation() {
  const errs = validateStep(currentStep);
  showStepAlerts(errs, errs.length > 0 && hardFailed);
  // clear the red border on fields the user has since fixed
  document.querySelectorAll("[data-path].bad").forEach(n => {
    if (!errs.some(e2 => e2.path === n.dataset.path)) n.classList.remove("bad");
  });
  const st = $("#stepper");
  if (st) st.querySelectorAll(".step").forEach((btn, i) => {
    const c = validateStep(i).length;
    let badge = btn.querySelector(".err-count");
    if (c && i !== currentStep) {
      if (!badge) { badge = el(`<span class="err-count">${c}</span>`); btn.appendChild(badge); }
      badge.textContent = c;
    } else if (badge) badge.remove();
  });
}

function guardStep(i) {
  const errs = validateStep(i);
  if (!errs.length) { hardFailed = false; showStepAlerts([], false); return true; }
  hardFailed = true;
  markBadPaths(errs);
  showStepAlerts(errs, true);
  const first = document.querySelector("[data-path].bad");
  if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
  else $("#step-alerts")?.scrollIntoView({ behavior: "smooth", block: "start" });
  return false;
}

/* Run the complete mandatory-field map (used on Finish / downloads). */
function guardAllSteps(actionLabel) {
  const failing = validateAllSteps();
  if (!failing.length) return true;
  const first = failing[0];
  currentStep = first.step;
  render();
  guardStep(currentStep);
  showModal(`Mandatory fields incomplete — cannot ${actionLabel}`,
    `<p>The following required fields must be completed before you can ${actionLabel}:</p>` +
    failing.map(f => `<p><b>Step ${f.step + 1} · ${esc(f.title)}</b></p>
      <ul>${[...new Set(f.errors.map(e2 => e2.msg))].map(m => `<li>${esc(m)}</li>`).join("")}</ul>`).join("") +
    `<p class="muted">You have been taken to the first step that needs attention — the
     incomplete fields are highlighted in red.</p>`);
  return false;
}

/* Highlight a field reported by the server (DataError path). */
function markServerField(path) {
  if (!path) return;
  const n = document.querySelector(`[data-path="${CSS.escape(path)}"]`);
  if (n) n.classList.add("bad");
}


/* ------------------------------------------------------------ tab: personal */
function renderPersonal(host) {
  const p = state.personal, a = p.address;
  let h = "";
  h += section("Personal information", `<div class="grid c3">
      ${txt("personal.first_name", "First name", { req: 1, max: 25 })}
      ${txt("personal.middle_name", "Middle name", { max: 25 })}
      ${txt("personal.last_name", "Last name (as per PAN)", { req: 1, max: 75 })}
      ${txt("personal.pan", "PAN", { req: 1, max: 10, ph: "ABCDE1234F" })}
      ${dateInp("personal.dob", "Date of birth", "Determines senior-citizen status (60+ if born on/before 01-04-1966)", true)}
      ${txt("personal.aadhaar", "Aadhaar number", { max: 12 })}
      ${sel("personal.employer_category", "Nature of employment", optsFrom(meta.employer_categories, p.employer_category, "Select…"), "", true)}
      ${sel("personal.status", "Status", opt("I", "Individual", p.status) + opt("H", "HUF", p.status), "", true)}
      ${sel("personal.residential_status", "Residential status",
        opt("RES", "RES - Resident", p.residential_status) +
        opt("NRI", "NRI - Non Resident", p.residential_status) +
        opt("NOR", "NOR - Resident but not Ordinarily Resident", p.residential_status),
        "drives form eligibility, 87A rebate and senior slabs — exactly as in the utility", true)}
    </div>`);
  h += section("Address & contact", `<div class="grid c3">
      ${txt("personal.address.residence_no", "Flat/Door/Block no.", { req: 1 })}
      ${txt("personal.address.residence_name", "Premises/Building")}
      ${txt("personal.address.road_or_street", "Road/Street")}
      ${txt("personal.address.locality_or_area", "Locality/Area", { req: 1 })}
      ${txt("personal.address.city_or_town_or_district", "City/Town/District", { req: 1 })}
      ${sel("personal.address.state_code", "State", optsFrom(meta.state_codes, a.state_code, "Select…"), "", true)}
      ${txt("personal.address.pin_code", "PIN code", { max: 6, req: 1 })}
      ${txt("personal.address.mobile_no", "Mobile no.", { max: 10, req: 1 })}
      ${txt("personal.address.email_address", "Email", { type: "email", req: 1 })}
    </div>
    <hr class="sep">${radioPair("personal.secondary_address_same", "Is the secondary/communication address same as above?")}
    ${p.secondary_address_same === false ? `<div class="grid c3">
      ${txt("personal.alternate_address.residence_no", "Alt. Flat/No")}
      ${txt("personal.alternate_address.locality_or_area", "Alt. Locality")}
      ${txt("personal.alternate_address.city_or_town_or_district", "Alt. City")}
      ${sel("personal.alternate_address.state_code", "Alt. State", optsFrom(meta.state_codes, getPath(state, "personal.alternate_address.state_code"), "Select…"))}
      ${txt("personal.alternate_address.pin_code", "Alt. PIN", { max: 6 })}
    </div>` : ""}`);

  const fs = state.filing;
  h += section("Filing status", `<div class="grid c3">
      ${sel("filing.return_section", "Return filed under section", optsFrom(meta.return_sections, fs.return_section))}
      ${dateInp("filing.filing_date", "Date of filing", "Used for fee u/s 234F and interest u/s 234A/234B")}
      <div class="field"><label>Tax regime</label><div class="inline">
        <label><input type="radio" name="regime" data-noauto data-action="regime" value="new" ${fs.opt_out_new_regime ? "" : "checked"}> New regime (default)</label>
        <label><input type="radio" name="regime" data-noauto data-action="regime" value="old" ${fs.opt_out_new_regime ? "checked" : ""}> Old regime</label>
      </div></div>
    </div>
    ${fs.return_section === 17 ? `<div class="grid c2" style="margin-top:.8rem">
      ${txt("filing.original_receipt_no", "Acknowledgement no. of original return (15 digits)")}
      ${dateInp("filing.original_filing_date", "Date of filing of original return")}
    </div>` : ""}
    ${[13, 14, 16, 18].includes(Number(fs.return_section)) ? `<div class="grid c2" style="margin-top:.8rem">
      ${txt("filing.notice_no", "Notice no.")}
      ${dateInp("filing.notice_date", "Date of notice/order")}
    </div>` : ""}
    <hr class="sep">
    ${radioPair("filing.is_representative", "Is this return being filed by a representative assessee?")}
    ${fs.is_representative ? `<div class="grid c3" style="margin-top:.6rem">
      ${txt("filing.representative.name", "Representative name")}
      ${txt("filing.representative.email", "Representative email")}
      ${txt("filing.representative.mobile", "Representative mobile")}
    </div>` : ""}
    <hr class="sep">
    <h3>Seventh proviso to section 139(1) — high-expenditure mandatory filing</h3>
    <div class="inline">${check("filing.seventh_provisio", "Applies (foreign travel > ₹2L / electricity > ₹1L / deposits > ₹1 Cr)")}</div>
    ${fs.seventh_provisio ? `<div class="grid c3" style="margin-top:.6rem">
      ${check("filing.foreign_travel_exceeds_2l", "Foreign travel expense > ₹2 lakh")}
      ${check("filing.electricity_exceeds_1l", "Electricity expense > ₹1 lakh")}
      ${check("filing.deposit_exceeds_1cr", "Current-account deposits > ₹1 crore")}
    </div>` : ""}`);
  host.appendChild(el(`<div>${h}</div>`));
}

/* -------------------------------------------------------------- tab: salary */
function renderSalary(host) {
  const s = state.salary;
  let h = section("Salary / pension income (as per Form 16)", `
    <p class="hint">💡 <b>Double-click</b> the 17(1) box (or use the button) to open the
    <b>monthly salary grid</b> — totals flow back here automatically.</p>
    <div class="inline" style="margin-bottom:.6rem">
      <button type="button" class="btn small ghost" data-action="monthly-open">🗓 Monthly break-up…</button>
      ${(s.monthly || []).some(r => r && Object.keys(r).length > 1)
        ? `<span class="badge green">${(s.monthly || []).length} month(s) filled</span>` : ""}
    </div>
    <div class="grid c3">
      ${money("salary.salary_17_1", "Salary u/s 17(1)", "Basic + allowances + bonus (before exemptions)", "data-monthly")}
      ${money("salary.perquisites_17_2", "Perquisites u/s 17(2)")}
      ${money("salary.profits_17_3", "Profits in lieu of salary u/s 17(3)")}
    </div>
    <p class="hint">Or skip these three boxes entirely and use the <b>salary-slip break-up</b>
    below — the engine aggregates it into 17(1)/17(2)/17(3) with per-item exemptions.</p>
    <div class="grid c3">
      ${money("salary.entertainment_allowance", "Entertainment allowance received u/s 16(ii)", "CG/SG/PSU only, old regime")}
      ${money("salary.professional_tax", "Professional tax u/s 16(iii)", "Old regime only")}
      ${money("salary.standard_deduction_override", "Standard deduction override (optional)")}
    </div>`);

  /* ---------------- employer + salary-slip break-up (advanced, multi-form) --- */
  const emp = s.employer || {};
  h += `<details class="adv" ${(s.components || []).length || emp.name ? "open" : ""}>
    <summary>🏢 Employer &amp; salary-slip break-up <span class="muted">(earnings grid as per payslip — feeds 17(1))</span></summary>
    <div class="adv-body">${
    section("Employer", `<div class="grid c4">
      ${txt("salary.employer.name", "Employer name")}
      ${txt("salary.employer.tan", "TAN of employer", { ph: "AAAA00000A" })}
      ${sel("salary.employer.category", "Employer category", optsFrom(meta.employer_categories, emp.category || "OTH", "Select…"))}
      ${txt("salary.employer.pin_code", "Office PIN", { max: "6" })}
    </div>
    <div class="grid c2">
      ${txt("salary.employer.address", "Office address")}
      ${txt("salary.employer.city", "City")}
    </div>`)}
    ${section("Earnings grid (as per salary slip)", (() => {
      const rowHtml = (row, i) => {
        let opts = "";
        for (const [code, info] of Object.entries(CATALOG.components || {})) {
          const label = Array.isArray(info) ? info[0] : (info.label || code);
          opts += opt(code, `${code} — ${label}`, row.code);
        }
        const inf = (CATALOG.components || {})[row.code];
        const sec = Array.isArray(inf) ? inf[1] : "";
        return `<div class="grid c4">
          ${sel(`salary.components.${i}.code`, "Component", opts)}
          ${money(`salary.components.${i}.amount`, "Amount (annual)")}
          ${row.code === "ARREARS"
            ? txt(`salary.components.${i}.arrear_fy`, "Belongs to FY", { ph: "2023-24" })
            : `<div class="field"><label>Exemption</label><div class="hint" style="padding-top:.45rem">${sec && sec !== "null" ? `u/s ${sec} (auto-limited)` : "fully taxable"}</div></div>`}
          ${row.code === "ARREARS"
            ? money(`salary.components.${i}.income_that_fy`, "Total income of that FY", "for §89 relief (Form 10E)")
            : ""}
        </div>`;
      };
      return repeater("salary.components", "Add earning component", rowHtml,
        { itemTitle: (r, i) => r.code ? `${r.code} — ₹${Number(r.amount || 0).toLocaleString("en-IN")}` : `Component ${i + 1}` }) +
      `<p class="hint" style="margin-top:.6rem">Exempt allowances (HRA aside) are reduced with their
      statutory limits automatically. Tag an <b>ARREARS</b> component with the earlier FY and that
      year's total income — relief u/s 89 (Form 10E) is computed on its own.</p>`;
    })())}
    </div>
  </details>`;

  /* ---------------- perquisites + profits ---------------------------------- */
  h += `<details class="adv" ${(s.perquisites || []).length || (s.profits_detail || []).length ? "open" : ""}>
    <summary>💼 Perquisites (17(2)) and profits in lieu (17(3))</summary>
    <div class="adv-body">${
    section("Perquisites u/s 17(2)", (() => {
      const rowHtml = (row, i) => {
        let opts = "";
        for (const [code, info] of Object.entries(CATALOG.perquisites || {})) {
          const label = typeof info === "string" ? info : (info.label || code);
          opts += opt(code, `${label}`, row.code);
        }
        return `<div class="grid c2">
          ${sel(`salary.perquisites.${i}.code`, "Perquisite", opts)}
          ${money(`salary.perquisites.${i}.amount`, "Taxable value (after rule 3 valuation)")}
        </div>`;
      };
      return repeater("salary.perquisites", "Add perquisite", rowHtml,
        { itemTitle: (r, i) => r.code || `Perquisite ${i + 1}` }) +
      `<p class="hint" style="margin-top:.6rem">Enter the <b>taxable</b> value (e.g. rent-free
      accommodation as per rule 3). These aggregate into 17(2).</p>`;
    })())}
    ${section("Profits in lieu of salary u/s 17(3)", (() => {
      const rowHtml = (row, i) => `<div class="grid c2">
        ${txt(`salary.profits_detail.${i}.label`, "Nature (e.g. keyman policy)")}
        ${money(`salary.profits_detail.${i}.amount`, "Amount")}
      </div>`;
      return repeater("salary.profits_detail", "Add profit-in-lieu item", rowHtml,
        { itemTitle: (r, i) => r.label || `Item ${i + 1}` });
    })())}
    </div>
  </details>`;

  /* ---------------- terminal benefits -------------------------------------- */
  const T = s.terminal || {};
  const tb = (key, title, inner) => `<details class="adv" ${(T[key] && Object.keys(T[key]).length) ? "open" : ""}>
    <summary>${title}</summary><div class="adv-body" style="padding-top:.7rem">${inner}</div></details>`;
  h += `<details class="adv" ${Object.keys(T).length ? "open" : ""}>
    <summary>🎖 Terminal benefits (retirement) — least-of-three exemptions auto-computed</summary>
    <div class="adv-body">
    ${tb("gratuity", "Gratuity u/s 10(10)", `<div class="grid c4">
      ${money("salary.terminal.gratuity.received", "Gratuity received")}
      ${money("salary.terminal.gratuity.last_drawn_basic_da", "Last drawn (basic + DA) / month")}
      ${money("salary.terminal.gratuity.years", "Years of service")}
      <div class="field"><label>&nbsp;</label>${check("salary.terminal.gratuity.covered_act", "Covered by Payment of Gratuity Act")}</div>
    </div>
    <div class="inline">${check("salary.terminal.gratuity.govt", "Government employee (fully exempt)")}</div>`)}
    ${tb("leave_encashment", "Leave encashment u/s 10(10AA)", `<div class="grid c4">
      ${money("salary.terminal.leave_encashment.received", "Leave encashment received")}
      ${money("salary.terminal.leave_encashment.avg_10m_salary", "Average salary of last 10 months")}
      ${moneyf("salary.terminal.leave_encashment.leave_months_standing", "Leave standing to credit (months)")}
      ${moneyf("salary.terminal.leave_encashment.years", "Years of service")}
    </div>
    <div class="inline">${check("salary.terminal.leave_encashment.govt", "Government employee (fully exempt)")}
    <span class="hint">Exempt = least of [received · ₹25L · 10×avg · months-capped leave salary]</span></div>`)}
    ${tb("commuted_pension", "Commuted pension u/s 10(10A)", `<div class="grid c2">
      ${money("salary.terminal.commuted_pension.received", "Commuted pension received")}
      <div class="field"><label>&nbsp;</label>${check("salary.terminal.commuted_pension.govt", "Government employee (fully exempt)")}</div>
    </div>`)}
    ${tb("vrs", "VRS / voluntary separation u/s 10(10C)", `<div class="grid c4">
      ${money("salary.terminal.vrs.received", "VRS compensation received")}
      ${money("salary.terminal.vrs.months_salary", "Salary per month (basic + DA)")}
      ${moneyf("salary.terminal.vrs.years_completed", "Years of service completed")}
      ${moneyf("salary.terminal.vrs.months_remaining", "Months left to retirement")}
    </div>
    <p class="hint">Exempt = least of [received · ₹5,00,000 · 3×salary×years · salary×months-remaining]</p>`)}
    ${tb("retrenchment", "Retrenchment compensation u/s 10(10B)", `<div class="grid c4">
      ${money("salary.terminal.retrenchment.received", "Compensation received")}
      ${money("salary.terminal.retrenchment.avg_3m_salary", "Average salary of last 3 months")}
      ${moneyf("salary.terminal.retrenchment.years", "Years of service")}
      ${moneyf("salary.terminal.retrenchment.months_extra", "Months in excess of full years")}
    </div>
    <p class="hint">Exempt = least of [received · ₹5,00,000 · 15/26 × avg × (years + extra/12)]</p>`)}
    </div>
  </details>`;

  /* ---------------- arrears / relief 89 ------------------------------------ */
  h += `<details class="adv" ${(s.arrears || []).length ? "open" : ""}>
    <summary>⏪ Arrears of earlier years — relief u/s 89 (Form 10E)</summary>
    <div class="adv-body">${(() => {
      const rowHtml = (row, i) => `<div class="grid c3">
        ${txt(`salary.arrears.${i}.fy`, "Arrear belongs to FY", { ph: "2023-24" })}
        ${money(`salary.arrears.${i}.amount`, "Arrear amount (received this year)")}
        ${money(`salary.arrears.${i}.income_that_fy`, "Total income of that FY (without arrear)")}
      </div>`;
      return repeater("salary.arrears", "Add arrear year", rowHtml,
        { itemTitle: (r, i) => r.fy || `Arrear ${i + 1}` }) +
      `<div class="inline" style="margin-top:.6rem">
        <button type="button" class="btn small ghost" data-action="relief89">Compute §89 relief…</button>
        <span class="badge" id="relief89-badge"></span>
      </div>
      <p class="hint">Relief is computed automatically in the return (no manual entry needed) and
      shown as the Form-10E annexure in the computation PDF. The arrear itself stays inside your
      17(1) / earnings grid — it is not added twice.</p>`;
    })()}</div>
  </details>`;

  h += section("House Rent Allowance — schedule 10(13A)", `
    <p class="hint">Exemption = least of (actual HRA, rent − 10% of basic+DA, 50%/40% of basic+DA). Old regime only.</p>
    <div class="inline">${check("salary.has_hra", "I received HRA and paid rent")}</div>
    ${hraEnabled() ? `<div class="grid c4" style="margin-top:.7rem">
      ${sel("salary.hra.place", "Place of work", opt("1", "Metro (Delhi/Mumbai/Chennai/Kolkata)", hra().place ?? "1") + opt("2", "Non-metro", hra().place ?? "1"))}
      ${money("salary.hra.basic_salary", "Basic salary")}
      ${money("salary.hra.da", "Dearness allowance")}
      ${money("salary.hra.hra_received", "HRA received")}
      ${money("salary.hra.rent_paid", "Rent paid")}
    </div>
    <p class="hint" id="hra-preview"></p>` : ""}`);

  const allwHtml = (row, i) => `<div class="grid c2">
      ${sel(`salary.exempt_allowances.${i}.section`, "Section", (() => {
        let opts = "";
        for (const [v, l] of Object.entries(meta.exempt_allowance_sections)) opts += opt(v, `${v} — ${l}`, row.section);
        return opts;
      })())}
      ${money(`salary.exempt_allowances.${i}.amount`, "Exempt amount")}
    </div>`;
  h += section("Exempt allowances u/s 10 (other than HRA)",
    repeater("salary.exempt_allowances", "Add exempt allowance", allwHtml,
      { itemTitle: (r, i) => r.section ? `${r.section}` : `Allowance ${i + 1}` }) +
    `<p class="hint" style="margin-top:.6rem">In the new regime only the (115BAC) variants and
    retirement-benefit exemptions (10(10), 10(10A), 10(10AA), 10(10C)) are allowed.</p>`);
  host.appendChild(el(`<div>${h}</div>`));
  updateHraPreview();
}
const hraEnabled = () => state.salary.has_hra === true || (state.salary.hra && Object.values(state.salary.hra).some(v => v));
const hra = () => (state.salary.hra = state.salary.hra || {});
function updateHraPreview() {
  const elp = $("#hra-preview");
  if (!elp || !hraEnabled()) return;
  const h = hra();
  const sal = (+h.basic_salary || 0) + (+h.da || 0);
  const a = +h.hra_received || 0, r = +h.rent_paid || 0;
  const limbs = [a, Math.max(0, r - Math.floor(sal / 10)), Math.floor(sal * (String(h.place ?? "1") === "1" ? 0.5 : 0.4))];
  elp.innerHTML = `Eligible exemption preview: <b>${inr(Math.min(...limbs))}</b>
    (HRA ${inr(a)} · rent−10% ${inr(limbs[1])} · ${String(h.place ?? "1") === "1" ? "50" : "40"}% ${inr(limbs[2])})`;
}

/* ---------------------------------------------------- monthly salary grid */
const MONTH_COLS = [["basic", "Basic"], ["da", "DA"], ["hra", "HRA recd"],
                    ["rent_paid", "Rent"], ["other_taxable", "Other taxable"],
                    ["other_exempt", "Other exempt"], ["pt", "Prof. tax"], ["tds", "TDS"]];
function defaultMonths() {
  const out = [];
  for (let m = 4; m <= 12; m++) out.push({ month: `2025-${String(m).padStart(2, "0")}` });
  for (let m = 1; m <= 3; m++) out.push({ month: `2026-${String(m).padStart(2, "0")}` });
  return out;
}
function openMonthlyModal() {
  const rows = (state.salary.monthly || []).length ? state.salary.monthly : defaultMonths();
  const tbl = `<p class="hint">Fill what your payslips show per month — totals go to
    17(1), the HRA schedule and professional tax automatically. Leave unused months empty.</p>
    <div style="overflow-x:auto"><table class="summary monthly-grid">
      <tr class="head"><th>Month</th>${MONTH_COLS.map(c => `<th>${c[1]}</th>`).join("")}</tr>
      ${rows.map((r, i) => `<tr>
        <td><b>${esc(r.month)}</b></td>
        ${MONTH_COLS.map(c => `<td><input type="number" min="0" step="1" class="mini"
            data-mfill="${i}.${c[0]}" value="${esc(r[c[0]] ?? "")}"></td>`).join("")}
      </tr>`).join("")}
    </table></div>
    <div class="inline" style="margin-top:.8rem">
      <button type="button" class="btn primary small" data-action="monthly-save">Save break-up</button>
      <span id="monthly-total" class="muted"></span>
    </div>`;
  showModal("Monthly salary break-up — FY 2025-26", tbl);
}
async function saveMonthlyModal() {
  const rows = defaultMonths();
  document.querySelectorAll("[data-mfill]").forEach(inp => {
    const [i, k] = inp.dataset.mfill.split(".");
    if (inp.value !== "") rows[+i][k] = Number(inp.value) || 0;
  });
  const res = await postJSON("/api/salary/monthly-fill", { salary: { monthly: rows } });
  if (res.error) { showModal("Monthly break-up", `<p>${esc(res.error)}</p>`); return; }
  state.salary.monthly = rows;
  const gross = res.basic + res.da + res.hra_received + res.other_taxable;
  if (!Number(state.salary.salary_17_1) && gross > 0) state.salary.salary_17_1 = gross;
  if (res.professional_tax && !Number(state.salary.professional_tax))
    state.salary.professional_tax = res.professional_tax;
  if (res.rent_paid && res.hra_received && !hraEnabled()) {
    state.salary.has_hra = true;
    const h = hra();
    h.basic_salary = res.basic; h.da = res.da;
    h.hra_received = res.hra_received; h.rent_paid = res.rent_paid;
  }
  $("#modal-overlay").classList.add("hidden");
  render();
  showModal("Monthly break-up saved ✓", `<table class="summary">
    <tr><td>Months filled</td><td class="num"><b>${res.months_filled}</b></td></tr>
    <tr><td>Basic + DA + HRA + other taxable</td><td class="num"><b>${inr(gross)}</b></td></tr>
    <tr><td>HRA received / rent paid (HRA schedule armed)</td><td class="num">${inr(res.hra_received)} / ${inr(res.rent_paid)}</td></tr>
    <tr><td>Professional tax</td><td class="num">${inr(res.professional_tax)}</td></tr>
    <tr><td>TDS as per payslips (also fill in Taxes paid)</td><td class="num">${inr(res.tds)}</td></tr>
  </table>`);
}

/* ---------------------------------------------------------------- tab: hp */
const blankProp = () => ({
  occupancy: "L", address_line: "", city: "", state_code: "", pin_code: "",
  owner: "SE", co_owned: false, assessee_share_pct: 100, co_owners: [], tenants: [],
  annual_lettable_value: "", rent_not_realized: "", municipal_taxes: "",
  interest_on_borrowed_capital: "", loans: [], arrears_unrealized_rent_received: "",
});
function renderHP(host) {
  const itemHtml = (p, i) => `
    <div class="grid c3">
      ${sel(`house_properties.${i}.occupancy`, "Type of property",
        opt("L", "Let out", p.occupancy) + opt("D", "Deemed let out", p.occupancy) + opt("S", "Self occupied", p.occupancy))}
      ${money(`house_properties.${i}.annual_lettable_value`, "Gross rent / lettable value (₹/year)")}
      ${money(`house_properties.${i}.rent_not_realized`, "Unrealised rent")}
      ${money(`house_properties.${i}.municipal_taxes`, "Municipal taxes paid")}
      ${money(`house_properties.${i}.interest_on_borrowed_capital`, "Interest on borrowed capital u/s 24(b)",
        p.occupancy === "S" ? "Self-occupied: capped at ₹2L (old), nil (new regime)" : "")}
      ${money(`house_properties.${i}.arrears_unrealized_rent_received`, "Arrears / unrealised rent received", "70% is taxable")}
    </div>
    <h3>Address</h3>
    <div class="grid c4">
      ${txt(`house_properties.${i}.address_line`, "Address")}
      ${txt(`house_properties.${i}.city`, "City")}
      ${sel(`house_properties.${i}.state_code`, "State", optsFrom(meta.state_codes, p.state_code, "Select…"))}
      ${txt(`house_properties.${i}.pin_code`, "PIN", { max: 6 })}
    </div>
    <h3>Ownership</h3>
    <div class="inline">
      ${sel(`house_properties.${i}.owner`, "Owned by", opt("SE", "Self", p.owner) + opt("MI", "Minor", p.owner) + opt("SP", "Spouse", p.owner) + opt("OT", "Others", p.owner))}
      ${check(`house_properties.${i}.co_owned`, "Co-owned property")}
      ${p.co_owned ? `<div class="field" style="width:130px"><label>Own share %</label>
        <input type="number" data-path="house_properties.${i}.assessee_share_pct" data-kind="float" value="${esc(p.assessee_share_pct)}"></div>` : ""}
    </div>
    ${p.co_owned ? "<h3>Co-owners</h3>" + repeater(`house_properties.${i}.co_owners`, "Add co-owner", (c, j) => `
      <div class="grid c4">
        ${txt(`house_properties.${i}.co_owners.${j}.name`, "Name")}
        ${txt(`house_properties.${i}.co_owners.${j}.pan`, "PAN")}
        <div class="field"><label>Share %</label><input type="number" data-path="house_properties.${i}.co_owners.${j}.share_pct" data-kind="float" value="${esc(c.share_pct ?? "")}"></div>
        ${txt(`house_properties.${i}.co_owners.${j}.aadhaar`, "Aadhaar (opt.)")}
      </div>`) : ""}
    ${p.occupancy !== "S" ? "<h3>Tenants</h3>" + repeater(`house_properties.${i}.tenants`, "Add tenant", (t, j) => `
      <div class="grid c3">
        ${txt(`house_properties.${i}.tenants.${j}.name`, "Tenant name")}
        ${txt(`house_properties.${i}.tenants.${j}.pan`, "Tenant PAN")}
        ${txt(`house_properties.${i}.tenants.${j}.aadhaar`, "Tenant Aadhaar (opt.)")}
      </div>`) : ""}
    <h3>Housing-loan details (schedule 24(b))</h3>
    ${repeater(`house_properties.${i}.loans`, "Add loan", (l, j) => loanEditor(`house_properties.${i}.loans.${j}`, l, "InterestUs24B"))}
  `;
  host.appendChild(el(section("Income from house property (max 2)",
    (state.house_properties.length ?
      repeater("house_properties", "Add property", itemHtml,
        { itemTitle: (p, i) => p.occupancy === "S" ? `Property ${i + 1} — self occupied` : `Property ${i + 1} — let out${p.occupancy === "D" ? " (deemed)" : ""}` })
      : `<p class="hint">No house property? Nothing to do here — otherwise add your property. ITR-1 allows up to two properties.</p>
         ${repeater("house_properties", "Add property", itemHtml)}`),
    "Where interest details are given per loan, their total replaces the single interest amount above.")));
}

function loanEditor(base, l, interestLabel = "interest") {
  return `<div class="grid c4">
    ${sel(`${base}.lender_type`, "Loan taken from", opt("B", "Bank", l.lender_type) + opt("I", "Institution / other", l.lender_type))}
    ${txt(`${base}.lender_name`, "Lender name")}
    ${txt(`${base}.loan_account_no`, "Loan account no.")}
    ${dateInp(`${base}.loan_date`, "Date of loan")}
    ${money(`${base}.loan_amount`, "Total loan amount")}
    ${money(`${base}.outstanding`, "Outstanding amount")}
    ${money(`${base}.interest`, "Interest paid during the year")}
  </div>`;
}

/* ------------------------------------------------------------- tab: other */
function renderOther(host) {
  const osHtml = (r, i) => `<div class="grid c3">
      ${sel(`other_sources.${i}.nature`, "Nature of income", optsFrom(meta.other_source_natures, r.nature, "Select…"))}
      ${money(`other_sources.${i}.amount`, "Amount")}
      ${r.nature === "OTH" ? txt(`other_sources.${i}.description`, "Description") : ""}
    </div>
    ${r.nature === "DIV" ? `<div class="grid c4" style="margin-top:.5rem">
      ${money(`other_sources.${i}.dividend_quarters.upto_15_06`, "Up to 15 Jun")}
      ${money(`other_sources.${i}.dividend_quarters.upto_15_09`, "Up to 15 Sep")}
      ${money(`other_sources.${i}.dividend_quarters.16_09_15_12`, "16 Sep – 15 Dec")}
      ${money(`other_sources.${i}.dividend_quarters.16_12_15_03`, "16 Dec – 15 Mar")}
      ${money(`other_sources.${i}.dividend_quarters.16_03_31_03`, "16 – 31 Mar")}
    </div>` : ""}`;
  let h = section("Income from other sources",
    repeater("other_sources", "Add income row", osHtml,
      { itemTitle: (r, i) => (meta.other_source_natures[r.nature] || `Row ${i + 1}`) }) +
    `<p class="hint" style="margin-top:.6rem">${FAMILY_PENSION_NOTE}</p>`);

  h += section("Long-term capital gains u/s 112A (up to ₹1.25 lakh)", `
    <div class="grid c2">
      ${money("ltcg112a.sale_consideration", "Total sale consideration")}
      ${money("ltcg112a.cost_of_acquisition", "Total cost of acquisition")}
    </div>
    <p class="hint">Gains up to ₹1,25,000 are exempt and can be reported in ITR-1. Higher gains require ITR-2.</p>`);

  // Special-rate income (Schedule SI of ITR-2) — taxed at fixed rates instead
  // of slab rates, exactly like the official utility's Schedule OS/SI.
  h += section("Special-rate income — taxed at special rates (Schedule SI)", `
    <div class="grid c3">
      ${money("other_sources_special.lottery_115bb", "Lottery / crossword / races / card games u/s 115BB (30%)")}
      ${money("other_sources_special.online_games_115bbj", "Winnings from online games u/s 115BBJ (30%)")}
      ${money("other_sources_special.unexplained_115bbe", "Unexplained income u/s 68–69D (115BBE, 60%)")}
    </div>
    <p class="hint">These amounts are moved out of normal other-sources income and taxed at the flat
    special rate in <b>ITR-2</b> (Schedule OS → Schedule SI), as the official utility does. Tax on
    unexplained income (115BBE) additionally carries a flat <b>25% surcharge</b>. Family-pension /
    savings / FD interest stay in the rows above at slab rate.</p>`);

  const exHtml = (r, i) => `<div class="grid c4">
      ${sel(`exempt_income.${i}.category`, "Category", optsFrom(EXEMPT_CATEGORIES, r.category, "Select…"))}
      ${sel(`exempt_income.${i}.sub_category`, "Nature", EXEMPT_SUBCATEGORIES.map(v => opt(v, v, r.sub_category)).join(""))}
      ${money(`exempt_income.${i}.amount`, "Amount")}
      ${txt(`exempt_income.${i}.description`, "Description (opt.)")}
    </div>`;
  h += section("Exempt income (reporting only)",
    repeater("exempt_income", "Add exempt income", exHtml) +
    `<p class="hint" style="margin-top:.6rem">Agricultural income exemption here is limited to ₹5,000 (rule 29); for more, use ITR-2.</p>`);
  host.appendChild(el(`<div>${h}</div>`));
}

/* ------------------------------------------------------- tab: capital gains */
const CG_SECTIONS_META = {
  equity_st: { n: "Listed equity / equity-MF — short term (≤12m, STT paid) 111A @ 20%",
               note: "Scrip-wise. Non-STT short sales go to “Other assets — short term”." },
  equity_lt: { n: "Listed equity / equity-MF — long term u/s 112A @ 12.5% (₹1.25L exempt)",
               note: "Grandfathering: for buys on/before 31-01-2018 enter the FMV — cost base = higher of cost and lower of FMV/sale." },
  land_building: { n: "Land / building — long term (≥24m)",
                   note: "Resident choice: 12.5% flat vs 20% with indexation (acquired before 23-07-2024) — 'auto' picks the cheaper tax per property." },
  other_st: { n: "Other assets — short term (slab rates)", note: "Gold, unlisted shares, property held <24m, crypto-unrelated…" },
  other_lt: { n: "Other assets — long term (12.5% / 20% indexed)", note: "Quoted bonds, unlisted equity held long, etc." },
};
function renderCapGains(host) {
  const cg = state.capital_gains;
  let h = section("Capital gains workbook", `
    <div class="inline" style="flex-wrap:wrap">
      <button class="btn small ghost" type="button" data-action="cg-export">⬇ Export to Excel</button>
      <button class="btn small ghost" type="button" data-action="cg-template">📋 Download JSON template</button>
      <button class="btn small ghost" type="button" data-action="cg-import">⬆ Import from Excel / 112A CSV</button>
      <input type="file" id="cg-import-file" accept=".xlsx,.csv" class="hidden" data-noauto>
      <span class="hint">Import merges rows into these grids (department CSV_112A format also understood).</span>
    </div>
    <p class="hint">Any amount entered here automatically moves your return to <b>ITR-2</b>
    (auto-detect on the Start tab shows it). Quarterly accrual for sec 234C is derived from the sale dates.</p>`);

  const rows = (kind, fields) => (r, i) => `
    <div class="grid c4">
      ${fields.map(f => {
        const [key, label, kind2] = f;
        const path = `capital_gains.${kind}.${i}.${key}`;
        if (kind2 === "date") return txt(path, label, { type: "date" });
        if (kind2 === "sel-idx") return sel(path, label,
          opt("auto", "Auto (lower tax)", r[key] ?? "auto") +
          opt("12.5", "12.5% without indexation", r[key]) +
          opt("20", "20% with indexation", r[key]));
        if (kind2 === "text") return txt(path, label);
        return money(path, label);
      }).join("")}
    </div>`;

  const detail = (kind, fields, extra = "") => {
    const total = (cg[kind] || []).reduce((s, r) =>
      s + (Number(r.sale_consideration) || 0) - (Number(r.cost_of_acquisition) || 0)
        - (Number(r.cost_of_improvement) || 0) - (Number(r.expenses) || 0), 0);
    return `<details class="adv" ${(cg[kind] || []).length ? "open" : ""}>
      <summary>${CG_SECTIONS_META[kind].n}
        ${(cg[kind] || []).length ? `<span class="badge ${total >= 0 ? "green" : "red"}">net ${inr(total)}</span>` : ""}</summary>
      <div class="adv-body">${extra}
      ${repeater(`capital_gains.${kind}`, "Add transaction", rows(kind, fields),
        { itemTitle: (r, i) => r.scrip || r.description || `Row ${i + 1}` })}
      <p class="hint" style="margin-top:.5rem">${CG_SECTIONS_META[kind].note}</p></div>
    </details>`;
  };

  h += detail("equity_st", [["scrip", "Scrip / fund name", "text"], ["buy_date", "Buy date", "date"],
    ["sale_date", "Sale date", "date"], ["sale_consideration", "Sale value"],
    ["cost_of_acquisition", "Cost of acquisition"], ["expenses", "Transfer expenses"]]);
  h += detail("equity_lt", [["scrip", "Scrip / fund name", "text"], ["isin", "ISIN", "text"],
    ["buy_date", "Buy date", "date"], ["sale_date", "Sale date", "date"],
    ["sale_consideration", "Sale value"], ["cost_of_acquisition", "Cost of acquisition"],
    ["fmv_31_01_2018", "FMV on 31-01-2018 (grandfathering)"], ["expenses", "Transfer expenses"]]);
  h += detail("land_building", [["description", "Property description", "text"],
    ["buy_date", "Buy date", "date"], ["sale_date", "Sale date", "date"],
    ["sale_consideration", "Sale value"], ["cost_of_acquisition", "Cost of acquisition"],
    ["cost_of_improvement", "Cost of improvement"], ["expenses", "Transfer expenses"],
    ["indexation_option", "Tax option", "sel-idx"]]);
  h += detail("other_st", [["description", "Asset description", "text"],
    ["buy_date", "Buy date", "date"], ["sale_date", "Sale date", "date"],
    ["sale_consideration", "Sale value"], ["cost_of_acquisition", "Cost of acquisition"],
    ["cost_of_improvement", "Cost of improvement"], ["expenses", "Transfer expenses"]]);
  h += detail("other_lt", [["description", "Asset description", "text"],
    ["buy_date", "Buy date", "date"], ["sale_date", "Sale date", "date"],
    ["sale_consideration", "Sale value"], ["cost_of_acquisition", "Cost of acquisition"],
    ["cost_of_improvement", "Cost of improvement"], ["expenses", "Transfer expenses"]]);

  const exHtml = (r, i) => `<div class="grid c3">
      ${sel(`capital_gains.exemptions.${i}.section`, "Section",
        [["54", "54 — new residential house (vs house LTCG)"],
         ["54B", "54B — agricultural land"],
         ["54EC", "54EC — NHAI/REC bonds (≤ ₹50 lakh)"],
         ["54EE", "54EE — notified fund units (≤ ₹50 lakh)"],
         ["54F", "54F — new residential house (vs other LTCG)"]]
          .map(([v, l]) => opt(v, l, r.section)).join(""))}
      ${money(`capital_gains.exemptions.${i}.amount`, "Amount invested")}
      ${txt(`capital_gains.exemptions.${i}.date`, "Date of investment", { type: "date" })}
    </div>`;
  h += section("Exemptions against LTCG (54 / 54B / 54EC / 54EE / 54F)",
    repeater("capital_gains.exemptions", "Add exemption claim", exHtml) +
    `<p class="hint" style="margin-top:.6rem">Applied in the engine order: 54EC/54EE → 20% & 12.5% & 112A pools;
    54 → 12.5%/20% pools; 54F → 112A pool. Caps are enforced automatically.</p>`);

  host.appendChild(el(`<div>${h}</div>`));
}

/* ------------------------------------------------------ tab: presumptive */
function renderPresumptive(host) {
  const p = state.presumptive;
  let h = section("Presumptive business / profession (44AD / 44ADA / 44AE)", `
    <p class="hint">For <b>resident</b> individuals carrying on small business or profession —
    sections 44AD/44ADA are not available to a <b>Non-Resident</b> (the official utility locks those
    cells for NR; HUFs cannot use 44ADA either, rule 212). The engine
    applies <b>max(declared, presumptive)</b> per undertaking and warns when declared income is below
    the presumptive floor. Entering anything here moves the return to <b>ITR-4</b> (total income up to
    ₹50 lakh); regular books of account need ITR-3 (use the department utility).</p>`);

  // Official nature-of-business dropdowns (315 codes for 44AD, 38 for 44ADA,
  // 7 for 44AE — straight from the CBDT utility lists) rendered as searchable
  // combo boxes: type any part of the code or label to filter, exactly like
  // the utility's dropdown filter.
  const codeList = (fam, id) => `
    <datalist id="${id}">${Object.entries((CATALOG.business_codes || {})[fam] || {})
      .map(([code, label]) => `<option value="${esc(code)}">${esc(`${code} — ${label}`)}</option>`).join("")}</datalist>`;
  const codeCombo = (path, label, fam, cur) => `
    <div class="field"><label>${esc(label)}</label>
      <input data-path="${path}" list="nob-${fam}" value="${esc(getPath(state, path) ?? cur ?? "")}"
             placeholder="type to search code…" autocomplete="off">
      <span class="sub">${esc(((((CATALOG.business_codes || {})[fam]) || {})[getPath(state, path)]) || `${Object.keys((CATALOG.business_codes || {})[fam] || {}).length} official codes — type to filter`)}</span></div>`;
  h += codeList("44ad", "nob-44ad") + codeList("44ada", "nob-44ada") + codeList("44ae", "nob-44ae");

  const adHtml = (r, i) => `<div class="grid c4">
      ${txt(`presumptive.44ad.${i}.name`, "Business name")}
      ${codeCombo(`presumptive.44ad.${i}.code`, "Business code", "44ad", r.code || "09028")}
      ${money(`presumptive.44ad.${i}.gross_turnover`, "Gross turnover (≤ ₹3 crore)")}
      ${money(`presumptive.44ad.${i}.turnover_bank`, "of which: via bank/digital")}
      ${money(`presumptive.44ad.${i}.turnover_cash`, "of which: cash")}
      ${money(`presumptive.44ad.${i}.turnover_other`, "of which: other non-cash")}
      ${money(`presumptive.44ad.${i}.income_declared`, "Income declared (if higher than presumptive)")}
    </div>
    <p class="hint">Presumptive income = 6% × digital turnover + 8% × cash turnover.</p>`;
  h += `<details class="adv" ${(p["44ad"] || []).length ? "open" : ""}>
    <summary>🏬 44AD — small business <span class="muted">(turnover ≤ ₹3 crore)</span></summary>
    <div class="adv-body">${repeater("presumptive.44ad", "Add 44AD business", adHtml,
      { itemTitle: (r, i) => r.name || `Business ${i + 1}` })}</div></details>`;

  const adaHtml = (r, i) => `<div class="grid c4">
      ${txt(`presumptive.44ada.${i}.name`, "Profession name")}
      ${codeCombo(`presumptive.44ada.${i}.code`, "Profession code", "44ada", r.code || "14001")}
      ${money(`presumptive.44ada.${i}.gross_receipts`, "Gross receipts (≤ ₹75 lakh)")}
      ${money(`presumptive.44ada.${i}.receipts_bank`, "of which: via bank/digital")}
      ${money(`presumptive.44ada.${i}.receipts_cash`, "of which: cash")}
      ${money(`presumptive.44ada.${i}.receipts_other`, "of which: other non-cash")}
      ${money(`presumptive.44ada.${i}.income_declared`, "Income declared (if higher than 50%)")}
    </div>
    <p class="hint">Presumptive income = 50% of gross receipts.</p>`;
  h += `<details class="adv" ${(p["44ada"] || []).length ? "open" : ""}>
    <summary>🎓 44ADA — specified profession <span class="muted">(receipts ≤ ₹75 lakh)</span></summary>
    <div class="adv-body">${repeater("presumptive.44ada", "Add 44ADA profession", adaHtml,
      { itemTitle: (r, i) => r.name || `Profession ${i + 1}` })}</div></details>`;

  const aeHtml = (r, i) => `<div class="grid c4">
      ${txt(`presumptive.44ae.${i}.reg_no`, "Vehicle registration no.")}
      ${sel(`presumptive.44ae.${i}.owned`, "Holding",
        opt("OWN", "Owned", r.owned ?? "OWN") + opt("LEASE", "Leased", r.owned) + opt("HIRED", "Hired", r.owned))}
      ${money(`presumptive.44ae.${i}.months`, "Months held (1–12)")}
      ${money(`presumptive.44ae.${i}.tonnage`, "Gross vehicle weight (tons; heavy = >12)")}
      <div class="field"><label>&nbsp;</label>${check(`presumptive.44ae.${i}.heavy`, "Heavy goods vehicle")}</div>
    </div>
    <p class="hint">Income = ₹7,500/month per vehicle (others) or ₹1,000/ton/month (heavy, >12 MT).</p>`;
  h += `<details class="adv" ${(p["44ae"] || []).length ? "open" : ""}>
    <summary>🚚 44AE — goods carriage <span class="muted">(up to 10 vehicles)</span></summary>
    <div class="adv-body">${repeater("presumptive.44ae", "Add vehicle", aeHtml,
      { itemTitle: (r, i) => r.reg_no || `Vehicle ${i + 1}` })}</div></details>`;

  const finMap = [["bank_balance", "Balance with banks"], ["cash_in_hand", "Cash in hand"],
    ["sundry_debtors", "Sundry debtors"], ["inventories", "Inventories / stock"],
    ["fixed_assets", "Fixed assets"], ["total_assets", "Total assets"],
    ["sundry_creditors", "Sundry creditors"], ["partner_capital", "Own capital"],
    ["total_capital_liabilities", "Total capital & liabilities"]];
  h += section("No-books financial particulars (Schedule BP — financials)",
    `<div class="grid c3">${finMap.map(([k, l]) => money(`presumptive.financials.${k}`, l)).join("")}</div>
     <p class="hint">Approximate balances as on 31-03-2026 — mandatory in ITR-4 when the presumptive scheme is used.</p>`);

  h += section("Regular business income (books of account)", `
    <div class="grid c2">${money("presumptive.business_income_manual", "Net profit as per P&L")}</div>
    <p class="hint">⚠ Filling this means you keep regular books — the correct form becomes
    <b>ITR-3</b>, which this suite does not generate (auto-detect will say so).</p>`);

  host.appendChild(el(`<div>${h}</div>`));
}

/* ---------------------------------------------------------- tab: imports */
let lastAis = null, last26as = null, lastPrevSummary = null;
function renderImports(host) {
  const lcf = state.losses_cf || {};
  let h = section("Import previous-year ITR JSON (prefill + comparison)", `
    <p class="hint">The JSON you downloaded from the e-Filing portal for <b>AY 2025-26</b>
    (or any earlier year). Personal details, address and bank accounts are prefilled (empty fields
    only) and brought-forward losses are captured below.</p>
    <div class="inline">
      <button class="btn small ghost" type="button" data-action="imp-itr-file">⬆ Upload ITR JSON…</button>
      <input type="file" id="imp-itr-json" accept=".json" class="hidden" data-noauto>
      <span id="imp-itr-status" class="badge"></span>
    </div>
    <div id="imp-itr-panel"></div>`);

  h += section("Import AIS (Annual Information Statement)", `
    <p class="hint">AIS JSON file — either plain <code>.json</code> or the password-protected
    <code>.zip</code> from the compliance portal. The zip password is
    <b>PAN (lowercase) + DOB (DDMMYYYY)</b> — it is tried automatically from your Personal tab.</p>
    <div class="inline">
      <button class="btn small ghost" type="button" data-action="imp-ais-file">⬆ Upload AIS json/zip…</button>
      <input type="file" id="imp-ais" accept=".json,.zip" class="hidden" data-noauto>
      <span id="imp-ais-status" class="badge"></span>
    </div>
    <div id="imp-ais-panel"></div>`);

  h += section("Import Form 26AS / Annual Tax Statement", `
    <p class="hint">The <code>.txt</code> or password-protected <code>.pdf</code> of Form 26AS.
    The PDF password is your <b>DOB (DDMMYYYY)</b> — auto-detected from your Personal tab.
    TDS/TCS/advance-tax entries can be pushed straight into the Taxes paid step.</p>
    <div class="inline">
      <button class="btn small ghost" type="button" data-action="imp-26as-file">⬆ Upload 26AS pdf/txt…</button>
      <input type="file" id="imp-26as" accept=".pdf,.txt" class="hidden" data-noauto>
      <span id="imp-26as-status" class="badge"></span>
    </div>
    <div id="imp-26as-panel"></div>`);

  h += section(`Brought-forward losses ${Object.values(lcf).some(v => Number(v)) ? "" : "(none entered)"}`, `
    <div class="grid c4">
      ${money("losses_cf.stcg_cf", "ST capital loss b/f")}
      ${money("losses_cf.ltcg_cf", "LT capital loss b/f")}
      ${money("losses_cf.hp_loss_cf", "House-property loss b/f")}
      ${money("losses_cf.business_loss_cf", "Business loss b/f")}
    </div>
    <p class="hint">Auto-filled when you import the previous-year ITR JSON above; editable.
    Amounts are <b>carried forward for reporting</b> — set-off against this year's income is flagged
    in the Review notes (file it on the portal after verifying the CFL schedule).</p>`);

  h += section("Year-on-year comparison", `
    <p class="hint">Once a previous-year ITR JSON is imported and a current-year summary exists,
    compare the two years head by head.</p>
    <div class="inline">
      <button class="btn small ghost" type="button" data-action="yoy-compare">📊 Compare AY 2026-27 vs imported year…</button>
    </div>
    <div id="yoy-panel"></div>`);

  host.appendChild(el(`<div>${h}</div>`));
}

/* -------------------------------------------------------- tab: deductions */
function renderDeductions(host) {
  const oldOnly = state.filing.opt_out_new_regime ? "" :
    `<div class="notice warn">You are currently on the <b>new regime</b>: Chapter VI-A deductions other than
    80CCD(2) and 80CCH are ignored. Switch to the old regime on the Review page if these deductions help you —
    the app will recommend the cheaper regime automatically.</div>`;
  let h = `<div class="card"><h2>Deductions under Chapter VI-A</h2>${oldOnly}</div>`;

  /* --- 80C block --- */
  h += section("80C / 80CCC / 80CCD — investments & NPS", `<div class="grid c3">
      ${money("deductions.80c.amount", "80C amount (LIC/PPF/ELSS/EPF…)", "Cap ₹1,50,000 combined with 80CCC+80CCD(1)")}
      ${money("deductions.80ccc.amount", "80CCC pension fund contribution")}
      ${money("deductions.80ccd1", "80CCD(1) — own NPS contribution", "≤10% of salary (or 20% of GTI for pensioners)")}
      ${money("deductions.80ccd1b", "80CCD(1B) — additional NPS", "Cap ₹50,000")}
      ${money("deductions.80ccd2", "80CCD(2) — employer NPS contribution", "New regime: ≤14% of salary. Allowed in both regimes.")}
      ${txt("deductions.prans", "PRAN number(s) (comma separated)", { ph: "12-digit PRAN" })}
    </div>
    <h3>80C supporting details (rule 224)</h3>
    ${repeater("deductions.80c.items", "Add 80C document", (r, i) => `<div class="grid c2">
      ${txt(`deductions.80c.items.${i}.identification_no`, "Document / identification no.")}
      ${money(`deductions.80c.items.${i}.amount`, "Amount")}
    </div>`)}`);

  /* --- 80D --- */
  const d80 = state.deductions["80d"] || {};
  const band = (label, base, val) => `<div class="rep-item"><div class="rep-title">${label}</div><div class="grid c3" style="margin-top:.5rem">
      ${money(`${base}.premium`, "Health-insurance premium")}
      ${money(`${base}.preventive`, "Preventive health check-up (≤₹5,000 all)")}
      ${val === "sr" ? money(`${base}.medical_exp`, "Medical expenditure (uninsured senior citizen)") : ""}
      ${txt(`${base}.insurer`, "Insurer name (mandatory for premium claim)")}
      ${txt(`${base}.policy_no`, "Policy number")}
    </div></div>`;
  h += section("80D — health insurance & medical expenditure", `
    <div class="grid c2">
      ${sel("deductions.80d.self_family_senior_citizen", "You / family (excl. parents): senior citizen?",
        opt("N", "No", d80.self_family_senior_citizen ?? "S") + opt("Y", "Yes", d80.self_family_senior_citizen ?? "S") + opt("S", "Not claiming for self/family", d80.self_family_senior_citizen ?? "S"))}
      ${sel("deductions.80d.parents_senior_citizen", "Parents: senior citizen?",
        opt("N", "No", d80.parents_senior_citizen ?? "P") + opt("Y", "Yes", d80.parents_senior_citizen ?? "P") + opt("P", "Not claiming for parents", d80.parents_senior_citizen ?? "P"))}
    </div>
    ${(d80.self_family_senior_citizen ?? "S") === "N" ? band("Self & family (non-senior) — cap ₹25,000", "deductions.80d.self_family") : ""}
    ${(d80.self_family_senior_citizen ?? "S") === "Y" ? band("Self & family incl. senior citizen — cap ₹50,000", "deductions.80d.self_family_sr", "sr") : ""}
    ${(d80.parents_senior_citizen ?? "P") === "N" ? band("Parents (non-senior) — cap ₹25,000", "deductions.80d.parents") : ""}
    ${(d80.parents_senior_citizen ?? "P") === "Y" ? band("Parents incl. senior citizen — cap ₹50,000", "deductions.80d.parents_sr", "sr") : ""}`);

  /* --- disability --- */
  h += section("80DD / 80U / 80DDB — disability & specified disease", `
    <div class="grid c2">
      <div class="rep-item"><div class="rep-title">80U — self with disability (₹75,000 / ₹1,25,000)</div>
        <div class="inline" style="margin:.5rem 0">${check("deductions.80u.enabled", "Claim 80U")}</div>
        ${dedEnabled("80u") ? `<div class="grid c2">
          ${sel("deductions.80u.nature", "Nature", opt("1", "Self with disability", "1") + opt("2", "Self with severe disability", getPath(state, "deductions.80u.nature")))}
          ${sel("deductions.80u.type", "Type", opt("2", "Other disabilities", getPath(state, "deductions.80u.type") ?? "2") + opt("1", "Autism/cerebral palsy/multiple", getPath(state, "deductions.80u.type") ?? "2"))}
          ${txt("deductions.80u.form10ia_ack", "Form 10IA ack no. (opt.)")}
          ${txt("deductions.80u.udid", "UDID no. (opt.)")}
        </div>` : ""}</div>
      <div class="rep-item"><div class="rep-title">80DD — dependent with disability</div>
        <div class="inline" style="margin:.5rem 0">${check("deductions.80dd.enabled", "Claim 80DD")}</div>
        ${dedEnabled("80dd") ? `<div class="grid c2">
          ${sel("deductions.80dd.nature", "Nature", opt("1", "Dependent with disability", getPath(state, "deductions.80dd.nature")) + opt("2", "Dependent with severe disability", getPath(state, "deductions.80dd.nature")))}
          ${sel("deductions.80dd.dependent_type", "Dependent", [["1", "Spouse"], ["2", "Son"], ["3", "Daughter"], ["4", "Father"], ["5", "Mother"], ["6", "Brother"], ["7", "Sister"]].map(x => opt(x[0], x[1], getPath(state, "deductions.80dd.dependent_type"))).join(""))}
          ${txt("deductions.80dd.dependent_pan", "Dependent PAN")}
          ${txt("deductions.80dd.dependent_aadhaar", "Dependent Aadhaar")}
          ${txt("deductions.80dd.form10ia_ack", "Form 10IA ack no. (opt.)")}
          ${txt("deductions.80dd.udid", "UDID no. (opt.)")}
        </div>` : ""}</div>
    </div>
    <h3>80DDB — medical treatment of specified disease</h3>
    <div class="grid c3">
      ${sel("deductions.80ddb.type", "Patient category", opt("1", "Self/dependent — cap ₹40,000", getPath(state, "deductions.80ddb.type")) + opt("2", "Senior citizen — cap ₹1,00,000", getPath(state, "deductions.80ddb.type")))}
      ${sel("deductions.80ddb.disease", "Specified disease", optsFrom(DDB_DISEASES, getPath(state, "deductions.80ddb.disease"), "Select…"))}
      ${money("deductions.80ddb.amount", "Amount")}
    </div>`);

  /* --- loans 80E/EE/EEA/EEB --- */
  const loanSched = (key, title, note) => `
    <div class="inline" style="margin-bottom:.4rem">${check(`deductions.${key}.enabled`, `Claim ${title}`)}</div>
    ${dedEnabled(key) ? repeater(`deductions.${key}.loans`, "Add loan", (l, i) => loanEditor(`deductions.${key}.loans.${i}`, l)) +
      (key === "80eea" ? `<div class="grid c2" style="margin-top:.5rem">${money("deductions.80eea.stamp_duty_value", "Stamp-duty value of property (≤ ₹45L)")}</div>` : "") +
      `<p class="hint" style="margin-top:.4rem">${note}</p>` : ""}`;
  h += section("80E / 80EE / 80EEA / 80EEB — loan interest", `
    ${loanSched("80e", "80E — education-loan interest", "No monetary cap; old regime only.")}
    <hr class="sep">${loanSched("80ee", "80EE — first home loan (FY 2016-17, loan ≤ ₹35L)", "Cap ₹50,000; loan sanctioned 01-04-2016 to 31-03-2017.")}
    <hr class="sep">${loanSched("80eea", "80EEA — affordable housing", "Cap ₹1,50,000; loan sanctioned 01-04-2019 to 31-03-2022; SDV ≤ ₹45L. 80EE and 80EEA are mutually exclusive.")}
    <hr class="sep">${loanSched("80eeb", "80EEB — electric-vehicle loan", "Cap ₹1,50,000; loan sanctioned 01-04-2019 to 31-03-2023.")}`);

  /* --- donations --- */
  const doneeHtml = (key) => (r, i) => `<div class="grid c3">
      ${txt(`${key}.${i}.name`, "Donee name")}
      ${txt(`${key}.${i}.pan`, "Donee PAN")}
      ${txt(`${key}.${i}.arn`, "ARN (donation reference, opt.)")}
      ${txt(`${key}.${i}.address.addr_detail`, "Donee address")}
      ${txt(`${key}.${i}.address.city`, "City")}
      ${sel(`${key}.${i}.address.state_code`, "State", optsFrom(meta.state_codes, getPath(state, `${key}.${i}.address.state_code`), "Select…"))}
      ${txt(`${key}.${i}.address.pin_code`, "PIN", { max: 6 })}
      ${money(`${key}.${i}.cash`, "Donation in cash (≤ ₹2,000 eligible)")}
      ${money(`${key}.${i}.other`, "Donation in other mode")}
      ${txt(`${key}.${i}.txn_ref`, "Transaction reference (other mode)")}
      ${txt(`${key}.${i}.ifsc`, "Donee IFSC (other mode)")}
    </div>`;
  h += section("80G / 80GGA / 80GGC — donations (old regime)", `
    <h3>80G — entitled to 100% deduction without qualifying limit</h3>${repeater("deductions.80g.don_100", "Add donee (100%)", doneeHtml("deductions.80g.don_100"))}
    <h3>80G — entitled to 50% deduction without qualifying limit</h3>${repeater("deductions.80g.don_50", "Add donee (50%)", doneeHtml("deductions.80g.don_50"))}
    <h3>80G — 100% subject to qualifying limit (10% of adjusted GTI)</h3>${repeater("deductions.80g.don_100_ql", "Add donee (100% QL)", doneeHtml("deductions.80g.don_100_ql"))}
    <h3>80G — 50% subject to qualifying limit</h3>${repeater("deductions.80g.don_50_ql", "Add donee (50% QL)", doneeHtml("deductions.80g.don_50_ql"))}
    <hr class="sep">
    <h3>80GGA — donations for scientific research / rural development</h3>
    ${repeater("deductions.80gga.donations", "Add 80GGA donation", (r, i) => `<div class="grid c3">
      ${sel(`deductions.80gga.donations.${i}.clause`, "Relevant clause", [
        ["80GGA2a", "80GGA(2)(a) Scientific research"], ["80GGA2aa", "80GGA(2)(aa) Social/statistical research"],
        ["80GGA2b", "80GGA(2)(b) Rural development"], ["80GGA2bb", "80GGA(2)(bb) Approved project"],
        ["80GGA2c", "80GGA(2)(c) Natural resources"], ["80GGA2cc", "80GGA(2)(cc) Afforestation"],
        ["80GGA2d", "80GGA(2)(d) Rural development fund"], ["80GGA2e", "80GGA(2)(e) NUPEF"],
      ].map(x => opt(x[0], x[1], r.clause)).join(""))}
      ${txt(`deductions.80gga.donations.${i}.name`, "Donee name")}
      ${txt(`deductions.80gga.donations.${i}.pan`, "Donee PAN")}
      ${txt(`deductions.80gga.donations.${i}.address.addr_detail`, "Address")}
      ${txt(`deductions.80gga.donations.${i}.address.city`, "City")}
      ${sel(`deductions.80gga.donations.${i}.address.state_code`, "State", optsFrom(meta.state_codes, getPath(state, `deductions.80gga.donations.${i}.address.state_code`), "Select…"))}
      ${txt(`deductions.80gga.donations.${i}.address.pin_code`, "PIN", { max: 6 })}
      ${money(`deductions.80gga.donations.${i}.cash`, "Cash (≤ ₹2,000 eligible)")}
      ${money(`deductions.80gga.donations.${i}.other`, "Other mode")}
    </div>`)}
    <hr class="sep">
    <h3>80GGC — contributions to political parties / electoral trusts</h3>
    ${repeater("deductions.80ggc.donations", "Add 80GGC contribution", (r, i) => `<div class="grid c3">
      ${dateInp(`deductions.80ggc.donations.${i}.date`, "Date of contribution")}
      ${txt(`deductions.80ggc.donations.${i}.party_name`, "Political party name")}
      ${txt(`deductions.80ggc.donations.${i}.party_pan`, "Party PAN")}
      ${money(`deductions.80ggc.donations.${i}.other`, "Contribution (other mode only)")}
      ${txt(`deductions.80ggc.donations.${i}.txn_ref`, "Transaction reference")}
      ${txt(`deductions.80ggc.donations.${i}.ifsc`, "IFSC")}
    </div>`)}`);

  /* --- remaining small ones --- */
  h += section("80GG / 80TTA / 80TTB / 80CCH", `<div class="grid c3">
      ${money("deductions.80gg.amount", "80GG — rent paid (no HRA)", "Min of ₹60,000 / 25% of adjusted income / rent−10%")}
      ${money("deductions.80gg.rent_paid", "Rent paid (for 80GG computation)")}
      ${txt("deductions.80gg.form10ba_ack", "Form 10BA ack no.")}
      ${money("deductions.80tta", "80TTA — savings-account interest", "Cap ₹10,000; below 60 years; old regime")}
      ${money("deductions.80ttb", "80TTB — interest (senior citizens)", "Cap ₹50,000; 60+ only; old regime")}
      ${money("deductions.80cch", "80CCH — Agniveer Corpus Fund", "CG employees (Agnipath); ≤46.2% of salary; both regimes")}
    </div>`);
  host.appendChild(el(`<div>${h}</div>`));
}
const dedEnabled = key => !!(state.deductions[key] && state.deductions[key].enabled) ||
  (state.deductions[key] && Array.isArray(state.deductions[key].loans) && state.deductions[key].loans.length > 0);

/* ----------------------------------------------------------- tab: taxpaid */
function renderTaxPaid(host) {
  const tp = state.tax_paid;
  let h = section("TDS on salary (Schedule TDS-1, as per Form 16)",
    repeater("tax_paid.tds_on_salary", "Add employer / deductor", (r, i) => `<div class="grid c4">
      ${txt(`tax_paid.tds_on_salary.${i}.tan`, "Employer TAN", { ph: "MUMB12345E" })}
      ${txt(`tax_paid.tds_on_salary.${i}.employer_name`, "Employer name")}
      ${money(`tax_paid.tds_on_salary.${i}.salary_income`, "Income chargeable under salaries")}
      ${money(`tax_paid.tds_on_salary.${i}.tds`, "TDS deducted")}
    </div>`));
  h += section("TDS on income other than salary (Schedule TDS-2, Form 16A)",
    repeater("tax_paid.tds_others", "Add deductor", (r, i) => `<div class="grid c3">
      ${txt(`tax_paid.tds_others.${i}.tan`, "Deductor TAN")}
      ${txt(`tax_paid.tds_others.${i}.deductor_name`, "Deductor name")}
      ${sel(`tax_paid.tds_others.${i}.section`, "TDS section", optsFrom(TDS_SECTIONS, r.section, "Select…"))}
      ${money(`tax_paid.tds_others.${i}.amount_paid`, "Amount paid/credited")}
      ${sel(`tax_paid.tds_others.${i}.deduction_year`, "Year of deduction", ["2025", "2024", "2023"].map(y => opt(y, y + "-" + String(+y + 1).slice(2), r.deduction_year ?? "2025")).join(""))}
      ${money(`tax_paid.tds_others.${i}.tds_deducted`, "TDS deducted")}
      ${money(`tax_paid.tds_others.${i}.tds_claimed`, "TDS claimed this year")}
    </div>`));
  h += section("TDS on rent (Schedule TDS-3, Form 16C) & TCS",
    `<h3>TDS-3</h3>` + repeater("tax_paid.tds3", "Add tenant TDS (194IB)", (r, i) => `<div class="grid c4">
      ${txt(`tax_paid.tds3.${i}.tenant_pan`, "Tenant PAN/TAN")}
      ${txt(`tax_paid.tds3.${i}.tenant_name`, "Tenant name")}
      ${money(`tax_paid.tds3.${i}.gross_receipt`, "Gross rent")}
      ${sel(`tax_paid.tds3.${i}.deduction_year`, "Year", opt("2025", "2025-26", r.deduction_year ?? "2025") + opt("2024", "2024-25", r.deduction_year ?? "2025"))}
      ${money(`tax_paid.tds3.${i}.tds_deducted`, "TDS deducted")}
      ${money(`tax_paid.tds3.${i}.tds_claimed`, "TDS claimed")}
    </div>`) +
    `<h3>TCS collected (Schedule TCS)</h3>` + repeater("tax_paid.tcs", "Add collector", (r, i) => `<div class="grid c3">
      ${txt(`tax_paid.tcs.${i}.tan`, "Collector TAN")}
      ${txt(`tax_paid.tcs.${i}.collector_name`, "Collector name")}
      ${money(`tax_paid.tcs.${i}.amount_subject`, "Amount received/debited")}
      ${sel(`tax_paid.tcs.${i}.collection_year`, "Year", opt("2025", "2025-26", r.collection_year ?? "2025") + opt("2024", "2024-25", r.collection_year ?? "2025"))}
      ${money(`tax_paid.tcs.${i}.tcs_collected`, "TCS collected")}
      ${money(`tax_paid.tcs.${i}.tcs_claimed`, "TCS claimed this year")}
    </div>`));
  h += section("Advance tax / self-assessment tax challans (Schedule IT)",
    repeater("tax_paid.challans", "Add challan", (r, i) => `<div class="grid c4">
      ${txt(`tax_paid.challans.${i}.bsr_code`, "BSR code", { ph: "7 digits" })}
      ${dateInp(`tax_paid.challans.${i}.date`, "Date of deposit", "Within FY → advance tax; after 31-03-2026 → self-assessment tax")}
      ${txt(`tax_paid.challans.${i}.challan_no`, "Challan serial no.", { max: 5 })}
      ${money(`tax_paid.challans.${i}.amount`, "Amount")}
    </div>`) +
    `<hr class="sep"><div class="grid c2">
      ${money("relief_89", "Relief u/s 89 (arrears/advance salary)", "Category D: furnish Form 10E on the portal")}
    </div>
    <h3>Interest / fee override (advanced)</h3>
    <p class="hint">Computed automatically by default — interest u/s 234A/234B/234C from your challan dates,
    fee u/s 234F (₹1,000/₹5,000) after the due date, fee u/s 234-I for revised returns after 31-12-2026.
    Enter a value only to override.</p>
    <div class="grid c4">
      ${money("interest_override.IntrstPayUs234A", "234A override")}
      ${money("interest_override.IntrstPayUs234B", "234B override")}
      ${money("interest_override.IntrstPayUs234C", "234C override")}
      ${money("interest_override.LateFilingFee234F", "234F fee override")}
    </div>`);
  host.appendChild(el(`<div>${h}</div>`));
}

/* -------------------------------------------------------------- tab: bank */
function renderBank(host) {
  let h = section("Bank accounts (refund account mandatory)", `
    <div class="repeater">${(state.bank_accounts || []).map((b, i) => `
      <div class="rep-item">
        <div class="rep-head"><span class="rep-title">Account ${i + 1}</span>
          <span>${state.bank_accounts.length > 1 ? `<button type="button" class="rep-del" data-del-rep="bank_accounts" data-idx="${i}">Remove</button>` : ""}
          <label style="font-size:.8rem"><input type="radio" name="refund_acct" data-action="refund-acct" data-noauto data-idx="${i}" ${b.use_for_refund ? "checked" : ""}> use for refund</label></span></div>
        <div class="grid c3">
          ${txt(`bank_accounts.${i}.ifsc`, "IFSC", { ph: "SBIN0000123", req: 1 })}
          ${txt(`bank_accounts.${i}.bank_name`, "Bank name", { req: 1 })}
          ${txt(`bank_accounts.${i}.account_no`, "Account number", { req: 1 })}
          ${sel(`bank_accounts.${i}.account_type`, "Account type",
            opt("SB", "Savings", b.account_type) + opt("CA", "Current", b.account_type) + opt("CC", "Cash credit", b.account_type) + opt("OD", "Overdraft", b.account_type) + opt("NRO", "NRO", b.account_type) + opt("OTH", "Other", b.account_type))}
        </div>
      </div>`).join("")}
      <button type="button" class="btn-add" data-add-rep="bank_accounts">+ Add bank account</button></div>`);
  h += section("Verification", `<div class="grid c3">
      ${txt("verification.name", "Name (declarant)", { req: 1 })}
      ${txt("verification.father_name", "Father's name", { req: 1 })}
      ${txt("verification.pan", "PAN (declarant)", { ph: "defaults to assessee PAN" })}
      ${txt("verification.place", "Place of filing", { req: 1 })}
      ${sel("verification.capacity", "Capacity", opt("S", "Self", state.verification.capacity) + opt("R", "Representative", state.verification.capacity))}
    </div>
    <p class="hint">By downloading the JSON you declare that the particulars are true and correct to the
    best of your knowledge and belief, as required by the verification section of the return.</p>`);
  host.appendChild(el(`<div>${h}</div>`));
}

/* ------------------------------------------------------------ tab: review */
async function renderReview(host) {
  const form = await effectiveForm();
  lastDetect = lastDetect || await postJSON("/api/forms/detect", state);

  /* form banner */
  host.appendChild(el(section("Return form", `
    <div class="inline" style="gap:.8rem;flex-wrap:wrap">
      <span class="badge green" style="font-size:1rem;padding:.35rem .8rem">📄 ${esc(effFormLabel())}</span>
      <span class="muted">${state.form === "auto"
        ? `auto-detected — ${esc(lastDetect?.why || "")}`
        : `manually selected on the Start tab (auto-detection suggests ${esc(lastDetect?.recommended || "")})`}</span>
    </div>
    <div class="inline" style="margin-top:.4rem">
      ${(state.losses_cf && Object.values(state.losses_cf).some(v => Number(v)))
        ? `<span class="badge red">Carry-forward losses entered — verify the CFL set-off on the portal after upload</span>` : ""}
    </div>`)));

  /* mandatory-field completeness map */
  const failing = validateAllSteps();
  if (failing.length) {
    host.appendChild(el(section("Mandatory fields — action needed", `
      <div class="notice warn"><b>${failing.reduce((n, f) => n + f.errors.length, 0)} required field(s)
      still incomplete.</b> The return cannot be generated until these are filled:</div>
      <ul class="checklist err">${failing.map(f =>
        `<li>Step ${f.step + 1} · <b>${esc(f.title)}</b>: ${esc([...new Set(f.errors.map(e2 => e2.msg))].slice(0, 4).join(" • "))}${f.errors.length > 4 ? ` • and ${f.errors.length - 4} more` : ""}</li>`).join("")}
      </ul>`)));
  } else {
    host.appendChild(el(section("Mandatory fields", `
      <p class="lp-ok">✓ All mandatory fields (per the official ${esc(effFormLabel())} schema) are filled.</p>`)));
  }

  /* regime selector card */
  const s = await postSummaryV2(form);
  const n = s?.new, o = s?.old;
  const rec = s?.recommended;
  const curOld = !!state.filing.opt_out_new_regime;
  const card = (key, label, d, sel_) => `
    <div class="regime-card ${sel_ ? "selected" : ""}" data-action="pick-regime" data-regime="${key}">
      ${key === rec ? '<span class="rc-recommend badge green">Recommended — lower tax</span>' : ""}
      <div class="rc-title">${label}</div>
      <div class="rc-amt">${inr(d?.total_tax_interest ?? 0)}</div>
      <p class="muted">Total tax + interest &nbsp;·&nbsp; taxable income ${inr(d?.total_income ?? 0)}</p>
      <table class="summary">
        <tr><td>Tax on total income</td><td class="num">${inr(d?.tax_on_total_income)}</td></tr>
        <tr><td>Rebate u/s 87A</td><td class="num">− ${inr(d?.rebate_87a)}</td></tr>
        <tr><td>H&E cess (4%)</td><td class="num">${inr(d?.cess)}</td></tr>
        <tr><td>Interest + fee</td><td class="num">${inr(d?.interest_fee)}</td></tr>
        <tr><td>Taxes paid (TDS/TCS/AT/SAT)</td><td class="num">− ${inr(d?.taxes_paid?.TotalTaxesPaid)}</td></tr>
        <tr class="total"><td>${(d?.refund_due ?? 0) > 0 ? "Refund due" : "Balance payable"}</td>
            <td class="num">${inr((d?.refund_due ?? 0) > 0 ? d.refund_due : d?.balance_payable)}</td></tr>
      </table>
    </div>`;
  host.appendChild(el(section("Pick your tax regime", `
    <div class="regime-cards">
      ${card("new", "New regime (default u/s 115BAC)", n, !curOld)}
      ${card("old", "Old regime (with deductions)", o, curOld)}
    </div>
    ${(n?.warnings?.length || o?.warnings?.length) ? `<div class="notice warn">${[...(n?.warnings || []), ...(o?.warnings || [])].map(w => `<div>• ${esc(w)}</div>`).join("")}</div>` : ""}`)));

  /* slab table */
  const chosen = curOld ? o : n;
  if (chosen?.slab_breakup?.length) {
    host.appendChild(el(section("Slab-wise breakup (chosen regime)", `
      <table class="summary"><tr class="head"><th>Slab</th><th>Rate</th><th class="num">Taxable</th><th class="num">Tax</th></tr>
      ${chosen.slab_breakup.map(r => `<tr><td>${esc(r.slab)}</td><td>${(r.rate * 100).toFixed(0)}%</td><td class="num">${inr(r.taxable)}</td><td class="num">${inr(r.tax)}</td></tr>`).join("")}
      </table>`)));
  }

  /* validation */
  const box = el(section("Validation", `<p class="muted">Checking JSON-schema compliance and the CBDT
    Category A/B/D rules…</p>`));
  host.appendChild(box);
  const gen = await postGenerateForm();
  if (gen.error) {
    const missing = validateAllSteps();
    box.innerHTML = `<h2>CBDT validation</h2>
      <div class="issue a"><span class="tag">ERR</span><div>Could not validate the return:
      <b>${esc(gen.error)}</b>${gen.field ? `<div class="path">Field: ${esc(gen.field)}</div>` : ""}
      <div class="path">${missing.length
        ? `Incomplete mandatory fields → ${missing.map(m => `Step ${m.step + 1} (${m.title}): ${m.errors.length}`).join(", ")}. Use the wizard steps above — missing fields are listed at the top of each step.`
        : "Check the highlighted field and try again."}</div></div></div>`;
    if (gen.path) markServerField(gen.path);
  } else {
    const issues = gen.issues || [];
    const aCnt = issues.filter(i => i.category === "A" || i.category === "SCHEMA").length;
    let html;
    if (!issues.length) html = `<div class="okbox">✓ The return passes the official JSON schema and all implemented CBDT rules — ready to upload.</div>`;
    else {
      html = issues.map(i => `
        <div class="issue ${i.category.toLowerCase()}"><span class="tag">${esc(i.category)}</span>
        <div><b>${esc(i.rule)}</b> — ${esc(i.message)}${i.path ? `<div class="path">${esc(i.path)}</div>` : ""}</div></div>`).join("");
      if (aCnt === 0) html = `<div class="notice">No blocking (Category A / schema) defects. Review the warnings below — the return can still be uploaded.</div>` + html;
    }
    box.innerHTML = `<h2>CBDT validation</h2>${html}`;
  }

  /* detailed computation view */
  const compBox = el(section("Detailed computation", `<p class="muted">Computing…</p>`));
  host.appendChild(compBox);
  const comp = await postComputationForm(form);
  compBox.innerHTML = `<h2>Computation of total income &amp; tax — in depth</h2>` +
    (comp.error ? `<div class="issue a"><span class="tag">ERR</span><div>${esc(comp.error)}</div></div>`
                : renderComputation(comp));
  const btns = document.createElement("div");
  btns.style.cssText = "display:flex;gap:.5rem;margin-top:.8rem;flex-wrap:wrap";
  btns.innerHTML = `
    <button class="btn primary small" data-action="download-pdf" type="button">📄 Download PDF</button>
    <button class="btn ghost small" data-action="download-docx" type="button">📝 Download Word (.docx)</button>`;
  compBox.appendChild(btns);

  /* final action */
  host.appendChild(el(section("Generate & download", `
    <p class="hint">The JSON below conforms to the official <b>${esc(effFormLabel())}</b> schema
    for AY 2026-27 (Ver1.x). Upload it on the e-Filing portal under “File Income Tax Return →
    Upload JSON”.</p>
    <div class="inline">
      <button class="btn primary" data-action="download-file" type="button">⬇ Download ${esc(effFormLabel())} JSON</button>
      <button class="btn ghost" data-action="preview-json" type="button">Preview JSON</button>
    </div>`)));
}

/* ------------------------------------------------ detailed computation */
function renderComputation(comp) {
  const moneyCell = v => (v === null || v === undefined) ? "" : `<td class="num">${inr(v)}</td>`;
  const rows = list => (list || []).map(x =>
    `<tr><td>${esc(x.label)}</td><td class="muted">${esc(x.note || "")}</td>${moneyCell(x.amount)}</tr>`).join("");
  const tbl = (rowsHtml, head = `<tr class="head"><th>Particulars</th><th>Working / basis</th><th class="num">Amount</th></tr>`) =>
    `<table class="summary">${head}${rowsHtml}</table>`;

  const annexures = r => {
    const out = [];
    const sw = r.salary_workings;
    if (sw) {
      const crows = (sw.components || []).map(c =>
        `<tr><td>${esc(c.label || c.code)}</td><td class="muted">${c.exempt ? `exempt u/s ${c.section}: ${inr(c.exempt)}` : ""}</td>${moneyCell((c.amount || 0) - (c.exempt || 0))}</tr>`).join("")
        + (sw.perquisites || []).map(p =>
          `<tr><td>Perquisite — ${esc(p.label)}</td><td class="muted">taxable u/s 17(2)</td>${moneyCell(p.amount)}</tr>`).join("")
        + (sw.profits || []).map(p =>
          `<tr><td>Profits in lieu — ${esc(p.label)}</td><td class="muted">taxable u/s 17(3)</td>${moneyCell(p.amount)}</tr>`).join("");
      if (crows) out.push(`<h3>Annexure — Salary break-up as per the salary slip</h3>` + tbl(crows,
        `<tr class="head"><th>Component</th><th>Working / basis</th><th class="num">Taxable</th></tr>`));
      if (sw.hra_rows?.length) out.push(`<h3>Annexure — HRA exemption u/s 10(13A) (least of three)</h3>` + tbl(rows(sw.hra_rows)));
      for (const tw of sw.terminal || []) {
        out.push(`<h3>Annexure — Terminal benefit: ${esc(tw.name)}</h3>` + tbl(
          (tw.lines || []).map(([l, v]) => `<tr><td>${esc(l)}</td><td></td>${moneyCell(v)}</tr>`).join("") +
          `<tr><td>Received</td><td></td>${moneyCell(tw.received)}</tr>` +
          `<tr><td>Less: exempt</td><td class="muted">${esc(tw.limit || "")}</td>${moneyCell(-(tw.exempt || 0))}</tr>` +
          `<tr class="total"><td>Taxable (added to salary)</td><td></td>${moneyCell(tw.taxable)}</tr>`));
      }
    }
    const cg = r.capital_gains;
    if (cg) {
      out.push(`<h3>Annexure — Capital gains computation</h3>` + tbl(rows(cg.rows)));
      if (cg.details?.length) out.push(`<h3>Annexure — Transaction-wise gains</h3>` + tbl(rows(cg.details)));
      if (cg.quarterly?.length) out.push(`<h3>Annexure — Accrual across the year (sec 234C)</h3>` + tbl(rows(cg.quarterly)));
    }
    if (r.business?.financials_rows?.length) {
      out.push(`<h3>Annexure — No-books financial particulars (sec 44AA(2))</h3>` + tbl(rows(r.business.financials_rows)));
    }
    if (r.form10e_rows?.length) {
      const fm = r.form10e_meta || {};
      out.push(`<h3>Annexure — Relief u/s 89 on salary arrears (Form 10E)</h3>` + tbl(
        r.form10e_rows.map(x =>
          `<tr><td>Arrear salary of FY ${esc(x.fy)}</td><td class="muted">arrear ${inr(x.amount)} · income that year ${inr(x.income_that_fy)}</td>${moneyCell(x.delta)}</tr>`).join("") +
        `<tr><td>Extra tax this year due to arrears</td><td></td>${moneyCell(fm.current_delta)}</tr>` +
        `<tr class="total"><td><b>Relief u/s 89(1)</b></td><td></td>${moneyCell(fm.relief89)}</tr>`));
    }
    return out.join("");
  };

  const per = (r, title) => {
    const salaryRows = rows(r.salary.rows) +
      (r.salary.exempt_rows || []).map(x =>
        `<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;${esc(x.label)}</td><td class="muted">exempt allowance u/s 10</td>${moneyCell(x.amount)}</tr>`).join("") +
      `<tr class="total"><td>Income from salaries</td><td></td>${moneyCell(r.salary.income)}</tr>`;
    const hpRows = (r.house_properties || []).map(hp =>
      `<tr><td colspan="3"><b>${esc(hp.title)}</b></td></tr>` + rows(hp.rows) +
      `<tr class="total"><td>Income — ${esc(hp.title)}</td><td></td>${moneyCell(hp.income)}</tr>`).join("") +
      (r.house_properties?.length ? `<tr class="total"><td>Total house property income</td><td class="muted">${esc(r.hp_note || "")}</td>${moneyCell(r.hp_income)}</tr>` : "");
    const osRows = rows(r.other_sources.rows) +
      `<tr class="total"><td>Income from other sources</td><td></td>${moneyCell(r.other_sources.income)}</tr>`;
    const bizRows = r.business ?
      `<tr><td colspan="3"><b>Income from business / profession (presumptive scheme)</b></td></tr>` +
      rows(r.business.rows) : "";
    const cgRow = r.capital_gains ?
      `<tr class="total"><td>Income from capital gains (annexure below)</td>
       <td class="muted">${esc(r.capital_gains.note || "")}</td>${moneyCell(r.capital_gains.total)}</tr>` : "";
    const gtiHtml = (r.gti_rows?.length) ? rows(r.gti_rows) :
      `<tr><td>Gross total income (before LTCG 112A)</td><td></td>${moneyCell(r.gti_excl_ltcg)}</tr>` +
      `<tr><td>LTCG u/s 112A (exempt up to ₹1,25,000)</td><td class="muted">${esc(r.ltcg_note || "")}</td>${moneyCell(r.ltcg)}</tr>` +
      `<tr class="total"><td>Gross total income (incl. LTCG)</td><td></td>${moneyCell(r.gti_incl_ltcg)}</tr>`;
    const partA = `<h3>PART A — Computation of total income</h3>` + tbl(
      salaryRows + hpRows + bizRows + osRows + cgRow + gtiHtml);
    const via = r.via_rows?.length ? (`<h3>Deductions under Chapter VI-A</h3>` +
      (r.via_caption ? `<p class="notice warn">${esc(r.via_caption)}</p>` : "") +
      `<table class="summary"><tr class="head"><th>Section</th><th>Details</th><th class="num">Claimed</th><th class="num">Allowable</th></tr>` +
      r.via_rows.map(x => `<tr><td><b>${esc(x.section)}</b></td><td class="muted">${esc(x.note)}</td><td class="num">${inr(x.claimed)}</td><td class="num"><b>${inr(x.eligible)}</b></td></tr>`).join("") +
      `<tr class="total"><td colspan="2">Total VI-A deductions</td><td></td><td class="num">${inr(r.via_total)}</td></tr></table>`) : "";
    const tiRow = tbl(`<tr class="total"><td><b>TOTAL INCOME</b></td><td class="muted">${esc(r.total_income_note)}</td>${moneyCell(r.total_income)}</tr>`);

    const slabs = r.slab_rows?.length ? `<h3>Slab-wise tax</h3>` +
      `<table class="summary"><tr class="head"><th>Slab (₹)</th><th>Rate × taxable</th><th class="num">Tax</th></tr>` +
      r.slab_rows.map(x => `<tr><td>${esc(x.slab)}</td><td>${(x.rate * 100).toFixed(0)}% × ${inr(x.taxable)}</td><td class="num">${inr(x.tax)}</td></tr>`).join("") +
      `</table>` : "";
    const interestRows = (r.interest || []).map(it =>
      `<tr><td>${esc(it.label)}</td><td class="muted">${esc(it.note || "")}</td>${moneyCell(it.amount)}</tr>` +
      (it.instalments || []).map(inst =>
        `<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;instalment due ${esc(inst.due)} (${inst.required_pct * 100}%)</td>
         <td class="muted">required ${inr(inst.required)}, paid ${inr(inst.paid)}, 1% × ${inst.months} mo.</td>
         ${moneyCell(inst.amount)}</tr>`).join("")).join("");
    const partB = `<h3>PART B — Computation of tax liability</h3>` + slabs + tbl(
      `<tr><td>Tax at slab rates (on normal-rate income)</td><td></td>${moneyCell(r.slab_tax)}</tr>` +
      (r.special_tax ? `<tr><td>Tax at special rates (111A @ 20% / LTCG @ 12.5% or 20%)</td><td class="muted">${esc(r.special_tax_note || "")}</td>${moneyCell(r.special_tax)}</tr>` : "") +
      `<tr><td>Tax on LTCG u/s 112A</td><td></td>${moneyCell(r.ltcg_tax)}</tr>` +
      `<tr><td>Surcharge</td><td class="muted">${esc(r.surcharge_note)}</td>${moneyCell(r.surcharge)}</tr>` +
      `<tr class="total"><td>Tax payable on total income</td><td></td>${moneyCell(r.tax_payable)}</tr>` +
      `<tr><td>Less: rebate u/s 87A</td><td class="muted">${esc(r.rebate_note)}</td>${moneyCell(-r.rebate)}</tr>` +
      `<tr><td>Tax after rebate</td><td></td>${moneyCell(r.tax_after_rebate)}</tr>` +
      `<tr><td>Add: H&E cess @ 4%</td><td class="muted">${esc(r.cess_note)}</td>${moneyCell(r.cess)}</tr>` +
      `<tr class="total"><td>Gross tax liability</td><td></td>${moneyCell(r.gross_tax_liability)}</tr>` +
      `<tr><td>Less: relief u/s 89</td><td class="muted">${esc(r.relief89_note || "")}</td>${moneyCell(-r.relief89)}</tr>` +
      `<tr><td>Net tax liability (288B rounded)</td><td></td>${moneyCell(r.net_tax_liability)}</tr>` +
      interestRows +
      `<tr class="total"><td>Total tax + interest &amp; fee</td><td></td>${moneyCell(r.tot_tax_plus_interest)}</tr>`);
    const tp = r.taxes_paid;
    const partC = `<h3>PART C — Taxes paid &amp; refund / payable</h3>` + tbl(
      `<tr><td>TDS (TDS-1+2+3)</td><td></td>${moneyCell(tp.tds)}</tr>` +
      `<tr><td>TCS</td><td></td>${moneyCell(tp.tcs)}</tr>` +
      `<tr><td>Advance tax</td><td class="muted">challans within FY 2025-26</td>${moneyCell(tp.advance_tax)}</tr>` +
      `<tr><td>Self-assessment tax</td><td class="muted">challans after 31-03-2026</td>${moneyCell(tp.self_assessment_tax)}</tr>` +
      `<tr class="total"><td>Total taxes paid</td><td></td>${moneyCell(tp.total)}</tr>` +
      `<tr class="total"><td>Balance payable</td><td></td>${moneyCell(r.balance_payable)}</tr>` +
      `<tr class="total"><td>Refund due</td><td></td>${moneyCell(r.refund_due)}</tr>`) +
      (r.result_note ? `<p class="notice">${esc(r.result_note)}</p>` : "");
    return `<details ${title.includes("New") === (comp.comparison.recommended === "new") ? "open" : ""}>
      <summary>${esc(title)} — total tax: <b>${inr(r.tot_tax_plus_interest)}</b></summary>
      ${r.age_note ? `<p class="muted">${esc(r.age_note)}</p>` : ""}
      ${partA}${via}${tiRow}${annexures(r)}${partB}${partC}</details>`;
  };

  return `
    <p class="muted">Full statutory working for both regimes — identical to the PDF / Word export below.</p>
    ${per(comp.new, comp.new.regime_title)}
    ${per(comp.old, comp.old.regime_title)}`;
}

/* ---------------------------------------------------- salary relief-89 view */
async function showRelief89() {
  const res = await postJSON("/api/salary/relief89", { data: state });
  if (res.error) { showModal("Relief u/s 89", `<p>${esc(res.error)}</p>`); return; }
  if (!res.relief89) {
    showModal("Relief u/s 89 (Form 10E)", `<p>${esc(res.message || "No relief available for the arrears entered.")}</p>`);
    return;
  }
  const rows = (res.rows || []).map(r => `<tr>
      <td>${esc(r.fy)}</td><td class="num">${inr(r.amount)}</td>
      <td class="num">${inr(r.income_that_fy)}</td>
      <td class="num">${inr(r.tax_with - r.tax_without)}</td></tr>`).join("");
  showModal("Relief u/s 89 — Form-10E working", `
    <table class="summary"><tr class="head"><th>Arrear of FY</th><th class="num">Arrear</th>
    <th class="num">Income that FY</th><th class="num">Extra tax that FY</th></tr>${rows}
    <tr><td>Extra tax this year (arrears included)</td><td colspan="2"></td>
        <td class="num">${inr(res.current_delta)}</td></tr>
    <tr class="total"><td><b>Relief u/s 89(1)</b> — auto-applied in the return</td><td colspan="2"></td>
        <td class="num"><b>${inr(res.relief89)}</b></td></tr></table>
    <p class="hint">Form 10E must be furnished on the portal before/with the return.</p>`);
}

/* ------------------------------------------------- capital-gain xlsx actions */
async function cgExportXlsx() {
  const res = await fetch("/api/cg/xlsx", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
  if (!res.ok) { showModal("Export failed", `<p>Server error ${res.status}</p>`); return; }
  const blob = await res.blob();
  const a = el(`<a href="${URL.createObjectURL(blob)}" download="capital_gains_ay2026-27.xlsx"></a>`);
  document.body.appendChild(a); a.click(); a.remove();
}
async function cgDownloadTemplate() {
  const res = await fetch("/api/cg/template");
  const tpl = await res.json();
  const blob = new Blob([JSON.stringify(tpl, null, 2)], { type: "application/json" });
  const a = el(`<a href="${URL.createObjectURL(blob)}" download="capital_gains_template.json"></a>`);
  document.body.appendChild(a); a.click(); a.remove();
}
async function handleCgImportFile(input) {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  const bytes = await file.arrayBuffer();
  const res = await fetch("/api/cg/import-xlsx", {
    method: "POST", headers: { "Content-Type": "application/octet-stream", "X-Filename": file.name },
    body: bytes });
  const body = await res.json();
  if (!res.ok || body.error) { showModal("Capital-gain import failed", `<p>${esc(body.error || `Server error ${res.status}`)}</p>`); return; }
  state.capital_gains = body.capital_gains || state.capital_gains;
  const counts = Object.entries(body.counts || {}).map(([k, v]) => `${k}: ${v}`).join(" · ") || "no rows";
  render();
  showModal("Capital gains imported ✓", `<p>${esc(counts)} — the grids on the Capital gains tab now hold the imported rows.</p>`);
}

/* ---------------------------------------------------------- import actions */
function panel(id) { return document.getElementById(id); }
function statusBadge(id, text, ok) {
  const b = panel(id); if (!b) return;
  b.textContent = text; b.className = `badge ${ok ? "green" : "red"}`;
}

async function uploadBinary(url, file, { apply } = {}) {
  const fd = new FormData();
  fd.append("file", file, file.name);
  fd.append("data", JSON.stringify(state));
  if (apply !== undefined) fd.append("apply", apply ? "1" : "0");
  const res = await fetch(url, { method: "POST", body: fd });
  let body = {};
  try { body = await res.json(); } catch (e) { /* non-JSON */ }
  return { ok: res.ok && !body.error, body };
}

async function handleItrJsonFile(input) {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  statusBadge("imp-itr-status", "parsing…", true);
  const { ok, body } = await uploadBinary("/api/import/itr-json", file, { apply: true });
  if (!ok) { statusBadge("imp-itr-status", "failed", false);
    showModal("ITR JSON import failed", `<p>${esc(body.error || "could not parse the file")}</p>`); return; }
  statusBadge("imp-itr-status", `✓ ${body.summary.form || "ITR"} AY ${body.summary.assessment_year || ""}`, true);
  lastPrevSummary = body.summary;
  const s = body.summary;
  const losses = Object.entries(s.losses || {}).filter(([, v]) => Number(v));
  panel("imp-itr-panel").innerHTML = `
    <table class="summary" style="margin-top:.7rem">
      <tr><td>Assessment year / form</td><td class="num"><b>${esc(s.assessment_year || "?")} · ${esc(s.form || "?")}</b></td></tr>
      <tr><td>Total income / tax paid</td><td class="num">${inr(s.total_income)} / ${inr(s.total_tax_paid)}</td></tr>
      <tr><td>Prefilled</td><td>${(body.filled_fields || []).map(f => `<span class="badge green">${esc(f)}</span> `).join("") || '<span class="muted">—</span>'}</td></tr>
      <tr><td>Losses carried forward</td><td>${losses.length ? losses.map(([k, v]) => `<span class="badge red">${esc(k)}: ${inr(v)}</span> `).join("") : '<span class="muted">none</span>'}</td></tr>
    </table>
    ${(body.notes || []).length ? `<div class="notice" style="margin-top:.5rem">${body.notes.map(n => `<div>• ${esc(n)}</div>`).join("")}</div>` : ""}
    <p class="hint" style="margin-top:.5rem">Personal/bank details were merged into empty fields only —
    your current entries were not overwritten. Losses appear in the “Brought-forward losses” panel below.</p>`;
  if (body.merged_data) state = Object.assign(blankState(), body.merged_data);
  lastDetect = null;
}

async function handleAisFile(input) {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  statusBadge("imp-ais-status", "parsing…", true);
  const { ok, body } = await uploadBinary("/api/import/ais", file);
  if (!ok) { statusBadge("imp-ais-status", "failed", false);
    showModal("AIS import failed", `<p>${esc(body.error || "could not parse the file")}</p>`); return; }
  lastAis = body;
  statusBadge("imp-ais-status",
    `✓ ${body.entry_count} entries${body.password_detected ? " · 🔓 password auto-detected (PAN+DOB)" : ""}`, true);
  const cmp = (body.comparison || []).map(row => `<tr>
      <td>${esc(row.head)}</td><td class="num">${inr(row.as_per_return)}</td>
      <td class="num">${inr(row.as_per_ais)}</td>
      <td><span class="badge ${row.status === "matched" ? "green" : row.status.startsWith("AIS higher") ? "red" : ""}">${esc(row.status)}</span></td></tr>`).join("");
  const cands = (body.income_candidates || []).map((c, i) => `<tr>
      <td>${esc(c.description || c.category)}</td><td class="num">${inr(c.amount)}</td>
      <td><button class="btn small primary" type="button" data-action="ais-add-income" data-idx="${i}">＋ Add income</button></td></tr>`).join("");
  const tc = body.tax_credits || {};
  const tcCount = (tc.tds_on_salary || []).length + (tc.tds_others || []).length +
                  (tc.tds3 || []).length + (tc.tcs || []).length + (tc.challans || []).length;
  panel("imp-ais-panel").innerHTML = `
    ${cmp ? `<h3>As per AIS vs as per this return</h3>
      <table class="summary"><tr class="head"><th>Head</th><th class="num">In return</th><th class="num">In AIS</th><th>Status</th></tr>${cmp}</table>` : ""}
    ${cands ? `<h3 style="margin-top:.8rem">Income found in AIS — click to add</h3>
      <table class="summary"><tr class="head"><th>Source</th><th class="num">Amount</th><th></th></tr>${cands}</table>` : ""}
    ${tcCount ? `<div class="inline" style="margin-top:.7rem">
      <button class="btn small primary" type="button" data-action="ais-apply-credits">⇩ Apply ${tcCount} tax-credit entries to Taxes paid</button></div>` : ""}`;
}

async function aisAddIncome(idx) {
  if (!lastAis) return;
  const cand = (lastAis.income_candidates || [])[+idx];
  if (!cand) return;
  const res = await postJSON("/api/import/ais/apply",
    { data: state, candidates: lastAis.income_candidates, income_ids: [cand.id] });
  if (res.error) { showModal("Could not add the income", `<p>${esc(res.error)}</p>`); return; }
  state = Object.assign(blankState(), res.merged_data);
  cand._added = true;
  lastDetect = null;
  render();
  showModal("Income added ✓", `<p>${esc(cand.description || cand.category)} — ${inr(cand.amount)} is now on the
    ${cand.category === "salary" ? "Salary" : "Other income"} tab. Added rows are tagged “(as per AIS)”.</p>`);
}

async function aisApplyCredits() {
  if (!lastAis || !lastAis.tax_credits) return;
  const res = await postJSON("/api/import/ais/apply",
    { data: state, tax_credits: lastAis.tax_credits });
  if (res.error) { showModal("Could not apply the credits", `<p>${esc(res.error)}</p>`); return; }
  state = Object.assign(blankState(), res.merged_data);
  lastDetect = null;
  render();
  showModal("Tax credits applied ✓", `<p>TDS/TCS/challan entries from AIS were merged into the Taxes paid step (duplicates skipped).</p>`);
}

async function handle26asFile(input) {
  const file = input.files?.[0];
  input.value = "";
  if (!file) return;
  statusBadge("imp-26as-status", "parsing…", true);
  const { ok, body } = await uploadBinary("/api/import/26as", file);
  if (!ok) { statusBadge("imp-26as-status", "failed", false);
    showModal("26AS import failed", `<p>${esc(body.error || "could not parse the file")}</p>`); return; }
  last26as = body;
  const c = body.counts || {};
  statusBadge("imp-26as-status",
    `✓ TDS ${(c.tds_salary || 0) + (c.tds_others || 0) + (c.tds3 || 0)} · TCS ${c.tcs || 0} · challans ${c.challans || 0}${body.password_detected ? " · 🔓 DOB password auto-detected" : ""}`, true);
  panel("imp-26as-panel").innerHTML = `
    <table class="summary" style="margin-top:.7rem">
      <tr><td>PART-A1 — TDS on salary (192)</td><td class="num">${c.tds_salary || 0}</td></tr>
      <tr><td>PART-A2 — TDS other than salary</td><td class="num">${c.tds_others || 0}</td></tr>
      <tr><td>TDS u/s 194-IA/194-IB (TDS-3)</td><td class="num">${c.tds3 || 0}</td></tr>
      <tr><td>PART-B — TCS</td><td class="num">${c.tcs || 0}</td></tr>
      <tr><td>Advance / self-assessment challans</td><td class="num">${c.challans || 0}</td></tr>
    </table>
    <div class="inline" style="margin-top:.7rem">
      <button class="btn small primary" type="button" data-action-26as-apply>⇩ Apply to Taxes paid (dedupe)</button>
    </div>
    <p class="hint">Rows are de-duplicated by deductor/section/amount, so re-importing is safe.</p>`;
  panel("imp-26as-panel").querySelector("[data-action-26as-apply]").addEventListener("click", async () => {
    const res = await uploadBinary("/api/import/26as", file, { apply: true });
    if (!res.ok) { showModal("Apply failed", `<p>${esc(res.body.error || "")}</p>`); return; }
    state = Object.assign(blankState(), res.body.merged_data);
    lastDetect = null;
    render();
    showModal("26AS applied ✓", `<p>TDS/TCS/challans from Form 26AS are now in the Taxes paid step.</p>`);
  });
}

async function yoyCompare() {
  if (!lastPrevSummary) { showModal("Year-on-year", `<p>Import a previous-year ITR JSON first.</p>`); return; }
  const res = await postJSON("/api/yoy", { data: state, prev_summary: lastPrevSummary });
  if (res.error) { showModal("Year-on-year", `<p>${esc(res.error)}</p>`); return; }
  const rows = (res.rows || []).map(r => {
    const arrow = r.delta > 0 ? "▲" : r.delta < 0 ? "▼" : "·";
    return `<tr><td>${esc(r.head)}</td><td class="num">${inr(r.previous)}</td>
      <td class="num"><b>${inr(r.current)}</b></td>
      <td class="num">${arrow} ${inr(Math.abs(r.delta))}${r.delta_pct !== null && r.delta_pct !== undefined ? ` (${r.delta_pct}%)` : ""}</td></tr>`;
  }).join("");
  panel("yoy-panel").innerHTML = `
    <table class="summary" style="margin-top:.7rem">
      <tr class="head"><th>Head</th><th class="num">AY ${esc(res.prev_year || "PY")}</th>
      <th class="num">AY 2026-27</th><th class="num">Change</th></tr>${rows}</table>
    <p class="hint">Current-year figures are the taxable summaries from the chosen regime.</p>`;
}

/* ------------------------------------------------------------------ APIs */
async function postJSON(url, body) {
  try {
    const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    let payload = null;
    try { payload = await res.json(); } catch (e) { /* non-JSON body */ }
    if (!res.ok) {
      return { error: (payload && payload.error) || `Server error ${res.status}` };
    }
    return payload || { error: "Empty response from server" };
  } catch (e) {
    return { error: `Cannot reach the server: ${e}` };
  }
}
const postSummary = () => postJSON("/api/summary", state).then(r => { if (!r.error) lastSummary = r; return r; });
const postGenerate = () => postJSON("/api/generate", state);
const postComputation = () => postJSON("/api/computation", state);

/* -------------------------------------------------- form-aware plumbing */
async function effectiveForm() {
  if (state.form && state.form !== "auto") return state.form;
  if (!lastDetect) lastDetect = await postJSON("/api/forms/detect", state);
  const rec = lastDetect?.recommended || "ITR1";
  return ["ITR1", "ITR2", "ITR4"].includes(rec) ? rec : "ITR2";
}
function effFormLabel() {
  return (state.form === "auto" ? (lastDetect?.recommended || "ITR1") : state.form)
    .replace("ITR1", "ITR-1").replace("ITR2", "ITR-2").replace("ITR4", "ITR-4");
}
const formEndpoints = f => f === "ITR1"
  ? { generate: "/api/generate", download: "/api/download" }
  : { generate: `/api/${f.toLowerCase()}/generate`, download: `/api/${f.toLowerCase()}/download` };
async function postSummaryV2(f) {
  return postJSON("/api/summary/v2", { data: state, form: f || await effectiveForm() });
}
async function postGenerateForm() {
  const f = await effectiveForm();
  return postJSON(formEndpoints(f).generate, state);
}
async function postComputationForm(f) {
  return postJSON("/api/computation", { data: state, form: f || await effectiveForm() });
}

async function downloadEndpoint(url, fallbackName) {
  const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
  if (!res.ok) {
    let msg = `Server error ${res.status}`;
    try { msg = (await res.json()).error || msg; } catch (e) {}
    showModal("Could not generate the file", `<p>${esc(msg)}</p>`);
    return;
  }
  const blob = await res.blob();
  const cd = res.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename=([^;]+)/);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = m ? m[1] : fallbackName;
  a.click();
  URL.revokeObjectURL(a.href);
}
const downloadFile = async () => {
  const f = await effectiveForm();
  const res = await fetch(formEndpoints(f).download, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(state) });
  if (!res.ok) {
    let msg = `Server error ${res.status}`;
    try { msg = (await res.json()).error || msg; } catch (e) {}
    showModal("Could not generate the JSON", `<p>${esc(msg)}</p>`);
    return;
  }
  const blob = await res.blob();
  const cd = res.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename=([^;]+)/);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = m ? m[1] : `${f}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
};
const downloadPdf = async () => downloadEndpointBody("/api/computation/pdf",
  "ITR_Computation.pdf", async () => ({ data: state, form: await effectiveForm() }));
const downloadDocx = async () => downloadEndpointBody("/api/computation/docx",
  "ITR_Computation.docx", async () => ({ data: state, form: await effectiveForm() }));
async function downloadEndpointBody(url, fallbackName, bodyFn) {
  const body = await bodyFn();
  const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!res.ok) {
    let msg = `Server error ${res.status}`;
    try { msg = (await res.json()).error || msg; } catch (e) {}
    showModal("Could not generate the file", `<p>${esc(msg)}</p>`);
    return;
  }
  const blob = await res.blob();
  const cd = res.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename=([^;]+)/);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = m ? m[1] : fallbackName;
  a.click();
  URL.revokeObjectURL(a.href);
}
function showModal(title, bodyHtml) {
  $("#modal-title").textContent = title;
  $("#modal-body").innerHTML = bodyHtml;
  $("#modal-overlay").classList.remove("hidden");
}

/* ------------------------------------------------------------ live panel */
let lpTimer = null;
function scheduleLive() {
  clearTimeout(lpTimer);
  lpTimer = setTimeout(updateLive, 450);
}
async function updateLive() {
  if (state.form === "auto") lastDetect = await postJSON("/api/forms/detect", state);
  const form = await effectiveForm();
  const s = await postSummaryV2(form);
  const host = $("#live-body");
  if (!s || s.error) {
    host.innerHTML = `<p class="muted">Fill in income details to see a real-time estimate.</p>` +
      (s && s.error ? `<p class="issue a" style="font-size:.75rem"><span class="tag">!</span>${esc(s.error)}</p>` : "");
    if (s && s.path) markServerField(s.path);
    return;
  }

  const chosenKey = state.filing.opt_out_new_regime ? "old" : "new";
  const d = s[chosenKey];

  /* --- main figures of the chosen regime --- */
  const slabRows = (d.slab_breakup || []).map(r =>
    `<div class="lp-row lp-slab"><span>${esc(r.slab)} @ ${(r.rate * 100).toFixed(0)}%</span><b>${inr(r.tax)}</b></div>`).join("");
  const row = (l, v, strong = false) =>
    `<div class="lp-row ${strong ? "lp-strong" : ""}"><span>${l}</span><b>${inr(v)}</b></div>`;
  const lp = `
    <div class="lp-caption">Live estimate — ${chosenKey === "new" ? "New regime" : "Old regime"}
      · <b>${esc(effFormLabel())}</b></div>
    ${row("Gross total income", d.gross_total_income)}
    ${row("Deductions (VI-A)", d.deductions)}
    ${row("Total income", d.total_income, true)}
    ${slabRows}
    ${row("Tax on total income", d.tax_on_total_income)}
    ${row("Rebate u/s 87A", -d.rebate_87a)}
    ${row("H&E cess 4%", d.cess)}
    ${row("Tax liability", d.gross_tax_liability, true)}
    ${row("Interest + fee", d.interest_fee)}
    ${row("TDS/TCS/AT/SAT paid", -d.taxes_paid.TotalTaxesPaid)}
    <div class="lp-result ${d.refund_due > 0 ? "lp-refund" : "lp-payable"}">
      ${d.refund_due > 0 ? `Refund due: ${inr(d.refund_due)}` : d.balance_payable > 0 ? `Balance payable: ${inr(d.balance_payable)}` : "No tax payable"}
    </div>`;

  /* --- regime comparison strip --- */
  const cmpRow = (name, d2, isRec, isChosen) => `
    <div class="lp-cmp-row ${isChosen ? "chosen" : ""}">
      <span>${name}${isRec ? ' <span class="badge green">best</span>' : ""}</span>
      <b>${inr(d2.total_tax_interest)}</b>
    </div>`;
  const cmp = `
    <div class="lp-caption">Both regimes — total tax + interest</div>
    ${cmpRow("New", s.new, s.recommended === "new", chosenKey === "new")}
    ${cmpRow("Old", s.old, s.recommended === "old", chosenKey === "old")}`;

  /* --- issues + warnings detected while typing --- */
  const issues = s.issues || [];
  const blocking = issues.filter(i => i.category === "A" || i.category === "SCHEMA");
  const warns = [...(d.warnings || [])];
  let alertHtml = "";
  if (blocking.length || warns.length) {
    alertHtml = `<div class="lp-caption lp-alert-title">⚠ Needs attention (${blocking.length + warns.length})</div>` +
      blocking.slice(0, 5).map(i => `<div class="lp-issue a"><b>${esc(i.rule)}</b> ${esc(i.message)}</div>`).join("") +
      warns.slice(0, 5).map(w => `<div class="lp-issue w">${esc(w)}</div>`).join("") +
      ((blocking.length + warns.length) > 10 ? `<p class="muted">…and more — see Review tab</p>` : "");
  } else {
    alertHtml = `<div class="lp-ok">✓ No rule violations detected so far</div>`;
  }

  host.innerHTML = lp + cmp + alertHtml;
}

/* ------------------------------------------------------------ rendering */
function render() {
  /* stepper */
  const st = $("#stepper");
  st.innerHTML = "";
  TABS.forEach((t, i) => {
    const b = el(`<button type="button" class="step ${i === currentStep ? "active" : ""} ${i < currentStep ? "done" : ""}">
      <span class="dot">${i < currentStep ? "✓" : i + 1}</span><span>${esc(t.title)}</span></button>`);
    b.addEventListener("click", () => {
      if (i > currentStep && !guardStep(currentStep)) return;
      currentStep = i; render();
    });
    st.appendChild(b);
  });

  const host = $("#tab-host");
  host.innerHTML = "";
  host.appendChild(el(`<div id="step-alerts" class="hidden"></div>`));
  hardFailed = false;
  const res = TABS[currentStep].render(host);
  if (res && res.then) res.catch(e => console.error(e));

  $("#step-indicator").textContent = `Step ${currentStep + 1} of ${TABS.length} — ${TABS[currentStep].title}`;
  $("#btn-prev").disabled = currentStep === 0;
  $("#btn-next").textContent = currentStep === TABS.length - 1 ? "Finish ✓" : "Next →";
  refreshSoftValidation();
  scheduleLive();
}

/* ----------------------------------------------------------------- events */
document.addEventListener("input", e => {
  const t = e.target;
  if (!t.dataset || t.dataset.noauto) return;
  const path = t.dataset.path;
  if (!path) return;
  let v;
  if (t.dataset.kind === "bool") v = t.checked;
  else if (t.dataset.kind === "radio-bool") v = t.value === "true";
  else if (t.dataset.kind === "int") v = t.value === "" || isNaN(parseInt(t.value, 10)) ? "" : parseInt(t.value, 10);
  else if (t.dataset.kind === "float") v = t.value === "" || isNaN(parseFloat(t.value)) ? "" : parseFloat(t.value);
  else v = t.value;
  if (typeof v === "string" && /\.(pan|tan|ifsc)$/.test(path)) { v = v.toUpperCase(); t.value = v; }
  setPath(state, path, v);
  refreshSoftValidation();

  if (path === "salary.has_hra") { if (!v) state.salary.hra = null; render(); return; }
  if (path?.startsWith("salary.hra.")) { updateHraPreview(); }
  if (path === "personal.secondary_address_same" || path.startsWith("filing.") && ["return_section", "is_representative", "seventh_provisio"].includes(path.split(".")[1]) ||
      path.startsWith("deductions.80d.") || /^deductions\.(80u|80dd|80e|80ee|80eea|80eeb)\.enabled$/.test(path) ||
      /^house_properties\.\d+\.co_owned$/.test(path) || /^house_properties\.\d+\.occupancy$/.test(path) ||
      /^salary\.components\.\d+\.code$/.test(path)) {
    render(); return;
  }
  scheduleLive();
});

/* File pickers (capital-gain workbook import + the three import uploads). */
document.addEventListener("change", e => {
  const t = e.target;
  if (t.id === "cg-import-file") { handleCgImportFile(t); return; }
  if (t.id === "imp-itr-json") { handleItrJsonFile(t); return; }
  if (t.id === "imp-ais") { handleAisFile(t); return; }
  if (t.id === "imp-26as") { handle26asFile(t); return; }
});
document.addEventListener("change", e => {
  const t = e.target;
  if (t.matches?.("[data-change-form]")) { state.form = t.value; render(); return; }
  if (!t.dataset?.noauto) return;
  if (t.dataset.action === "regime") {
    state.filing.opt_out_new_regime = t.value === "old";
    render();
  }
});
document.addEventListener("click", async e => {
  const t = e.target.closest("[data-add-rep],[data-del-rep],[data-action]");
  if (!t) return;
  if (t.dataset.addRep) {
    const factories = {
      "house_properties": blankProp,
      "salary.exempt_allowances": () => ({ section: "10(5)", amount: "" }),
      "other_sources": () => ({ nature: "SAV", description: "", amount: "" }),
      "exempt_income": () => ({ category: "AGRI", sub_category: "10(1)", description: "", amount: "" }),
      "deductions.80c.items": () => ({ identification_no: "", amount: "" }),
      "deductions.80g.don_100": blankDonee, "deductions.80g.don_50": blankDonee,
      "deductions.80g.don_100_ql": blankDonee, "deductions.80g.don_50_ql": blankDonee,
      "deductions.80gga.donations": () => ({ clause: "80GGA2b", name: "", pan: "", address: {}, cash: "", other: "" }),
      "deductions.80ggc.donations": () => ({ date: "", other: "", txn_ref: "", ifsc: "", party_name: "", party_pan: "" }),
      "tax_paid.tds_on_salary": () => ({ tan: "", employer_name: "", salary_income: "", tds: "" }),
      "tax_paid.tds_others": () => ({ tan: "", deductor_name: "", section: "94A", amount_paid: "", deduction_year: "2025", tds_deducted: "", tds_claimed: "" }),
      "tax_paid.tds3": () => ({ tenant_pan: "", tenant_name: "", gross_receipt: "", deduction_year: "2025", tds_deducted: "", tds_claimed: "" }),
      "tax_paid.tcs": () => ({ tan: "", collector_name: "", amount_subject: "", collection_year: "2025", tcs_collected: "", tcs_claimed: "" }),
      "tax_paid.challans": () => ({ bsr_code: "", date: "", challan_no: "", amount: "" }),
      "salary.components": () => ({ code: "BASIC", amount: "" }),
      "salary.perquisites": () => ({ code: "RFA", amount: "" }),
      "salary.profits_detail": () => ({ label: "", amount: "" }),
      "salary.arrears": () => ({ fy: "2024-25", amount: "", income_that_fy: "" }),
      "capital_gains.equity_st": () => ({ scrip: "", buy_date: "", sale_date: "", sale_consideration: "", cost_of_acquisition: "", expenses: "" }),
      "capital_gains.equity_lt": () => ({ scrip: "", isin: "", buy_date: "", sale_date: "", sale_consideration: "", cost_of_acquisition: "", fmv_31_01_2018: "", expenses: "" }),
      "capital_gains.land_building": () => ({ description: "", buy_date: "", sale_date: "", sale_consideration: "", cost_of_acquisition: "", cost_of_improvement: "", expenses: "", indexation_option: "auto" }),
      "capital_gains.other_st": () => ({ description: "", buy_date: "", sale_date: "", sale_consideration: "", cost_of_acquisition: "", cost_of_improvement: "", expenses: "" }),
      "capital_gains.other_lt": () => ({ description: "", buy_date: "", sale_date: "", sale_consideration: "", cost_of_acquisition: "", cost_of_improvement: "", expenses: "" }),
      "capital_gains.exemptions": () => ({ section: "54EC", amount: "", date: "" }),
      "presumptive.44ad": () => ({ name: "", code: "09028", gross_turnover: "", turnover_bank: "", turnover_cash: "", turnover_other: "", income_declared: "" }),
      "presumptive.44ada": () => ({ name: "", code: "19003", gross_receipts: "", receipts_bank: "", receipts_cash: "", receipts_other: "", income_declared: "" }),
      "presumptive.44ae": () => ({ reg_no: "", owned: "OWN", tonnage: "", months: "", heavy: false }),
      "bank_accounts": () => ({ ifsc: "", bank_name: "", account_no: "", account_type: "SB", use_for_refund: false }),
    };
    const key = t.dataset.addRep;
    let factory = factories[key];
    if (!factory && /loans$/.test(key)) factory = () => ({ lender_type: "B", lender_name: "", loan_account_no: "", loan_date: "", loan_amount: "", outstanding: "", interest: "" });
    if (!factory && /co_owners$/.test(key)) factory = () => ({ name: "", pan: "", share_pct: "", aadhaar: "" });
    if (!factory && /tenants$/.test(key)) factory = () => ({ name: "", pan: "", aadhaar: "" });
    if (!factory && /deductions\.(80e|80ee|80eea|80eeb)\.loans/.test(key)) factory = () => ({ lender_type: "B", lender_name: "", loan_account_no: "", loan_date: "", loan_amount: "", outstanding: "", interest: "" });
    if (factory) addRepeaterItem(key, factory);
    return;
  }
  if (t.dataset.delRep) { delRepeaterItem(t.dataset.delRep, +t.dataset.idx); return; }

  switch (t.dataset.action) {
    case "pick-regime":
      state.filing.opt_out_new_regime = t.dataset.regime === "old";
      render(); break;
    case "refund-acct":
      state.bank_accounts.forEach((b, i) => b.use_for_refund = (i === +t.dataset.idx));
      render(); break;
    case "download-file": await downloadFile(); break;
    case "download-pdf": await downloadPdf(); break;
    case "download-docx": await downloadDocx(); break;
    case "preview-json": {
      const g = await postGenerateForm();
      if (g.error) { showModal("Could not generate the JSON", `<p>${esc(g.error)}</p>`); break; }
      showModal(`Official ${effFormLabel()} JSON — AY 2026-27`, `<pre>${esc(JSON.stringify(g?.json, null, 2))}</pre>`);
      break;
    }
    case "monthly-open": openMonthlyModal(); break;
    case "monthly-save": await saveMonthlyModal(); break;
    case "relief89": await showRelief89(); break;
    case "cg-export": await cgExportXlsx(); break;
    case "cg-template": await cgDownloadTemplate(); break;
    case "cg-import": $("#cg-import-file").click(); break;
    case "imp-itr-file": $("#imp-itr-json").click(); break;
    case "imp-ais-file": $("#imp-ais").click(); break;
    case "imp-26as-file": $("#imp-26as").click(); break;
    case "ais-add-income": await aisAddIncome(t.dataset.idx); break;
    case "ais-apply-credits": await aisApplyCredits(); break;
    case "yoy-compare": await yoyCompare(); break;
    case "regime": /* handled on change */ break;
  }
});

/* Double-click on the big salary boxes opens the monthly grid. */
document.addEventListener("dblclick", e => {
  if (e.target.closest?.("[data-monthly]")) openMonthlyModal();
});

const blankDonee = () => ({ name: "", pan: "", address: {}, cash: "", other: "", txn_ref: "", ifsc: "", arn: "" });

$("#btn-prev").addEventListener("click", () => { if (currentStep > 0) { currentStep--; render(); } });
$("#btn-next").addEventListener("click", () => {
  if (!guardStep(currentStep)) return;
  if (currentStep < TABS.length - 1) { currentStep++; render(); }
  else finishAndDownload();
});
function finishAndDownload() {
  if (guardAllSteps("download the ITR JSON")) downloadFile();
}
$("#btn-load-sample").addEventListener("click", async () => {
  const s = await (await fetch("/api/sample")).json();
  state = Object.assign(blankState(), s);
  lastDetect = null;
  currentStep = 1;
  render();
});
$("#btn-download-top").addEventListener("click", finishAndDownload);
const bp = $("#btn-pdf-top"); if (bp) bp.addEventListener("click", () => { if (guardAllSteps("download the PDF")) downloadPdf(); });
const bw = $("#btn-word-top"); if (bw) bw.addEventListener("click", () => { if (guardAllSteps("download the Word file")) downloadDocx(); });
$("#modal-close").addEventListener("click", () => $("#modal-overlay").classList.add("hidden"));
$("#modal-overlay").addEventListener("click", e => { if (e.target.id === "modal-overlay") e.currentTarget.classList.add("hidden"); });

/* ------------------------------------------------------------------- boot */
(async function boot() {
  await loadMeta();
  await loadCatalog();
  render();
})();
