/* Node harness: exercises the wizard's mandatory-field mapping (STEP_CHECKS)
 * with a stub DOM.  Run:  node tests/wizard_steps_check.js  → prints PASS/FAIL. */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const appPath = path.join(__dirname, "..", "web", "static", "app.js");
let code = fs.readFileSync(appPath, "utf8");
// expose internals for testing
code += `\n;globalThis.__t = {
  validateStep, validateAllSteps, blankState, TABS,
  set: s => { state = s; }, get: () => state,
};`;

const dummyEl = () => {
  const el2 = {
    innerHTML: "", textContent: "", disabled: false, value: "",
    style: {}, dataset: {}, firstElementChild: null,
    classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
    appendChild() {}, remove() {}, addEventListener() {}, querySelector: () => null,
    querySelectorAll: () => [], scrollIntoView() {}, closest: () => null,
  };
  el2.content = { firstElementChild: el2 };
  return el2;
};
const sandbox = {
  console,
  document: {
    querySelector: () => dummyEl(),
    querySelectorAll: () => [],
    createElement: () => dummyEl(),
    addEventListener() {},
  },
  fetch: async () => ({ json: async () => ({}), ok: true }),
  CSS: { escape: s => s },
  setTimeout, clearTimeout,
  URL: { createObjectURL: () => "", revokeObjectURL() {} },
  Blob: function () {},
};
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

let failures = 0;
function check(name, cond, detail = "") {
  if (cond) console.log(`PASS  ${name}`);
  else { failures++; console.log(`FAIL  ${name}  ${detail}`); }
}

vm.runInContext(code, sandbox, { filename: appPath });

setTimeout(() => {
  const T = sandbox.__t;
  if (!T) { console.log("FAIL  harness could not expose app internals"); process.exit(1); }

  /* ---- blank form: personal + bank steps must flag every mandatory field */
  T.set(T.blankState());
  const stepIdx = id => T.TABS.findIndex(t => t.id === id);
  const pErrs = T.validateStep(stepIdx("personal"));
  const pMsgs = pErrs.map(e => e.msg).join("\n");
  for (const m of ["PAN is required", "Date of birth is required", "Flat/Door/Block no.",
                   "Locality/Area is required", "City/Town/District is required",
                   "State is required", "PIN code is required", "Mobile no. is required",
                   "Email address is required"])
    check(`blank personal flags: ${m}`, pMsgs.includes(m), `got:\n${pMsgs}`);

  const bErrs = T.validateStep(stepIdx("bank"));
  const bMsgs = bErrs.map(e => e.msg).join("\n");
  for (const m of ["IFSC", "bank name", "account number", "name of declarant",
                   "father's name", "place of filing"])
    check(`blank bank flags: ${m}`, bMsgs.includes(m), `got:\n${bMsgs}`);

  check("salary step has no mandatory errors on blank", T.validateStep(stepIdx("salary")).length === 0);
  check("hp step empty on blank", T.validateStep(stepIdx("hp")).length === 0);

  const all = T.validateAllSteps();
  const failingIds = all.map(f => f.title);
  check("validateAllSteps maps Personal & filing", failingIds.includes("Personal & filing"), JSON.stringify(failingIds));
  check("validateAllSteps maps Bank & verification", failingIds.includes("Bank & verification"), JSON.stringify(failingIds));

  /* ---- formats ---- */
  const s1 = T.blankState();
  s1.personal.pan = "abcde1234f";                    // lowercase (before normalisation)
  s1.personal.dob = "2030-05-01";                    // future
  s1.personal.address.pin_code = "0123";
  s1.personal.address.mobile_no = "123";
  s1.personal.address.email_address = "not-an-email";
  T.set(s1);
  const fMsgs = T.validateStep(stepIdx("personal")).map(e => e.msg).join("\n");
  check("bad PAN format flagged", fMsgs.includes("invalid"), fMsgs);
  check("future DOB flagged", fMsgs.includes("cannot be after 01-04-2026"), fMsgs);
  check("bad PIN flagged", fMsgs.includes("6 digits"), fMsgs);
  check("bad mobile flagged", fMsgs.includes("10 digits"), fMsgs);
  check("bad email flagged", fMsgs.includes("name@example.com"), fMsgs);

  /* ---- fully filled mandatory fields → clean ---- */
  const s2 = T.blankState();
  Object.assign(s2.personal, { last_name: "Sharma", pan: "BQKPS1234F", dob: "1985-06-15" });
  Object.assign(s2.personal.address, {
    residence_no: "A-1", locality_or_area: "Adajan", city_or_town_or_district: "Surat",
    state_code: "24", pin_code: "395009", mobile_no: "9876543210",
    email_address: "ravi@example.com",
  });
  Object.assign(s2.bank_accounts[0], { ifsc: "SBIN0000123", bank_name: "SBI", account_no: "12345678901" });
  Object.assign(s2.verification, { name: "Ravi Sharma", father_name: "Suresh Sharma", place: "Surat" });
  T.set(s2);
  check("filled personal step clean", T.validateStep(stepIdx("personal")).length === 0,
        JSON.stringify(T.validateStep(stepIdx("personal"))));
  check("filled bank step clean", T.validateStep(stepIdx("bank")).length === 0,
        JSON.stringify(T.validateStep(stepIdx("bank"))));
  check("validateAllSteps empty when complete", T.validateAllSteps().length === 0,
        JSON.stringify(T.validateAllSteps()));

  /* ---- conditional mandatory: revised return, representative, alt address ---- */
  const s3 = T.blankState();
  s3.filing.return_section = 17;                     // 139(5) revised
  s3.filing.is_representative = true;
  s3.personal.secondary_address_same = false;
  T.set(s3);
  const cMsgs = T.validateStep(stepIdx("personal")).map(e => e.msg).join("\n");
  check("revised return needs original ack no.", cMsgs.includes("Acknowledgement no."), cMsgs);
  check("revised return needs original filing date", cMsgs.includes("original return"), cMsgs);
  check("representative name required", cMsgs.includes("Representative name"), cMsgs);
  check("alternate address required", cMsgs.includes("Alternate address Flat/No"), cMsgs);

  /* ---- challan must be complete-or-empty ---- */
  const s4 = T.blankState();
  s4.tax_paid.challans = [{ bsr_code: "0001234", date: "", challan_no: "", amount: 5000 }];
  T.set(s4);
  const chMsgs = T.validateStep(stepIdx("taxpaid")).map(e => e.msg).join("\n");
  check("partial challan flagged for date", chMsgs.includes("date of deposit"), chMsgs);
  check("partial challan flagged for serial", chMsgs.includes("serial"), chMsgs);
  s4.tax_paid.challans = [{ bsr_code: "0001234", date: "2025-09-10", challan_no: "12", amount: 5000 }];
  T.set(s4);
  check("complete challan clean", T.validateStep(stepIdx("taxpaid")).length === 0,
        JSON.stringify(T.validateStep(stepIdx("taxpaid"))));

  /* ---- tds claimed > deducted flagged ---- */
  const s5 = T.blankState();
  s5.tax_paid.tds_others = [{ tan: "AAAA00000A", deductor_name: "Bank", amount_paid: 10000,
                              deduction_year: "2025", tds_deducted: 500, tds_claimed: 900 }];
  T.set(s5);
  const tMsgs = T.validateStep(stepIdx("taxpaid")).map(e => e.msg).join("\n");
  check("claimed > deducted flagged", tMsgs.includes("cannot exceed TDS deducted"), tMsgs);

  /* ---- co-owner share range ---- */
  const s6 = T.blankState();
  s6.house_properties = [{ occupancy: "S", co_owned: true, assessee_share_pct: 150,
                           co_owners: [{ name: "", share_pct: "" }] }];
  T.set(s6);
  const hMsgs = T.validateStep(stepIdx("hp")).map(e => e.msg).join("\n");
  check("assessee share >100 flagged", hMsgs.includes("between 1 and 100"), hMsgs);
  check("co-owner name required", hMsgs.includes("Co-owner 1 name"), hMsgs);

  /* ---- let-out needs rent ---- */
  const s7 = T.blankState();
  s7.house_properties = [{ occupancy: "L", co_owned: false, loans: [] }];
  T.set(s7);
  check("let-out without rent flagged",
        T.validateStep(stepIdx("hp")).some(e => e.msg.includes("gross rent is required")),
        JSON.stringify(T.validateStep(stepIdx("hp"))));

  /* ---- new V2 steps exist and are optional on blank ---- */
  for (const id of ["capgains", "presumptive", "imports"]) {
    check(`step '${id}' exists`, stepIdx(id) >= 0, JSON.stringify(T.TABS.map(t => t.id)));
    check(`step '${id}' clean on blank`, T.validateStep(stepIdx(id)).length === 0,
          JSON.stringify(T.validateStep(stepIdx(id))));
  }

  /* ---- capital-gain row checks ---- */
  const s8 = T.blankState();
  s8.capital_gains.equity_lt = [{ scrip: "", buy_date: "2019-01-01", sale_date: "2024-08-01",
                                  sale_consideration: 500000, cost_of_acquisition: 300000 }];
  T.set(s8);
  const cgMsgs = T.validateStep(stepIdx("capgains")).map(e => e.msg).join("\n");
  check("CG scrip name required", cgMsgs.includes("scrip name"), cgMsgs);
  check("CG sale must be in FY 2025-26", cgMsgs.includes("2025-04-01 to 2026-03-31"), cgMsgs);

  const s9 = T.blankState();
  s9.capital_gains.land_building = [{ description: "Plot", buy_date: "2019-05-05",
                                      sale_date: "2025-11-20", sale_consideration: 5200000,
                                      cost_of_acquisition: 2600000 }];
  T.set(s9);
  check("complete land row clean", T.validateStep(stepIdx("capgains")).length === 0,
        JSON.stringify(T.validateStep(stepIdx("capgains"))));

  /* ---- presumptive checks ---- */
  const s10 = T.blankState();
  s10.presumptive["44ad"] = [{ name: "", code: "", gross_turnover: 35000000 }];
  s10.presumptive["44ae"] = [{ reg_no: "", months: 15 }];
  T.set(s10);
  const prMsgs = T.validateStep(stepIdx("presumptive")).map(e => e.msg).join("\n");
  check("44AD name required", prMsgs.includes("44AD row 1: business name"), prMsgs);
  check("44AD ceiling flagged", prMsgs.includes("₹3 crore"), prMsgs);
  check("44AE months range flagged", prMsgs.includes("months (1–12)"), prMsgs);

  const s11 = T.blankState();
  s11.presumptive["44ada"] = [{ name: "Clinic", code: "19003", gross_receipts: 4000000 }];
  T.set(s11);
  check("complete 44ADA row clean", T.validateStep(stepIdx("presumptive")).length === 0,
        JSON.stringify(T.validateStep(stepIdx("presumptive"))));

  /* ---- review aggregates CG/presumptive blockers ---- */
  const s12 = T.blankState();
  s12.capital_gains.other_st = [{ description: "", sale_consideration: 1000 }];
  T.set(s12);
  const rMsgs = T.validateStep(stepIdx("review")).map(e => e.msg).join("\n");
  check("review aggregates CG errors", rMsgs.includes("other_st 1"), rMsgs);

  console.log(failures ? `\n${failures} FAILURE(S)` : "\nALL PASS");
  process.exit(failures ? 1 : 0);
}, 100);
