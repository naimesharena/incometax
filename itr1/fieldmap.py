"""Crosswalk between the official Excel-utility field maps (extracted under
``extracted/xlsm/<FORM>/fields.json``) and this software's input model +
official JSON output paths.

Every entry answers for one utility input field:

* ``name``     - the utility's own defined name (e.g. ``sheet1.PAN``), exactly
                 as read by the official VBA
* ``source``   - the friendly input key this software reads (dotted);
                 ``None`` for exception fields (deliberately not ingested)
* ``path``     - the official JSON path our builders emit for it (from the
                 document root), ``None`` for exceptions
* ``label``    - fragment of the on-screen label in the utility (checked)
* ``kind``     - "mapped" | "partial" (mapped with documented difference) |
                 "exception" (utility-only field, reason in ``note``)
* ``sent``     - (value, expected) pair used by the sentinel flow test:
                 the value injected at ``source`` and the value expected at
                 ``path`` in the generated document.  ``("INJECT", {...})``
                 merges extra input first (for conditional nodes).

``tests/test_field_mapping.py`` enforces all of this against both the
extracted field maps and the generated documents.
"""
from __future__ import annotations

ITR1 = ("ITR", "ITR1")
ITR2 = ("ITR", "ITR2")
ITR4 = ("ITR", "ITR4")

PI1 = ("PersonalInfo",)
PI2 = ("PartA_GEN1", "PersonalInfo")
FS1 = ("FilingStatus",)
FS2 = ("PartA_GEN1", "FilingStatus")
ADDR = ("PersonalInfo", "Address")
ADDR2 = ("PartA_GEN1", "PersonalInfo", "Address")
ALT1 = ("PersonalInfo", "AlternateAddress")
ALT2 = ("PartA_GEN1", "PersonalInfo", "AlternateAddress")
ALT4 = ("PersonalInfo", "AlternateAddress")
NAME1 = ("PersonalInfo", "AssesseeName")
NAME2 = ("PartA_GEN1", "PersonalInfo", "AssesseeName")


def F(name, source, root, suffix, label, kind="mapped", note="", sent=None,
      extra_inject=None):
    return {"name": name, "source": source,
            "path": tuple(root) + tuple(suffix) if suffix else None,
            "label": label, "kind": kind, "note": note,
            "sent": sent, "extra_inject": extra_inject}


# ---------------------------------------------------------------------------
# ITR-1 (Sahaj) - utility identity sheet: "Income Details"
# ---------------------------------------------------------------------------
ITR1_FIELDS = [
    # identity
    F("sheet1.PAN", "personal.pan", ITR1, PI1 + ("PAN",), "PAN",
      sent=("AABCR1234C", "AABCR1234C")),
    F("sheet1.FirstName", "personal.first_name", ITR1, NAME1 + ("FirstName",),
      "First Name", sent=("Zubin", "Zubin")),
    F("sheet1.MiddleName", "personal.middle_name", ITR1, NAME1 + ("MiddleName",),
      "Middle Name", sent=("Manek", "Manek")),
    F("sheet1.SurNameOrOrgName", "personal.last_name", ITR1,
      NAME1 + ("SurNameOrOrgName",), "Last Name", sent=("Sorabjee", "Sorabjee")),
    F("sheet1.DOB", "personal.dob", ITR1, PI1 + ("DOB",), "Date of Birth",
      sent=("1988-02-14", "1988-02-14")),
    F("Sheet1.Aadhaar", "personal.aadhaar", ITR1, PI1 + ("AadhaarCardNo",),
      None, sent=("999988887777", "999988887777")),
    F("sheet1.EmployerCategory1", "personal.employer_category", ITR1,
      PI1 + ("EmployerCategory",), "Nature of Employment",
      sent=("PSU", "PSU")),
    F("sheet1.Status", None, None, None, None, kind="exception",
      note="ITR-1 is individual-only; status shown on the sheet only to gate eligibility."),
    F("sheet1.ResidentialStatus1", None, None, None, None, kind="exception",
      note="ITR-1 is resident-only; the utility uses this solely to gate eligibility (no schema node)."),
    F("Sheet1.AadhaarEnrol", None, None, None, None, kind="exception",
      note="Aadhaar enrolment id not modelled (Aadhaar number itself is mapped)."),
    F("sheet1.Gender1", None, None, None, None, kind="exception",
      note="Gender is not an AY 2026-27 schema field for ITR-1."),
    # primary address
    F("sheet1.ResidenceNo", "personal.address.residence_no", ITR1,
      ADDR + ("ResidenceNo",), "Flat / Door",
      sent=("7-B", "7-B")),
    F("sheet1.ResidenceName", "personal.address.residence_name", ITR1,
      ADDR + ("ResidenceName",), "Name of Premises",
      sent=("Krishna Kunj", "Krishna Kunj")),
    F("sheet1.RoadOrStreet", "personal.address.road_or_street", ITR1,
      ADDR + ("RoadOrStreet",), "Road / Street",
      sent=("Ring Road", "Ring Road")),
    F("sheet1.LocalityOrArea", "personal.address.locality_or_area", ITR1,
      ADDR + ("LocalityOrArea",), "Area / Locality",
      sent=("Adajan", "Adajan")),
    F("sheet1.CityOrTownOrDistrict", "personal.address.city_or_town_or_district",
      ITR1, ADDR + ("CityOrTownOrDistrict",), "Town / City",
      sent=("Surat", "Surat")),
    F("sheet1.StateCode1", "personal.address.state_code", ITR1,
      ADDR + ("StateCode",), "State", sent=("06", "06")),
    F("sheet1.Country", "personal.address.country_code", ITR1,
      ADDR + ("CountryCode",), "Country", sent=("91", "91")),
    F("sheet1.PinCode", "personal.address.pin_code", ITR1,
      ADDR + ("PinCode",), "PIN Code", sent=(395009, 395009)),
    F("sheet1.ZipCode", None, None, None, None, kind="exception",
      note="Foreign ZIP (non-Indian address) not modelled; PinCode is mapped."),
    # secondary address (utility: 'Is the secondary address same as primary?')
    F("Secondary_Address", "personal.secondary_address_same", ITR1,
      PI1 + ("SecondaryAdd",), None,
      sent=(False, "N"),
      extra_inject={"personal": {"alternate_address": {
          "residence_no": "22", "locality_or_area": "Vesu",
          "city_or_town_or_district": "Surat", "state_code": "06",
          "pin_code": 395007}}},
      note="Y = same as primary (no AlternateAddress), N = different address follows."),
    F("sheet1.ResidenceNo1", "personal.alternate_address.residence_no", ITR1,
      ALT1 + ("ResidenceNo",), "Flat / Door", sent=("22", "22"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.LocalityOrArea1", "personal.alternate_address.locality_or_area",
      ITR1, ALT1 + ("LocalityOrArea",), "Area / Locality", sent=("Vesu", "Vesu"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.StateCode2", "personal.alternate_address.state_code", ITR1,
      ALT1 + ("StateCode",), "State", sent=("06", "06"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.PinCode1", "personal.alternate_address.pin_code", ITR1,
      ALT1 + ("PinCode",), "PIN Code", sent=(395007, 395007),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    # contact
    F("sheet1.EmailAddress", "personal.address.email_address", ITR1,
      ADDR + ("EmailAddress",), "Primary Email",
      sent=("zubin.s@example.in", "zubin.s@example.in")),
    F("sheet1.EmailAddress1", "personal.address.email_address_sec", ITR1,
      ADDR + ("EmailAddressSec",), "Secondary Email",
      sent=("z.alt@example.in", "z.alt@example.in")),
    F("sheet1.Mobileno", "personal.address.mobile_no", ITR1,
      ADDR + ("MobileNo",), None, sent=(9825098250, 9825098250)),
    F("sheet1.MobileCountryCode", "personal.address.country_code_mobile", ITR1,
      ADDR + ("CountryCodeMobile",), "Primary Mobile", sent=(91, 91)),
    F("sheet1.Mobileno1", "personal.address.mobile_no_sec", ITR1,
      ADDR + ("MobileNoSec",), None, sent=(9820098200, 9820098200)),
    F("sheet1.MobileCountryCode1", "personal.address.country_code_mobile_sec",
      ITR1, ADDR + ("CountryCodeMobileNoSec",), "Secondary Mobile",
      sent=(91, 91), extra_inject={"personal": {"address": {"mobile_no_sec": 9820098200}}}),
    F("sheet1.STDcode", None, None, None, None, kind="exception",
      note="Residential/office landline (STD code + number) is optional in practice and not modelled."),
    # filing status
    F("sheet1.ReturnFileSec", "filing.return_section", ITR1,
      FS1 + ("ReturnFileSec",), "Filed u/s", sent=(14, 14),
      extra_inject={"filing": {"notice_no": "ITBA/142(1)/2025-26/1",
                               "notice_date": "2026-01-10"}},
      note="schema section codes (11=139(1), 12=139(4), 14=142(1)...) - same as utility dropdown."),
    F("sheet1.ReceiptNo", "filing.original_receipt_no", ITR1,
      FS1 + ("ReceiptNo",), None, sent=("912345678901230", "912345678901230"),
      extra_inject={"filing": {"return_section": 17,
                               "original_filing_date": "2026-06-01"}}),
    F("sheet1.OrigRetFiledDate", "filing.original_filing_date", ITR1,
      FS1 + ("OrigRetFiledDate",), None, sent=("2026-06-01", "2026-06-01"),
      extra_inject={"filing": {"return_section": 17,
                               "original_receipt_no": "912345678901230"}}),
    F("sheet1.NoticeNo", "filing.notice_no", ITR1,
      FS1 + ("NoticeNo",), None, sent=("ITBA/142(1)/2025-26/1", "ITBA/142(1)/2025-26/1"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_date": "2026-01-10"}}),
    F("sheet1.NoticeDate", "filing.notice_date", ITR1,
      FS1 + ("NoticeDateUnderSec",), None, sent=("2026-01-10", "2026-01-10"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_no": "ITBA/142(1)/2025-26/1"}}),
    F("sheet1.RepAssessee", "filing.is_representative", ITR1,
      FS1 + ("AsseseeRepFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"representative": {
          "name": "Rep Name", "pan": "AABCR1234C", "address": "Addr",
          "mobile": 9825098250, "email": "rep@example.in"}}}),
    F("Sheet1.DepositAmountFlag", "filing.deposit_exceeds_1cr", ITR1,
      FS1 + ("clauseiv7provisio139i",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("Sheet1.AggrigateAmountFlag", "filing.foreign_travel_exceeds_2l", ITR1,
      FS1 + ("IncrExpAggAmt2LkTrvFrgnCntryFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("Sheet1.AggrigateAmountFlag1", "filing.electricity_exceeds_1l", ITR1,
      FS1 + ("IncrExpAggAmt1LkElctrctyPrYrFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("clauseiv7provisio139iFlg_1", None, None, None, None, kind="exception",
      note="Utility's per-country granular rows of the 7th proviso; the aggregate flags above are what the schema carries."),
    F("sheet1.SeventhProvisoFlag", "filing.seventh_provisio", ITR1,
      FS1 + ("SeventhProvisio139",), None, sent=(True, "Y")),
    F("taxcheck", None, None, None, None, kind="partial",
      note="Utility's display-only tax-regime toggles; regime selection is step 1 of our wizard (filing.opt_out_new_regime -> FilingStatus/OptOutNewTaxRegime)."),
]

# ---------------------------------------------------------------------------
# ITR-2 - utility identity sheet: "PART A - General"
# ---------------------------------------------------------------------------
ITR2_FIELDS = [
    F("sheet1.PAN", "personal.pan", ITR2, PI2 + ("PAN",), "PAN",
      sent=("AABCR1234C", "AABCR1234C")),
    F("sheet1.FirstName", "personal.first_name", ITR2, NAME2 + ("FirstName",),
      "First Name", sent=("Zubin", "Zubin")),
    F("sheet1.MiddleName", "personal.middle_name", ITR2, NAME2 + ("MiddleName",),
      "Middle Name", sent=("Manek", "Manek")),
    F("sheet1.SurNameOrOrgName", "personal.last_name", ITR2,
      NAME2 + ("SurNameOrOrgName",), "Last Name", sent=("Sorabjee", "Sorabjee")),
    F("sheet1.DOB", "personal.dob", ITR2, PI2 + ("DOB",), "Date of Birth",
      sent=("1988-02-14", "1988-02-14")),
    F("sheet1.adhaarno", "personal.aadhaar", ITR2, PI2 + ("AadhaarCardNo",),
      None, sent=("999988887777", "999988887777")),
    F("sheet1.Status", "personal.status", ITR2, PI2 + ("Status",),
      "Status (I-Individual,H-HUF)", kind="partial",
      note="value passed through (default IIndividual); the computation engine is tuned for individuals - HUF nuances (87A etc.) are documented scope.",
      sent=("I", "I")),
    F("sheet1.DOC", None, None, None, None, kind="exception",
      note="Date of commencement of business - ITR-3-only; ITR-2 has no such schema node."),
    F("sheet1.Foreign_Exchange", None, None, None, None, kind="exception",
      note="Utility display helper, no schema node."),
    # residential status & address
    F("sheet1.ResidentialStatus1", "personal.residential_status", ITR2,
      FS2 + ("ResidentialStatus",), "Residential Status",
      sent=("NRI", "NRI")),
    F("sheet1.ResidentialStatus1_Condition", None, None, None, None,
      kind="exception",
      note="Utility-internal RNOR condition helper (derived from stay day-counts)."),
    F("Resident_Button", None, None, None, None, kind="exception",
      note="Stay-day-count grid for RNOR determination - utility helper, no schema node."),
    F("NRI_TotPeriodofStay_PreviousYear", None, None, None, None,
      kind="exception", note="RNOR helper, see Resident_Button."),
    F("NRI_TotPeriodofStay_PrecedingYear", None, None, None, None,
      kind="exception", note="RNOR helper."),
    F("sheet1.ResidenceNo", "personal.address.residence_no", ITR2,
      ADDR2 + ("ResidenceNo",), "Flat / Door", sent=("7-B", "7-B")),
    F("sheet1.ResidenceName", "personal.address.residence_name", ITR2,
      ADDR2 + ("ResidenceName",), "Name of Premises",
      sent=("Krishna Kunj", "Krishna Kunj")),
    F("sheet1.RoadOrStreet", "personal.address.road_or_street", ITR2,
      ADDR2 + ("RoadOrStreet",), "Road / Street", sent=("Ring Road", "Ring Road")),
    F("sheet1.LocalityOrArea", "personal.address.locality_or_area", ITR2,
      ADDR2 + ("LocalityOrArea",), "Area / Locality", sent=("Adajan", "Adajan")),
    F("sheet1.CityOrTownOrDistrict", "personal.address.city_or_town_or_district",
      ITR2, ADDR2 + ("CityOrTownOrDistrict",), "Town/City", sent=("Surat", "Surat")),
    F("sheet1.StateCode1", "personal.address.state_code", ITR2,
      ADDR2 + ("StateCode",), "State", sent=("06", "06")),
    F("sheet1.CountryCode", "personal.address.country_code", ITR2,
      ADDR2 + ("CountryCode",), "Country", sent=("91", "91")),
    F("sheet1.PinCode", "personal.address.pin_code", ITR2,
      ADDR2 + ("PinCode",), "Pin Code", sent=(395009, 395009)),
    F("sheet1.ZipCode", None, None, None, None, kind="exception",
      note="Foreign ZIP not modelled."),
    F("Secondary_Address", "personal.secondary_address_same", ITR2,
      PI2 + ("SecondaryAdd",), None, sent=(False, "N"),
      extra_inject={"personal": {"alternate_address": {
          "residence_no": "22", "locality_or_area": "Vesu",
          "city_or_town_or_district": "Surat", "state_code": "06",
          "pin_code": 395007}}}),
    F("sheet1.ResidenceNo2", "personal.alternate_address.residence_no", ITR2,
      ALT2 + ("ResidenceNo",), "Flat / Door", sent=("22", "22"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.LocalityOrArea2", "personal.alternate_address.locality_or_area",
      ITR2, ALT2 + ("LocalityOrArea",), "Area / Locality", sent=("Vesu", "Vesu"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.StateCode2", "personal.alternate_address.state_code", ITR2,
      ALT2 + ("StateCode",), "State", sent=("06", "06"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.PinCode2", "personal.alternate_address.pin_code", ITR2,
      ALT2 + ("PinCode",), "Pin Code", sent=(395007, 395007),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    # contact
    F("sheet1.EmailAddress", "personal.address.email_address", ITR2,
      ADDR2 + ("EmailAddress",), "Primary Email",
      sent=("zubin.s@example.in", "zubin.s@example.in")),
    F("sheet1.EmailAddressSec", "personal.address.email_address_sec", ITR2,
      ADDR2 + ("EmailAddressSec",), "Secondary Email",
      sent=("z.alt@example.in", "z.alt@example.in")),
    F("Countrycode1", "personal.address.country_code_mobile", ITR2,
      ADDR2 + ("CountryCodeMobile",), "Primary Mobile", sent=(91, 91)),
    F("sheet1.MobileNo", "personal.address.mobile_no", ITR2,
      ADDR2 + ("MobileNo",), None, sent=(9825098250, 9825098250)),
    F("Countrycode2", "personal.address.country_code_mobile_sec", ITR2,
      ADDR2 + ("CountryCodeMobileNoSec",), "Secondary Mobile",
      sent=(91, 91),
      extra_inject={"personal": {"address": {"mobile_no_sec": 9820098200}}}),
    F("sheet1.MobileNoSec", "personal.address.mobile_no_sec", ITR2,
      ADDR2 + ("MobileNoSec",), None, sent=(9820098200, 9820098200),
      extra_inject=None),
    F("sheet1.STDcode", None, None, None, None, kind="exception",
      note="Landline field not modelled."),
    F("sheet1.PhoneNo", None, None, None, None, kind="exception",
      note="Landline field not modelled."),
    # filing status
    F("sheet1.ReturnFileSec", "filing.return_section", ITR2,
      FS2 + ("ReturnFileSec",), "Filed in response to", sent=(14, 14),
      extra_inject={"filing": {"notice_no": "ITBA/142(1)/2025-26/1",
                               "notice_date": "2026-01-10"}}),
    F("OptingNewTaxRegime", None, None, None, None, kind="partial",
      note="regime is wizard step 1: filing.opt_out_new_regime -> FilingStatus/OptOutNewTaxRegime."),
    F("BacValueDropdown", None, None, None, None, kind="exception",
      note="regime radio helper (same as OptingNewTaxRegime)."),
    F("Radio_Button_115", None, None, None, None, kind="exception",
      note="regime radio helper."),
    F("sheet1.ReceiptNo", "filing.original_receipt_no", ITR2,
      FS2 + ("ReceiptNo",), None, sent=("912345678901230", "912345678901230"),
      extra_inject={"filing": {"return_section": 17,
                               "original_filing_date": "2026-06-01"}}),
    F("sheet1.OrigRetFiledDate", "filing.original_filing_date", ITR2,
      FS2 + ("OrigRetFiledDate",), "Date of filing of Original",
      sent=("2026-06-01", "2026-06-01"),
      extra_inject={"filing": {"return_section": 17,
                               "original_receipt_no": "912345678901230"}}),
    F("Sheet1.UniqueNo", "filing.notice_no", ITR2,
      FS2 + ("NoticeNo",), None, sent=("ITBA/142(1)/2025-26/1", "ITBA/142(1)/2025-26/1"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_date": "2026-01-10"}}),
    F("sheet1.NoticeDate", "filing.notice_date", ITR2,
      FS2 + ("NoticeDate",), "Date or if filed u/s 92CD",
      sent=("2026-01-10", "2026-01-10"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_no": "ITBA/142(1)/2025-26/1"}}),
    F("sheet1.SeventhProvisoFlag", "filing.seventh_provisio", ITR2,
      FS2 + ("SeventhProvisio139",), None, sent=(True, "Y")),
    F("Sheet1.DepositAmountFlag", "filing.deposit_exceeds_1cr", ITR2,
      FS2 + ("DepAmtAggAmtExcd1CrPrYrFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("Sheet1.AggrigateAmountFlag", "filing.foreign_travel_exceeds_2l", ITR2,
      FS2 + ("IncrExpAggAmt2LkTrvFrgnCntryFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("Sheet1.AggrigateAmountFlag1", "filing.electricity_exceeds_1l", ITR2,
      FS2 + ("IncrExpAggAmt1LkElctrctyPrYrFlg",), None, sent=(True, "Y"),
      extra_inject={"filing": {"seventh_provisio": True}}),
    F("clauseiv7provisio139iFlg", None, None, None, None, kind="exception",
      note="per-country granular grid of the 7th proviso; aggregate flags are mapped."),
    F("sheet1.capacityassessee", None, None, None, None, kind="partial",
      note="representative-capacity detail rows - supported on ITR-1 (filing.is_representative + representative block), pending on ITR-2."),
    F("acknowledgmentnumber", None, None, None, None, kind="exception",
      note="display-only helper."),
]

# ---------------------------------------------------------------------------
# ITR-4 (Sugam) - utility identity sheet: "Income Details"
# ---------------------------------------------------------------------------
PI4 = ("PersonalInfo",)
ADDR4 = ("PersonalInfo", "Address")
NAME4 = ("PersonalInfo", "AssesseeName")
FS4 = ("FilingStatus",)

ITR4_FIELDS = [
    F("sheet1.PAN", "personal.pan", ITR4, PI4 + ("PAN",),
      "Permanent Account Number", sent=("AABCR1234C", "AABCR1234C")),
    F("sheet1.FirstName", "personal.first_name", ITR4, NAME4 + ("FirstName",),
      "First Name", sent=("Zubin", "Zubin")),
    F("sheet1.MiddleName", "personal.middle_name", ITR4, NAME4 + ("MiddleName",),
      "Middle Name", sent=("Manek", "Manek"),
      extra_inject=None),
    F("sheet1.SurNameOrOrgName", "personal.last_name", ITR4,
      NAME4 + ("SurNameOrOrgName",), "Last Name", sent=("Sorabjee", "Sorabjee")),
    F("sheet1.DOB", "personal.dob", ITR4, PI4 + ("DOB",),
      "Date of Birth/Formation", sent=("1988-02-14", "1988-02-14")),
    F("sheet1.Aadhaar", "personal.aadhaar", ITR4, PI4 + ("AadhaarCardNo",),
      None, sent=("999988887777", "999988887777")),
    F("sheet1.AadhaarEnrol", None, None, None, None, kind="exception",
      note="Aadhaar enrolment id not modelled."),
    F("sheet1.Status", "personal.status", ITR4, PI4 + ("Status",), "Status",
      kind="partial",
      note="passed through (default I); computation tuned for individuals.",
      sent=("I", "I")),
    F("sheet1.NatureofEmployment", "personal.employer_category", ITR4,
      PI4 + ("EmployerCategory",), "Nature of Employment", sent=("OTH", "OTH")),
    F("sheet1.PortugeseCC5A", None, None, None, None, kind="exception",
      note="Portuguese Civil Code 5A income-apportionment regime (Goa) - not implemented (documented scope; spouse PAN field pairs with it)."),
    F("sheet1.PANOFSPOUSE", None, None, None, None, kind="exception",
      note="spouse PAN - relevant only with the 5A regime, see PortugeseCC5A."),
    F("sheet1.DesigOfficerWardorCircle", None, None, None, None,
      kind="exception",
      note="IT ward/circle designation - not a mandatory schema node; not captured in v2."),
    # addresses
    F("sheet1.ResidenceNo", "personal.address.residence_no", ITR4,
      ADDR4 + ("ResidenceNo",), "Flat/ Door", sent=("7-B", "7-B")),
    F("sheet1.ResidenceName", "personal.address.residence_name", ITR4,
      ADDR4 + ("ResidenceName",), "Name of Premises",
      sent=("Krishna Kunj", "Krishna Kunj")),
    F("sheet1.RoadOrStreet", "personal.address.road_or_street", ITR4,
      ADDR4 + ("RoadOrStreet",), "Road / Street", sent=("Ring Road", "Ring Road")),
    F("sheet1.LocalityOrArea", "personal.address.locality_or_area", ITR4,
      ADDR4 + ("LocalityOrArea",), "Area / Locality", sent=("Adajan", "Adajan")),
    F("sheet1.CityOrTownOrDistrict", "personal.address.city_or_town_or_district",
      ITR4, ADDR4 + ("CityOrTownOrDistrict",), "Town/City", sent=("Surat", "Surat")),
    F("sheet1.StateCode1", "personal.address.state_code", ITR4,
      ADDR4 + ("StateCode",), "State", sent=("06", "06")),
    F("sheet1.CountryCode", "personal.address.country_code", ITR4,
      ADDR4 + ("CountryCode",), "Country", sent=("91", "91")),
    F("sheet1.PinCode", "personal.address.pin_code", ITR4,
      ADDR4 + ("PinCode",), "PIN Code", sent=(395009, 395009)),
    F("sheet1.ZipCode", None, None, None, None, kind="exception",
      note="Foreign ZIP not modelled."),
    F("Secondary_Address", "personal.secondary_address_same", ITR4,
      PI4 + ("SecondaryAdd",), None, sent=(False, "N"),
      extra_inject={"personal": {"alternate_address": {
          "residence_no": "22", "locality_or_area": "Vesu",
          "city_or_town_or_district": "Surat", "state_code": "06",
          "pin_code": 395007}}}),
    F("sheet1.ResidenceNo2", "personal.alternate_address.residence_no", ITR4,
      ALT4 + ("ResidenceNo",), "Flat/ Door", sent=("22", "22"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.LocalityOrArea2", "personal.alternate_address.locality_or_area",
      ITR4, ALT4 + ("LocalityOrArea",), "Area / Locality", sent=("Vesu", "Vesu"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.StateCode2", "personal.alternate_address.state_code", ITR4,
      ALT4 + ("StateCode",), "State", sent=("06", "06"),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    F("sheet1.PinCode2", "personal.alternate_address.pin_code", ITR4,
      ALT4 + ("PinCode",), "PIN Code", sent=(395007, 395007),
      extra_inject={"personal": {"secondary_address_same": False,
                                 "alternate_address": {}}}),
    # contact
    F("sheet1.EmailAddress", "personal.address.email_address", ITR4,
      ADDR4 + ("EmailAddress",), "Primary Email",
      sent=("zubin.s@example.in", "zubin.s@example.in")),
    F("sheet1.EmailAddress2", "personal.address.email_address_sec", ITR4,
      ADDR4 + ("EmailAddressSec",), "Secondary Email",
      sent=("z.alt@example.in", "z.alt@example.in")),
    F("sheet1.Mobileno", "personal.address.mobile_no", ITR4,
      ADDR4 + ("MobileNo",), None, sent=(9825098250, 9825098250)),
    F("sheet1.mobileCountryCode", "personal.address.country_code_mobile", ITR4,
      ADDR4 + ("CountryCodeMobile",), "Primary Mobile", sent=(91, 91)),
    F("sheet1.MobileNoSec", "personal.address.mobile_no_sec", ITR4,
      ADDR4 + ("MobileNoSec",), None, sent=(9820098200, 9820098200)),
    F("sheet1.mobileCountryCode2", "personal.address.country_code_mobile_sec",
      ITR4, ADDR4 + ("CountryCodeMobileNoSec",), "Secondary Mobile",
      sent=(91, 91),
      extra_inject={"personal": {"address": {"mobile_no_sec": 9820098200}}}),
    F("sheet1.STDcode", None, None, None, None, kind="exception",
      note="Landline field not modelled."),
    F("sheet1.PhoneNo", None, None, None, None, kind="exception",
      note="Landline field not modelled."),
    # filing status
    F("sheet1.ReturnFileSec", "filing.return_section", ITR4,
      FS4 + ("ReturnFileSec",), "Filed u/s", sent=(14, 14),
      extra_inject={"filing": {"notice_no": "ITBA/142(1)/2025-26/1",
                               "notice_date": "2026-01-10"}}),
    F("sheet1.ReturnType1", None, None, None, None, kind="partial",
      note="original/revised is implied by the return-section code itself."),
    F("Sheet1.115BAC", None, None, None, None, kind="partial",
      note="regime is wizard step 1; the utility's Form-10IEA history grid (BAC115.* rows) is captured as documentation - we file the current-year election only."),
    F("sheet1.ReceiptNo1", "filing.original_receipt_no", ITR4,
      FS4 + ("ReceiptNo",), None, sent=("912345678901230", "912345678901230"),
      extra_inject={"filing": {"return_section": 17,
                               "original_filing_date": "2026-06-01"}}),
    F("sheet1.OrigRetFiledDate1", "filing.original_filing_date", ITR4,
      FS4 + ("OrigRetFiledDate",), "Date of filing of original",
      sent=("2026-06-01", "2026-06-01"),
      extra_inject={"filing": {"return_section": 17,
                               "original_receipt_no": "91234567891230"}}),
    F("sheet1.NoticeNo", "filing.notice_no", ITR4,
      FS4 + ("NoticeNo",), None, sent=("ITBA/142(1)/2025-26/1", "ITBA/142(1)/2025-26/1"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_date": "2026-01-10"}}),
    F("sheet1.NoticeDate", "filing.notice_date", ITR4,
      FS4 + ("NoticeDateUnderSec",), "Date of such Notice",
      sent=("2026-01-10", "2026-01-10"),
      extra_inject={"filing": {"return_section": 14,
                               "notice_no": "ITBA/142(1)/2025-26/1"}}),
    F("Radiobuttoncheck", None, None, None, None, kind="exception",
      note="utility-internal return-section radio helper."),
]

FIELD_MAP = {"ITR1": ITR1_FIELDS, "ITR2": ITR2_FIELDS, "ITR4": ITR4_FIELDS}

IDENTITY_SHEET = {"ITR1": "Income Details", "ITR2": "PART A - General",
                  "ITR3": "PART A - General", "ITR4": "Income Details"}


def coverage(form: str) -> dict:
    fields = FIELD_MAP[form]
    out = {"mapped": 0, "partial": 0, "exception": 0, "total": len(fields)}
    for f in fields:
        out[f["kind"]] = out.get(f["kind"], 0) + 1
    out["mapped_incl_partial"] = out["mapped"] + out["partial"]
    return out
