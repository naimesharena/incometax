"""Import hub tests: ITR JSON prefill/YoY/losses, AIS JSON + password ZIP,
26AS text + password PDF, and the credit/income apply flows."""

import io
import json
import struct
import zipfile

import pytest

from itr1 import imports
from itr1.errors import DataError

USER = {
    "personal": {"first_name": "Aarav", "last_name": "Sharma", "pan": "BQXPS1234F",
                 "dob": "1985-06-15", "employer_category": "OTH",
                 "address": {"residence_no": "B-204", "locality_or_area": "Adajan",
                             "city_or_town_or_district": "Surat", "state_code": "06",
                             "country_code": "91", "pin_code": 395009,
                             "mobile_no": 9876543210,
                             "email_address": "aarav@example.com"}},
}


# ---------------------------------------------------------------------------
# tiny classic-ZipCrypto writer (stdlib can only *decrypt*)
# ---------------------------------------------------------------------------

def _crc_table():
    table = []
    for i in range(256):
        c = i
        for _ in range(8):
            c = (c >> 1) ^ 0xEDB88320 if c & 1 else c >> 1
        table.append(c)
    return table


_CRC = _crc_table()


def _zipcrypto_encrypt(data: bytes, password: str) -> bytes:
    """Traditional PKWARE cipher, mirroring zipfile's local crc32 primitive."""
    import zlib
    key0, key1, key2 = 0x12345678, 0x23456789, 0x34567890

    def crc32_byte(ch, crc):
        return ((crc >> 8) ^ _CRC[(crc ^ ch) & 0xFF]) & 0xFFFFFFFF

    def update(c):
        nonlocal key0, key1, key2
        key0 = crc32_byte(c, key0)
        key1 = (key1 + (key0 & 0xFF)) & 0xFFFFFFFF
        key1 = (key1 * 134775813 + 1) & 0xFFFFFFFF
        key2 = crc32_byte((key1 >> 24) & 0xFF, key2)

    for b in password.encode():
        update(b)

    crc = zlib.crc32(data) & 0xFFFFFFFF
    header = bytes(range(8, 19)) + bytes([(crc >> 24) & 0xFF])   # exactly 12 bytes
    out = bytearray()
    for plain in header + data:
        k = (key2 | 2) & 0xFFFFFFFF
        out.append(plain ^ (((k * (k ^ 1)) >> 8) & 0xFF))   # XOR, not subtraction
        update(plain)          # keys advance on the PLAINTEXT byte
    return crc, bytes(out)


def make_encrypted_zip(entries: dict[str, bytes], password: str) -> bytes:
    """Build a stored-method ZipCrypto archive (the layout the portal ships)."""
    buf = io.BytesIO()
    central = []
    offset = 0
    for name, data in entries.items():
        crc, enc = _zipcrypto_encrypt(data, password)
        nb = name.encode()
        local = struct.pack(
            "<IHHHHHIIIHH", 0x04034B50, 20, 0x01, 0, 0, 0,
            crc, len(enc), len(data), len(nb), 0)
        buf.write(local + nb + enc)
        central.append((name, crc, len(enc), len(data), offset))
        offset += len(local) + len(nb) + len(enc)
    cd_start = offset
    for name, crc, csize, usize, off in central:
        nb = name.encode()
        rec = struct.pack(
            "<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, 0x01, 0, 0, 0,
            crc, csize, usize, len(nb), 0, 0, 0, 0, 0, off)
        buf.write(rec + nb)
        offset += len(rec) + len(nb)
    cd_size = offset - cd_start
    buf.write(struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, len(central),
                          len(central), cd_size, cd_start, 0))
    return buf.getvalue()


# ---------------------------------------------------------------------------
# passwords
# ---------------------------------------------------------------------------

def test_password_candidates():
    pw = imports.passwords_for(USER, "ais")
    assert pw[0] == "bqxps1234f15061985"      # pan.lower() + DDMMYYYY
    assert "BQXPS1234F15061985" in pw
    pw26 = imports.passwords_for(USER, "26as")
    assert pw26[0] == "15061985"              # DDMMYYYY first for 26AS


# ---------------------------------------------------------------------------
# ITR JSON (previous-year) parsing
# ---------------------------------------------------------------------------

PREV_ITR2 = {
    "ITR": {"ITR2": {
        "CreationInfo": {"SWVersionNo": "1.0", "SWCreatedBy": "SW00000001",
                         "JSONCreatedBy": "SW00000001", "IntermediaryCity": "Delhi"},
        "Form_ITR2": {"FormName": "ITR-2", "Description": "x", "AssessmentYear": "2025"},
        "PartA_GEN1": {
            "PersonalInfo": {
                "AssesseeName": {"FirstName": "Aarav", "SurNameOrOrgName": "Sharma"},
                "PAN": "BQXPS1234F", "DOB": "1985-06-15",
                "Address": {"ResidenceNo": "B-204", "LocalityOrArea": "Adajan",
                            "CityOrTownOrDistrict": "Surat", "StateCode": "06",
                            "CountryCode": "91", "PinCode": 395009,
                            "MobileNo": 9876543210,
                            "EmailAddress": "aarav@example.com", "CountryCodeMobile": 91},
            },
            "FilingStatus": {"ReturnFileSec": 11},
        },
        "PartB-TI": {
            "Salaries": 1850000, "IncomeFromHP": -42000,
            "CapGain": {"ShortTerm": {"TotalShortTerm": 30000},
                        "LongTerm": {"TotalLongTerm": 210000},
                        "TotalCapGains": 240000},
            "IncFromOS": {"TotIncFromOS": 88000},
            "GrossTotalIncome": 2146000, "TotalIncome": 1881000,
            "LossesOfCurrentYearCarriedFwd": 50000,
        },
        "ScheduleCFL": {"STCG": {"TotalAmtSTCGCL": 50000},
                        "LTCG": {"TotalAmtLTCGCL": 120000}},
        "PartB_TTI": {
            "ComputationOfTaxLiability": {"AggregateTaxInterestLiability": 380000},
            "TaxPaid": {"TaxesPaid": {"AdvanceTax": 100000, "TDS": 265000, "TCS": 0,
                                      "SelfAssessmentTax": 35000},
                        "BalTaxPayable": 0},
            "Refund": {"RefundDue": 20000,
                       "BankAccountDtls": {"AddtnlBankDetails": [
                           {"IFSCCode": "HDFC0001234", "BankName": "HDFC Bank",
                            "BankAccountNo": "50100123456789", "AccountType": "SB",
                            "UseForRefund": "true"}]}},
        },
        "Verification": {"Declaration": {"AssesseeVerName": "Aarav Sharma",
                                         "FatherName": "Rajesh Sharma",
                                         "AssesseeVerPAN": "BQXPS1234F"},
                         "Capacity": "S", "Place": "Surat"},
    }}}


def test_parse_itr_json_summary():
    s = imports.parse_itr_json(json.dumps(PREV_ITR2))
    assert s["form"] == "ITR2"
    assert s["assessment_year"] == "2025"
    assert s["pan"] == "BQXPS1234F"
    assert s["heads"]["salary"] == 1850000
    assert s["heads"]["cg"] == 240000
    assert s["heads"]["total_income"] == 1881000
    assert s["tax"]["tds"] == 265000
    assert s["tax"]["refund"] == 20000
    assert s["losses"]["stcg_cf"] == 50000
    assert s["losses"]["ltcg_cf"] == 120000
    assert s["bank_count"] == 1


def test_parse_itr_json_rejects_garbage():
    with pytest.raises(DataError):
        imports.parse_itr_json("not json at all")
    with pytest.raises(DataError):
        imports.parse_itr_json('{"foo": 1}')


def test_apply_itr_prefill_fills_only_empty():
    data = {"personal": {"pan": "XXXXX0000X"}}
    summary = imports.parse_itr_json(PREV_ITR2)
    out, notes = imports.apply_itr_prefill(data, summary)
    # PAN kept (not empty), DOB + address filled
    assert out["personal"]["pan"] == "XXXXX0000X"
    assert out["personal"]["dob"] == "1985-06-15"
    assert out["personal"]["address"]["pin_code"] == 395009
    assert out["bank_accounts"][0]["ifsc"] == "HDFC0001234"
    assert out["losses_cf"]["stcg_cf"] == 50000
    assert out["losses_cf"]["ltcg_cf"] == 120000


def test_year_on_year():
    summary = imports.parse_itr_json(PREV_ITR2)
    rows = imports.year_on_year({"salary": 2100000, "total_income": 2300000}, summary)
    by = {r["head"]: r for r in rows}
    assert by["Salary income (u/s 16)"]["delta"] == 250000
    assert by["Total income"]["trend"] == "up"
    assert by["Total income"]["delta_pct"] == pytest.approx(22.3, abs=0.1)


# ---------------------------------------------------------------------------
# AIS
# ---------------------------------------------------------------------------

AIS_DOC = {
    "assessmentYear": "2026",
    "PART-B": {
        "TDS Information": {
            "records": [
                {"informationDescription": "TDS u/s 192 by employer",
                 "informationSource": "ACME SOFTWARE PVT LTD", "amount": 240000,
                 "transactionDate": "2026-03-31"},
            ],
        },
        "Tax Payments": {
            "records": [
                {"informationDescription": "Self Assessment Tax paid",
                 "amount": 60000, "transactionDate": "2026-03-14"},
            ],
        },
    },
    "SFT Information": {
        "records": [
            {"informationDescription": "Interest from savings bank",
             "informationSource": "STATE BANK OF INDIA", "amount": 35120,
             "transactionDate": "2026-03-31"},
            {"informationDescription": "Interest from fixed deposit",
             "informationSource": "HDFC BANK", "amount": 85200,
             "transactionDate": "2026-03-31"},
            {"informationDescription": "Dividend from equity shares",
             "informationSource": "TCS LTD", "amount": 42000,
             "transactionDate": "2025-11-10"},
            {"informationDescription": "Salary reported by employer SFT-013",
             "informationSource": "ACME SOFTWARE PVT LTD", "amount": 2210000,
             "transactionDate": "2026-03-31"},
        ],
    },
}


def test_parse_ais_json_and_candidates():
    ais = imports.parse_ais(json.dumps(AIS_DOC).encode(), "ais.json", USER)
    assert ais["assessment_year"] == "2026"
    assert len(ais["entries"]) == 6
    cands = imports.ais_income_candidates(ais)
    kinds = {(c["category"], c["amount"]) for c in cands}
    assert ("interest_savings", 35120) in kinds
    assert ("interest_deposit", 85200) in kinds
    assert ("dividend", 42000) in kinds
    assert ("salary", 2210000) in kinds


def test_ais_zip_password_autodetect():
    password = "bqxps1234f15061985"
    blob = make_encrypted_zip({"ais.json": json.dumps(AIS_DOC).encode()}, password)
    ais = imports.parse_ais(blob, "AIS_bqxps1234f.zip", USER)
    assert len(ais["entries"]) == 6
    # wrong profile -> clean DataError
    bad = {"personal": {"pan": "AAAAA1111A", "dob": "2000-01-01"}}
    with pytest.raises(DataError):
        imports.parse_ais(blob, "AIS.zip", bad)


def test_apply_ais_items_adds_and_dedupes():
    ais = imports.parse_ais(json.dumps(AIS_DOC).encode(), "ais.json", USER)
    cands = imports.ais_income_candidates(ais)
    ids = [c["id"] for c in cands if c["category"] == "interest_savings"]
    out, notes = imports.apply_ais_items({}, cands, ids)
    assert out["other_sources"][0]["nature"] == "SAV"
    assert out["other_sources"][0]["amount"] == 35120
    # applying again must not duplicate
    out2, notes2 = imports.apply_ais_items(out, cands, ids)
    assert len(out2["other_sources"]) == 1


def test_compare_with_ais():
    ais = imports.parse_ais(json.dumps(AIS_DOC).encode(), "ais.json", USER)
    data = {
        "salary": {"components": [{"code": "BASIC", "amount": 2210000}]},
        "other_sources": [{"nature": "SAV", "amount": 35120}],
        "house_properties": [],
    }
    rows = imports.compare_with_ais(data, ais, {})
    by = {r["head"]: r for r in rows}
    assert by["Salary"]["status"] == "matched"
    assert by["Interest (savings + deposits)"]["status"] == "AIS higher — review"
    assert by["Interest (savings + deposits)"]["difference"] == 35120 - 120320
    assert by["Dividend"]["status"] == "missing in return"


# ---------------------------------------------------------------------------
# 26AS
# ---------------------------------------------------------------------------

TXT_26AS = """
Form No. 26AS (Annual Tax Statement)   PAN: BQXPS1234F   A.Y. 2026-27
PART-A1  Details of Tax Deducted at Source on Salary
Sr. No.  Name of Deductor               TAN          Section  Amount Paid  Tax Deducted
1        ACME SOFTWARE PRIVATE LIMITED  SRTE01234A   192      22,10,000    2,40,000

PART-A2  Details of Tax Deducted at Source (Other than Salary)
Sr. No.  Name of Deductor               TAN          Section  Amount Paid  Tax Deducted
2        HDFC BANK LIMITED              MUMB09876B   194A     8,52,000     8,520

PART-B   Details of Tax Collected at Source
Sr. No.  Name of Collector              TAN          Section  Amount       Tax Collected
3        FOREX TOURS LTD                MUMC12345C   206C     8,00,000     4,000

PART-C   Details of Tax Paid (Other than TDS/TCS)
Sr. No.  BSR Code    Date of Deposit    Challan Serial No.   Amount
4        0004321     14-Mar-2026        00101                60,000
"""


def test_parse_26as_text():
    out = imports.parse_26as(TXT_26AS.encode(), "form26as.txt", USER)
    assert out["tds_salary"][0]["tan"] == "SRTE01234A"
    assert out["tds_salary"][0]["tax"] == 240000
    assert out["tds_others"][0]["section"] == "194A"
    assert out["tcs"][0]["tax"] == 4000
    assert out["challans"][0]["amount"] == 60000
    assert out["challans"][0]["date"] == "2026-03-14"


def test_parse_26as_password_pdf():
    from reportlab.pdfgen import canvas
    from pypdf import PdfReader, PdfWriter

    buf = io.BytesIO()
    cv = canvas.Canvas(buf)
    cv.drawString(40, 780, "PART-A2  Details of Tax Deducted at Source (Other than Salary)")
    cv.drawString(40, 760, "1  HDFC BANK LIMITED  MUMB09876B  194A  8,52,000  8,520")
    cv.save()
    w = PdfWriter()
    w.append(PdfReader(buf))
    w.encrypt("15061985", algorithm="AES-256")     # DOB ddmmyyyy
    out = io.BytesIO()
    w.write(out)

    parsed = imports.parse_26as(out.getvalue(), "26as.pdf", USER)
    assert parsed["tds_others"][0]["tan"] == "MUMB09876B"
    assert parsed["password_used"] is True
    with pytest.raises(DataError):
        imports.parse_26as(out.getvalue(), "26as.pdf",
                           {"personal": {"pan": "AAAAA1111A", "dob": "2000-01-01"}})


def test_apply_26as_merges_tax_paid():
    parsed = imports.parse_26as(TXT_26AS.encode(), "form26as.txt", USER)
    out, notes = imports.apply_26as({}, parsed)
    tp = out["tax_paid"]
    assert tp["tds_on_salary"][0]["tds"] == 240000
    assert tp["tds_others"][0]["tds_deducted"] == 8520
    assert tp["tcs"][0]["tcs_collected"] == 4000
    assert tp["challans"][0]["amount"] == 60000
    out2, _ = imports.apply_26as(out, parsed)
    assert len(out2["tax_paid"]["tds_on_salary"]) == 1   # de-duplicated
