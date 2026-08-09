"""End-to-end tests of the V2 web API: form detection, per-form generation,
capital-gain Excel round-trip, imports (ITR JSON / AIS / 26AS), YoY and the
salary helper endpoints."""

import io
import json
import os

import pytest

from itr1.webapp import create_app

import test_imports as TI   # shared fixtures (AIS doc, 26AS text, zip writer)

HERE = os.path.dirname(__file__)


def _sample(name):
    with open(os.path.join(HERE, os.pardir, "samples", name), encoding="utf-8") as fh:
        return json.load(fh)


@pytest.fixture()
def client():
    return create_app().test_client()


@pytest.fixture()
def d1():
    return _sample("sample_taxpayer.json")


@pytest.fixture()
def d2():
    return _sample("sample_itr2.json")


@pytest.fixture()
def d4():
    return _sample("sample_itr4.json")


# ---------------------------------------------------------------------------
# form detection + per-form summary
# ---------------------------------------------------------------------------

def test_detect_recommends_right_form(client, d1, d2, d4):
    for data, want in ((d1, "ITR1"), (d2, "ITR2"), (d4, "ITR4")):
        r = client.post("/api/forms/detect", json=data)
        assert r.status_code == 200
        body = r.get_json()
        assert body["recommended"] == want
        assert set(body["eligible"]) == {"ITR1", "ITR2", "ITR3", "ITR4"}
        assert body["eligible"][want]["ok"] is True


def test_detect_handles_empty(client):
    r = client.post("/api/forms/detect", json={})
    assert r.status_code == 200
    assert r.get_json()["recommended"] == "ITR1"


def test_summary_v2_reports_both_regimes(client, d2, d4):
    r = client.post("/api/summary/v2", json={"data": d2, "form": "ITR2"})
    assert r.status_code == 200
    body = r.get_json()
    assert body["form"] == "ITR2"
    assert body["old"]["total_income"] > 0
    assert body["new"]["total_income"] > 0
    assert body["recommended"] in ("old", "new")

    r = client.post("/api/summary/v2", json={"data": d4, "form": "ITR4"})
    body = r.get_json()
    assert body["form"] == "ITR4"
    # presumptive business shows up in the warnings/summary path, not an error
    assert body["old"]["total_income"] > 0


def test_summary_v2_itr3_friendly_error(client, d1):
    r = client.post("/api/summary/v2", json={"data": d1, "form": "ITR3"})
    assert r.status_code == 400
    assert "ITR-3" in json.dumps(r.get_json())


# ---------------------------------------------------------------------------
# per-form generation / download
# ---------------------------------------------------------------------------

def test_itr2_generate_and_download(client, d2):
    r = client.post("/api/itr2/generate", json=d2)
    assert r.status_code == 200
    body = r.get_json()
    assert body["schema_errors"] == []
    assert body["json"]["ITR"]["ITR2"]["PartA_GEN1"]["PersonalInfo"]["PAN"]
    # Form-10E relief entirely automatic from the arrear component
    assert body["report"]["form10e"]["relief89"] > 0
    assert body["report"]["tax"]["Section89"] == body["report"]["form10e"]["relief89"]

    r = client.post("/api/itr2/download", json=d2)
    assert r.status_code == 200
    assert "attachment" in r.headers.get("Content-Disposition", "")
    doc = json.loads(r.data)
    assert "ITR2" in doc["ITR"]


def test_itr4_generate_and_download(client, d4):
    r = client.post("/api/itr4/generate", json=d4)
    assert r.status_code == 200
    body = r.get_json()
    assert body["schema_errors"] == []
    assert "ITR4" in body["json"]["ITR"]

    r = client.post("/api/itr4/download", json=d4)
    assert r.status_code == 200


def test_itr2_generate_partial_data_not_500(client):
    r = client.post("/api/itr2/generate", json={"personal": {}, "salary": {"salary_17_1": 700000}})
    assert r.status_code in (200, 400)          # never a 500
    if r.status_code == 400:
        assert "error" in r.get_json()


# ---------------------------------------------------------------------------
# capital-gain Excel round trip
# ---------------------------------------------------------------------------

def test_cg_xlsx_round_trip(client, d2):
    r = client.post("/api/cg/xlsx", json=d2)
    assert r.status_code == 200
    assert r.data[:2] == b"PK"
    r = client.post("/api/cg/import-xlsx", data=r.data,
                    headers={"X-Filename": "cg.xlsx"},
                    content_type="application/octet-stream")
    assert r.status_code == 200
    counts = r.get_json()["counts"]
    assert counts.get("equity_lt", 0) >= 1


def test_cg_template_downloads(client):
    r = client.get("/api/cg/template")
    assert r.status_code == 200
    tpl = json.loads(r.data)
    assert "capital_gains" in tpl


# ---------------------------------------------------------------------------
# imports: previous-year ITR JSON
# ---------------------------------------------------------------------------

def test_import_itr_json_prefill(client):
    buf = io.BytesIO(json.dumps(TI.PREV_ITR2).encode())
    r = client.post("/api/import/itr-json",
                    data={"file": (buf, "prev.json"), "data": json.dumps({"personal": {}})},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    body = r.get_json()
    assert body["summary"]["form"] == "ITR2"
    assert body["summary"]["losses"]["stcg_cf"] == 50000
    assert body["merged_data"]["personal"]["dob"] == "1985-06-15"
    assert body["merged_data"]["losses_cf"]["stcg_cf"] == 50000
    assert body["notes"]


def test_import_itr_json_without_file_400(client):
    r = client.post("/api/import/itr-json", data={"data": "{}"},
                    content_type="multipart/form-data")
    assert r.status_code == 400
    assert "error" in r.get_json()


# ---------------------------------------------------------------------------
# imports: AIS (plain JSON + password ZIP) and click-to-add
# ---------------------------------------------------------------------------

def _post_ais(client, blob, filename, data=None):
    return client.post("/api/import/ais",
                       data={"file": (io.BytesIO(blob), filename),
                             "data": json.dumps(data if data is not None else TI.USER)},
                       content_type="multipart/form-data")


def test_import_ais_plain_json(client):
    r = _post_ais(client, json.dumps(TI.AIS_DOC).encode(), "ais.json")
    assert r.status_code == 200
    body = r.get_json()
    assert body["entry_count"] == 6
    assert body["password_detected"] is False
    assert body["categories"]["salary"] == 1
    kinds = {c["category"] for c in body["income_candidates"]}
    assert {"interest_savings", "interest_deposit", "dividend", "salary"} <= kinds
    heads = {row["head"]: row for row in body["comparison"]}
    assert "Salary" in heads and "Dividend" in heads


def test_import_ais_password_zip_autodetect(client):
    blob = TI.make_encrypted_zip({"ais.json": json.dumps(TI.AIS_DOC).encode()},
                                 "bqxps1234f15061985")
    r = _post_ais(client, blob, "AIS_BQXPS1234F.zip")
    assert r.status_code == 200
    body = r.get_json()
    assert body["password_detected"] is True
    assert body["entry_count"] == 6


def test_import_ais_wrong_profile_clean_error(client):
    blob = TI.make_encrypted_zip({"ais.json": json.dumps(TI.AIS_DOC).encode()},
                                 "bqxps1234f15061985")
    r = _post_ais(client, blob, "AIS.zip",
                  {"personal": {"pan": "AAAAA1111A", "dob": "2000-01-01"}})
    assert r.status_code == 400
    assert "password" in json.dumps(r.get_json()).lower()


def test_ais_apply_income_and_credits(client):
    r = _post_ais(client, json.dumps(TI.AIS_DOC).encode(), "ais.json")
    body = r.get_json()
    cands = body["income_candidates"]
    ids = [c["id"] for c in cands if c["category"] in ("interest_savings", "dividend")]
    r = client.post("/api/import/ais/apply",
                    json={"data": {}, "candidates": cands, "income_ids": ids})
    assert r.status_code == 200
    merged = r.get_json()["merged_data"]
    assert len(merged["other_sources"]) == 2

    r = client.post("/api/import/ais/apply",
                    json={"data": {}, "tax_credits": body["tax_credits"]})
    assert r.status_code == 200
    tp = r.get_json()["merged_data"].get("tax_paid", {})
    assert tp.get("tds_on_salary") or tp.get("challans")


# ---------------------------------------------------------------------------
# imports: Form 26AS (TXT and AES-encrypted PDF)
# ---------------------------------------------------------------------------

def test_import_26as_text_with_apply(client):
    buf = io.BytesIO(TI.TXT_26AS.encode())
    r = client.post("/api/import/26as",
                    data={"file": (buf, "form26as.txt"), "data": json.dumps(TI.USER),
                          "apply": "1"},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    body = r.get_json()
    assert body["counts"]["tds_salary"] == 1
    assert body["counts"]["tcs"] == 1
    assert body["counts"]["challans"] == 1
    tp = body["merged_data"]["tax_paid"]
    assert tp["tds_on_salary"][0]["tds"] == 240000
    assert tp["challans"][0]["amount"] == 60000


def test_import_26as_aes_pdf_password_autodetect(client):
    from reportlab.pdfgen import canvas
    from pypdf import PdfReader, PdfWriter

    buf = io.BytesIO()
    cv = canvas.Canvas(buf)
    cv.drawString(40, 780, "PART-A2  Details of Tax Deducted at Source (Other than Salary)")
    cv.drawString(40, 760, "1  HDFC BANK LIMITED  MUMB09876B  194A  8,52,000  8,520")
    cv.save()
    writer = PdfWriter()
    writer.append(PdfReader(buf))
    writer.encrypt("15061985", algorithm="AES-256")         # DOB ddmmyyyy
    out = io.BytesIO()
    writer.write(out)

    r = client.post("/api/import/26as",
                    data={"file": (io.BytesIO(out.getvalue()), "26as.pdf"),
                          "data": json.dumps(TI.USER), "apply": "0"},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    body = r.get_json()
    assert body["password_detected"] is True
    assert body["counts"]["tds_others"] == 1


# ---------------------------------------------------------------------------
# YoY + salary helpers
# ---------------------------------------------------------------------------

def test_yoy_endpoint(client, d2):
    prev = TI.imports.parse_itr_json(TI.PREV_ITR2)
    r = client.post("/api/yoy", json={"data": d2, "prev_summary": prev})
    assert r.status_code == 200
    body = r.get_json()
    assert len(body["rows"]) >= 5
    by = {row["head"]: row for row in body["rows"]}
    assert any("Salary" in h for h in by)
    assert body["prev"]["assessment_year"] == "2025"


def test_salary_monthly_fill_aliases(client):
    rows = ([{"month": f"2025-{m:02d}", "basic": 90000, "hra": 36000,
              "professional_tax": 200, "tds": 9000} for m in range(4, 13)] +
            [{"month": f"2026-{m:02d}", "basic": 95000, "hra": 36000,
              "professional_tax": 200, "tds": 9500} for m in range(1, 4)])
    r = client.post("/api/salary/monthly-fill", json={"salary": {"monthly": rows}})
    assert r.status_code == 200
    body = r.get_json()
    assert body["months_filled"] == 12
    assert body["basic"] == 90000 * 9 + 95000 * 3
    assert body["hra_received"] == 36000 * 12          # 'hra' alias accepted
    assert body["professional_tax"] == 2400            # 'professional_tax' alias


def test_salary_relief89_from_component_arrears(client, d2):
    r = client.post("/api/salary/relief89", json={"data": d2})
    assert r.status_code == 200
    body = r.get_json()
    assert body["relief89"] > 0
    assert body["rows"][0]["fy"] == "2024-25"
    assert body["rows"][0]["amount"] == 150000


def test_salary_catalog(client):
    r = client.get("/api/salary/catalog")
    assert r.status_code == 200
    body = r.get_json()
    assert "BASIC" in body["components"]
    assert "44ad" in body["business_codes"]
    assert "44ada" in body["business_codes"]
    assert body["perquisites"]["RFA"]
    assert "gratuity" in body["terminal"]
