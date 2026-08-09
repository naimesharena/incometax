"""End-to-end tests of the web API, including partial-input robustness."""

import json
import os

import pytest

from itr1.webapp import create_app

SAMPLE = os.path.join(os.path.dirname(__file__), os.pardir,
                      "samples", "sample_taxpayer.json")


@pytest.fixture()
def client():
    app = create_app()
    return app.test_client()


@pytest.fixture()
def sample():
    with open(SAMPLE, encoding="utf-8") as fh:
        return json.load(fh)


def test_summary_with_partial_data_no_dob(client):
    """Original bug: typing salary before DOB crashed the live computation."""
    partial = {"salary": {"salary_17_1": 900000}, "personal": {},
               "filing": {"opt_out_new_regime": False}}
    r = client.post("/api/summary", json=partial)
    assert r.status_code == 200
    body = r.get_json()
    assert body["new"]["total_income"] == 825_000
    assert body["new"]["gross_tax_liability"] == 0      # rebate up to 12L
    assert "issues" in body and "issue_summary" in body


def test_generate_with_partial_data_returns_json_not_500(client):
    partial = {"salary": {"salary_17_1": 900000}, "personal": {}, "filing": {}}
    r = client.post("/api/generate", json=partial)
    assert r.status_code == 200
    body = r.get_json()
    assert "issues" in body and "json" in body


def test_summary_full_sample_and_live_issues(client, sample):
    r = client.post("/api/summary", json=sample)
    assert r.status_code == 200
    body = r.get_json()
    assert body["recommended"] == "new"
    assert body["chosen"] == "new"
    assert body["new"]["refund_due"] == 36_500
    assert body["issue_summary"]["ok"] is True


def test_summary_live_issue_detection(client, sample):
    # force a violation: new regime with a belated-return old-regime conflict
    sample["filing"]["opt_out_new_regime"] = True
    sample["filing"]["return_section"] = 12          # belated return cannot use old regime
    sample["filing"]["filing_date"] = "2026-08-02"
    r = client.post("/api/summary", json=sample)
    assert r.status_code == 200
    body = r.get_json()
    rules = [i["rule"] for i in body["issues"]]
    assert any("151" in rl for rl in rules)
    assert body["issue_summary"]["A"] >= 1


def test_computation_endpoint(client, sample):
    r = client.post("/api/computation", json=sample)
    assert r.status_code == 200
    body = r.get_json()
    assert body["comparison"]["new_total"] == 0
    assert body["new"]["total_income"] == 912_500
    assert body["old"]["via_rows"]


def test_pdf_download(client, sample):
    r = client.post("/api/computation/pdf", json=sample)
    assert r.status_code == 200
    assert r.data.startswith(b"%PDF")
    assert "filename=ITR1_Computation_AY2026-27_BQKPS1234F.pdf" in r.headers["Content-Disposition"]


def test_docx_download(client, sample):
    r = client.post("/api/computation/docx", json=sample)
    assert r.status_code == 200
    assert r.data.startswith(b"PK")
    assert r.headers["Content-Type"].startswith(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


def test_pdf_endpoint_partial_data(client):
    r = client.post("/api/computation/pdf",
                    json={"salary": {"salary_17_1": 700000}, "personal": {}, "filing": {}})
    assert r.status_code == 200
    assert r.data.startswith(b"%PDF")


def test_download_json_still_works(client, sample):
    r = client.post("/api/download", json=sample)
    assert r.status_code == 200
    doc = json.loads(r.data.decode())
    assert "ITR" in doc


def test_index_and_static(client):
    assert client.get("/").status_code == 200
    assert client.get("/static/style.css").status_code == 200
    assert client.get("/static/app.js").status_code == 200
