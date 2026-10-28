from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from fastapi.testclient import TestClient
from pypdf import PdfReader

import app as app_module
from app import Settings, create_app
from integrations import JobRecord, LocalJobTracker, LocalStorage, utc_now
from report_generator import InputValidationError, analyze_attendance, generate_pdf, process_attendance_file


def attendance_rows() -> pd.DataFrame:
    return pd.DataFrame([
        {"Panchayat LGD": 1, "Mandal Name": "Mandal A", "Panchayat Name": "GP One",
         "First Attendence Status": "Present", "First Attendence Datetime": "28-10-2026 09:00"},
        {"Panchayat LGD": 2, "Mandal Name": "Mandal A", "Panchayat Name": "GP Two",
         "First Attendence Status": "Present", "First Attendence Datetime": "28-10-2026 08:59"},
        {"Panchayat LGD": 3, "Mandal Name": "Mandal A", "Panchayat Name": "GP Three",
         "First Attendence Status": "Present", "First Attendence Datetime": "28-10-2026 09:01"},
        {"Panchayat LGD": 4, "Mandal Name": "Mandal B", "Panchayat Name": "GP Four",
         "First Attendence Status": "", "First Attendence Datetime": ""},
        {"Panchayat LGD": 5, "Mandal Name": "Mandal B", "Panchayat Name": "GP Five",
         "First Attendence Status": "On Leave", "First Attendence Datetime": "28-10-2026 08:30"},
        {"Panchayat LGD": 6, "Mandal Name": "Mandal B", "Panchayat Name": "GP Six",
         "First Attendence Status": "Training", "First Attendence Datetime": "28-10-2026 09:05"},
        {"Panchayat LGD": 7, "Mandal Name": "Mandal B", "Panchayat Name": "GP Seven",
         "First Attendence Status": "", "First Attendence Datetime": ""},
        {"Panchayat LGD": 7, "Mandal Name": "Mandal C", "Panchayat Name": "GP Seven Conflict",
         "First Attendence Status": "Meeting", "First Attendence Datetime": "28-10-2026 09:00"},
    ])


def bytes_for(frame: pd.DataFrame, extension: str) -> bytes:
    buffer = BytesIO()
    if extension == ".xlsx":
        frame.to_excel(buffer, index=False, engine="openpyxl")
        return buffer.getvalue()
    if extension == ".csv":
        return frame.to_csv(index=False).encode("utf-8")
    if extension == ".xls":
        return ("<html><body>" + frame.to_html(index=False) + "</body></html>").encode("utf-8")
    raise AssertionError(extension)


def make_settings(tmp_path: Path) -> Settings:
    return Settings(
        api_token="test-only-token-value-that-is-at-least-32-chars",
        data_dir=tmp_path / "data",
        local_input_dir=None,
        max_upload_bytes=30 * 1024 * 1024,
        credentials_path="",
        drive_upload_folder_id="",
        drive_output_folder_id="",
        spreadsheet_id="",
        spreadsheet_tab="AttendanceReportJobs",
        tasks_project="",
        tasks_location="",
        tasks_queue="",
        public_base_url="",
        task_invoker_service_account="",
    )


def test_valid_xlsx_generates_unique_mandal_abstract_and_categories():
    report = process_attendance_file("synthetic.xlsx", bytes_for(attendance_rows(), ".xlsx"))

    assert report.unique_gp_count == 7
    assert report.duplicate_lgd_rows == 1
    assert len(report.conflicts) == 1
    assert report.conflicts[0]["lgd"] == "7"
    assert report.conflicts[0]["mandals"] == ["Mandal B", "Mandal C"]
    assert report.totals == {
        "unique_gps": 7,
        "before_9": 4,
        "on_leave": 1,
        "meeting_training": 2,
        "pending": 1,
        "after_9": 2,
    }
    assert [row["mandal"] for row in report.mandal_rows] == ["Mandal B", "Mandal A"]
    assert report.district_total["before_9_pct"] == pytest.approx(4 / 7 * 100)
    assert report.mandal_rows[0]["before_9_pct"] == 50


def test_csv_with_header_variations_is_supported():
    frame = attendance_rows().iloc[:1].rename(columns={
        "Panchayat LGD": " panchayat-lgd ",
        "Mandal Name": "MANDAL NAME",
        "First Attendence Status": "First Attendance Status",
        "First Attendence Datetime": "First Attendance Date Time",
    })
    report = process_attendance_file("synthetic.csv", bytes_for(frame, ".csv"))
    assert report.unique_gp_count == 1
    assert report.totals["before_9"] == 1


def test_literal_na_status_is_submitted_not_pending():
    frame = attendance_rows().iloc[:1].copy()
    frame.loc[frame.index[0], "First Attendence Status"] = "NA"
    frame.loc[frame.index[0], "First Attendence Datetime"] = ""
    report = process_attendance_file("synthetic.csv", bytes_for(frame, ".csv"))
    assert report.totals["pending"] == 0
    assert report.totals["before_9"] == 0
    assert report.after_9_rows == []


def test_html_table_saved_as_xls_is_supported():
    report = process_attendance_file("portal-export.xls", bytes_for(attendance_rows().iloc[:2], ".xls"))
    assert report.unique_gp_count == 2
    assert report.totals["before_9"] == 2


def test_missing_required_columns_return_a_useful_error():
    source = pd.DataFrame({"Mandal": ["A"], "GP": ["one"]}).to_csv(index=False).encode()
    with pytest.raises(InputValidationError, match=r"Required column\(s\) not found") as error:
        process_attendance_file("bad.csv", source)
    assert "Panchayat LGD" in str(error.value)
    assert "First Attendence Datetime" in str(error.value)


def test_duplicate_lgd_prefers_a_nonblank_status_and_flags_conflicting_names():
    report = analyze_attendance(attendance_rows())
    assert report.unique_gp_count == 7
    assert report.duplicate_lgd_rows == 1
    assert report.conflicts[0]["mandals"] == ["Mandal B", "Mandal C"]
    assert report.conflicts[0]["panchayats"] == ["GP Seven", "GP Seven Conflict"]
    meeting = next(row for row in report.meeting_training_rows if row["gp"] == "GP Seven")
    assert meeting["status"] == "Meeting"
    assert meeting["timestamp"]


def test_time_boundary_blank_leave_meeting_training_and_after_nine_rules():
    report = analyze_attendance(attendance_rows())
    after_nine = {row["gp"] for row in report.after_9_rows}
    assert "GP One" not in after_nine  # 09:00 exactly belongs to before-9.
    assert "GP Two" not in after_nine  # 08:59.
    assert {"GP Three", "GP Six"} == after_nine
    assert report.totals["pending"] == 1
    assert [row["gp"] for row in report.leave_rows] == ["GP Five"]
    assert {row["gp"] for row in report.meeting_training_rows} == {"GP Seven", "GP Six"}
    assert all(row["status"] for row in report.leave_rows + report.meeting_training_rows)


def test_empty_sections_are_rendered_and_district_totals_remain_correct(tmp_path: Path):
    frame = attendance_rows().iloc[[0]].copy()
    report = analyze_attendance(frame)
    assert report.totals["unique_gps"] == 1
    assert report.district_total["total_gps"] == 1
    assert report.district_total["before_9"] == 1
    assert report.pending_rows == []
    assert report.leave_rows == []
    assert report.meeting_training_rows == []
    assert report.after_9_rows == []

    pdf_path = generate_pdf(report, tmp_path / "empty-sections.pdf")
    reader = PdfReader(str(pdf_path))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "KHAMMAM DISTRICT" in text
    assert text.count("No records found for this section.") == 4
    assert "DISTRICT TOTAL" in text


def test_pdf_generation_includes_all_report_section_titles(tmp_path: Path):
    report = analyze_attendance(attendance_rows())
    pdf_path = generate_pdf(report, tmp_path / "attendance.pdf")
    reader = PdfReader(str(pdf_path))
    assert len(reader.pages) >= 5
    assert float(reader.pages[0].mediabox.width) == pytest.approx(595.28, abs=0.1)
    assert float(reader.pages[0].mediabox.height) == pytest.approx(841.89, abs=0.1)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Mandal-wise Attendance Abstract" in text
    assert "ATTENDANCE PENDING LIST" in text
    assert "PS ON LEAVE LIST" in text
    assert "TRAINING / MEETING LIST" in text
    assert "AFTER 9 AM" in text
    assert "GP Seven Conflict" in text or "conflicting names" in text


def test_mandal_abstract_repeats_column_headers_when_continued(tmp_path: Path):
    frame = pd.DataFrame([
        {
            "Panchayat LGD": index + 1,
            "Mandal Name": f"Mandal {index:02d}",
            "Panchayat Name": f"GP {index:02d}",
            "First Attendence Status": "Present",
            "First Attendence Datetime": "28-10-2026 08:00",
        }
        for index in range(70)
    ])
    report = analyze_attendance(frame)
    pdf_path = generate_pdf(report, tmp_path / "continued-abstract.pdf")
    reader = PdfReader(str(pdf_path))
    assert len(reader.pages) > 5
    # The mandal abstract extends beyond page one, so its repeated header is
    # visible at the beginning of page two.
    page_two = reader.pages[1].extract_text()
    assert "Sl. No." in page_two
    assert "No. of GPs" in page_two


def test_api_upload_status_download_auth_and_duplicate_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    config = make_settings(tmp_path)
    api = create_app(config)
    client = TestClient(api)
    monkeypatch.setattr(app_module, "utc_now", lambda: "2026-10-28T03:04:05+00:00")
    token = {"x-api-token": config.api_token}
    workbook = bytes_for(attendance_rows(), ".xlsx")
    calls = 0
    real_process = app_module.process_attendance_file

    def count_process(filename: str, contents: bytes):
        nonlocal calls
        calls += 1
        return real_process(filename, contents)

    monkeypatch.setattr(app_module, "process_attendance_file", count_process)
    response = client.post(
        "/api/reports",
        headers=token,
        data={"JobID": "job-test-001", "RequestedBy": "tester@example.org"},
        files={"file": ("synthetic.xlsx", workbook, "application/octet-stream")},
    )
    assert response.status_code == 200
    assert response.json()["status"] == "Completed"
    assert response.json()["job_id"] == "job-test-001"
    assert calls == 1

    status = client.get("/api/reports/job-test-001", headers=token)
    assert status.status_code == 200
    assert status.json()["status"] == "Completed"
    download = client.get("/api/reports/job-test-001/download", headers=token)
    assert download.status_code == 200
    assert download.headers["content-type"] == "application/pdf"
    assert download.headers["content-disposition"] == (
        'attachment; filename="PS_Attendance_Report_2026-10-28_08-34-05.pdf"'
    )
    assert download.content.startswith(b"%PDF")
    assert (config.data_dir / "reports" / "PS_Attendance_Report_2026-10-28_08-34-05.pdf").is_file()

    duplicate = client.post(
        "/api/reports",
        headers=token,
        data={"JobID": "job-test-001", "RequestedBy": "tester@example.org"},
        files={"file": ("synthetic.xlsx", b"not an xlsx", "application/octet-stream")},
    )
    assert duplicate.status_code == 200
    assert duplicate.json()["status"] == "Completed"
    assert calls == 1

    assert client.get("/api/reports/job-test-001").status_code == 401
    assert client.get("/api/reports/job-test-001/download", headers={"x-api-token": "wrong"}).status_code == 401
    unauthorized_upload = client.post(
        "/api/reports", data={"JobID": "unauthorized"},
        files={"file": ("synthetic.xlsx", workbook, "application/octet-stream")},
    )
    assert unauthorized_upload.status_code == 401


def test_api_rejects_invalid_types_and_reports_missing_columns(tmp_path: Path):
    config = make_settings(tmp_path)
    client = TestClient(create_app(config))
    token = {"x-api-token": config.api_token}

    invalid_type = client.post(
        "/api/reports", headers=token, data={"JobID": "bad-type"},
        files={"file": ("payload.exe", b"data", "application/octet-stream")},
    )
    assert invalid_type.status_code == 415

    bad_csv = b"Mandal,GP\nA,One\n"
    missing = client.post(
        "/api/reports", headers=token, data={"JobID": "bad-columns"},
        files={"file": ("bad.csv", bad_csv, "text/csv")},
    )
    assert missing.status_code == 422
    assert "Required column(s) not found" in missing.json()["detail"]["message"]
    assert missing.json()["detail"]["job_id"] == "bad-columns"


def test_appsheet_submitted_sheet_row_is_claimed_and_processed(tmp_path: Path):
    config = make_settings(tmp_path)
    tracker = LocalJobTracker(config.data_dir / "attendance_jobs.sqlite3")
    storage = LocalStorage(config.data_dir)
    source_path = storage.input_dir / "attendance.xlsx"
    source_path.write_bytes(bytes_for(attendance_rows().iloc[:2], ".xlsx"))
    tracker.create(JobRecord(
        job_id="appsheet-job-1",
        uploaded_file="attendance.xlsx",
        requested_by="tester@example.org",
        requested_at=utc_now(),
        status="Submitted",
    ))

    client = TestClient(create_app(config, tracker=tracker, storage=storage))
    response = client.post(
        "/api/reports/from-drive",
        headers={"x-api-token": config.api_token},
        json={
            "JobID": "appsheet-job-1",
            "UploadedFile": "AttendanceUploads/attendance.xlsx",
            "RequestedBy": "tester@example.org",
        },
    )
    assert response.status_code == 200
    assert response.json()["status"] == "Completed"
    assert tracker.get("appsheet-job-1").status == "Completed"


def test_from_drive_requires_stable_appsheet_job_id(tmp_path: Path):
    config = make_settings(tmp_path)
    client = TestClient(create_app(config))
    response = client.post(
        "/api/reports/from-drive",
        headers={"x-api-token": config.api_token},
        json={"UploadedFile": "attendance.xlsx"},
    )
    assert response.status_code == 400
    assert "JobID" in response.json()["detail"]


def test_health_is_public_and_unconfigured_auth_fails_closed(tmp_path: Path):
    config = make_settings(tmp_path)
    config = Settings(**{**config.__dict__, "api_token": ""})
    client = TestClient(create_app(config))
    assert client.get("/health").status_code == 200
    protected = client.get("/api/reports/missing")
    assert protected.status_code == 503


def test_root_page_opens_directly_with_upload_and_download_workflow(tmp_path: Path):
    client = TestClient(create_app(make_settings(tmp_path)))

    response = client.get("/")

    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "Khammam PS Attendance Analytics" in response.text
    assert 'id="login-form"' not in response.text
    assert "APP_API_TOKEN" not in response.text
    assert 'id="upload-form"' in response.text
    assert 'accept=".xls,.xlsx,.csv"' in response.text
    assert "Generate and download report" in response.text
    assert "/api/web/reports/" in response.text
    assert "Content-Disposition" in response.text
    assert "Khammam_PS_Attendance_Report_${jobId}" not in response.text


def test_public_web_upload_download_works_without_token_and_keeps_job_private(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    config = make_settings(tmp_path)
    client = TestClient(create_app(config), base_url="https://testserver")
    monkeypatch.setattr(app_module, "utc_now", lambda: "2026-10-28T03:04:05+00:00")

    upload = client.post(
        "/api/web/reports",
        data={"RequestedBy": "tester"},
        files={"file": ("synthetic.xlsx", bytes_for(attendance_rows(), ".xlsx"), "application/octet-stream")},
    )
    assert upload.status_code == 200
    job_id = upload.json()["job_id"]
    assert upload.json()["status"] == "Completed"
    cookie = upload.headers["set-cookie"]
    assert "httponly" in cookie.lower()
    assert "secure" in cookie.lower()
    assert "samesite=strict" in cookie.lower()
    assert client.get(f"/api/web/reports/{job_id}").status_code == 200

    download = client.get(f"/api/web/reports/{job_id}/download")
    assert download.status_code == 200
    assert download.content.startswith(b"%PDF")
    assert download.headers["content-disposition"] == (
        'attachment; filename="PS_Attendance_Report_2026-10-28_08-34-05.pdf"'
    )

    other_client = TestClient(create_app(make_settings(tmp_path / "other")), base_url="https://testserver")
    assert other_client.get(f"/api/web/reports/{job_id}").status_code == 401
    assert other_client.get(f"/api/web/reports/{job_id}/download").status_code == 401
    assert client.get(f"/api/reports/{job_id}").status_code == 401


def test_public_web_upload_endpoint_limits_requests_by_client(tmp_path: Path):
    client = TestClient(create_app(make_settings(tmp_path)))
    invalid_upload = {
        "files": {"file": ("not-allowed.exe", b"no executable", "application/octet-stream")},
    }

    for _ in range(app_module.WEB_UPLOADS_PER_HOUR):
        assert client.post("/api/web/reports", **invalid_upload).status_code == 415
    limited = client.post("/api/web/reports", **invalid_upload)
    assert limited.status_code == 429
    assert limited.headers["retry-after"] == "3600"


def test_local_tracker_claims_once_and_recovers_a_stale_processing_job(tmp_path: Path):
    tracker = LocalJobTracker(tmp_path / "jobs.sqlite3")
    tracker.create(JobRecord(
        job_id="claim-test",
        uploaded_file="attendance.xlsx",
        requested_by="tester@example.org",
        requested_at=utc_now(),
    ))
    assert tracker.claim("claim-test") is True
    assert tracker.claim("claim-test") is False
    tracker.update("claim-test", notes="Processing started at 2000-01-01T00:00:00+00:00")
    assert tracker.claim("claim-test") is True
