# Khammam PS Attendance Analytics

A FastAPI web application for uploading Gram Panchayat attendance files and downloading formatted A4 PDF reports. The application uses local SQLite for report history and the webserver filesystem for generated PDFs.

## Features

- Accepts .xls, .xlsx, .csv, and HTML tables saved as .xls.
- Deduplicates Gram Panchayat records by Panchayat LGD and flags conflicting names.
- Shows both submitted-attendance percentage and attendance-before-9-AM percentage in the Mandal abstract.
- Builds a DSR abstract using GPs where `First Attendence Status` or `Second Attendence Status` is `Present`.
- Counts DSR submissions only for eligible present PSs; Leave, Meeting, Training, and other statuses are excluded.
- Sorts the DSR abstract by DSR-pending GP count in ascending order.
- Places the DSR-pending GP list immediately after the attendance-pending list.
- Downloads the PDF automatically as PS_Attendance_Report_YYYY-MM-DD_HH-MM-SS.pdf.
- Keeps each report download private to the browser session that uploaded it.

The uploaded source must include `First Attendence Status`, `Second Attendence Status`, `DSR Entry Status`, and the existing Panchayat LGD, Mandal Name, Panchayat Name, and attendance timestamp fields. Column names are matched case-insensitively and common DSR status/date aliases are accepted. If no DSR field can be detected, the upload returns a validation error instead of producing misleading DSR totals.

## Run locally on Windows

Use Python 3.11 or newer. In PowerShell, from the repository folder:

~~~powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements-dev.txt
Copy-Item .env.example .env
~~~

Generate a session secret and add it to .env as WEB_SESSION_SECRET:

~~~powershell
python -c "import secrets; print(secrets.token_hex(32))"
~~~

Start the server:

~~~powershell
uvicorn app:app --reload --port 8080 --env-file .env
~~~

Open http://127.0.0.1:8080/ and upload an attendance file. The PDF downloads when report generation completes. /health provides the service health check; API documentation routes are disabled.

Jobs are stored in data/attendance_jobs.sqlite3; generated reports are stored in data/reports/. Set DATA_DIR in .env to move this local webserver data.

## Test

Tests use synthetic attendance data only:

~~~powershell
pytest -q
~~~

## Docker

Build and run the webserver locally:

~~~powershell
docker build -t khammam-ps-attendance .
docker run --rm -p 8080:8080 -e WEB_SESSION_SECRET=replace-with-a-random-secret-of-at-least-32-characters -e DATA_DIR=/tmp/khammam-attendance-data khammam-ps-attendance
~~~

## Deploy to Render

The repository includes a render.yaml Blueprint for the kmmdsr web service on the Free plan. For exact deployment details and storage limits, see README_Render.md.

Render Free uses ephemeral local storage. Reports and SQLite job history can be removed on restart, redeploy, or spin-down. Download reports and save copies. For durable server-side storage, attach an appropriate persistent storage solution and review its current cost before enabling it.

## Files

- app.py — browser upload, report job handling, session cookies, and health endpoint.
- report_generator.py — spreadsheet parsing, attendance analytics, and PDF generation.
- integrations.py — local SQLite and filesystem storage only.
- tests/ — synthetic-data processing, PDF, and web upload tests.
- Dockerfile, render.yaml, .env.example — container, Render, and runtime configuration.
