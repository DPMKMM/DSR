"""FastAPI webserver for Khammam PS attendance reports."""

from __future__ import annotations

import hmac
import hashlib
import json
import logging
import os
import re
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

from integrations import (
    ALLOWED_EXTENSIONS,
    JobRecord,
    JobTracker,
    LocalJobTracker,
    LocalStorage,
    Storage,
    utc_now,
)
from report_generator import InputValidationError, generate_pdf, process_attendance_file


logger = logging.getLogger("khammam_attendance")
logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO").upper(), format="%(message)s")
MAX_UPLOAD_BYTES = 30 * 1024 * 1024
JOB_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,100}$")
WEB_JOB_COOKIE_NAME = "khammam_web_report"
WEB_JOB_TTL_SECONDS = 60 * 60
WEB_UPLOADS_PER_HOUR = 20
REPORT_FILENAME_PATTERN = re.compile(r"^PS_Attendance_Report_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.pdf$")

UPLOAD_PAGE_HTML = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Khammam PS Attendance Analytics</title>
  <style>
    :root { color-scheme: light; font-family: Inter, Segoe UI, Arial, sans-serif; }
    body { margin: 0; background: #f3f7f5; color: #15332b; }
    main { max-width: 780px; margin: 6vh auto; padding: 0 20px; }
    .card { background: white; border: 1px solid #d9e6df; border-radius: 18px;
      padding: clamp(24px, 5vw, 42px); box-shadow: 0 14px 40px #173b2b12; }
    .eyebrow { color: #247558; font-weight: 700; letter-spacing: .08em;
      text-transform: uppercase; font-size: .8rem; }
    h1 { margin: 12px 0; font-size: clamp(2rem, 5vw, 3rem); line-height: 1.1; }
    p { line-height: 1.6; color: #486158; }
    label { display: block; margin: 18px 0 7px; font-weight: 650; }
    input[type="file"], input[type="text"] { width: 100%;
      box-sizing: border-box; padding: 12px; border: 1px solid #bfd1c7;
      border-radius: 9px; background: white; color: #15332b; }
    button, a.button { display: inline-block; border: 0; border-radius: 10px;
      background: #167451; color: white; padding: 12px 18px; font-weight: 700;
      text-decoration: none; cursor: pointer; }
    button:disabled { opacity: .6; cursor: wait; }
    .actions { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-top: 22px; }
    a.secondary, button.secondary { background: #eaf3ee; color: #176247; }
    .notice { margin: 20px 0; padding: 12px 14px; border-radius: 10px;
      background: #fff5dc; color: #684f08; line-height: 1.5; }
    .message { min-height: 1.5em; margin-top: 14px; color: #8b2635; }
    .message.success { color: #176247; }
    .fine-print { border-top: 1px solid #e5eee9; margin-top: 24px; padding-top: 18px; font-size: .92rem; }
  </style>
</head>
<body>
  <main>
    <section class="card">
      <div class="eyebrow">Khammam District · Telangana</div>
      <h1>PS Attendance Analytics</h1>
      <p>Select the latest attendance and DSR Excel file. The service will create your formatted report and download the PDF.</p>
      <div class="notice">
        Reports and job history are stored on this webserver. If hosted on Render Free, local files can be cleared by a restart or redeploy, so download and save each report.
      </div>

      <form id="upload-form">
        <label for="attendance-file">Attendance data file</label>
        <input id="attendance-file" name="file" type="file" required>
        <label for="requested-by">Your name or office (optional)</label>
        <input id="requested-by" name="RequestedBy" type="text" maxlength="320"
          autocomplete="name" placeholder="Name or office">
        <div class="actions"><button id="generate-button" type="submit">Generate and download report</button></div>
        <div id="report-message" class="message" role="status" aria-live="polite"></div>
      </form>

      <p class="fine-print">Choose a file from your device. Supported report data: Excel (.xls, .xlsx), CSV, or an HTML table saved as .xls. The report downloads automatically when processing finishes.</p>
    </section>
  </main>
  <script>
    (() => {
      const uploadForm = document.getElementById('upload-form');
      const generateButton = document.getElementById('generate-button');
      const reportMessage = document.getElementById('report-message');

      function showMessage(element, message, success = false) {
        element.textContent = message;
        element.classList.toggle('success', success);
      }

      async function readResponse(response) {
        const body = await response.json().catch(() => ({}));
        if (!response.ok) {
          const detail = body.detail;
          const message = typeof detail === 'string' ? detail : (detail?.message || 'The request failed.');
          throw new Error(message);
        }
        return body;
      }

      async function apiJson(url, options = {}) {
        const response = await fetch(url, { credentials: 'same-origin', ...options });
        return readResponse(response);
      }

      uploadForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        const fileInput = document.getElementById('attendance-file');
        const file = fileInput.files[0];
        if (!file) return;
        generateButton.disabled = true;
        showMessage(reportMessage, 'Uploading and processing the attendance file…');
        try {
          const data = new FormData();
          data.append('file', file);
          const requestedBy = document.getElementById('requested-by').value.trim();
          if (requestedBy) data.append('RequestedBy', requestedBy);
          let job = await apiJson('/api/web/reports', { method: 'POST', body: data });
          const jobId = job.job_id;
          if (!jobId) throw new Error('The service did not return a report job ID.');
          let attempts = 0;
          while (job.status === 'Submitted' || job.status === 'Processing') {
            showMessage(reportMessage, `Report ${job.status.toLowerCase()}…`);
            if (attempts++ > 180) throw new Error('The report is taking too long. Please try again later.');
            await new Promise(resolve => setTimeout(resolve, 2000));
            job = await apiJson(`/api/web/reports/${encodeURIComponent(jobId)}`);
          }
          if (job.status !== 'Completed') throw new Error(job.notes || 'The report job did not complete.');
          showMessage(reportMessage, 'Report ready. Downloading PDF…', true);
          const pdfResponse = await fetch(`/api/web/reports/${encodeURIComponent(jobId)}/download`, {
            credentials: 'same-origin'
          });
          if (!pdfResponse.ok) await readResponse(pdfResponse);
          const blob = await pdfResponse.blob();
          const objectUrl = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = objectUrl;
          const disposition = pdfResponse.headers.get('Content-Disposition') || '';
          const fileName = /filename="?([^";]+)"?/i.exec(disposition)?.[1];
          link.download = fileName || 'PS_Attendance_Report.pdf';
          document.body.appendChild(link);
          link.click();
          link.remove();
          setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
          showMessage(reportMessage, 'Report downloaded. You can upload another file.', true);
          uploadForm.reset();
        } catch (error) {
          showMessage(reportMessage, error.message || 'Report generation failed.');
        } finally {
          generateButton.disabled = false;
        }
      });

    })();
  </script>
</body>
</html>"""


@dataclass(frozen=True)
class Settings:
    session_secret: str
    data_dir: Path
    max_upload_bytes: int

    @classmethod
    def from_env(cls) -> "Settings":
        try:
            max_bytes = int(os.getenv("MAX_UPLOAD_BYTES", str(MAX_UPLOAD_BYTES)))
        except ValueError as exc:
            raise RuntimeError("MAX_UPLOAD_BYTES must be an integer.") from exc
        if max_bytes < 1 or max_bytes > 100 * 1024 * 1024:
            raise RuntimeError("MAX_UPLOAD_BYTES must be between 1 byte and 100 MB.")
        session_secret = os.getenv("WEB_SESSION_SECRET", "").strip()
        if not session_secret:
            # Keep existing Render deployments working while the environment
            # variable is migrated from the former API-token name.
            session_secret = os.getenv("APP_API_TOKEN", "").strip()
        return cls(
            session_secret=session_secret,
            data_dir=Path(os.getenv("DATA_DIR", "./data")).resolve(),
            max_upload_bytes=max_bytes,
        )
def _clean_filename(value: str | None) -> str:
    raw = (value or "attendance.xls").replace("\\", "/").split("/")[-1].strip()
    return raw or "attendance.xls"


def _validate_job_id(value: str | None) -> str:
    candidate = (value or uuid4().hex).strip()
    if not JOB_ID_PATTERN.fullmatch(candidate):
        raise HTTPException(status_code=400, detail="JobID must contain 1-100 letters, numbers, underscores, or hyphens.")
    return candidate


def _make_report_filename(report_date: date, requested_at: str) -> str:
    try:
        uploaded_at = datetime.fromisoformat(requested_at.replace("Z", "+00:00"))
    except (AttributeError, TypeError, ValueError):
        uploaded_at = datetime.now(timezone.utc)
    if uploaded_at.tzinfo is None:
        uploaded_at = uploaded_at.replace(tzinfo=timezone.utc)
    india_time = uploaded_at.astimezone(timezone(timedelta(hours=5, minutes=30)))
    return f"PS_Attendance_Report_{report_date:%Y-%m-%d}_{india_time:%H-%M-%S}.pdf"


def _record_report_filename(record: JobRecord) -> str:
    match = re.search(r"(?:^|;\s*)Download filename: ([A-Za-z0-9_-]+\.pdf)(?:$|;)", record.notes)
    if match and REPORT_FILENAME_PATTERN.fullmatch(match.group(1)):
        return match.group(1)
    if record.local_pdf_path:
        candidate = Path(record.local_pdf_path).name
        if REPORT_FILENAME_PATTERN.fullmatch(candidate):
            return candidate
    return "PS_Attendance_Report.pdf"


def _safe_error(exc: Exception) -> tuple[int, str]:
    if isinstance(exc, InputValidationError):
        return 422, str(exc)
    if isinstance(exc, PermissionError):
        return 403, str(exc)
    if isinstance(exc, FileNotFoundError):
        return 404, str(exc)
    if isinstance(exc, ValueError):
        return 400, str(exc)
    return 500, "Report processing failed. Check the service logs for details."


def _read_upload(file: UploadFile, max_bytes: int) -> bytes:
    filename = _clean_filename(file.filename)
    if Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=415, detail="Allowed file types are .xls, .xlsx, and .csv.")
    output = bytearray()
    while True:
        chunk = file.file.read(1024 * 1024)
        if not chunk:
            break
        output.extend(chunk)
        if len(output) > max_bytes:
            raise HTTPException(status_code=413, detail=f"File exceeds the {max_bytes // (1024 * 1024)} MB upload limit.")
    if not output:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    return bytes(output)


def create_app(
    settings: Settings | None = None,
    tracker: JobTracker | None = None,
    storage: Storage | None = None,
) -> FastAPI:
    config = settings or Settings.from_env()
    if config.session_secret and len(config.session_secret) < 32:
        raise RuntimeError("WEB_SESSION_SECRET must contain at least 32 characters.")

    storage = storage or LocalStorage(config.data_dir)
    tracker = tracker or LocalJobTracker(config.data_dir / "attendance_jobs.sqlite3")
    api = FastAPI(
        title="Khammam PS Attendance Analytics",
        version="3.0.0",
        description="Upload attendance files and download formatted reports.",
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )
    api.state.settings = config
    api.state.tracker = tracker
    api.state.storage = storage
    api.state.web_upload_attempts = {}

    @api.get("/", response_class=HTMLResponse, include_in_schema=False)
    def home() -> HTMLResponse:
        return HTMLResponse(UPLOAD_PAGE_HTML)

    @api.middleware("http")
    async def limit_content_length(request: Request, call_next: Any) -> Response:
        if request.url.path == "/api/web/reports" and request.method == "POST":
            raw_length = request.headers.get("content-length")
            if raw_length:
                try:
                    if int(raw_length) > config.max_upload_bytes + 1024 * 1024:
                        return JSONResponse(status_code=413, content={"detail": "Multipart request is too large."})
                except ValueError:
                    return JSONResponse(status_code=400, content={"detail": "Invalid Content-Length header."})
        return await call_next(request)

    def _check_web_upload_rate_limit(request: Request) -> None:
        now = time.monotonic()
        cutoff = now - 60 * 60
        attempts_by_ip: dict[str, list[float]] = api.state.web_upload_attempts
        for address, attempts in list(attempts_by_ip.items()):
            recent = [attempt for attempt in attempts if attempt > cutoff]
            if recent:
                attempts_by_ip[address] = recent
            else:
                attempts_by_ip.pop(address, None)
        address = request.client.host if request.client else "unknown"
        attempts = attempts_by_ip.setdefault(address, [])
        if len(attempts) >= WEB_UPLOADS_PER_HOUR:
            raise HTTPException(
                status_code=429,
                detail="Upload limit reached. Please try again later.",
                headers={"Retry-After": "3600"},
            )
        attempts.append(now)

    def _make_web_job_cookie(job_id: str) -> str:
        issued_at = str(int(time.time()))
        payload = f"{job_id}.{issued_at}"
        signature = hmac.new(
            config.session_secret.encode("utf-8"), payload.encode("ascii"), hashlib.sha256,
        ).hexdigest()
        return f"{payload}.{signature}"

    def _web_job_cookie_is_valid(request: Request, job_id: str) -> bool:
        cookie = request.cookies.get(WEB_JOB_COOKIE_NAME, "")
        try:
            cookie_job_id, issued_text, supplied_signature = cookie.split(".", maxsplit=2)
            issued_at = int(issued_text)
        except (ValueError, TypeError):
            return False
        now = int(time.time())
        if cookie_job_id != job_id or issued_at > now + 60 or now - issued_at > WEB_JOB_TTL_SECONDS:
            return False
        payload = f"{cookie_job_id}.{issued_text}"
        expected_signature = hmac.new(
            config.session_secret.encode("utf-8"), payload.encode("ascii"), hashlib.sha256,
        ).hexdigest()
        return hmac.compare_digest(supplied_signature, expected_signature)

    def _get_web_job(job_id: str, request: Request) -> JobRecord:
        stable_id = _validate_job_id(job_id)
        if not config.session_secret:
            raise HTTPException(status_code=503, detail="Report sessions are not configured.")
        if not _web_job_cookie_is_valid(request, stable_id):
            raise HTTPException(status_code=401, detail="This report session is invalid or expired.")
        record = tracker.get(stable_id)
        if not record:
            raise HTTPException(status_code=404, detail="Report job was not found.")
        return record

    def _new_record(job_id: str, filename: str, requested_by: str) -> JobRecord:
        return JobRecord(
            job_id=job_id,
            uploaded_file=filename,
            requested_by=(requested_by or "unknown")[:320],
            requested_at=utc_now(),
        )

    def _process_job(job_id: str, source: tuple[str, bytes]) -> JobRecord:
        current = tracker.get(job_id)
        if not current:
            raise FileNotFoundError("Report job was not found.")
        if current.status == "Completed":
            return current
        if not tracker.claim(job_id):
            current = tracker.get(job_id) or current
            if current.status == "Completed":
                return current
            raise RuntimeError("Report job is already being processed.")

        try:
            filename, raw_bytes = source
            report = process_attendance_file(filename, raw_bytes)
            download_filename = _make_report_filename(report.report_date, current.requested_at)
            with tempfile.TemporaryDirectory(prefix="khammam_attendance_") as directory:
                output_path = Path(directory) / download_filename
                generate_pdf(report, output_path)
                output_reference = storage.save_pdf(job_id, output_path)
            completed = tracker.update(
                job_id,
                status="Completed",
                output_pdf=output_reference,
                local_pdf_path=str(config.data_dir / "reports" / download_filename),
                notes=(
                    f"Unique GPs: {report.unique_gp_count}; duplicate LGD rows excluded: "
                    f"{report.duplicate_lgd_rows}; rows without LGD excluded: {report.missing_lgd_rows}; "
                    f"conflicting LGDs: {len(report.conflicts)}; Download filename: {download_filename}"
                ),
            )
            logger.info(json.dumps({
                "event": "job_completed",
                "job_id": job_id,
                "unique_gps": report.unique_gp_count,
                "duplicate_lgd_rows": report.duplicate_lgd_rows,
            }))
            return completed
        except Exception as exc:
            _, safe_message = _safe_error(exc)
            tracker.update(job_id, status="Failed", notes=safe_message[:1000])
            logger.error(json.dumps({
                "event": "job_failed",
                "job_id": job_id,
                "error_type": type(exc).__name__,
                "message": safe_message,
            }))
            raise

    @api.get("/health")
    def health() -> dict[str, str]:
        return {
            "status": "ok",
            "service": "Khammam PS Attendance Analytics",
            "storage": "local",
            "job_tracking": "local_sqlite",
        }

    @api.post("/api/web/reports", status_code=200, include_in_schema=False)
    def create_web_report(
        request: Request,
        file: UploadFile = File(...),
        requested_by: str = Form(default="", alias="RequestedBy"),
    ) -> JSONResponse:
        if not config.session_secret:
            raise HTTPException(status_code=503, detail="Report sessions are not configured.")
        _check_web_upload_rate_limit(request)
        filename = _clean_filename(file.filename)
        raw = _read_upload(file, config.max_upload_bytes)
        job_id = uuid4().hex
        record, _ = tracker.create(_new_record(job_id, filename, requested_by))
        try:
            completed = _process_job(job_id, (filename, raw))
        except Exception as exc:
            status_code, safe_message = _safe_error(exc)
            failed = tracker.get(job_id)
            raise HTTPException(status_code=status_code, detail={
                "job_id": job_id,
                "status": failed.status if failed else "Failed",
                "message": safe_message,
            }) from None
        response = JSONResponse(completed.public_dict())
        response.set_cookie(
            WEB_JOB_COOKIE_NAME,
            _make_web_job_cookie(job_id),
            max_age=WEB_JOB_TTL_SECONDS,
            httponly=True,
            secure=request.url.scheme == "https",
            samesite="strict",
            path="/api/web/reports",
        )
        return response

    @api.get("/api/web/reports/{job_id}", include_in_schema=False)
    def get_web_report(job_id: str, request: Request) -> dict[str, str]:
        return _get_web_job(job_id, request).public_dict()

    @api.get("/api/web/reports/{job_id}/download", include_in_schema=False)
    def download_web_report(job_id: str, request: Request) -> Response:
        record = _get_web_job(job_id, request)
        if record.status != "Completed" or not record.output_pdf:
            raise HTTPException(status_code=409, detail=f"Report is not ready (status: {record.status}).")
        try:
            pdf_bytes = storage.read_pdf(record.output_pdf, record.local_pdf_path)
        except Exception as exc:
            status_code, message = _safe_error(exc)
            raise HTTPException(status_code=status_code, detail=message) from None
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{_record_report_filename(record)}"'},
        )

    return api
app = create_app()
