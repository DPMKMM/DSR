"""FastAPI service for authorized Khammam PS attendance report jobs."""

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
from pathlib import Path
from typing import Any
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, Response

from integrations import (
    ALLOWED_EXTENSIONS,
    GoogleCloudTasksDispatcher,
    GoogleDriveStorage,
    GoogleSheetsJobTracker,
    JobRecord,
    JobTracker,
    LocalJobTracker,
    LocalStorage,
    Storage,
    utc_now,
    verify_cloud_task_token,
)
from report_generator import InputValidationError, generate_pdf, process_attendance_file


logger = logging.getLogger("khammam_attendance")
logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO").upper(), format="%(message)s")
MAX_UPLOAD_BYTES = 30 * 1024 * 1024
JOB_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,100}$")
SESSION_COOKIE_NAME = "khammam_attendance_session"
BROWSER_SESSION_TTL_SECONDS = 8 * 60 * 60

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
    input[type="password"], input[type="file"], input[type="text"] { width: 100%;
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
    .docs { color: #176247; }
  </style>
</head>
<body>
  <main>
    <section class="card">
      <div class="eyebrow">Khammam District · Telangana</div>
      <h1>PS Attendance Analytics</h1>
      <p>Upload the daily Gram Panchayat attendance file to generate and download the formatted district report.</p>
      <div id="storage-warning" class="notice" hidden>
        Reports are currently stored on temporary service storage. Download each report when it is generated.
      </div>

      <form id="login-form">
        <label for="api-token">Access token</label>
        <input id="api-token" type="password" autocomplete="current-password" required
          placeholder="Enter the Render APP_API_TOKEN">
        <div class="actions"><button id="login-button" type="submit">Unlock report upload</button></div>
        <div id="login-message" class="message" role="status" aria-live="polite"></div>
      </form>

      <form id="upload-form" hidden>
        <label for="attendance-file">Attendance file</label>
        <input id="attendance-file" name="file" type="file" accept=".xls,.xlsx,.csv" required>
        <label for="requested-by">Submitted by (optional)</label>
        <input id="requested-by" name="RequestedBy" type="text" maxlength="320"
          autocomplete="name" placeholder="Name or office">
        <div class="actions">
          <button id="generate-button" type="submit">Generate and download PDF</button>
          <button id="logout-button" class="secondary" type="button">Sign out</button>
        </div>
        <div id="report-message" class="message" role="status" aria-live="polite"></div>
      </form>

      <p class="fine-print">Your access token is exchanged for a short-lived, secure browser session and is not saved in this page. Report API routes remain protected.</p>
      <a class="docs" href="/docs">Open API documentation</a>
    </section>
  </main>
  <script>
    (() => {
      const loginForm = document.getElementById('login-form');
      const uploadForm = document.getElementById('upload-form');
      const tokenInput = document.getElementById('api-token');
      const loginButton = document.getElementById('login-button');
      const generateButton = document.getElementById('generate-button');
      const loginMessage = document.getElementById('login-message');
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

      async function refreshSession() {
        try {
          await apiJson('/api/session');
          loginForm.hidden = true;
          uploadForm.hidden = false;
        } catch (_) {
          uploadForm.hidden = true;
          loginForm.hidden = false;
        }
      }

      async function showStorageMode() {
        try {
          const response = await fetch('/health', { credentials: 'same-origin' });
          const health = await response.json();
          document.getElementById('storage-warning').hidden =
            health.storage === 'google_drive' && health.job_tracking === 'google_sheets';
        } catch (_) { /* Keep the page usable if the optional health hint is unavailable. */ }
      }

      loginForm.addEventListener('submit', async (event) => {
        event.preventDefault();
        loginButton.disabled = true;
        showMessage(loginMessage, 'Checking access…');
        try {
          await apiJson('/api/session', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ api_token: tokenInput.value })
          });
          tokenInput.value = '';
          showMessage(loginMessage, 'Access granted.', true);
          await refreshSession();
        } catch (error) {
          tokenInput.value = '';
          showMessage(loginMessage, error.message || 'Could not verify the access token.');
        } finally {
          loginButton.disabled = false;
        }
      });

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
          let job = await apiJson('/api/reports', { method: 'POST', body: data });
          const jobId = job.job_id;
          if (!jobId) throw new Error('The service did not return a report job ID.');
          let attempts = 0;
          while (job.status === 'Submitted' || job.status === 'Processing') {
            showMessage(reportMessage, `Report ${job.status.toLowerCase()}…`);
            if (attempts++ > 180) throw new Error('The report is taking too long. Check the job status in API documentation.');
            await new Promise(resolve => setTimeout(resolve, 2000));
            job = await apiJson(`/api/reports/${encodeURIComponent(jobId)}`);
          }
          if (job.status !== 'Completed') throw new Error(job.notes || 'The report job did not complete.');
          showMessage(reportMessage, 'Report ready. Downloading PDF…', true);
          const pdfResponse = await fetch(`/api/reports/${encodeURIComponent(jobId)}/download`, {
            credentials: 'same-origin'
          });
          if (!pdfResponse.ok) await readResponse(pdfResponse);
          const blob = await pdfResponse.blob();
          const objectUrl = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = objectUrl;
          link.download = `Khammam_PS_Attendance_Report_${jobId}.pdf`;
          document.body.appendChild(link);
          link.click();
          link.remove();
          setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
          showMessage(reportMessage, 'Report downloaded. You can upload another file.', true);
          uploadForm.reset();
        } catch (error) {
          if (error.message === 'Authentication required.') await refreshSession();
          showMessage(reportMessage, error.message || 'Report generation failed.');
        } finally {
          generateButton.disabled = false;
        }
      });

      document.getElementById('logout-button').addEventListener('click', async () => {
        try { await apiJson('/api/session', { method: 'DELETE' }); } catch (_) {}
        await refreshSession();
      });

      refreshSession();
      showStorageMode();
    })();
  </script>
</body>
</html>"""


@dataclass(frozen=True)
class Settings:
    api_token: str
    data_dir: Path
    local_input_dir: Path | None
    max_upload_bytes: int
    credentials_path: str
    drive_upload_folder_id: str
    drive_output_folder_id: str
    spreadsheet_id: str
    spreadsheet_tab: str
    tasks_project: str
    tasks_location: str
    tasks_queue: str
    public_base_url: str
    task_invoker_service_account: str

    @classmethod
    def from_env(cls) -> "Settings":
        input_dir = os.getenv("LOCAL_INPUT_DIR", "").strip()
        try:
            max_bytes = int(os.getenv("MAX_UPLOAD_BYTES", str(MAX_UPLOAD_BYTES)))
        except ValueError as exc:
            raise RuntimeError("MAX_UPLOAD_BYTES must be an integer.") from exc
        if max_bytes < 1 or max_bytes > 100 * 1024 * 1024:
            raise RuntimeError("MAX_UPLOAD_BYTES must be between 1 byte and 100 MB.")
        return cls(
            api_token=os.getenv("APP_API_TOKEN", "").strip(),
            data_dir=Path(os.getenv("DATA_DIR", "./data")).resolve(),
            local_input_dir=Path(input_dir).resolve() if input_dir else None,
            max_upload_bytes=max_bytes,
            credentials_path=os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "").strip(),
            drive_upload_folder_id=os.getenv("DRIVE_UPLOAD_FOLDER_ID", "").strip(),
            drive_output_folder_id=os.getenv("DRIVE_OUTPUT_FOLDER_ID", "").strip(),
            spreadsheet_id=os.getenv("SHEETS_SPREADSHEET_ID", "").strip(),
            spreadsheet_tab=os.getenv("SHEETS_JOBS_TAB", "AttendanceReportJobs").strip(),
            tasks_project=os.getenv("CLOUD_TASKS_PROJECT", "").strip(),
            tasks_location=os.getenv("CLOUD_TASKS_LOCATION", "").strip(),
            tasks_queue=os.getenv("CLOUD_TASKS_QUEUE", "").strip(),
            public_base_url=os.getenv("PUBLIC_BASE_URL", "").strip().rstrip("/"),
            task_invoker_service_account=os.getenv("TASK_INVOKER_SERVICE_ACCOUNT", "").strip(),
        )

    @property
    def tasks_enabled(self) -> bool:
        return bool(self.tasks_queue)


def _clean_filename(value: str | None) -> str:
    raw = (value or "attendance.xls").replace("\\", "/").split("/")[-1].strip()
    return raw or "attendance.xls"


def _validate_job_id(value: str | None) -> str:
    candidate = (value or uuid4().hex).strip()
    if not JOB_ID_PATTERN.fullmatch(candidate):
        raise HTTPException(status_code=400, detail="JobID must contain 1-100 letters, numbers, underscores, or hyphens.")
    return candidate


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
    dispatcher: Any | None = None,
) -> FastAPI:
    config = settings or Settings.from_env()
    if config.api_token and len(config.api_token) < 32:
        raise RuntimeError("APP_API_TOKEN must contain at least 32 characters.")
    drive_configured = bool(config.drive_upload_folder_id or config.drive_output_folder_id)
    if drive_configured and not (config.drive_upload_folder_id and config.drive_output_folder_id):
        raise RuntimeError("Set both DRIVE_UPLOAD_FOLDER_ID and DRIVE_OUTPUT_FOLDER_ID to enable Drive storage.")

    if storage is None:
        if drive_configured:
            storage = GoogleDriveStorage(
                config.drive_upload_folder_id, config.drive_output_folder_id,
                config.credentials_path, config.max_upload_bytes,
            )
        else:
            storage = LocalStorage(config.data_dir, config.local_input_dir, config.max_upload_bytes)
    if tracker is None:
        if config.spreadsheet_id:
            tracker = GoogleSheetsJobTracker(
                config.spreadsheet_id, config.credentials_path, config.spreadsheet_tab,
            )
        else:
            tracker = LocalJobTracker(config.data_dir / "attendance_jobs.sqlite3")

    task_settings = [config.tasks_project, config.tasks_location, config.tasks_queue,
                     config.public_base_url, config.task_invoker_service_account]
    if any(task_settings) and not all(task_settings):
        raise RuntimeError(
            "Cloud Tasks requires CLOUD_TASKS_PROJECT, CLOUD_TASKS_LOCATION, CLOUD_TASKS_QUEUE, "
            "PUBLIC_BASE_URL, and TASK_INVOKER_SERVICE_ACCOUNT."
        )
    if config.tasks_enabled:
        if not storage.is_drive or not config.spreadsheet_id:
            raise RuntimeError("Cloud Tasks mode requires Google Drive storage and SHEETS_SPREADSHEET_ID.")
        if dispatcher is None:
            dispatcher = GoogleCloudTasksDispatcher(
                config.tasks_project, config.tasks_location, config.tasks_queue,
                config.public_base_url, config.task_invoker_service_account,
            )

    api = FastAPI(
        title="Khammam PS Attendance Analytics API",
        version="2.0.0",
        description="Protected attendance report jobs for Khammam District.",
    )
    api.state.settings = config
    api.state.tracker = tracker
    api.state.storage = storage
    api.state.dispatcher = dispatcher

    @api.get("/", response_class=HTMLResponse, include_in_schema=False)
    def home() -> HTMLResponse:
        return HTMLResponse(UPLOAD_PAGE_HTML)

    @api.middleware("http")
    async def limit_content_length(request: Request, call_next: Any) -> Response:
        if request.url.path == "/api/reports" and request.method == "POST":
            raw_length = request.headers.get("content-length")
            if raw_length:
                try:
                    if int(raw_length) > config.max_upload_bytes + 1024 * 1024:
                        return JSONResponse(status_code=413, content={"detail": "Multipart request is too large."})
                except ValueError:
                    return JSONResponse(status_code=400, content={"detail": "Invalid Content-Length header."})
        return await call_next(request)

    def _browser_session_is_valid(request: Request) -> bool:
        cookie = request.cookies.get(SESSION_COOKIE_NAME, "")
        try:
            issued_text, supplied_signature = cookie.split(".", maxsplit=1)
            issued_at = int(issued_text)
        except (ValueError, TypeError):
            return False
        now = int(time.time())
        if issued_at > now + 60 or now - issued_at > BROWSER_SESSION_TTL_SECONDS:
            return False
        expected_signature = hmac.new(
            config.api_token.encode("utf-8"), issued_text.encode("ascii"), hashlib.sha256,
        ).hexdigest()
        return hmac.compare_digest(supplied_signature, expected_signature)

    @api.post("/api/session", include_in_schema=False)
    async def create_browser_session(request: Request) -> JSONResponse:
        if not config.api_token:
            raise HTTPException(status_code=503, detail="API authentication is not configured.")
        try:
            payload = await request.json()
        except Exception:
            raise HTTPException(status_code=400, detail="Request must contain a JSON access token.") from None
        supplied_token = payload.get("api_token") if isinstance(payload, dict) else None
        if not isinstance(supplied_token, str) or not hmac.compare_digest(supplied_token, config.api_token):
            raise HTTPException(status_code=401, detail="Access token is incorrect.")
        issued_text = str(int(time.time()))
        signature = hmac.new(
            config.api_token.encode("utf-8"), issued_text.encode("ascii"), hashlib.sha256,
        ).hexdigest()
        response = JSONResponse({"authenticated": True})
        response.set_cookie(
            SESSION_COOKIE_NAME,
            f"{issued_text}.{signature}",
            max_age=BROWSER_SESSION_TTL_SECONDS,
            httponly=True,
            secure=True,
            samesite="strict",
            path="/",
        )
        return response

    @api.get("/api/session", include_in_schema=False)
    def browser_session_status(request: Request) -> dict[str, bool]:
        if not config.api_token or not _browser_session_is_valid(request):
            raise HTTPException(status_code=401, detail="Authentication required.")
        return {"authenticated": True}

    @api.delete("/api/session", include_in_schema=False)
    def delete_browser_session() -> Response:
        response = Response(status_code=204)
        response.delete_cookie(SESSION_COOKIE_NAME, path="/", secure=True, httponly=True, samesite="strict")
        return response

    def authorize(
        request: Request,
        x_api_token: str | None = Header(default=None),
    ) -> None:
        if not config.api_token:
            raise HTTPException(status_code=503, detail="API authentication is not configured.")
        header_is_valid = bool(x_api_token and hmac.compare_digest(x_api_token, config.api_token))
        if not header_is_valid and not _browser_session_is_valid(request):
            raise HTTPException(status_code=401, detail="Authentication required.")

    def _new_record(job_id: str, uploaded_file: str, requested_by: str) -> JobRecord:
        return JobRecord(
            job_id=job_id,
            uploaded_file=uploaded_file,
            requested_by=(requested_by or "unknown")[:320],
            requested_at=utc_now(),
        )

    def _enqueue_or_process(job: JobRecord, source: tuple[str, bytes] | None = None) -> JobRecord:
        if job.status in {"Processing", "Completed", "Failed"}:
            return job
        if config.tasks_enabled:
            if source is not None:
                filename, raw = source
                with tempfile.TemporaryDirectory(prefix="khammam_attendance_") as directory:
                    local_source = Path(directory) / _clean_filename(filename)
                    local_source.write_bytes(raw)
                    source_ref = storage.save_input(local_source, local_source.name)
                job = tracker.update(job.job_id, uploaded_file=source_ref)
            try:
                dispatcher.enqueue(job.job_id)
            except Exception as exc:
                logger.error(json.dumps({"event": "task_enqueue_failed", "job_id": job.job_id,
                                         "error_type": type(exc).__name__}))
                failed = tracker.update(job.job_id, status="Failed", notes="Could not queue this report job.")
                raise HTTPException(status_code=503, detail={
                    "job_id": job.job_id, "status": failed.status, "message": failed.notes,
                }) from None
            return tracker.get(job.job_id) or job

        try:
            return _process_job(job.job_id, allow_failed=False, source=source)
        except Exception as exc:
            status_code, safe_message = _safe_error(exc)
            record = tracker.get(job.job_id)
            if record and record.status in {"Processing", "Completed"}:
                return record
            raise HTTPException(status_code=status_code, detail={
                "job_id": job.job_id,
                "status": record.status if record else "Failed",
                "message": safe_message,
            }) from None

    def _process_job(
        job_id: str,
        allow_failed: bool,
        source: tuple[str, bytes] | None = None,
    ) -> JobRecord:
        current = tracker.get(job_id)
        if not current:
            raise FileNotFoundError("Report job was not found.")
        if current.status == "Completed":
            return current
        if not tracker.claim(job_id, allow_failed=allow_failed):
            current = tracker.get(job_id) or current
            if current.status == "Processing":
                raise RuntimeError("Report job is already being processed.")
            if current.status == "Completed":
                return current
            raise RuntimeError("Report job is not available for processing.")

        try:
            if source is None:
                filename, raw_bytes = storage.read_uploaded(current.uploaded_file)
            else:
                filename, raw_bytes = source
            report = process_attendance_file(filename, raw_bytes)
            with tempfile.TemporaryDirectory(prefix="khammam_attendance_") as directory:
                output_path = Path(directory) / "Khammam_PS_Attendance_Report.pdf"
                generate_pdf(report, output_path)
                output_reference = storage.save_pdf(job_id, output_path)
            update_values: dict[str, str] = {
                "status": "Completed",
                "output_pdf": output_reference,
                "notes": (
                    f"Unique GPs: {report.unique_gp_count}; duplicate LGD rows excluded: "
                    f"{report.duplicate_lgd_rows}; rows without LGD excluded: {report.missing_lgd_rows}; "
                    f"conflicting LGDs: {len(report.conflicts)}."
                ),
            }
            if not storage.is_drive:
                update_values["local_pdf_path"] = str(
                    config.data_dir / "reports" / f"{_safe_local_job_name(job_id)}.pdf"
                )
            completed = tracker.update(job_id, **update_values)
            logger.info(json.dumps({
                "event": "job_completed", "job_id": job_id, "unique_gps": report.unique_gp_count,
                "duplicate_lgd_rows": report.duplicate_lgd_rows,
            }))
            return completed
        except Exception as exc:
            _, safe_message = _safe_error(exc)
            tracker.update(job_id, status="Failed", notes=safe_message[:1000])
            logger.error(json.dumps({"event": "job_failed", "job_id": job_id,
                                     "error_type": type(exc).__name__, "message": safe_message}))
            raise

    @api.get("/health")
    def health() -> dict[str, Any]:
        return {
            "status": "ok",
            "service": "Khammam PS Attendance Analytics API",
            "storage": "google_drive" if storage.is_drive else "local",
            "job_tracking": "google_sheets" if config.spreadsheet_id else "local_sqlite",
            "queued_processing": config.tasks_enabled,
        }

    @api.post("/api/reports", status_code=200, dependencies=[Depends(authorize)])
    def create_report(
        file: UploadFile = File(...),
        job_id: str | None = Form(default=None, alias="JobID"),
        requested_by: str = Form(default="", alias="RequestedBy"),
    ) -> dict[str, Any]:
        filename = _clean_filename(file.filename)
        if Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise HTTPException(status_code=415, detail="Allowed file types are .xls, .xlsx, and .csv.")
        stable_id = _validate_job_id(job_id)
        existing = tracker.get(stable_id)
        if existing and existing.status != "Submitted":
            return existing.public_dict()
        raw = _read_upload(file, config.max_upload_bytes)
        if existing:
            record = existing
        else:
            record, _ = tracker.create(_new_record(stable_id, filename, requested_by))
        return _enqueue_or_process(record, source=(filename, raw)).public_dict()

    @api.post("/api/reports/from-drive", status_code=200, dependencies=[Depends(authorize)])
    def create_report_from_drive(payload: dict[str, Any]) -> dict[str, Any]:
        requested_job_id = str(payload.get("JobID") or payload.get("job_id") or "").strip()
        if not requested_job_id:
            raise HTTPException(status_code=400, detail="AppSheet payload must include JobID.")
        stable_id = _validate_job_id(requested_job_id)
        upload_ref = str(
            payload.get("file_id") or payload.get("UploadedFile") or payload.get("uploaded_file") or ""
        ).strip()
        if not upload_ref:
            raise HTTPException(status_code=400, detail="Provide file_id or UploadedFile.")
        original_name = _clean_filename(str(payload.get("original_name") or "attendance.xls"))
        existing = tracker.get(stable_id)
        if existing and existing.status != "Submitted":
            return existing.public_dict()
        if existing:
            record = existing
            # The sheet row is the source of truth once AppSheet has submitted it.
            upload_ref = existing.uploaded_file
        else:
            record, _ = tracker.create(_new_record(stable_id, upload_ref, str(
                payload.get("RequestedBy") or payload.get("requested_by") or "unknown"
            )))
        if config.tasks_enabled:
            return _enqueue_or_process(record).public_dict()
        try:
            source = storage.read_uploaded(upload_ref, original_name)
        except Exception as exc:
            status_code, safe_message = _safe_error(exc)
            tracker.update(stable_id, status="Failed", notes=safe_message[:1000])
            logger.error(json.dumps({"event": "job_failed", "job_id": stable_id,
                                     "error_type": type(exc).__name__, "message": safe_message}))
            raise HTTPException(status_code=status_code, detail={
                "job_id": stable_id, "status": "Failed", "message": safe_message,
            }) from None
        return _enqueue_or_process(record, source=source).public_dict()

    @api.get("/api/reports/{job_id}", dependencies=[Depends(authorize)])
    def get_report(job_id: str) -> dict[str, Any]:
        stable_id = _validate_job_id(job_id)
        record = tracker.get(stable_id)
        if not record:
            raise HTTPException(status_code=404, detail="Report job was not found.")
        return record.public_dict()

    @api.get("/api/reports/{job_id}/download", dependencies=[Depends(authorize)])
    def download_report(job_id: str) -> Response:
        stable_id = _validate_job_id(job_id)
        record = tracker.get(stable_id)
        if not record:
            raise HTTPException(status_code=404, detail="Report job was not found.")
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
            headers={"Content-Disposition": 'attachment; filename="Khammam_PS_Attendance_Report.pdf"'},
        )

    @api.post("/internal/tasks/process-report")
    def process_report_task(request: Request, payload: dict[str, Any]) -> dict[str, str]:
        if not config.tasks_enabled:
            raise HTTPException(status_code=404, detail="Not found.")
        authorization = request.headers.get("authorization", "")
        token = authorization[7:].strip() if authorization.lower().startswith("bearer ") else ""
        if not token or not verify_cloud_task_token(
            token, config.public_base_url, config.task_invoker_service_account,
        ):
            raise HTTPException(status_code=401, detail="Task authentication required.")
        stable_id = _validate_job_id(str(payload.get("job_id") or ""))
        try:
            record = _process_job(stable_id, allow_failed=True)
            return {"job_id": record.job_id, "status": record.status}
        except FileNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from None
        except Exception:
            raise HTTPException(status_code=500, detail="Report processing failed.") from None

    return api


def _safe_local_job_name(job_id: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", job_id)[:100] or "job"


app = create_app()
