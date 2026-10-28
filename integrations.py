"""Local and Google-backed job tracking and file storage adapters."""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import sqlite3
import threading
from contextlib import closing
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Protocol
from urllib.parse import parse_qs, urlparse


SHEET_HEADERS = ["JobID", "UploadedFile", "RequestedBy", "RequestedAt", "Status", "OutputPDF", "Notes"]
ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".csv"}
PROCESSING_STALE_SECONDS = 15 * 60


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _processing_note() -> str:
    return f"Processing started at {utc_now()}"


def _processing_is_stale(note: str) -> bool:
    prefix = "Processing started at "
    if not note.startswith(prefix):
        return False
    try:
        started = datetime.fromisoformat(note[len(prefix):])
        return (datetime.now(timezone.utc) - started).total_seconds() > PROCESSING_STALE_SECONDS
    except (TypeError, ValueError):
        return False


@dataclass
class JobRecord:
    job_id: str
    uploaded_file: str
    requested_by: str
    requested_at: str
    status: str = "Submitted"
    output_pdf: str = ""
    notes: str = ""
    local_pdf_path: str = ""

    def public_dict(self) -> dict[str, str]:
        return {
            "job_id": self.job_id,
            "uploaded_file": self.uploaded_file,
            "requested_by": self.requested_by,
            "requested_at": self.requested_at,
            "status": self.status,
            "output_pdf": self.output_pdf,
            "notes": self.notes,
        }


class JobTracker(Protocol):
    def get(self, job_id: str) -> JobRecord | None: ...
    def create(self, record: JobRecord) -> tuple[JobRecord, bool]: ...
    def update(self, job_id: str, **values: str) -> JobRecord: ...
    def claim(self, job_id: str, allow_failed: bool = False) -> bool: ...


class LocalJobTracker:
    """SQLite tracker suitable for local development and a single local process."""

    def __init__(self, database_path: str | Path):
        self.path = Path(database_path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with closing(self._connect()) as connection:
            connection.execute("""
                CREATE TABLE IF NOT EXISTS attendance_jobs (
                    job_id TEXT PRIMARY KEY,
                    uploaded_file TEXT NOT NULL,
                    requested_by TEXT NOT NULL,
                    requested_at TEXT NOT NULL,
                    status TEXT NOT NULL,
                    output_pdf TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    local_pdf_path TEXT NOT NULL DEFAULT ''
                )
            """)

    @staticmethod
    def _record(row: sqlite3.Row | None) -> JobRecord | None:
        return JobRecord(**dict(row)) if row else None

    def get(self, job_id: str) -> JobRecord | None:
        with closing(self._connect()) as connection:
            row = connection.execute("SELECT * FROM attendance_jobs WHERE job_id = ?", (job_id,)).fetchone()
        return self._record(row)

    def create(self, record: JobRecord) -> tuple[JobRecord, bool]:
        with closing(self._connect()) as connection:
            connection.execute("""
                INSERT OR IGNORE INTO attendance_jobs
                (job_id, uploaded_file, requested_by, requested_at, status, output_pdf, notes, local_pdf_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, tuple(asdict(record).values()))
            created = connection.execute("SELECT changes()").fetchone()[0] == 1
        return self.get(record.job_id) or record, created

    def update(self, job_id: str, **values: str) -> JobRecord:
        allowed = {"uploaded_file", "status", "output_pdf", "notes", "local_pdf_path"}
        if not values or not set(values).issubset(allowed):
            raise ValueError("Unsupported job update field.")
        assignments = ", ".join(f"{key} = ?" for key in values)
        with closing(self._connect()) as connection:
            cursor = connection.execute(
                f"UPDATE attendance_jobs SET {assignments} WHERE job_id = ?",
                (*values.values(), job_id),
            )
            if cursor.rowcount != 1:
                raise KeyError(job_id)
        return self.get(job_id)  # type: ignore[return-value]

    def claim(self, job_id: str, allow_failed: bool = False) -> bool:
        accepted = ("Submitted", "Failed") if allow_failed else ("Submitted",)
        with closing(self._connect()) as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT status, notes FROM attendance_jobs WHERE job_id = ?", (job_id,),
            ).fetchone()
            eligible = bool(row and (row["status"] in accepted or (
                row["status"] == "Processing" and _processing_is_stale(row["notes"])
            )))
            cursor = connection.execute(
                "UPDATE attendance_jobs SET status = 'Processing', notes = ? WHERE job_id = ?",
                (_processing_note(), job_id),
            ) if eligible else None
            connection.execute("COMMIT")
        return bool(cursor and cursor.rowcount == 1)


class GoogleSheetsJobTracker:
    """Job tracker backed by the configured AttendanceReportJobs sheet tab."""

    def __init__(self, spreadsheet_id: str, credentials_path: str = "", tab: str = "AttendanceReportJobs"):
        from googleapiclient.discovery import build

        credentials = _google_credentials(credentials_path, [
            "https://www.googleapis.com/auth/spreadsheets",
        ])
        self.service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
        self.spreadsheet_id = spreadsheet_id
        self.tab = tab
        self._lock = threading.RLock()
        self._ensure_headers()

    def _range(self, suffix: str) -> str:
        escaped_tab = self.tab.replace("'", "''")
        return f"'{escaped_tab}'!{suffix}"

    def _ensure_headers(self) -> None:
        response = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=self._range("A1:G1"),
        ).execute()
        rows = response.get("values", [])
        if not rows or not any(rows[0]):
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id, range=self._range("A1:G1"),
                valueInputOption="RAW", body={"values": [SHEET_HEADERS]},
            ).execute()
        elif rows[0][:len(SHEET_HEADERS)] != SHEET_HEADERS:
            raise RuntimeError(
                "AttendanceReportJobs sheet headers must be: " + ", ".join(SHEET_HEADERS)
            )

    def _all_rows(self) -> list[list[str]]:
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=self._range("A2:G"),
        ).execute()
        return result.get("values", [])

    @staticmethod
    def _to_record(row: list[str]) -> JobRecord:
        padded = (row + [""] * 7)[:7]
        return JobRecord(
            job_id=padded[0], uploaded_file=padded[1], requested_by=padded[2],
            requested_at=padded[3], status=padded[4], output_pdf=padded[5], notes=padded[6],
        )

    def get(self, job_id: str) -> JobRecord | None:
        with self._lock:
            for row in self._all_rows():
                if row and row[0] == job_id:
                    return self._to_record(row)
        return None

    def create(self, record: JobRecord) -> tuple[JobRecord, bool]:
        with self._lock:
            existing = self.get(record.job_id)
            if existing:
                return existing, False
            values = [[record.job_id, record.uploaded_file, record.requested_by,
                       record.requested_at, record.status, record.output_pdf, record.notes]]
            self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id, range=self._range("A:G"),
                valueInputOption="RAW", insertDataOption="INSERT_ROWS", body={"values": values},
            ).execute()
            return record, True

    def update(self, job_id: str, **values: str) -> JobRecord:
        allowed = {"uploaded_file", "status", "output_pdf", "notes"}
        if not values or not set(values).issubset(allowed):
            raise ValueError("Unsupported job update field.")
        with self._lock:
            rows = self._all_rows()
            for offset, row in enumerate(rows, start=2):
                if row and row[0] == job_id:
                    record = self._to_record(row)
                    for key, value in values.items():
                        setattr(record, key, value)
                    output = [[record.job_id, record.uploaded_file, record.requested_by,
                               record.requested_at, record.status, record.output_pdf, record.notes]]
                    self.service.spreadsheets().values().update(
                        spreadsheetId=self.spreadsheet_id, range=self._range(f"A{offset}:G{offset}"),
                        valueInputOption="RAW", body={"values": output},
                    ).execute()
                    return record
        raise KeyError(job_id)

    def claim(self, job_id: str, allow_failed: bool = False) -> bool:
        # Cloud Tasks concurrency should be set to one for this sheet-backed mode.
        # The local lock prevents duplicate claims within a single API instance.
        with self._lock:
            record = self.get(job_id)
            accepted = {"Submitted", "Failed"} if allow_failed else {"Submitted"}
            if not record or not (
                record.status in accepted
                or (record.status == "Processing" and _processing_is_stale(record.notes))
            ):
                return False
            self.update(job_id, status="Processing", notes=_processing_note())
            return True


class Storage(Protocol):
    is_drive: bool
    def read_uploaded(self, reference: str, original_name: str = "") -> tuple[str, bytes]: ...
    def save_pdf(self, job_id: str, pdf_path: str | Path) -> str: ...
    def read_pdf(self, reference: str, local_pdf_path: str = "") -> bytes: ...
    def save_input(self, local_path: str | Path, filename: str) -> str: ...


class LocalStorage:
    """Local development storage; Drive references resolve only inside input_dir."""

    is_drive = False

    def __init__(self, data_dir: str | Path, input_dir: str | Path | None = None, max_file_bytes: int = 30 * 1024 * 1024):
        self.data_dir = Path(data_dir).resolve()
        self.input_dir = Path(input_dir or (self.data_dir / "inputs")).resolve()
        self.report_dir = self.data_dir / "reports"
        self.input_dir.mkdir(parents=True, exist_ok=True)
        self.report_dir.mkdir(parents=True, exist_ok=True)
        self.max_file_bytes = max_file_bytes

    def read_uploaded(self, reference: str, original_name: str = "") -> tuple[str, bytes]:
        raw = str(reference).strip().replace("\\", "/")
        candidate = Path(raw)
        if not candidate.is_absolute():
            candidate = self.input_dir / PurePosixPath(raw)
        resolved = candidate.resolve()
        try:
            resolved.relative_to(self.input_dir)
        except ValueError as exc:
            raise PermissionError("Local input references must point inside the configured input folder.") from exc
        if not resolved.is_file():
            raise FileNotFoundError("The referenced file was not found in the local input folder.")
        if resolved.suffix.lower() not in ALLOWED_EXTENSIONS:
            raise ValueError("Allowed file types are .xls, .xlsx, and .csv.")
        if resolved.stat().st_size > self.max_file_bytes:
            raise ValueError("The referenced file exceeds the 30 MB upload limit.")
        return resolved.name, resolved.read_bytes()

    def save_pdf(self, job_id: str, pdf_path: str | Path) -> str:
        destination = (self.report_dir / Path(pdf_path).name).resolve()
        try:
            destination.relative_to(self.report_dir.resolve())
        except ValueError as exc:
            raise ValueError("Generated report filename is invalid.") from exc
        shutil.copyfile(pdf_path, destination)
        return f"/api/reports/{job_id}/download"

    def read_pdf(self, reference: str, local_pdf_path: str = "") -> bytes:
        source = Path(local_pdf_path).resolve() if local_pdf_path else None
        try:
            if source is None:
                candidate = Path(reference)
                source = candidate.resolve()
            source.relative_to(self.report_dir.resolve())
        except (ValueError, OSError) as exc:
            raise FileNotFoundError("The generated report is not available in local storage.") from exc
        if not source.is_file():
            raise FileNotFoundError("The generated report is not available in local storage.")
        return source.read_bytes()

    def save_input(self, local_path: str | Path, filename: str) -> str:
        destination = self.input_dir / Path(filename).name
        shutil.copyfile(local_path, destination)
        return destination.name


def _google_credentials(credentials_path: str, scopes: list[str]) -> Any:
    import google.auth
    from google.oauth2 import service_account

    if credentials_path:
        return service_account.Credentials.from_service_account_file(credentials_path, scopes=scopes)
    credentials, _ = google.auth.default(scopes=scopes)
    return credentials


class GoogleDriveStorage:
    """Private Drive storage scoped to configured upload and output folders."""

    is_drive = True

    def __init__(
        self,
        upload_folder_id: str,
        output_folder_id: str,
        credentials_path: str = "",
        max_file_bytes: int = 30 * 1024 * 1024,
    ):
        from googleapiclient.discovery import build

        if not upload_folder_id or not output_folder_id:
            raise ValueError("Drive mode requires both DRIVE_UPLOAD_FOLDER_ID and DRIVE_OUTPUT_FOLDER_ID.")
        self.upload_folder_id = upload_folder_id
        self.output_folder_id = output_folder_id
        self.max_file_bytes = max_file_bytes
        scopes = ["https://www.googleapis.com/auth/drive"]
        credentials = _google_credentials(credentials_path, scopes)
        self.service = build("drive", "v3", credentials=credentials, cache_discovery=False)

    @staticmethod
    def _extract_id(reference: str) -> str:
        value = str(reference).strip()
        parsed = urlparse(value)
        if parsed.scheme in {"http", "https"} and parsed.netloc.endswith("google.com"):
            query_id = parse_qs(parsed.query).get("id", [""])[0]
            match = re.search(r"/file/d/([^/]+)", parsed.path)
            value = query_id or (match.group(1) if match else "")
        return value if re.fullmatch(r"[A-Za-z0-9_-]{10,}", value) else ""

    def _get_metadata(self, file_id: str) -> dict[str, Any]:
        try:
            return self.service.files().get(
                fileId=file_id, supportsAllDrives=True,
                fields="id,name,mimeType,size,parents,trashed",
            ).execute()
        except Exception as exc:
            raise FileNotFoundError("The requested Drive file could not be accessed.") from exc

    def _resolve_file(self, reference: str, original_name: str = "") -> dict[str, Any]:
        file_id = self._extract_id(reference)
        if file_id:
            metadata = self._get_metadata(file_id)
            if metadata.get("trashed"):
                raise FileNotFoundError("The requested Drive file is in the trash.")
            if self.upload_folder_id not in metadata.get("parents", []):
                raise PermissionError("The Drive file must be inside the configured upload folder.")
            return metadata

        relative = str(reference or original_name).strip().replace("\\", "/")
        filename = PurePosixPath(relative).name
        if not filename or filename in {".", ".."}:
            raise ValueError("Provide a Drive file ID or an uploaded file path/name.")
        if Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise ValueError("Allowed file types are .xls, .xlsx, and .csv.")
        escaped = filename.replace("\\", "\\\\").replace("'", "\\'")
        query = (
            f"'{self.upload_folder_id}' in parents and name = '{escaped}' "
            "and trashed = false"
        )
        matches = self.service.files().list(
            q=query, supportsAllDrives=True, includeItemsFromAllDrives=True,
            fields="files(id,name,mimeType,size,parents,trashed)", pageSize=10,
        ).execute().get("files", [])
        if not matches:
            raise FileNotFoundError("The uploaded file was not found in the configured Drive upload folder.")
        if len(matches) > 1:
            raise ValueError("More than one uploaded file has that name; pass its Drive file ID instead.")
        return matches[0]

    def _download(self, file_id: str, size_hint: Any = None) -> bytes:
        from googleapiclient.http import MediaIoBaseDownload

        try:
            size = int(size_hint) if size_hint is not None else None
        except (ValueError, TypeError):
            size = None
        if size is not None and size > self.max_file_bytes:
            raise ValueError("The referenced file exceeds the 30 MB upload limit.")
        request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
        sink = io.BytesIO()
        downloader = MediaIoBaseDownload(sink, request, chunksize=1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk()
            if sink.tell() > self.max_file_bytes:
                raise ValueError("The referenced file exceeds the 30 MB upload limit.")
        return sink.getvalue()

    def read_uploaded(self, reference: str, original_name: str = "") -> tuple[str, bytes]:
        metadata = self._resolve_file(reference, original_name)
        filename = str(metadata.get("name") or original_name or "attendance.xls")
        if Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise ValueError("Allowed file types are .xls, .xlsx, and .csv.")
        return filename, self._download(metadata["id"], metadata.get("size"))

    def save_pdf(self, job_id: str, pdf_path: str | Path) -> str:
        from googleapiclient.http import MediaFileUpload

        destination = Path(pdf_path)
        media = MediaFileUpload(str(destination), mimetype="application/pdf", resumable=True)
        result = self.service.files().create(
            body={
                "name": destination.name,
                "mimeType": "application/pdf",
                "parents": [self.output_folder_id],
            },
            media_body=media,
            supportsAllDrives=True,
            fields="id,name",
        ).execute()
        # This is a restricted Drive URL; no public permission is created.
        return f"https://drive.google.com/file/d/{result['id']}/view"

    def read_pdf(self, reference: str, local_pdf_path: str = "") -> bytes:
        file_id = self._extract_id(reference)
        if not file_id:
            raise FileNotFoundError("The generated Drive report reference is invalid.")
        metadata = self._get_metadata(file_id)
        if self.output_folder_id not in metadata.get("parents", []):
            raise PermissionError("The generated file is outside the configured output folder.")
        return self._download(file_id, metadata.get("size"))

    def save_input(self, local_path: str | Path, filename: str) -> str:
        from googleapiclient.http import MediaFileUpload

        media = MediaFileUpload(str(local_path), resumable=True)
        result = self.service.files().create(
            body={"name": Path(filename).name, "parents": [self.upload_folder_id]},
            media_body=media,
            supportsAllDrives=True,
            fields="id,name",
        ).execute()
        return result["id"]


class GoogleCloudTasksDispatcher:
    """Enqueue durable report jobs for Cloud Run deployments."""

    def __init__(
        self,
        project_id: str,
        location: str,
        queue_name: str,
        base_url: str,
        invoker_service_account: str,
    ):
        from google.cloud import tasks_v2

        self.client = tasks_v2.CloudTasksClient()
        self.parent = self.client.queue_path(project_id, location, queue_name)
        self.base_url = base_url.rstrip("/")
        self.invoker_service_account = invoker_service_account

    def enqueue(self, job_id: str) -> None:
        from google.api_core.exceptions import AlreadyExists
        from google.cloud import tasks_v2

        import hashlib

        task_id = "report-" + hashlib.sha256(job_id.encode("utf-8")).hexdigest()[:32]
        task_name = f"{self.parent}/tasks/{task_id}"
        payload = json.dumps({"job_id": job_id}).encode("utf-8")
        task = {
            "name": task_name,
            "dispatch_deadline": {"seconds": 900},
            "http_request": {
                "http_method": tasks_v2.HttpMethod.POST,
                "url": f"{self.base_url}/internal/tasks/process-report",
                "headers": {"Content-Type": "application/json"},
                "body": payload,
                "oidc_token": {
                    "service_account_email": self.invoker_service_account,
                    "audience": self.base_url,
                },
            },
        }
        try:
            self.client.create_task(request={"parent": self.parent, "task": task})
        except AlreadyExists:
            # A retry of the same AppSheet JobID cannot create a second named task.
            return


def verify_cloud_task_token(token: str, audience: str, expected_email: str) -> bool:
    try:
        from google.auth.transport.requests import Request
        from google.oauth2 import id_token

        claims = id_token.verify_oauth2_token(token, Request(), audience=audience)
        return bool(claims.get("email_verified")) and claims.get("email") == expected_email
    except Exception:
        return False
