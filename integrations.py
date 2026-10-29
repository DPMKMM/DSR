"""Local SQLite job tracking and filesystem report storage for the web app."""

from __future__ import annotations

import shutil
import sqlite3
from contextlib import closing
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Protocol


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
    """SQLite report history for the single-instance web service."""

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
            row = connection.execute(
                "SELECT * FROM attendance_jobs WHERE job_id = ?", (job_id,),
            ).fetchone()
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
            eligible = bool(row and (
                row["status"] in accepted
                or (row["status"] == "Processing" and _processing_is_stale(row["notes"]))
            ))
            cursor = connection.execute(
                "UPDATE attendance_jobs SET status = 'Processing', notes = ? WHERE job_id = ?",
                (_processing_note(), job_id),
            ) if eligible else None
            connection.execute("COMMIT")
        return bool(cursor and cursor.rowcount == 1)


class Storage(Protocol):
    def save_pdf(self, job_id: str, pdf_path: str | Path) -> str: ...
    def read_pdf(self, reference: str, local_pdf_path: str = "") -> bytes: ...


class LocalStorage:
    """Store generated PDFs under DATA_DIR/reports."""

    def __init__(self, data_dir: str | Path):
        self.data_dir = Path(data_dir).resolve()
        self.report_dir = self.data_dir / "reports"
        self.report_dir.mkdir(parents=True, exist_ok=True)

    def save_pdf(self, job_id: str, pdf_path: str | Path) -> str:
        source = Path(pdf_path)
        destination = (self.report_dir / source.name).resolve()
        try:
            destination.relative_to(self.report_dir.resolve())
        except ValueError as exc:
            raise ValueError("Generated report filename is invalid.") from exc
        shutil.copyfile(source, destination)
        return f"/api/web/reports/{job_id}/download"

    def read_pdf(self, reference: str, local_pdf_path: str = "") -> bytes:
        try:
            source = Path(local_pdf_path).resolve() if local_pdf_path else Path(reference).resolve()
            source.relative_to(self.report_dir.resolve())
        except (ValueError, OSError) as exc:
            raise FileNotFoundError("The generated report is not available in local storage.") from exc
        if not source.is_file():
            raise FileNotFoundError("The generated report is not available in local storage.")
        return source.read_bytes()
