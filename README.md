# Khammam PS Attendance Analytics

FastAPI service that accepts Gram Panchayat attendance exports, deduplicates records by Panchayat LGD, generates a formatted A4 PDF, and tracks each report in local SQLite or a Google Sheet. Google Drive and Sheets use Application Default Credentials on Cloud Run; no credentials are stored in the source tree.

## What it does

- Reads `.xls`, `.xlsx`, `.csv`, and HTML tables saved with an `.xls` extension.
- Normalizes capitalization, punctuation, spacing, and common `Attendance` / `Attendence` header spellings.
- Uses one record per nonblank Panchayat LGD. Duplicate rows do not increase totals; conflicting Mandal or GP names are flagged in the PDF and job notes.
- Counts any nonblank status as submitted. Attendance at exactly 09:00 is counted before 9 AM; after-9-AM is strictly later than 09:00. Blank status alone is pending. Leave and meeting/training remain submitted.
- Generates a portrait A4 report with a Mandal abstract, district totals, percentage bars, pending, leave, training/meeting, and after-9-AM lists.
- Protects report endpoints with `x-api-token`. `GET /health` is public. Cloud Tasks requests use verified OIDC identity.
- Supports local SQLite/filesystem mode, Google Sheets/Drive mode, and optional durable asynchronous processing with Google Cloud Tasks.

## Run locally on Windows

Use Python 3.11 or newer. From PowerShell in this folder:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements-dev.txt
Copy-Item .env.example .env
```

Edit `.env` and replace `APP_API_TOKEN` with a random value of at least 32 characters. Keep the other Google and Cloud Tasks settings blank for local mode. Then start the server:

```powershell
uvicorn app:app --reload --port 8080 --env-file .env
```

Open `http://127.0.0.1:8080/docs`. Authenticate with the `x-api-token` header. The interactive API page can upload test files. Local jobs are stored in `data/attendance_jobs.sqlite3`, PDFs in `data/reports/`, and any referenced local inputs must be placed under `data/inputs/` (or the folder configured as `LOCAL_INPUT_DIR`). The `data/` folder is ignored by Git.

Example direct upload using `curl.exe`:

```powershell
curl.exe -X POST "http://127.0.0.1:8080/api/reports" `
  -H "x-api-token: replace-with-a-long-random-token-from-env-file" `
  -F "JobID=local-demo-001" `
  -F "RequestedBy=operator@example.org" `
  -F "file=@C:\path\to\synthetic-attendance.xlsx"
```

The response contains `job_id`, `status`, and `output_pdf`. With local synchronous processing, a successful response is `Completed`. Poll `GET /api/reports/{job_id}` and download at `GET /api/reports/{job_id}/download`. Reusing a JobID returns its existing job and does not generate a second report.

Run the automated synthetic-data test suite:

```powershell
pytest -q
```

## API

All routes below except health require `x-api-token`.

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/health` | Public liveness and configured-mode check |
| `POST` | `/api/reports` | Multipart upload (`file`, optional `JobID`, `RequestedBy`) |
| `POST` | `/api/reports/from-drive` | AppSheet JSON webhook with `JobID`, `UploadedFile` or `file_id`, and `RequestedBy` |
| `GET` | `/api/reports/{job_id}` | Read status and output reference |
| `GET` | `/api/reports/{job_id}/download` | Download the generated PDF |

Jobs transition through `Submitted`, `Processing`, `Completed`, or `Failed`. A failed validation response includes the JobID and safe, actionable message; internal exception details and credentials are not returned. In local mode and without Cloud Tasks, report generation is synchronous. For production webhook requests, enable Cloud Tasks so AppSheet receives `Submitted` promptly while the backend processes and updates the sheet asynchronously.

### AppSheet Drive reference payload

```json
{
  "JobID": "the-AppSheet-JobID",
  "UploadedFile": "AttendanceUploads/attendance.xlsx",
  "RequestedBy": "operator@example.org"
}
```

The backend accepts a Drive file ID, a standard Drive file URL, or AppSheet's uploaded path/name. Path/name references are resolved by exact filename inside `DRIVE_UPLOAD_FOLDER_ID`. If more than one file has the same name, pass its Drive file ID. A supplied file must belong directly to the configured upload folder. Generated PDFs are created in the configured output folder without public permissions.

## Google Drive and Sheets configuration

Set these environment variables to switch from local mode:

```text
DRIVE_UPLOAD_FOLDER_ID=<private Drive folder ID for AppSheet uploads>
DRIVE_OUTPUT_FOLDER_ID=<private Drive folder ID for generated PDFs>
SHEETS_SPREADSHEET_ID=<spreadsheet ID>
SHEETS_JOBS_TAB=AttendanceReportJobs
```

The sheet tab must have this exact first row, in this order:

```text
JobID | UploadedFile | RequestedBy | RequestedAt | Status | OutputPDF | Notes
```

Share the upload folder with the Cloud Run service account as Viewer and the output folder as Editor. Share the spreadsheet with the service account as Editor. Share the output folder with the authorized report viewers as Viewer if they need to open PDFs; do not enable public link access. When running locally with a service-account key, set `GOOGLE_APPLICATION_CREDENTIALS` to its path outside the repository. On Cloud Run, attach the service account and leave that variable blank.

`AttendanceReportJobs` must be the same sheet AppSheet uses. AppSheet creates the submitted row first; the backend claims that `Submitted` row, then updates status, output URL, and notes in place. Retries with the same JobID do not create another report. The Google Sheets API does not provide a transactional compare-and-swap; the optional Cloud Tasks deployment uses queue concurrency one for its worker path.

## Durable asynchronous processing with Cloud Tasks

Cloud Tasks is optional in local development and recommended for AppSheet production webhooks. It requires Drive and Sheets mode. Create a queue with a single concurrent dispatch:

```powershell
gcloud tasks queues create attendance-report-jobs `
  --location=asia-south1 `
  --max-concurrent-dispatches=1 `
  --max-attempts=5 `
  --min-backoff=10s `
  --max-backoff=300s
```

Set all of the following on the Cloud Run service:

```text
CLOUD_TASKS_PROJECT=<Google Cloud project ID>
CLOUD_TASKS_LOCATION=asia-south1
CLOUD_TASKS_QUEUE=attendance-report-jobs
PUBLIC_BASE_URL=https://<Cloud Run service URL>
TASK_INVOKER_SERVICE_ACCOUNT=<OIDC caller service account email>
```

The attached Cloud Run runtime service account needs `roles/cloudtasks.enqueuer` and `roles/iam.serviceAccountUser` on `TASK_INVOKER_SERVICE_ACCOUNT` so it can create tasks using that identity. The OIDC caller service account must be in the queue's project and needs `roles/run.invoker` on the Cloud Run service. Grant the Cloud Tasks service agent (`service-PROJECT_NUMBER@gcp-sa-cloudtasks.iam.gserviceaccount.com`) `roles/iam.serviceAccountUser` on the OIDC caller service account so Cloud Tasks can mint the token. The internal worker verifies the audience and caller email. Cloud Tasks retries failed processing requests; each task is deterministically named from JobID, and completed jobs are idempotent.

## Docker and Cloud Run deployment

For Render deployment, see [README_Render.md](README_Render.md). The Render Blueprint uses a free plan and requires Google Drive/Sheets configuration before operational use.

The container runs as an unprivileged user and listens on Cloud Run's `PORT`. Build and deploy from this directory:

```powershell
gcloud config set project YOUR_PROJECT_ID
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com `
  drive.googleapis.com sheets.googleapis.com secretmanager.googleapis.com cloudtasks.googleapis.com

gcloud run deploy khammam-attendance `
  --source . `
  --region asia-south1 `
  --port 8080 `
  --timeout 900 `
  --concurrency 1 `
  --max 3 `
  --cpu 2 `
  --memory 2Gi `
  --service-account khammam-attendance-runtime@YOUR_PROJECT_ID.iam.gserviceaccount.com `
  --allow-unauthenticated `
  --set-env-vars "DATA_DIR=/tmp/khammam-attendance-data,DRIVE_UPLOAD_FOLDER_ID=UPLOAD_FOLDER_ID,DRIVE_OUTPUT_FOLDER_ID=OUTPUT_FOLDER_ID,SHEETS_SPREADSHEET_ID=SPREADSHEET_ID,SHEETS_JOBS_TAB=AttendanceReportJobs"
```

The service is reachable over HTTPS so AppSheet can call it; every report endpoint still requires the application token. `--allow-unauthenticated` controls Cloud Run IAM ingress, not application authorization. Keep the API token in Secret Manager, never in source or an AppSheet expression. Create a secret version from a file outside the repo, then bind it to `APP_API_TOKEN` with `--set-secrets APP_API_TOKEN=khammam-api-token:latest`:

```powershell
gcloud secrets create khammam-api-token --replication-policy=automatic
gcloud secrets versions add khammam-api-token --data-file=C:\secure\khammam-api-token.txt
gcloud secrets add-iam-policy-binding khammam-api-token `
  --member="serviceAccount:khammam-attendance-runtime@YOUR_PROJECT_ID.iam.gserviceaccount.com" `
  --role="roles/secretmanager.secretAccessor"
```

Then deploy/update the service with the `--set-secrets` option. If using Cloud Tasks, first deploy with queue variables unset, copy the HTTPS service URL, create/configure the queue and service accounts, then update the service with the five Cloud Tasks variables above. Cloud Run stores only temporary local files; production reports are persisted in Drive and jobs in Sheets.

### Google Cloud setup and least privilege

1. Enable Cloud Run, Cloud Build, Artifact Registry, Drive API, Sheets API, Secret Manager, and Cloud Tasks API when using queued processing.
2. For `gcloud run deploy --source`, grant the deployer Cloud Run Source Developer, Service Usage Consumer, and Service Account User on the runtime identity. Grant Cloud Run Builder to the Cloud Build service account.
3. Create a runtime service account for the Cloud Run revision. Grant it Secret Manager Secret Accessor for the API-token secret and Cloud Tasks Enqueuer if queueing is enabled.
4. Share only the upload folder, output folder, and jobs spreadsheet with the runtime service account. Use Viewer on uploads, Editor on output and the spreadsheet.
5. Create a Cloud Tasks OIDC caller service account in the queue project. Grant it Cloud Run Invoker on this service; grant the Cloud Tasks service agent and the Cloud Run runtime account Service Account User on the caller account. Grant the runtime account Cloud Tasks Enqueuer.
6. Share the private output folder with authorized AppSheet users or an authorized Google Group as Viewer. Do not create public Drive links.
7. Set `APP_API_TOKEN` from Secret Manager and verify `/health`, an authorized synthetic upload, status polling, and PDF download before operational use.

No cloud resources are created by this repository or its tests. Deployment is a manual operator action.

## Security and operational notes

- Upload names are reduced to a basename; file types are allow-listed and contents are parsed as data only. Uploads are limited to 30 MiB by default (`MAX_UPLOAD_BYTES`).
- Drive reads are restricted to direct children of the configured upload folder; generated file reads are restricted to the output folder.
- Service account keys are not required on Cloud Run. `.gitignore` excludes `.env`, credential files, local databases, temporary data, and generated PDFs.
- `RequestedBy` is an audit value supplied by AppSheet. Enforce the actual user allowlist and row ownership in AppSheet security filters; the shared API token authenticates the AppSheet automation service.
- Google Sheet rows are the durable job log. Local SQLite/filesystem mode is for development only and is not shared across Cloud Run instances.
- Do not configure a Google Sheets service account key in `GOOGLE_APPLICATION_CREDENTIALS` on Cloud Run; attach a service account instead.

## Troubleshooting

- **401 Authentication required:** send the exact configured `x-api-token`; check that AppSheet's webhook task has the header and is calling the HTTPS service URL.
- **503 API authentication is not configured:** set `APP_API_TOKEN` (via Secret Manager in production).
- **Drive file not found:** ensure the AppSheet upload folder matches `DRIVE_UPLOAD_FOLDER_ID`, the service account can view the file, and the file is directly inside that folder. Use a Drive file ID when names collide.
- **Drive permission denied:** share the input/output folders with the attached Cloud Run service account; share the output folder with the people who should view reports.
- **Missing required columns:** check the detected-column detail in the job response/Notes and confirm the export has LGD, Mandal, Panchayat, status, and datetime columns.
- **XLS read error:** HTML exports saved as `.xls` are supported; genuine binary `.xls` requires `xlrd`, included in the dependencies.
- **Task remains Submitted:** check Cloud Tasks queue state, its OIDC caller identity, Cloud Run invoker binding, and the Cloud Run request logs for `task_enqueue_failed` or `job_failed` events.
- **AppSheet does not show a new status:** sync the app; the backend updates the shared Google Sheet directly. Confirm the app is using the same spreadsheet and tab.
- **Output link opens access denied:** grant the authorized users or group Viewer access to the private output folder; do not make the file public.

## Project files

- `app.py` — protected FastAPI API, validation, job orchestration, and optional Cloud Tasks worker endpoint.
- `report_generator.py` — reusable input readers, attendance analysis, and PDF report generation.
- `integrations.py` — local SQLite/filesystem adapters, Google Drive/Sheets adapters, and Cloud Tasks dispatcher.
- `README_AppSheet_Setup.txt` — step-by-step AppSheet setup and webhook expressions.
- `tests/` — synthetic input and API workflow tests.
- `Dockerfile`, `.dockerignore`, `requirements.txt`, `requirements-dev.txt`, `.env.example`, `.gitignore` — runtime, development, and configuration files.

