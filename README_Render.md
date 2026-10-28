# Deploying to Render

This repository includes a Render Blueprint (`render.yaml`) for the FastAPI service. Render deploys web services from a connected Git provider; it does not deploy this local folder directly. Render supports GitHub, GitLab, and Bitbucket repositories. See [Connect your Git Provider to Render](https://render.com/docs/git-provider).

## Current deployment

- Service: [khammam-ps-attendance-api](https://khammam-ps-attendance-api.onrender.com)
- Render Dashboard: <https://dashboard.render.com/web/srv-dbh2l6kipgvc73djn3l0>
- Plan/region: Free, Singapore
- Source: `https://github.com/DPMKMM/DSR`
- Deploy behavior: manual deploys. The service was created from the public repository URL with automatic deploys disabled.
- Storage status: local filesystem and SQLite until the Google Drive/Sheets settings below are supplied.

The service home page is a direct browser upload form: select an `.xls`, `.xlsx`, or `.csv` attendance file and it generates and downloads the PDF. Download names use the attendance date from the spreadsheet and the upload time in India Standard Time, in `PS_Attendance_Report_YYYY-MM-DD_HH-MM-SS.pdf` format. This page does not ask for the API token. Its browser upload endpoint is intentionally public and applies an in-memory limit of 20 upload attempts per client IP per hour; report status and download are restricted to the short-lived secure cookie issued for that upload. The AppSheet endpoints and the regular `/api/reports` API remain protected by `x-api-token`. Configure private Drive/Sheets storage before using the service for operational records; local Render storage can be cleared on restart or redeploy.

This service was created directly in Render and is not currently managed by syncing `render.yaml`. The Blueprint is a reproducible alternative; do not create another service from it unless you intend to run a second service.

## 1. Put the project in a private Git repository

For a new deployment, create a **private** GitHub, GitLab, or Bitbucket repository. The current `DPMKMM/DSR` repository is public. From PowerShell:

```powershell
Set-Location C:\Users\DPM\Desktop\REPOR
git init
git add .
git status --short
git commit -m "Prepare Khammam attendance reporting service"
git branch -M main
git remote add origin https://github.com/YOUR_ACCOUNT/khammam-ps-attendance.git
git push -u origin main
```

Review `git status --short` before committing. `.gitignore` excludes local `.env` files, credential files, generated reports, temporary uploads, and local job data. Do not add service-account JSON or API tokens to Git.

## 2. Create or recreate the Render service

1. Sign in to Render and connect the Git provider that contains the private repository.
2. In the Render Dashboard, choose **New > Blueprint** and select the repository.
3. Render reads `render.yaml`. Confirm the service name and region (`Singapore`) and create the service.
4. The Blueprint generates `APP_API_TOKEN` as a secret. Keep it private; the AppSheet webhook will need it later. The current service already has a randomly generated token stored in its Render environment.
5. Wait for the deploy to become live. Check `https://<service-name>.onrender.com/health`.

The Blueprint starts in local storage mode so the first deploy can boot without Google credentials. Do not use that mode for operational attendance data: Render's free service filesystem is ephemeral, so local uploads, PDFs, and SQLite jobs can be lost on restart, redeploy, or spin-down.

## 3. Configure durable Google Drive and Sheets storage

Before connecting AppSheet or uploading real attendance records:

1. In the Render service, open **Environment > Secret Files** and add `google-service-account.json` with the service-account JSON contents. Render mounts Docker service secret files under `/etc/secrets/` at runtime; this matches the Blueprint's `GOOGLE_APPLICATION_CREDENTIALS` setting. See [Render Docker secrets](https://render.com/docs/docker-secrets).
2. Add these environment variables in the service dashboard:
   - `DRIVE_UPLOAD_FOLDER_ID`: private Drive folder where AppSheet stores source files.
   - `DRIVE_OUTPUT_FOLDER_ID`: private Drive folder for generated PDFs.
   - `SHEETS_SPREADSHEET_ID`: spreadsheet ID containing the `AttendanceReportJobs` tab.
3. Ensure the jobs tab's first row is:

   ```text
   JobID | UploadedFile | RequestedBy | RequestedAt | Status | OutputPDF | Notes
   ```

4. Grant the service account Viewer on the upload folder, Editor on the output folder, and Editor on the jobs spreadsheet. Give the report viewers access to the output folder as Viewer; keep it private.
5. Save the environment changes and redeploy. Confirm `/health` reports Drive storage and Google Sheets tracking before use.

`APP_API_TOKEN` is stored in the Render service environment and is used in the `x-api-token` header for protected API routes. Configure the current value as a secret in the AppSheet webhook task; if you cannot retrieve it, replace it in Render with a new random value of at least 32 characters, save, and redeploy. Never put it in a public AppSheet expression or in source control.

## 4. Connect AppSheet

Use the service URL shown on the Render service page:

- Webhook URL: `https://<service-name>.onrender.com/api/reports/from-drive`
- Method: `POST`
- Content type: `application/json`
- Header: `x-api-token: <APP_API_TOKEN>`
- Body: pass the AppSheet `JobID`, uploaded Drive file reference, and `RequestedBy`, following `README_AppSheet_Setup.txt`.

The backend accepts a Drive file ID, URL, or uploaded path/name. Prefer the Drive file ID if filenames can repeat. The generated PDF stays in the private output folder; report viewers need Drive access to open it.

## Free plan behavior

The Blueprint selects Render's **Free** web-service plan to avoid selecting a paid plan without your approval. Free web services spin down after 15 minutes without traffic, can take about a minute to wake, and have ephemeral filesystems. Durable Drive/Sheets configuration is required even on the free plan. For an always-on production service, choose a paid plan in Render after reviewing its current pricing and service limits. See [Render free service limitations](https://render.com/docs/free).

This service processes requests synchronously unless Cloud Tasks is configured. If AppSheet reports webhook timeouts for large inputs, configure the documented Cloud Tasks workflow or use a paid always-on service and review the request size/processing time.

## Updating the deployment

The repository includes `.github/workflows/render-deploy.yml`, which calls the Render deploy hook for application and dependency changes pushed to `main`. To enable it:

1. Regenerate the hook in the Render service's **Settings > Deploy Hook** section because the previous URL was shared in chat.
2. In GitHub, open **Settings > Secrets and variables > Actions > New repository secret**.
3. Name the secret `RENDER_DEPLOY_HOOK_URL` and paste the newly generated hook URL as its value.
4. Save it. Future pushes to `main` that change the application, Docker, or dependency files will trigger a deploy. You can also run the workflow manually from the GitHub Actions tab.

The workflow does not contain the hook URL; it skips safely and warns if the secret has not been configured. Until the secret is added, deploy the latest commit manually from the Render Dashboard. Do not commit `.env`, API tokens, deploy-hook URLs, service-account JSON, uploaded attendance data, or generated PDFs.

## Local validation

From PowerShell in the project folder:

```powershell
docker build -t khammam-ps-attendance .
docker run --rm -p 8080:8080 `
  -e APP_API_TOKEN=replace-with-a-random-secret-of-at-least-32-characters `
  -e DATA_DIR=/tmp/khammam-attendance-data `
  khammam-ps-attendance
```

Then open `http://localhost:8080/health`. The Render Docker service uses the same Dockerfile and listens on Render's assigned `PORT`.
