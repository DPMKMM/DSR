KHAMMAM PS ATTENDANCE ANALYTICS — APPSHEET SETUP
================================================

Architecture
------------
AppSheet stores a job row and the uploaded file in Google Drive. A Bot calls
the protected HTTPS FastAPI endpoint with the AppSheet JobID and UploadedFile
path. The backend reads the file from a configured private Drive folder,
generates the PDF, saves it into a private output folder, and updates the same
Google Sheet row. Python runs only in the FastAPI service; AppSheet does not
execute Python.

Before creating the app
-----------------------
1. Deploy the API and set APP_API_TOKEN in Secret Manager. Keep the token out of
   app formulas and user-facing columns.
2. Create private Drive folders named AttendanceUploads and AttendanceReports.
   Record each folder's Drive ID. Set DRIVE_UPLOAD_FOLDER_ID and
   DRIVE_OUTPUT_FOLDER_ID on the backend. Share the upload folder with the
   backend service account as Viewer and the output folder as Editor.
3. Share AttendanceReports with the authorized user group as Viewer so report
   links work without public sharing.
4. Create a Google spreadsheet and a tab named AttendanceReportJobs. Put these
   exact headers in Row 1, in order:

   JobID | UploadedFile | RequestedBy | RequestedAt | Status | OutputPDF | Notes

5. Share the spreadsheet with the backend service account as Editor. Set its
   ID as SHEETS_SPREADSHEET_ID. The backend preserves these columns and writes
   status/output/notes back to the existing row.

Create the AppSheet application
-------------------------------
1. In AppSheet, choose Create > App > Start with existing data and select the
   spreadsheet. Add the AttendanceReportJobs tab as the main table.
2. In Data > Columns, set:
   - JobID: Text, Key = ON, Initial value = UNIQUEID(), Editable = OFF.
   - UploadedFile: File, Required = ON. Set the file folder/path to
     AttendanceUploads (the same Drive folder represented by
     DRIVE_UPLOAD_FOLDER_ID). Confirm a test upload is stored in that folder.
   - RequestedBy: Email, Initial value = USEREMAIL(), Editable = OFF.
   - RequestedAt: DateTime, Initial value = NOW(), Editable = OFF.
   - Status: Enum with values Submitted, Processing, Completed, Failed; leave
     it optional with a blank initial value. Users submit only through the
     Generate Report action described below.
   - OutputPDF: URL, Editable = OFF.
   - Notes: LongText, Editable = OFF.
3. Require users to sign in. Add a separate AttendanceReportUsers table with
   columns Email (key), Active (Yes/No), and Role (for example User/Admin).
   List only authorized accounts there.
4. Apply this security filter to AttendanceReportJobs to limit visibility to
   active users and show each user their own jobs; admins can see all jobs:

   AND(
     IN(USEREMAIL(), SELECT(AttendanceReportUsers[Email], [Active] = TRUE)),
     OR(
       [RequestedBy] = USEREMAIL(),
       LOOKUP(USEREMAIL(), "AttendanceReportUsers", "Email", "Role") = "Admin"
     )
   )

   Apply a corresponding filter to AttendanceReportUsers so users see only
   their own access record while admins can manage the full list. Keep user
   administration restricted to app owners/admins.

Create the upload form and Generate Report action
------------------------------------------------
1. Add a Form view over AttendanceReportJobs. Include UploadedFile and the
   read-only requester/date fields. Hide JobID, Status, OutputPDF, and Notes.
2. Add an action on AttendanceReportJobs:
   - Name: Generate Report
   - Action type: Data: set the values of some columns in this row
   - Set Status to "Submitted" and Notes to ""
   - Only if this condition is true: ISBLANK([Status])
   - Prominence: Display prominently
3. Add a detail view with the Generate Report action. Users upload and save a
   job row, then tap Generate Report. This transition to Submitted triggers
   the Bot. The action is unavailable after submission, so the same JobID is
   not sent repeatedly by the UI.

Create the automation webhook
-----------------------------
1. In Automation > Bots, create a Bot named Generate Khammam Attendance Report.
2. Event: Data change on AttendanceReportJobs, Updates only. Use this event
   condition so only the button transition triggers the webhook:

   AND(
     [Status] = "Submitted",
     [_THISROW_BEFORE].[Status] <> "Submitted"
   )

3. Add a Webhook task:
   - HTTP method: POST
   - URL: https://YOUR-CLOUD-RUN-URL/api/reports/from-drive
   - HTTP content type: application/json
   - Header: x-api-token = the APP_API_TOKEN stored in the API's Secret Manager
     configuration. Restrict editing this Bot to app administrators.
   - Body:

     {
       "JobID": "<<[JobID]>>",
       "UploadedFile": "<<[UploadedFile]>>",
       "RequestedBy": "<<[RequestedBy]>>",
       "original_name": "<<[UploadedFile]>>"
     }

   The service accepts a Drive file ID, a Google Drive URL, or AppSheet's
   relative uploaded file path. For a path/name, it searches for an exact
   filename inside DRIVE_UPLOAD_FOLDER_ID. If filenames can repeat, configure
   an automation helper to resolve the file ID and send it as `file_id`.
4. Test with a synthetic file. The API response is Submitted when Cloud Tasks
   is enabled and Completed in local/synchronous mode. In production the
   backend updates the shared sheet directly; AppSheet displays the latest
   status after sync.

Report history and access views
-------------------------------
1. Create a "My Report History" table/deck view from AttendanceReportJobs with
   a slice filter `[RequestedBy] = USEREMAIL()`. Show RequestedAt, UploadedFile,
   Status, Notes, and OutputPDF.
2. Create an "All Reports" admin view using a slice condition such as:

   LOOKUP(USEREMAIL(), "AttendanceReportUsers", "Email", "Role") = "Admin"

3. Configure OutputPDF as a URL so AppSheet renders a link. The link points to
   a private Drive file; authorized users need Viewer access to
   AttendanceReports. Never make this folder or its files public.
4. Add a sync action or rely on normal AppSheet sync so status and output URL
   changes written by the backend appear in the app.

Important deployment notes
--------------------------
- AppSheet's File value may be a relative path, not a Drive file ID. The backend
  resolves its basename only within the configured upload folder. Confirm the
  actual folder and file path with a test upload.
- The shared API token authenticates the AppSheet automation. AppSheet sign-in,
  security filters, and the AttendanceReportUsers table enforce user access.
- Cloud Tasks is recommended for production so webhook calls return promptly.
  Set CLOUD_TASKS_PROJECT, CLOUD_TASKS_LOCATION, CLOUD_TASKS_QUEUE,
  PUBLIC_BASE_URL, and TASK_INVOKER_SERVICE_ACCOUNT. The queue should have
  max-concurrent-dispatches=1. The README.md has Cloud Run and IAM commands.
- Do not upload service-account key JSON files to the repository. Cloud Run
  should use an attached service account and Application Default Credentials.
- The backend writes private Drive URLs. It does not grant anyone public access.
