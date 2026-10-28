"""Attendance data validation, analysis, and A4 PDF report generation.

This module is deliberately safe to import.  The command-line entry point at
the bottom is only invoked when this file is executed as a script.
"""

from __future__ import annotations

import argparse
import html
import io
import re
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Flowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


REQUIRED_COLUMNS = {
    "Panchayat LGD": {
        "panchayatlgd", "panchayatlgdcode", "lgd", "lgdcode",
    },
    "Mandal Name": {"mandalname", "mandal"},
    "Panchayat Name": {
        "panchayatname", "grampanchayat", "grampanchayatname", "gramapanchayatname",
    },
    "First Attendence Status": {
        "firstattendencestatus", "firstattendancestatus", "attendencestatus",
        "attendancestatus", "attendance", "status",
    },
    "First Attendence Datetime": {
        "firstattendencedatetime", "firstattendancedatetime", "attendencedatetime",
        "attendancedatetime", "attendancedateandtime", "attendancetime",
    },
}


class InputValidationError(ValueError):
    """A safe, user-facing validation error for an uploaded attendance file."""


@dataclass
class ReportData:
    report_date: date
    source_row_count: int
    unique_gp_count: int
    missing_lgd_rows: int
    duplicate_lgd_rows: int
    conflicts: list[dict[str, Any]]
    mandal_rows: list[dict[str, Any]]
    district_total: dict[str, Any]
    pending_rows: list[dict[str, Any]]
    leave_rows: list[dict[str, Any]]
    meeting_training_rows: list[dict[str, Any]]
    after_9_rows: list[dict[str, Any]]

    @property
    def totals(self) -> dict[str, int]:
        return {
            "unique_gps": self.unique_gp_count,
            "before_9": int(self.district_total["before_9"]),
            "on_leave": int(self.district_total["on_leave"]),
            "meeting_training": int(self.district_total["meeting_training"]),
            "pending": int(self.district_total["pending"]),
            "after_9": len(self.after_9_rows),
        }


def normalize_column_name(value: Any) -> str:
    """Normalize punctuation, case, and whitespace for tolerant header matching."""
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def _canonicalize_columns(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    frame = frame.loc[:, ~frame.columns.duplicated()].copy()
    aliases = {
        alias: canonical
        for canonical, names in REQUIRED_COLUMNS.items()
        for alias in names
    }
    rename: dict[Any, str] = {}
    found: set[str] = set()
    for column in frame.columns:
        canonical = aliases.get(normalize_column_name(column))
        if canonical and canonical not in found:
            rename[column] = canonical
            found.add(canonical)
    return frame.rename(columns=rename), [name for name in REQUIRED_COLUMNS if name not in found]


def _read_html_table(raw: bytes) -> pd.DataFrame:
    try:
        tables = pd.read_html(io.BytesIO(raw))
    except Exception as exc:
        raise InputValidationError("The uploaded HTML file does not contain a readable table.") from exc
    if not tables:
        raise InputValidationError("No tables were found in the uploaded HTML file.")
    for candidate in tables:
        standardized, missing = _canonicalize_columns(candidate.copy())
        if not missing:
            return standardized

    # Some portal exports put a title row above the actual column headings.
    for candidate in tables:
        frame = candidate.copy()
        for row_index in range(min(len(frame), 20)):
            values = frame.iloc[row_index].tolist()
            tentative = frame.iloc[row_index + 1 :].copy()
            tentative.columns = values
            standardized, missing = _canonicalize_columns(tentative)
            if not missing:
                return standardized
    largest = max(tables, key=lambda table: table.shape[1])
    return largest.copy()


def read_source_file(filename: str, raw_bytes: bytes) -> pd.DataFrame:
    """Read CSV, XLS/XLSX, or an HTML table exported with an .xls extension."""
    extension = Path(filename).suffix.lower()
    if extension not in {".xls", ".xlsx", ".csv"}:
        raise InputValidationError("Allowed file types are .xls, .xlsx, and .csv.")
    if not raw_bytes:
        raise InputValidationError("The uploaded file is empty.")

    prefix = raw_bytes[:8192].lstrip().lower()
    looks_html = (
        prefix.startswith(b"<!doctype html")
        or prefix.startswith(b"<html")
        or (b"<table" in prefix and (b"<tr" in prefix or b"<th" in prefix))
    )
    if looks_html:
        frame = _read_html_table(raw_bytes)
    elif extension == ".csv":
        try:
            try:
                frame = pd.read_csv(
                    io.BytesIO(raw_bytes), sep=None, engine="python", encoding="utf-8-sig", keep_default_na=False,
                )
            except UnicodeDecodeError:
                frame = pd.read_csv(
                    io.BytesIO(raw_bytes), sep=None, engine="python", encoding="cp1252", keep_default_na=False,
                )
        except Exception as exc:
            raise InputValidationError("Could not read the CSV file. Check its encoding and delimiter.") from exc
    else:
        try:
            frame = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=0)
        except Exception as exc:
            raise InputValidationError(
                "Could not read the upload as HTML or Excel. Upload a valid .xls/.xlsx workbook "
                "or an HTML table saved with an .xls extension."
            ) from exc

    frame.columns = [str(column).replace("\n", " ").strip() for column in frame.columns]
    frame, missing = _canonicalize_columns(frame)
    if missing:
        detected = ", ".join(str(column) for column in frame.columns[:40]) or "(none)"
        raise InputValidationError(
            "Required column(s) not found: " + ", ".join(missing) + ". Detected columns: " + detected
        )
    return frame


def _clean_string(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _clean_lgd(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    # Excel numeric cells are often represented as e.g. "12345.0".
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _parse_timestamp(value: Any) -> Any:
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, time):
        return pd.Timestamp(datetime.combine(date.today(), value))
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel stores a time-only cell as a fraction of a day.
        if 0 <= float(value) < 1:
            seconds = round(float(value) * 24 * 60 * 60)
            return pd.Timestamp(datetime.combine(date.today(), time.min)) + pd.Timedelta(seconds=seconds)
        # Also accept Excel serial dates found in CSV exports.
        if 10000 <= float(value) <= 100000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def _format_timestamp(value: Any) -> str:
    if pd.isna(value) or value == "":
        return ""
    try:
        stamp = pd.Timestamp(value)
        return stamp.strftime("%d-%m-%Y %I:%M %p")
    except (TypeError, ValueError, OverflowError):
        return _clean_string(value)


def _report_date(filename: str, timestamps: pd.Series) -> date:
    valid = timestamps.dropna()
    if not valid.empty:
        try:
            return valid.dt.date.mode().iloc[0]
        except (AttributeError, TypeError):
            pass
    match = re.search(r"(\d{1,2})[-_ ]([A-Za-z]{3})[-_ ](\d{4})", filename)
    if match:
        try:
            return datetime.strptime(match.group(0).replace("_", "-"), "%d-%b-%Y").date()
        except ValueError:
            pass
    return date.today()


def analyze_attendance(frame: pd.DataFrame, filename: str = "attendance.xlsx") -> ReportData:
    """Validate and aggregate one unique record per nonblank Panchayat LGD."""
    frame, missing = _canonicalize_columns(frame.copy())
    if missing:
        raise InputValidationError("Required column(s) not found: " + ", ".join(missing))
    frame["Panchayat LGD"] = frame["Panchayat LGD"].map(_clean_lgd)
    frame["Mandal Name"] = frame["Mandal Name"].map(_clean_string)
    frame["Panchayat Name"] = frame["Panchayat Name"].map(_clean_string)
    frame["First Attendence Status"] = frame["First Attendence Status"].map(_clean_string)
    frame["_timestamp"] = frame["First Attendence Datetime"].map(_parse_timestamp)

    valid = frame.loc[frame["Panchayat LGD"].ne("")].copy()
    if valid.empty:
        raise InputValidationError("No nonblank Panchayat LGD codes were found; no GP records can be reported.")

    unique_rows: list[dict[str, Any]] = []
    conflicts: list[dict[str, Any]] = []
    duplicate_count = 0
    for lgd, group in valid.groupby("Panchayat LGD", sort=False, dropna=False):
        duplicate_count += max(0, len(group) - 1)
        def distinct_names(column: str) -> list[str]:
            found: dict[str, str] = {}
            for value in group[column]:
                if value and value.casefold() not in found:
                    found[value.casefold()] = value
            return list(found.values())

        mandals = distinct_names("Mandal Name")
        panchayats = distinct_names("Panchayat Name")
        if len(mandals) > 1 or len(panchayats) > 1:
            conflicts.append({"lgd": str(lgd), "mandals": mandals, "panchayats": panchayats})

        submitted = group.loc[group["First Attendence Status"].ne("")]
        chosen = submitted.iloc[0] if not submitted.empty else group.iloc[0]
        unique_rows.append({
            "lgd": str(lgd),
            "mandal": mandals[0] if mandals else "(Unspecified)",
            "gp": panchayats[0] if panchayats else "(Unnamed GP)",
            "status": chosen["First Attendence Status"],
            "timestamp": chosen["First Attendence Datetime"],
            "parsed_timestamp": chosen["_timestamp"],
        })

    records = pd.DataFrame(unique_rows)
    records["submitted"] = records["status"].ne("")
    cutoff = time(9, 0)

    def is_before_9(stamp: Any) -> bool:
        if pd.isna(stamp):
            return False
        return stamp.time() <= cutoff

    records["before_9"] = records["submitted"] & records["parsed_timestamp"].map(is_before_9)
    records["after_9"] = records["submitted"] & records["parsed_timestamp"].map(
        lambda value: (not pd.isna(value)) and value.time() > cutoff
    )
    records["pending"] = ~records["submitted"]
    status_lower = records["status"].str.casefold()
    records["on_leave"] = status_lower.str.contains("leave", regex=False)
    records["meeting_training"] = (
        status_lower.str.contains("meeting", regex=False)
        | status_lower.str.contains("training", regex=False)
    )
    records["timestamp_text"] = records["timestamp"].map(_format_timestamp)

    mandal_rows: list[dict[str, Any]] = []
    for mandal, group in records.groupby("mandal", sort=False):
        total = len(group)
        before = int(group["before_9"].sum())
        mandal_rows.append({
            "mandal": mandal,
            "total_gps": total,
            "before_9": before,
            "before_9_pct": (before / total * 100) if total else 0.0,
            "on_leave": int(group["on_leave"].sum()),
            "meeting_training": int(group["meeting_training"].sum()),
            "pending": int(group["pending"].sum()),
        })
    mandal_rows.sort(key=lambda item: (item["before_9_pct"], item["mandal"].casefold()))
    for index, item in enumerate(mandal_rows, start=1):
        item["serial"] = index

    district_gps = len(records)
    district_before = int(records["before_9"].sum())
    district_total = {
        "total_gps": district_gps,
        "before_9": district_before,
        "before_9_pct": district_before / district_gps * 100 if district_gps else 0.0,
        "on_leave": int(records["on_leave"].sum()),
        "meeting_training": int(records["meeting_training"].sum()),
        "pending": int(records["pending"].sum()),
    }

    def make_rows(mask: pd.Series) -> list[dict[str, Any]]:
        selected = records.loc[mask].sort_values(["mandal", "gp"], key=lambda column: column.str.casefold(), kind="stable")
        return [
            {
                "serial": index,
                "mandal": row.mandal,
                "gp": row.gp,
                "status": row.status,
                "timestamp": row.timestamp_text,
            }
            for index, row in enumerate(selected.itertuples(index=False), start=1)
        ]

    pending_rows = make_rows(records["pending"])
    leave_rows = make_rows(records["on_leave"])
    meeting_training_rows = make_rows(records["meeting_training"])
    after_9_rows = make_rows(records["after_9"])
    # Use all available recorded dates to infer the report date.
    report_date = _report_date(filename, valid["_timestamp"])
    return ReportData(
        report_date=report_date,
        source_row_count=len(frame),
        unique_gp_count=district_gps,
        missing_lgd_rows=len(frame) - len(valid),
        duplicate_lgd_rows=duplicate_count,
        conflicts=conflicts,
        mandal_rows=mandal_rows,
        district_total=district_total,
        pending_rows=pending_rows,
        leave_rows=leave_rows,
        meeting_training_rows=meeting_training_rows,
        after_9_rows=after_9_rows,
    )


def process_attendance_file(filename: str, raw_bytes: bytes) -> ReportData:
    """Read and analyze an uploaded file."""
    return analyze_attendance(read_source_file(filename, raw_bytes), filename)


class PercentBar(Flowable):
    """A small data bar and label that are clipped to their table cell width."""

    def __init__(self, percent: float, width: float, height: float = 22):
        super().__init__()
        self.percent = max(0.0, min(100.0, float(percent)))
        self.width = width
        self.height = height

    def wrap(self, avail_width: float, avail_height: float) -> tuple[float, float]:
        self.width = min(self.width, max(1, avail_width))
        return self.width, self.height

    def draw(self) -> None:
        self.canv.saveState()
        self.canv.setFont("Helvetica-Bold", 6.5)
        self.canv.setFillColor(HexColor("#1F2937"))
        self.canv.drawCentredString(self.width / 2, self.height - 8, f"{self.percent:.2f}%")
        y, height = 3, 4
        self.canv.setFillColor(HexColor("#E8EDF3"))
        self.canv.roundRect(0, y, self.width, height, 2, fill=1, stroke=0)
        fill_width = min(self.width, self.width * self.percent / 100)
        if fill_width > 0:
            color = "#2E8B57" if self.percent >= 75 else "#E0A11A" if self.percent >= 50 else "#D9534F"
            self.canv.setFillColor(HexColor(color))
            self.canv.roundRect(0, y, fill_width, height, 2, fill=1, stroke=0)
        self.canv.restoreState()


def generate_pdf(report: ReportData, output_path: str | Path) -> Path:
    """Create a professional, printable A4 portrait report."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    page_width, _ = A4
    margin = 12 * mm
    available_width = page_width - 2 * margin
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=15, leading=18, alignment=TA_CENTER, textColor=HexColor("#17365D"), spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        name="ReportSubtitle", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8.5, leading=11, alignment=TA_CENTER, textColor=HexColor("#444444"), spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="SectionTitle", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=11, leading=14, textColor=HexColor("#17365D"), spaceBefore=2, spaceAfter=5,
    ))
    styles.add(ParagraphStyle(
        name="CellLeft", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=8.4,
        alignment=TA_LEFT, wordWrap="CJK",
    ))
    styles.add(ParagraphStyle(
        name="CellCenter", parent=styles["CellLeft"], alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        name="TableHeader", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6.5,
        leading=7.5, alignment=TA_CENTER, textColor=colors.white, wordWrap="CJK",
    ))
    styles.add(ParagraphStyle(
        name="TableHeaderLeft", parent=styles["TableHeader"], alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        name="Note", parent=styles["Normal"], fontName="Helvetica", fontSize=7, leading=9,
        textColor=HexColor("#444444"),
    ))

    def safe(value: Any) -> str:
        return html.escape("" if value is None else str(value))

    def make_table(
        headers: list[str], rows: list[list[Any]], widths: list[float], left_columns: tuple[int, ...] = (),
        total_row: bool = False,
    ) -> Table:
        data: list[list[Any]] = [[
            Paragraph(
                safe(value),
                styles["TableHeaderLeft"] if column_index in left_columns else styles["TableHeader"],
            )
            for column_index, value in enumerate(headers)
        ]]
        for row in rows:
            formatted = []
            for column_index, value in enumerate(row):
                if isinstance(value, Flowable):
                    formatted.append(value)
                else:
                    cell_style = styles["CellLeft"] if column_index in left_columns else styles["CellCenter"]
                    formatted.append(Paragraph(safe(value), cell_style))
            data.append(formatted)
        table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT", splitByRow=1)
        commands: list[tuple[Any, ...]] = [
            ("BACKGROUND", (0, 0), (-1, 0), HexColor("#1F4E78")),
            ("GRID", (0, 0), (-1, -1), 0.35, HexColor("#A6A6A6")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2.2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2.2),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, HexColor("#EAF2F8")]),
        ]
        for column_index in left_columns:
            commands.append(("ALIGN", (column_index, 1), (column_index, -1), "LEFT"))
        if total_row:
            commands.extend([
                ("BACKGROUND", (0, -1), (-1, -1), HexColor("#D9EAF7")),
                ("LINEABOVE", (0, -1), (-1, -1), 0.8, HexColor("#1F4E78")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ])
        table.setStyle(TableStyle(commands))
        return table

    def footer(canvas: Any, document: Any) -> None:
        canvas.saveState()
        width, _ = A4
        canvas.setStrokeColor(HexColor("#B7C9D6"))
        canvas.line(margin, 10 * mm, width - margin, 10 * mm)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(HexColor("#555555"))
        canvas.drawString(margin, 6 * mm, f"Khammam District | PS Attendance Report | {report.report_date:%d-%m-%Y}")
        canvas.drawRightString(width - margin, 6 * mm, f"Page {document.page}")
        canvas.restoreState()

    story: list[Any] = []
    story.extend([
        Paragraph("KHAMMAM DISTRICT", styles["ReportTitle"]),
        Paragraph("PANCHAYAT SECRETARY ATTENDANCE REPORT", styles["ReportTitle"]),
        Paragraph(f"Report Date: {report.report_date:%d-%m-%Y} | Unique GPs: {report.unique_gp_count:,}", styles["ReportSubtitle"]),
        Paragraph("1. Mandal-wise Attendance Abstract", styles["SectionTitle"]),
    ])

    abstract_headers = [
        "Sl. No.", "Mandal", "No. of GPs",
        "No. of GPs submitted attendance before 9 AM",
        "% of GPs submitted attendance before 9 AM", "No. of PS on leave",
        "No. of PS in training or meeting", "No. of GPs with attendance pending",
    ]
    widths = [10 * mm, 26 * mm, 14 * mm, 33 * mm, 27 * mm, 17 * mm, 32 * mm, 27 * mm]
    # Keep all eight columns inside the printable A4 width even if margins change.
    scale = min(1.0, available_width / sum(widths))
    widths = [width * scale for width in widths]
    abstract_rows: list[list[Any]] = []
    for item in report.mandal_rows:
        abstract_rows.append([
            item["serial"], item["mandal"], item["total_gps"], item["before_9"],
            PercentBar(item["before_9_pct"], widths[4] - 5), item["on_leave"],
            item["meeting_training"], item["pending"],
        ])
    total = report.district_total
    abstract_rows.append([
        "", "DISTRICT TOTAL", total["total_gps"], total["before_9"],
        PercentBar(total["before_9_pct"], widths[4] - 5), total["on_leave"],
        total["meeting_training"], total["pending"],
    ])
    story.append(make_table(abstract_headers, abstract_rows, widths, left_columns=(1,), total_row=True))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Method: A nonblank attendance status counts as submitted. Attendance at or before 09:00 counts "
        "as before 9 AM; after 09:00 counts as after 9 AM. Pending means the status is blank. "
        "All totals use one record per nonblank Panchayat LGD code.", styles["Note"],
    ))
    if report.duplicate_lgd_rows:
        story.append(Paragraph(
            f"Validation: Excluded {report.duplicate_lgd_rows} duplicate LGD row(s).", styles["Note"],
        ))
    if report.missing_lgd_rows:
        story.append(Paragraph(
            f"Validation: Excluded {report.missing_lgd_rows} source row(s) without a Panchayat LGD code.",
            styles["Note"],
        ))
    if report.conflicts:
        details = "; ".join(
            f"{item['lgd']} (Mandal: {', '.join(item['mandals']) or '-'}; GP: {', '.join(item['panchayats']) or '-'})"
            for item in report.conflicts[:8]
        )
        suffix = f"; and {len(report.conflicts) - 8} more" if len(report.conflicts) > 8 else ""
        story.append(Paragraph(
            f"Validation warning: {len(report.conflicts)} LGD code(s) have conflicting names: {safe(details + suffix)}.",
            styles["Note"],
        ))

    def append_list_section(title: str, total_label: str, rows: list[dict[str, Any]]) -> None:
        story.append(PageBreak())
        story.append(Paragraph(title, styles["SectionTitle"]))
        story.append(Paragraph(f"{total_label}: {len(rows):,}", styles["ReportSubtitle"]))
        headers = ["Sl. No.", "Mandal", "Gram Panchayat", "Attendance Status", "Attendance Date & Time"]
        list_widths = [12 * mm, 38 * mm, 45 * mm, 39 * mm, available_width - 134 * mm]
        list_rows: list[list[Any]] = [
            [item["serial"], item["mandal"], item["gp"], item["status"] or "Pending", item["timestamp"]]
            for item in rows
        ]
        if not list_rows:
            list_rows = [["-", "-", "No records found for this section.", "-", "-"]]
        table = make_table(headers, list_rows, list_widths, left_columns=(1, 2, 3, 4))
        if not rows:
            table.setStyle(TableStyle([("SPAN", (2, 1), (4, 1)), ("ALIGN", (2, 1), (4, 1), "LEFT")]))
        story.append(table)

    append_list_section("2. ATTENDANCE PENDING LIST", "Total pending GPs", report.pending_rows)
    append_list_section("3. PS ON LEAVE LIST", "Total PS on leave", report.leave_rows)
    append_list_section(
        "4. TRAINING / MEETING LIST", "Total training/meeting records", report.meeting_training_rows,
    )
    append_list_section(
        "5. GPs THAT SUBMITTED ATTENDANCE AFTER 9 AM", "Total GPs after 9 AM", report.after_9_rows,
    )

    doc = SimpleDocTemplate(
        str(output), pagesize=A4, rightMargin=margin, leftMargin=margin,
        topMargin=13 * mm, bottomMargin=14 * mm,
        title="Khammam District PS Attendance Report", author="Khammam District",
    )
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate Khammam PS attendance PDF report")
    parser.add_argument("input_file", help="Path to attendance raw-data file")
    parser.add_argument("--output", default=None, help="Optional output PDF path")
    args = parser.parse_args()
    source = Path(args.input_file)
    try:
        report = process_attendance_file(source.name, source.read_bytes())
        destination = Path(args.output) if args.output else Path(
            f"Khammam_PS_Attendance_Report_{report.report_date:%d-%b-%Y}.pdf"
        )
        generate_pdf(report, destination)
    except InputValidationError as exc:
        parser.error(str(exc))
    print(f"PDF report created: {destination}")
    print(f"Unique GPs: {report.unique_gp_count:,}; duplicate LGD rows excluded: {report.duplicate_lgd_rows:,}")
    print(f"Totals: {report.totals}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
