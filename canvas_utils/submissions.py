"""Submission timestamp extraction logic for Canvas assignments."""

from __future__ import annotations

import csv
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Sequence

from canvas_utils.api_client import CanvasClient

CORE_HEADERS = [
    "Assignment_ID",
    "Student_ID",
    "Attempt_Number",
    "Submitted_At",
    "Canvas_Late_Flag",
]

OPTIONAL_FIELD_SPECS = {
    "student_name": ("Student_Name", lambda submission: _submission_user(submission).get("name")),
    "score": ("Score", lambda submission: submission.get("score")),
    "grade": ("Grade", lambda submission: submission.get("grade")),
    "entered_grade": ("Entered_Grade", lambda submission: submission.get("entered_grade")),
    "workflow_state": ("Workflow_State", lambda submission: submission.get("workflow_state")),
    "missing": ("Missing", lambda submission: submission.get("missing")),
    "excused": ("Excused", lambda submission: submission.get("excused")),
    "late_policy_status": (
        "Late_Policy_Status",
        lambda submission: submission.get("late_policy_status"),
    ),
    "posted_at": ("Posted_At", lambda submission: submission.get("posted_at")),
    "seconds_late": ("Seconds_Late", lambda submission: submission.get("seconds_late")),
    "user_login_id": (
        "Student_Login_ID",
        lambda submission: _submission_user(submission).get("login_id"),
    ),
    "user_sis_id": (
        "Student_SIS_ID",
        lambda submission: _submission_user(submission).get("sis_user_id"),
    ),
    "user_short_name": (
        "Student_Short_Name",
        lambda submission: _submission_user(submission).get("short_name"),
    ),
    "submitted_at_excel": (
        "Submitted_At_Excel",
        lambda submission: _format_excel_datetime(submission.get("submitted_at")),
    ),
    "submitted_at_brisbane": (
        "Submitted_At_Brisbane",
        lambda submission: _format_brisbane_datetime(submission.get("submitted_at")),
    ),
}


def extract_submission_timestamps(
    client: CanvasClient,
    course_id: int,
    assignment_id: int | None,
    output: str | Path = ".",
    optional_fields: Sequence[str] | None = None,
) -> list[Path]:
    """Extract submissions and write one CSV file per assignment.

    Args:
        client: Authenticated Canvas API client.
        course_id: Canvas course ID.
        assignment_id: Target assignment ID, or None to process all assignments.
        output: Output directory or CSV file path for single-assignment mode.
        optional_fields: Additional Canvas submission fields to append to the CSV.

    Returns:
        A list of written CSV file paths.
    """
    selected_optional_fields = resolve_optional_fields(optional_fields)
    headers = build_csv_headers(selected_optional_fields)
    target_assignments = _resolve_assignment_records(client, course_id, assignment_id)
    output_path = Path(output)
    written_files: list[Path] = []

    for current_assignment_id, assignment_name in target_assignments:
        submissions = client.list_submissions(
            course_id=course_id,
            assignment_id=current_assignment_id,
            include_user=True,
        )
        rows = [
            _map_submission(submission, current_assignment_id, selected_optional_fields)
            for submission in submissions
        ]
        csv_path = _resolve_csv_output_path(
            output_path,
            assignment_id,
            current_assignment_id,
            assignment_name,
        )
        _write_csv(csv_path, rows, headers)
        written_files.append(csv_path)

    return written_files


def build_csv_headers(optional_fields: Sequence[str]) -> list[str]:
    """Build the ordered CSV header row for the requested fields."""
    headers = list(CORE_HEADERS)
    headers.extend(OPTIONAL_FIELD_SPECS[field][0] for field in optional_fields)
    return headers


def resolve_optional_fields(optional_fields: Sequence[str] | None) -> tuple[str, ...]:
    """Validate and normalize optional field selections."""
    if optional_fields is None:
        return tuple()

    normalized: list[str] = []
    seen: set[str] = set()
    for raw_field in optional_fields:
        field = raw_field.strip().lower()
        if not field:
            continue
        if field not in OPTIONAL_FIELD_SPECS:
            raise ValueError(
                f"Unsupported optional submission field '{raw_field}'. "
                f"Supported fields: {', '.join(sorted(OPTIONAL_FIELD_SPECS))}."
            )
        if field in seen:
            continue
        seen.add(field)
        normalized.append(field)

    return tuple(normalized)


def _resolve_assignment_records(
    client: CanvasClient,
    course_id: int,
    assignment_id: int | None,
) -> list[tuple[int, str]]:
    assignments = client.list_assignments(course_id)
    assignment_records = [
        (
            current_id,
            _assignment_name(assignment),
        )
        for assignment in assignments
        if (current_id := _coerce_positive_int(assignment.get("id"))) is not None
    ]

    if assignment_id is not None:
        for current_id, name in assignment_records:
            if current_id == assignment_id:
                return [(current_id, name)]
        return [(assignment_id, f"assignment-{assignment_id}")]

    return assignment_records


def _map_submission(
    submission: dict[str, Any],
    fallback_assignment_id: int,
    optional_fields: Sequence[str],
) -> dict[str, Any]:
    assignment_id = _coerce_positive_int(submission.get("assignment_id")) or fallback_assignment_id

    row = {
        "Assignment_ID": assignment_id,
        "Student_ID": _coerce_positive_int(submission.get("user_id")),
        "Attempt_Number": _coerce_non_negative_int(submission.get("attempt")),
        "Submitted_At": submission.get("submitted_at"),
        "Canvas_Late_Flag": bool(submission.get("late")),
    }

    for field in optional_fields:
        header, extractor = OPTIONAL_FIELD_SPECS[field]
        row[header] = extractor(submission)

    return row


def _resolve_csv_output_path(
    output_path: Path,
    assignment_id: int | None,
    current_assignment_id: int,
    assignment_name: str,
) -> Path:
    file_name = f"canvas_submissions_{current_assignment_id}_{_slugify_filename_part(assignment_name)}.csv"

    if assignment_id is None:
        if output_path.suffix.lower() == ".csv":
            raise ValueError("When --assignment is 'all', --output must be a directory, not a CSV file path.")
        output_path.mkdir(parents=True, exist_ok=True)
        return output_path / file_name

    if output_path.suffix.lower() == ".csv":
        output_path.parent.mkdir(parents=True, exist_ok=True)
        return output_path

    output_path.mkdir(parents=True, exist_ok=True)
    return output_path / file_name


def _write_csv(csv_path: Path, rows: list[dict[str, Any]], headers: Sequence[str]) -> None:
    with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=list(headers))
        writer.writeheader()
        writer.writerows(rows)


def _coerce_positive_int(value: object) -> int | None:
    if isinstance(value, int) and value > 0:
        return value
    if isinstance(value, str) and value.isdigit():
        parsed = int(value)
        return parsed if parsed > 0 else None
    return None


def _coerce_non_negative_int(value: object) -> int | None:
    if isinstance(value, int) and value >= 0:
        return value
    if isinstance(value, str) and value.isdigit():
        return int(value)
    return None


def _submission_user(submission: dict[str, Any]) -> dict[str, Any]:
    user = submission.get("user")
    return user if isinstance(user, dict) else {}


def _assignment_name(assignment: dict[str, Any]) -> str:
    name = assignment.get("name")
    if isinstance(name, str) and name.strip():
        return name.strip()
    return "assignment"


def _slugify_filename_part(value: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", value.strip()).strip("_")
    return slug.lower() or "assignment"


def _format_excel_datetime(value: object) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None

    normalized = value.strip().replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    return parsed.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _format_brisbane_datetime(value: object) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None

    normalized = value.strip().replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    brisbane = timezone(timedelta(hours=10))
    return parsed.astimezone(brisbane).strftime("%Y-%m-%d %H:%M:%S")
