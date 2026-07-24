"""Student roster export logic for Canvas courses."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from canvas_utils.api_client import CanvasClient

STUDENT_HEADERS = ["Student", "Email", "Student Number", "ID"]


def export_course_students(
    client: CanvasClient,
    course_id: int,
    output: str | Path = ".",
    *,
    active_only: bool = False,
) -> list[Path]:
    """Export a course student roster to XLSX."""
    student_records = client.list_course_students(course_id, active_only=active_only)
    rows = [_map_student_row(record) for record in student_records]

    xlsx_path = _resolve_xlsx_output_path(Path(output), course_id)
    _write_workbook(xlsx_path, rows)
    return [xlsx_path]


def _map_student_row(student: dict[str, Any]) -> dict[str, Any]:
    return {
        "Student": _string_or_none(student.get("name")),
        "Email": _string_or_none(student.get("email")) or _string_or_none(student.get("login_id")),
        "Student Number": _string_or_none(student.get("sis_user_id")),
        "ID": _coerce_positive_int(student.get("id")),
    }


def _resolve_xlsx_output_path(output_path: Path, course_id: int) -> Path:
    export_date = datetime.now().strftime("%Y-%m-%d")
    file_name = f"canvas_students_{course_id}_{export_date}.xlsx"

    if output_path.suffix.lower() == ".xlsx":
        output_path.parent.mkdir(parents=True, exist_ok=True)
        return output_path

    output_path.mkdir(parents=True, exist_ok=True)
    return output_path / file_name


def _write_workbook(xlsx_path: Path, rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Students"

    sheet.append(STUDENT_HEADERS)
    for row in rows:
        sheet.append([row.get(header) for header in STUDENT_HEADERS])

    _autosize_columns(sheet, STUDENT_HEADERS)
    workbook.save(xlsx_path)


def _autosize_columns(sheet: Any, headers: list[str]) -> None:
    for index, header in enumerate(headers, start=1):
        max_length = len(header)
        for row_index in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_index, column=index).value
            cell_length = len(str(cell_value)) if cell_value is not None else 0
            if cell_length > max_length:
                max_length = cell_length
        sheet.column_dimensions[get_column_letter(index)].width = max_length + 2


def _string_or_none(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _coerce_positive_int(value: object) -> int | None:
    if isinstance(value, int) and value > 0:
        return value
    if isinstance(value, str) and value.isdigit():
        parsed = int(value)
        return parsed if parsed > 0 else None
    return None
