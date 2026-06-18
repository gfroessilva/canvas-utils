from __future__ import annotations

import csv
import io
import re
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path

from canvas_utils.submissions import build_csv_headers, extract_submission_timestamps, resolve_optional_fields
from run import build_parser, parse_optional_fields


class FakeCanvasClient:
    def list_assignments(self, course_id: int) -> list[dict[str, object]]:
        return [{"id": 101, "name": "Assessment 1: Intro / Basics"}]

    def list_submissions(
        self,
        course_id: int,
        assignment_id: int,
        include_user: bool = True,
    ) -> list[dict[str, object]]:
        return [
            {
                "assignment_id": assignment_id,
                "user_id": 7,
                "attempt": 2,
                "submitted_at": "2026-05-24T08:00:00Z",
                "late": True,
                "score": 88.5,
                "workflow_state": "graded",
                "user": {"name": "Jane Doe", "login_id": "jdoe", "sis_user_id": "S123"},
            }
        ]


class SubmissionExtractionTests(unittest.TestCase):
    def test_optional_fields_are_normalized_and_headers_are_ordered(self) -> None:
        optional_fields = resolve_optional_fields(["score", "user_login_id", "score"])

        self.assertEqual(optional_fields, ("score", "user_login_id"))
        self.assertEqual(
            build_csv_headers(optional_fields),
            [
                "Assignment_ID",
                "Student_ID",
                "Attempt_Number",
                "Submitted_At",
                "Canvas_Late_Flag",
                "Score",
                "Student_Login_ID",
            ],
        )

    def test_extract_submission_timestamps_writes_student_name_and_optional_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            written_files = extract_submission_timestamps(
                client=FakeCanvasClient(),
                course_id=123,
                assignment_id=101,
                output=Path(temp_dir),
                optional_fields=("score", "workflow_state", "user_sis_id"),
            )

            self.assertEqual(len(written_files), 1)
            csv_path = written_files[0]
            self.assertTrue(csv_path.exists())
            self.assertEqual(
                bool(re.fullmatch(r"canvas_submissions_101_assessment_1_intro_basics_\d{4}-\d{2}-\d{2}\.csv", csv_path.name)),
                True,
            )

            with csv_path.open("r", newline="", encoding="utf-8") as handle:
                reader = csv.DictReader(handle)
                rows = list(reader)

        self.assertEqual(
            reader.fieldnames,
            [
                "Assignment_ID",
                "Student_ID",
                "Attempt_Number",
                "Submitted_At",
                "Canvas_Late_Flag",
                "Score",
                "Workflow_State",
                "Student_SIS_ID",
            ],
        )
        self.assertEqual(
            rows[0],
            {
                "Assignment_ID": "101",
                "Student_ID": "7",
                "Attempt_Number": "2",
                "Submitted_At": "2026-05-24T08:00:00Z",
                "Canvas_Late_Flag": "True",
                "Score": "88.5",
                "Workflow_State": "graded",
                "Student_SIS_ID": "S123",
            },
        )

    def test_student_name_and_timestamp_variants_are_optional_fields(self) -> None:
        headers = build_csv_headers(("student_name", "submitted_at_excel", "submitted_at_brisbane"))
        self.assertEqual(
            headers,
            [
                "Assignment_ID",
                "Student_ID",
                "Attempt_Number",
                "Submitted_At",
                "Canvas_Late_Flag",
                "Student_Name",
                "Submitted_At_Excel",
                "Submitted_At_Brisbane",
            ],
        )

    def test_parse_optional_fields_accepts_comma_separated_values(self) -> None:
        self.assertEqual(parse_optional_fields("score, user_login_id"), ("score", "user_login_id"))

    def test_cli_uses_submissions_subcommand(self) -> None:
        parser = build_parser()
        args = parser.parse_args(["submissions", "--course", "123"])
        self.assertEqual(args.command, "submissions")
        self.assertEqual(args.course, 123)
        self.assertEqual(args.assignment, "all")

    def test_cli_requires_subcommand(self) -> None:
        parser = build_parser()
        with redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                parser.parse_args([])


if __name__ == "__main__":
    unittest.main()
