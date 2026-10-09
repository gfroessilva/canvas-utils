from __future__ import annotations

import unittest

from canvas_utils.api_client import CanvasClient, _format_http_error


class RecordingCanvasClient(CanvasClient):
    def __init__(self) -> None:
        super().__init__(base_url="https://canvas.example.edu", token="token")
        self.last_path_or_url: str | None = None
        self.last_params: dict[str, object] | None = None

    def get_paginated_json(
        self,
        path_or_url: str,
        params: dict[str, object] | None = None,
    ) -> list[object]:
        self.last_path_or_url = path_or_url
        self.last_params = params
        return []


class CanvasApiClientErrorFormattingTests(unittest.TestCase):
    def test_404_for_user_lookup_mentions_hidden_or_deleted_user(self) -> None:
        message = _format_http_error(
            "https://canvas.example.edu/api/v1/users/74955",
            404,
            "Not Found",
            '{"errors":[{"message":"The specified resource does not exist."}]}',
        )

        self.assertIn("Canvas user not found (404 Not Found).", message)
        self.assertIn("deleted, hidden from your API token", message)

    def test_404_for_submission_lookup_mentions_course_and_assignment_access(self) -> None:
        message = _format_http_error(
            "https://canvas.example.edu/api/v1/courses/24224/assignments/214358/submissions",
            404,
            "Not Found",
            "",
        )

        self.assertIn("Canvas submissions resource not found (404 Not Found).", message)
        self.assertIn("confirm your token can access submissions", message)

    def test_generic_404_still_mentions_access_and_ids(self) -> None:
        message = _format_http_error(
            "https://canvas.example.edu/api/v1/somewhere/else",
            404,
            "Not Found",
            "",
        )

        self.assertIn("Canvas resource not found (404 Not Found).", message)
        self.assertIn("referenced Canvas IDs", message)


class CanvasApiClientStudentsTests(unittest.TestCase):
    def test_list_course_students_uses_student_enrollment_filter(self) -> None:
        client = RecordingCanvasClient()

        students = client.list_course_students(course_id=321)

        self.assertEqual(students, [])
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/users")
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "enrollment_type[]": ["student"],
            },
        )

    def test_list_course_students_active_only_adds_enrollment_state_filter(self) -> None:
        client = RecordingCanvasClient()

        students = client.list_course_students(course_id=321, active_only=True)

        self.assertEqual(students, [])
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/users")
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "enrollment_type[]": ["student"],
                "enrollment_state[]": ["active"],
            },
        )

    def test_list_course_groups_includes_member_users(self) -> None:
        client = RecordingCanvasClient()

        groups = client.list_course_groups(course_id=321)

        self.assertEqual(groups, [])
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/groups")
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "include[]": ["users"],
            },
        )


if __name__ == "__main__":
    unittest.main()
