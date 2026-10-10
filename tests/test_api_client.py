from __future__ import annotations

import unittest
from unittest.mock import patch

from canvas_utils.api_client import (
    CanvasClient,
    CanvasTransportError,
    _ResponseData,
    _format_http_error,
)


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


class SingleResponseCanvasClient(CanvasClient):
    def __init__(self) -> None:
        super().__init__(base_url="https://canvas.example.edu", token="token")
        self.last_path_or_url: str | None = None
        self.last_method: str | None = None
        self.last_params: dict[str, object] | None = None

    def request_json(
        self,
        path_or_url: str,
        method: str = "GET",
        params: dict[str, object] | None = None,
        form_data: dict[str, object] | None = None,
    ) -> _ResponseData:
        self.last_path_or_url = path_or_url
        self.last_method = method
        self.last_params = params
        return _ResponseData(payload={"id": 654, "rubric": []}, headers={})


class SubmissionUpdateCanvasClient(CanvasClient):
    def __init__(self) -> None:
        super().__init__(base_url="https://canvas.example.edu", token="token")
        self.last_path_or_url: str | None = None
        self.last_method: str | None = None
        self.last_params: dict[str, object] | None = None
        self.last_form_data: dict[str, object] | None = None

    def request_json(
        self,
        path_or_url: str,
        method: str = "GET",
        params: dict[str, object] | None = None,
        form_data: dict[str, object] | None = None,
    ) -> _ResponseData:
        self.last_path_or_url = path_or_url
        self.last_method = method
        self.last_params = params
        self.last_form_data = form_data
        return _ResponseData(payload={"id": 7}, headers={})


class GroupRecordingCanvasClient(CanvasClient):
    def __init__(self) -> None:
        super().__init__(base_url="https://canvas.example.edu", token="token")
        self.calls: list[tuple[str, dict[str, object] | None]] = []

    def get_paginated_json(
        self,
        path_or_url: str,
        params: dict[str, object] | None = None,
    ) -> list[object]:
        self.calls.append((path_or_url, params))
        if path_or_url == "/api/v1/courses/321/groups":
            return [{"id": 44, "users": [{"id": 1}]}]
        return [{"id": user_id} for user_id in range(1, 102)]


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

    def test_mutating_transport_failure_is_marked_indeterminate(self) -> None:
        client = CanvasClient(base_url="https://canvas.example.edu", token="token")

        with patch("canvas_utils.api_client.urlopen", side_effect=TimeoutError("timed out")):
            with self.assertRaises(CanvasTransportError) as raised:
                client.request_json("/api/v1/example", method="PUT", form_data={"value": "test"})

        self.assertTrue(raised.exception.request_may_have_succeeded)


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

    def test_list_course_groups_fetches_users_per_group(self) -> None:
        client = GroupRecordingCanvasClient()

        groups = client.list_course_groups(course_id=321)

        self.assertEqual(len(groups[0]["users"]), 101)
        self.assertEqual(
            client.calls,
            [
                ("/api/v1/courses/321/groups", {"per_page": 100}),
                ("/api/v1/groups/44/users", {"per_page": 100}),
            ],
        )

    def test_list_submissions_retains_user_and_adds_rubric_assessment(self) -> None:
        client = RecordingCanvasClient()

        submissions = client.list_submissions(
            course_id=321,
            assignment_id=654,
            include_user=True,
            include_rubric_assessment=True,
        )

        self.assertEqual(submissions, [])
        self.assertEqual(
            client.last_path_or_url,
            "/api/v1/courses/321/assignments/654/submissions",
        )
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "include[]": ["user", "rubric_assessment"],
            },
        )

    def test_list_assignments_adds_rubric_include(self) -> None:
        client = RecordingCanvasClient()

        assignments = client.list_assignments(course_id=321, include_rubric=True)

        self.assertEqual(assignments, [])
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/assignments")
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "include[]": ["rubric"],
            },
        )


class CanvasApiClientGradesTests(unittest.TestCase):
    def test_list_submissions_can_include_rubric_assessments(self) -> None:
        client = RecordingCanvasClient()

        submissions = client.list_submissions(
            course_id=321,
            assignment_id=654,
            include_rubric_assessment=True,
        )

        self.assertEqual(submissions, [])
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/assignments/654/submissions")
        self.assertEqual(
            client.last_params,
            {
                "per_page": 100,
                "include[]": ["user", "rubric_assessment"],
            },
        )

    def test_get_assignment_can_include_rubric(self) -> None:
        client = SingleResponseCanvasClient()

        assignment = client.get_assignment(course_id=321, assignment_id=654, include_rubric=True)

        self.assertEqual(assignment, {"id": 654, "rubric": []})
        self.assertEqual(client.last_path_or_url, "/api/v1/courses/321/assignments/654")
        self.assertEqual(client.last_method, "GET")
        self.assertEqual(client.last_params, {"include[]": ["rubric"]})

    def test_grade_submission_sends_grade_comment_and_rubric_form_fields(self) -> None:
        client = SubmissionUpdateCanvasClient()

        submission = client.grade_submission(
            course_id=321,
            assignment_id=654,
            user_id=7,
            posted_grade=85,
            submission_comment="Well done.",
            submission_comment_attempt=3,
            rubric_assessment={
                "_criterion-1": {
                    "rating_id": "rating-strong",
                    "points": 5,
                    "comments": "Well supported.",
                }
            },
        )

        self.assertEqual(submission, {"id": 7})
        self.assertEqual(
            client.last_path_or_url,
            "/api/v1/courses/321/assignments/654/submissions/7",
        )
        self.assertEqual(client.last_method, "PUT")
        self.assertEqual(
            client.last_form_data,
            {
                "submission[posted_grade]": 85,
                "comment[text_comment]": "Well done.",
                "comment[attempt]": 3,
                "rubric_assessment[_criterion-1][rating_id]": "rating-strong",
                "rubric_assessment[_criterion-1][points]": 5,
                "rubric_assessment[_criterion-1][comments]": "Well supported.",
            },
        )

    def test_grade_submission_rejects_empty_mutations(self) -> None:
        client = SubmissionUpdateCanvasClient()

        with self.assertRaisesRegex(ValueError, "grade, comment, or rubric"):
            client.grade_submission(course_id=321, assignment_id=654, user_id=7)

    def test_get_submission_can_include_rubric_assessment(self) -> None:
        client = SubmissionUpdateCanvasClient()

        submission = client.get_submission(
            course_id=321,
            assignment_id=654,
            user_id=7,
            include_rubric_assessment=True,
        )

        self.assertEqual(submission, {"id": 7})
        self.assertEqual(
            client.last_path_or_url,
            "/api/v1/courses/321/assignments/654/submissions/7",
        )
        self.assertEqual(client.last_method, "GET")
        self.assertEqual(client.last_params, {"include[]": ["rubric_assessment"]})
        self.assertIsNone(client.last_form_data)


if __name__ == "__main__":
    unittest.main()
