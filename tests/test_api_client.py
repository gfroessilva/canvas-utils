from __future__ import annotations

import unittest

from canvas_utils.api_client import _format_http_error


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


if __name__ == "__main__":
    unittest.main()
