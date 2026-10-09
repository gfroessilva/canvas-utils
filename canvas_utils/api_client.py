"""Canvas API client abstractions with auth, error handling, and pagination."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from http.client import HTTPException
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urljoin, urlparse
from urllib.request import Request, urlopen


class CanvasAPIError(RuntimeError):
    """Raised when Canvas API requests fail or payloads are invalid."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code


class CanvasTransportError(CanvasAPIError):
    """Raised when a request transport failure may leave a mutation indeterminate."""

    def __init__(self, message: str, *, request_may_have_succeeded: bool) -> None:
        super().__init__(message)
        self.request_may_have_succeeded = request_may_have_succeeded


@dataclass(frozen=True)
class _ResponseData:
    """Represents a JSON HTTP response payload and headers."""

    payload: Any
    headers: dict[str, str]


class CanvasClient:
    """HTTP client for Canvas API requests.

    Args:
        base_url: Base Canvas instance URL, e.g., https://canvas.example.edu.
        token: Canvas API token used in the Authorization Bearer header.
        timeout: Per-request timeout in seconds.
    """

    def __init__(self, base_url: str, token: str, timeout: int = 30) -> None:
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.timeout = timeout

    def list_assignments(self, course_id: int) -> list[dict[str, Any]]:
        """Return all assignments for a course."""
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/assignments",
            params={"per_page": 100},
        )
        return self._expect_list(payload, "assignments")

    def get_assignment(
        self,
        course_id: int,
        assignment_id: int,
        *,
        include_rubric: bool = False,
    ) -> dict[str, Any]:
        """Return a specific assignment, optionally including its rubric."""
        params: dict[str, Any] | None = None
        if include_rubric:
            params = {"include[]": ["rubric"]}

        response = self.request_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}",
            method="GET",
            params=params,
        )
        return self._expect_dict(response.payload, "assignment")

    def list_submissions(
        self,
        course_id: int,
        assignment_id: int,
        include_user: bool = True,
        *,
        include_rubric_assessment: bool = False,
    ) -> list[dict[str, Any]]:
        """Return all submissions for a specific assignment."""
        params: dict[str, Any] = {"per_page": 100}
        includes: list[str] = []
        if include_user:
            includes.append("user")
        if include_rubric_assessment:
            includes.append("rubric_assessment")
        if includes:
            params["include[]"] = includes

        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions",
            params=params,
        )
        return self._expect_list(payload, "submissions")

    def grade_submission(
        self,
        course_id: int,
        assignment_id: int,
        user_id: int,
        *,
        posted_grade: str | int | float | None = None,
        submission_comment: str | None = None,
        submission_comment_attempt: int | None = None,
        rubric_assessment: dict[str, dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Grade or comment on one submission, optionally with rubric criteria."""
        form_data: dict[str, Any] = {}
        if posted_grade is not None:
            form_data["submission[posted_grade]"] = posted_grade
        if submission_comment is not None:
            form_data["comment[text_comment]"] = submission_comment
        if submission_comment_attempt is not None:
            if submission_comment is None:
                raise ValueError("A submission comment is required when setting a comment attempt.")
            form_data["comment[attempt]"] = submission_comment_attempt
        if rubric_assessment is not None:
            for criterion_id, assessment in rubric_assessment.items():
                if not isinstance(assessment, dict):
                    raise ValueError(
                        f"Rubric assessment for criterion '{criterion_id}' must be a mapping."
                    )
                for field in ("rating_id", "points", "comments"):
                    value = assessment.get(field)
                    if value is not None:
                        form_data[f"rubric_assessment[{criterion_id}][{field}]"] = value

        if not form_data:
            raise ValueError("A submission grade, comment, or rubric assessment is required.")

        response = self.request_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions/{user_id}",
            method="PUT",
            form_data=form_data,
        )
        if not isinstance(response.payload, dict):
            raise CanvasTransportError(
                "Canvas returned an unexpected response after updating a submission. "
                "The update may have been applied.",
                request_may_have_succeeded=True,
            )
        return response.payload

    def get_submission(
        self,
        course_id: int,
        assignment_id: int,
        user_id: int,
        *,
        include_rubric_assessment: bool = False,
    ) -> dict[str, Any]:
        """Return one student's submission, optionally including its rubric assessment."""
        params: dict[str, Any] | None = None
        if include_rubric_assessment:
            params = {"include[]": ["rubric_assessment"]}
        response = self.request_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions/{user_id}",
            method="GET",
            params=params,
        )
        return self._expect_dict(response.payload, "submission")

    def list_course_students(self, course_id: int, *, active_only: bool = False) -> list[dict[str, Any]]:
        """Return course users with student enrollments."""
        params: dict[str, Any] = {
            "per_page": 100,
            "enrollment_type[]": ["student"],
        }
        if active_only:
            params["enrollment_state[]"] = ["active"]

        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/users",
            params=params,
        )
        return self._expect_list(payload, "course students")

    def list_course_groups(self, course_id: int) -> list[dict[str, Any]]:
        """Return course groups with their member users."""
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/groups",
            params={
                "per_page": 100,
                "include[]": ["users"],
            },
        )
        return self._expect_list(payload, "course groups")

    def list_pages(self, course_id: int) -> list[dict[str, Any]]:
        """Return all course pages."""
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/pages",
            params={"per_page": 100},
        )
        return self._expect_list(payload, "pages")

    def delete_page(self, course_id: int, page_url: str) -> None:
        """Delete a specific course page by URL slug."""
        self.request_json(
            f"/api/v1/courses/{course_id}/pages/{quote(page_url, safe='')}",
            method="DELETE",
        )

    def delete_assignment(self, course_id: int, assignment_id: int) -> None:
        """Delete a specific assignment."""
        self.request_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}",
            method="DELETE",
        )

    def list_modules(self, course_id: int) -> list[dict[str, Any]]:
        """Return all modules in a course."""
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/modules",
            params={"per_page": 100},
        )
        return self._expect_list(payload, "modules")

    def delete_module(self, course_id: int, module_id: int) -> None:
        """Delete a specific module."""
        self.request_json(
            f"/api/v1/courses/{course_id}/modules/{module_id}",
            method="DELETE",
        )

    def list_files(self, course_id: int) -> list[dict[str, Any]]:
        """Return all files in a course."""
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/files",
            params={"per_page": 100},
        )
        return self._expect_list(payload, "files")

    def delete_file(self, file_id: int) -> None:
        """Delete a specific file."""
        self.request_json(
            f"/api/v1/files/{file_id}",
            method="DELETE",
        )

    def list_discussion_topics(
        self,
        course_id: int,
        *,
        only_announcements: bool = False,
    ) -> list[dict[str, Any]]:
        """Return discussion topics for a course."""
        params: dict[str, Any] = {"per_page": 100}
        if only_announcements:
            params["only_announcements"] = True
        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/discussion_topics",
            params=params,
        )
        return self._expect_list(payload, "discussion topics")

    def delete_discussion_topic(self, course_id: int, topic_id: int) -> None:
        """Delete a discussion topic or announcement."""
        self.request_json(
            f"/api/v1/courses/{course_id}/discussion_topics/{topic_id}",
            method="DELETE",
        )

    def get_course(
        self,
        course_id: int,
        *,
        include_syllabus_body: bool = False,
    ) -> dict[str, Any]:
        """Return course details."""
        params: dict[str, Any] | None = None
        if include_syllabus_body:
            params = {"include[]": ["syllabus_body"]}
        response = self.request_json(
            f"/api/v1/courses/{course_id}",
            method="GET",
            params=params,
        )
        return self._expect_dict(response.payload, "course")

    def update_course(self, course_id: int, updates: dict[str, Any]) -> dict[str, Any]:
        """Update course attributes."""
        form_data = {f"course[{key}]": value for key, value in updates.items()}
        response = self.request_json(
            f"/api/v1/courses/{course_id}",
            method="PUT",
            form_data=form_data,
        )
        return self._expect_dict(response.payload, "course update")

    def list_calendar_events(self, course_id: int) -> list[dict[str, Any]]:
        """Return calendar events scoped to a course."""
        payload = self.get_paginated_json(
            "/api/v1/calendar_events",
            params={
                "context_codes[]": [f"course_{course_id}"],
                "all_events": True,
                "per_page": 100,
            },
        )
        return self._expect_list(payload, "calendar events")

    def get_user(self, user_id: int) -> dict[str, Any]:
        """Return a Canvas user record by ID."""
        response = self.request_json(
            f"/api/v1/users/{user_id}",
            method="GET",
        )
        return self._expect_dict(response.payload, "user")

    def get_course_user(self, course_id: int, user_id: int) -> dict[str, Any]:
        """Return a Canvas user record within a specific course context."""
        response = self.request_json(
            f"/api/v1/courses/{course_id}/users/{user_id}",
            method="GET",
        )
        return self._expect_dict(response.payload, "course user")

    def delete_calendar_event(self, event_id: int) -> None:
        """Delete a calendar event."""
        self.request_json(
            f"/api/v1/calendar_events/{event_id}",
            method="DELETE",
        )

    def get_paginated_json(
        self,
        path_or_url: str,
        params: dict[str, Any] | None = None,
    ) -> list[Any]:
        """Collect and return all items across paginated Canvas responses."""
        all_items: list[Any] = []
        next_url: str | None = path_or_url
        next_params = params

        while next_url is not None:
            response = self.request_json(next_url, method="GET", params=next_params)
            page_items = self._expect_list(response.payload, "paginated response")
            all_items.extend(page_items)
            next_url = _extract_next_link(response.headers.get("Link"))
            next_params = None

        return all_items

    def request_json(
        self,
        path_or_url: str,
        method: str = "GET",
        params: dict[str, Any] | None = None,
        form_data: dict[str, Any] | None = None,
    ) -> _ResponseData:
        """Perform an authenticated Canvas request and parse its JSON response."""
        url = self._build_url(path_or_url, params)
        encoded_form = urlencode(form_data, doseq=True).encode("utf-8") if form_data else None

        headers = {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/json",
        }
        if encoded_form is not None:
            headers["Content-Type"] = "application/x-www-form-urlencoded"

        request_method = method.upper()
        request = Request(url=url, method=request_method, data=encoded_form, headers=headers)

        try:
            with urlopen(request, timeout=self.timeout) as response:
                body = response.read().decode("utf-8")
                payload = json.loads(body) if body else None
                return _ResponseData(payload=payload, headers=dict(response.headers.items()))
        except HTTPError as exc:
            response_body = exc.read().decode("utf-8", errors="replace")
            raise CanvasAPIError(
                _format_http_error(url, exc.code, exc.reason, response_body),
                status_code=exc.code,
            ) from exc
        except URLError as exc:
            raise CanvasTransportError(
                f"Network error while contacting Canvas at {url}: {exc.reason}",
                request_may_have_succeeded=_request_may_have_succeeded(request_method),
            ) from exc
        except (TimeoutError, ConnectionError, OSError, HTTPException) as exc:
            raise CanvasTransportError(
                f"Network error while contacting Canvas at {url}: {exc}",
                request_may_have_succeeded=_request_may_have_succeeded(request_method),
            ) from exc
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise CanvasTransportError(
                f"Canvas returned an invalid response for {url}: {exc}",
                request_may_have_succeeded=_request_may_have_succeeded(request_method),
            ) from exc

    def _build_url(self, path_or_url: str, params: dict[str, Any] | None = None) -> str:
        if path_or_url.startswith("http://") or path_or_url.startswith("https://"):
            url = path_or_url
        else:
            url = urljoin(f"{self.base_url}/", path_or_url.lstrip("/"))

        if not params:
            return url

        query = urlencode(params, doseq=True)
        separator = "&" if "?" in url else "?"
        return f"{url}{separator}{query}"

    @staticmethod
    def _expect_list(payload: Any, label: str) -> list[Any]:
        if not isinstance(payload, list):
            raise CanvasAPIError(f"Expected a list payload for {label}, but received {type(payload).__name__}.")
        return payload

    @staticmethod
    def _expect_dict(payload: Any, label: str) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise CanvasAPIError(f"Expected an object payload for {label}, but received {type(payload).__name__}.")
        return payload


def _extract_next_link(link_header: str | None) -> str | None:
    """Return the rel=next URL from a Canvas Link header, if present."""
    if not link_header:
        return None

    for match in re.finditer(r'<([^>]+)>;\s*rel="([^"]+)"', link_header):
        if match.group(2) == "next":
            return match.group(1)
    return None


def _request_may_have_succeeded(method: str) -> bool:
    """Return whether a failed request could have changed Canvas state."""
    return method not in {"GET", "HEAD", "OPTIONS"}


def _format_http_error(url: str, status_code: int, reason: str, response_body: str) -> str:
    """Return a user-friendly message for common Canvas HTTP failures."""
    snippet = response_body.strip()
    if snippet:
        snippet = f" Response: {snippet}"

    if status_code == 401:
        return (
            "Canvas API authentication failed (401 Unauthorized). "
            "Check CANVAS_API_TOKEN or your configured token and ensure it has the required permissions."
            f" URL: {url}.{snippet}"
        )

    if status_code == 404:
        return f"{_describe_not_found_url(url)} URL: {url}.{snippet}"

    return f"Canvas API request failed ({status_code} {reason}) for {url}.{snippet}"


def _describe_not_found_url(url: str) -> str:
    """Return an endpoint-specific explanation for Canvas 404 errors."""
    path = urlparse(url).path

    if re.search(r"/api/v1/users/\d+$", path):
        return (
            "Canvas user not found (404 Not Found). "
            "The referenced user may be deleted, hidden from your API token, or only valid in a different Canvas context."
        )

    if re.search(r"/api/v1/courses/\d+/assignments/\d+/submissions(?:/\d+)?$", path):
        return (
            "Canvas submissions resource not found (404 Not Found). "
            "Verify the base URL, course ID, and assignment ID, and confirm your token can access submissions for that assignment."
        )

    if re.search(r"/api/v1/courses/\d+/assignments(?:/\d+)?$", path):
        return (
            "Canvas assignment resource not found (404 Not Found). "
            "Verify the base URL, course ID, and assignment ID, and confirm the assignment exists in that course."
        )

    if re.search(r"/api/v1/courses/\d+$", path):
        return (
            "Canvas course not found (404 Not Found). "
            "Verify the base URL and course ID, and confirm your token can access that course."
        )

    return (
        "Canvas resource not found (404 Not Found). "
        "Verify the base URL and any referenced Canvas IDs, and confirm your token has access to that resource."
    )
