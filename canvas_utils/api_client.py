"""Canvas API client abstractions with auth, error handling, and pagination."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urljoin
from urllib.request import Request, urlopen


class CanvasAPIError(RuntimeError):
    """Raised when Canvas API requests fail or payloads are invalid."""


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

    def list_submissions(
        self,
        course_id: int,
        assignment_id: int,
        include_user: bool = True,
    ) -> list[dict[str, Any]]:
        """Return all submissions for a specific assignment."""
        params: dict[str, Any] = {"per_page": 100}
        if include_user:
            params["include[]"] = ["user"]

        payload = self.get_paginated_json(
            f"/api/v1/courses/{course_id}/assignments/{assignment_id}/submissions",
            params=params,
        )
        return self._expect_list(payload, "submissions")

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

        request = Request(url=url, method=method.upper(), data=encoded_form, headers=headers)

        try:
            with urlopen(request, timeout=self.timeout) as response:
                body = response.read().decode("utf-8")
                payload = json.loads(body) if body else None
                return _ResponseData(payload=payload, headers=dict(response.headers.items()))
        except HTTPError as exc:
            response_body = exc.read().decode("utf-8", errors="replace")
            raise CanvasAPIError(_format_http_error(url, exc.code, exc.reason, response_body)) from exc
        except URLError as exc:
            raise CanvasAPIError(f"Network error while contacting Canvas at {url}: {exc.reason}") from exc
        except json.JSONDecodeError as exc:
            raise CanvasAPIError(f"Canvas returned invalid JSON for {url}: {exc}") from exc

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
        return (
            "Canvas resource not found (404 Not Found). "
            "Verify the base URL, course ID, and assignment ID are correct."
            f" URL: {url}.{snippet}"
        )

    return f"Canvas API request failed ({status_code} {reason}) for {url}.{snippet}"
