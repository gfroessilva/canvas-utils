from __future__ import annotations

import unittest

from canvas_utils.cleanup import ALL_CLEANUP_TARGETS, cleanup_course_resources, resolve_cleanup_targets
from run import build_parser, parse_cleanup_targets


class FakeCleanupClient:
    def __init__(self) -> None:
        self.deleted_pages: list[str] = []
        self.deleted_assignments: list[int] = []
        self.deleted_modules: list[int] = []
        self.deleted_files: list[int] = []
        self.deleted_topics: list[int] = []
        self.deleted_events: list[int] = []
        self.updated_courses: list[tuple[int, dict[str, object]]] = []

    def list_pages(self, course_id: int) -> list[dict[str, object]]:
        return [{"url": "page-1", "title": "Page 1"}]

    def delete_page(self, course_id: int, page_url: str) -> None:
        self.deleted_pages.append(page_url)

    def list_assignments(self, course_id: int) -> list[dict[str, object]]:
        return [{"id": 101, "name": "Assignment 1"}]

    def delete_assignment(self, course_id: int, assignment_id: int) -> None:
        self.deleted_assignments.append(assignment_id)

    def list_modules(self, course_id: int) -> list[dict[str, object]]:
        return [{"id": 201, "name": "Module 1"}]

    def delete_module(self, course_id: int, module_id: int) -> None:
        self.deleted_modules.append(module_id)

    def list_files(self, course_id: int) -> list[dict[str, object]]:
        return [{"id": 301, "display_name": "File 1"}]

    def delete_file(self, file_id: int) -> None:
        self.deleted_files.append(file_id)

    def list_discussion_topics(
        self,
        course_id: int,
        *,
        only_announcements: bool = False,
    ) -> list[dict[str, object]]:
        all_topics = [
            {"id": 401, "title": "Discussion 1", "announcement": False},
            {"id": 402, "title": "Announcement 1", "announcement": True},
        ]
        if only_announcements:
            return [topic for topic in all_topics if bool(topic.get("announcement"))]
        return all_topics

    def delete_discussion_topic(self, course_id: int, topic_id: int) -> None:
        self.deleted_topics.append(topic_id)

    def get_course(
        self,
        course_id: int,
        *,
        include_syllabus_body: bool = False,
    ) -> dict[str, object]:
        return {"id": course_id, "syllabus_body": "<p>Syllabus content</p>"}

    def update_course(self, course_id: int, updates: dict[str, object]) -> dict[str, object]:
        self.updated_courses.append((course_id, updates))
        return {"id": course_id, **updates}

    def list_calendar_events(self, course_id: int) -> list[dict[str, object]]:
        return [
            {"id": 501, "title": "Standalone event"},
            {"id": 502, "title": "Assignment date", "assignment_id": 101},
        ]

    def delete_calendar_event(self, event_id: int) -> None:
        self.deleted_events.append(event_id)


class CleanupTests(unittest.TestCase):
    def test_resolve_cleanup_targets_supports_all(self) -> None:
        self.assertEqual(resolve_cleanup_targets(["all"]), ALL_CLEANUP_TARGETS)

    def test_parse_cleanup_targets_from_csv(self) -> None:
        self.assertEqual(parse_cleanup_targets("pages,modules"), ("pages", "modules"))

    def test_parse_cleanup_targets_rejects_invalid_target(self) -> None:
        with self.assertRaises(ValueError):
            parse_cleanup_targets("pages,unknown")

    def test_cleanup_dry_run_does_not_delete(self) -> None:
        client = FakeCleanupClient()
        results = cleanup_course_resources(
            client=client,
            course_id=123,
            targets=("pages", "assessments"),
            dry_run=True,
        )

        self.assertEqual(len(results), 2)
        self.assertEqual(client.deleted_pages, [])
        self.assertEqual(client.deleted_assignments, [])
        self.assertEqual(results[0].found, 1)
        self.assertEqual(results[0].deleted, 0)

    def test_cleanup_execute_deletes_only_requested_targets(self) -> None:
        client = FakeCleanupClient()
        results = cleanup_course_resources(
            client=client,
            course_id=123,
            targets=("discussions",),
            dry_run=False,
        )

        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].target, "discussions")
        self.assertEqual(results[0].deleted, 1)
        self.assertEqual(client.deleted_topics, [401])

    def test_cleanup_execute_supports_syllabus_and_calendar_targets(self) -> None:
        client = FakeCleanupClient()
        results = cleanup_course_resources(
            client=client,
            course_id=123,
            targets=("syllabus", "calendar"),
            dry_run=False,
        )

        self.assertEqual(len(results), 2)
        self.assertEqual(results[0].target, "syllabus")
        self.assertEqual(results[0].deleted, 1)
        self.assertEqual(client.updated_courses, [(123, {"syllabus_body": ""})])
        self.assertEqual(results[1].target, "calendar")
        self.assertEqual(results[1].found, 1)
        self.assertEqual(results[1].deleted, 1)
        self.assertEqual(client.deleted_events, [501])

    def test_cli_exposes_cleanup_subcommand(self) -> None:
        parser = build_parser()
        args = parser.parse_args(["cleanup", "--course", "123", "--targets", "all", "--execute"])
        self.assertEqual(args.command, "cleanup")
        self.assertEqual(args.course, 123)
        self.assertEqual(args.targets, "all")
        self.assertTrue(args.execute)


if __name__ == "__main__":
    unittest.main()
