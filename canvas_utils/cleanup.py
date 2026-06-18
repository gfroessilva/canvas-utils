"""Progressive course cleanup (deletion) workflows."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Sequence

from canvas_utils.api_client import CanvasAPIError, CanvasClient

ALL_CLEANUP_TARGETS = (
    "pages",
    "assessments",
    "modules",
    "files",
    "discussions",
    "announcements",
    "syllabus",
    "calendar",
)

ProgressCallback = Callable[[str], None]


@dataclass(frozen=True)
class CleanupResult:
    """Summary of cleanup execution for a specific target group."""

    target: str
    found: int
    deleted: int
    failed: int


def resolve_cleanup_targets(raw_targets: Sequence[str] | None) -> tuple[str, ...]:
    """Validate and normalize target groups for cleanup execution."""
    if raw_targets is None:
        return ALL_CLEANUP_TARGETS

    normalized: list[str] = []
    seen: set[str] = set()

    for raw_target in raw_targets:
        target = raw_target.strip().lower()
        if not target:
            continue
        if target == "all":
            return ALL_CLEANUP_TARGETS
        if target not in ALL_CLEANUP_TARGETS:
            raise ValueError(
                f"Unsupported cleanup target '{raw_target}'. "
                f"Supported targets: {', '.join(ALL_CLEANUP_TARGETS)}, all."
            )
        if target in seen:
            continue
        seen.add(target)
        normalized.append(target)

    if not normalized:
        raise ValueError("At least one cleanup target is required.")

    return tuple(normalized)


def cleanup_course_resources(
    client: CanvasClient,
    course_id: int,
    targets: Sequence[str],
    *,
    dry_run: bool = True,
    progress_callback: ProgressCallback | None = None,
) -> list[CleanupResult]:
    """Delete selected Canvas course resources progressively.

    Args:
        client: Authenticated Canvas API client.
        course_id: Canvas course ID.
        targets: One or more target groups to delete.
        dry_run: If True, report what would be deleted without deleting.
        progress_callback: Optional callback for progressive status logs.

    Returns:
        A list of per-target cleanup summaries.
    """
    resolved_targets = resolve_cleanup_targets(targets)
    progress = progress_callback or (lambda _: None)

    handlers = {
        "pages": _cleanup_pages,
        "assessments": _cleanup_assessments,
        "modules": _cleanup_modules,
        "files": _cleanup_files,
        "discussions": _cleanup_discussions,
        "announcements": _cleanup_announcements,
        "syllabus": _cleanup_syllabus,
        "calendar": _cleanup_calendar,
    }

    results: list[CleanupResult] = []
    for target in resolved_targets:
        results.append(handlers[target](client, course_id, dry_run=dry_run, progress=progress))

    return results


def _cleanup_pages(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    pages = client.list_pages(course_id)
    found = len(pages)
    if dry_run:
        progress(f"[pages] Would delete {found} page(s).")
        return CleanupResult(target="pages", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, page in enumerate(pages, start=1):
        page_url = page.get("url")
        page_label = str(page.get("title") or page_url or "<unknown>")
        if not isinstance(page_url, str) or not page_url:
            failed += 1
            progress(f"[pages] Failed {index}/{found}: missing page URL for {page_label}.")
            continue
        try:
            client.delete_page(course_id, page_url)
            deleted += 1
            progress(f"[pages] Deleted {index}/{found}: {page_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[pages] Failed {index}/{found}: {page_label} ({exc})")

    return CleanupResult(target="pages", found=found, deleted=deleted, failed=failed)


def _cleanup_assessments(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    assignments = client.list_assignments(course_id)
    found = len(assignments)
    if dry_run:
        progress(f"[assessments] Would delete {found} assessment(s).")
        return CleanupResult(target="assessments", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, assignment in enumerate(assignments, start=1):
        assignment_id = _coerce_positive_int(assignment.get("id"))
        assignment_label = str(assignment.get("name") or assignment_id or "<unknown>")
        if assignment_id is None:
            failed += 1
            progress(f"[assessments] Failed {index}/{found}: missing assignment ID for {assignment_label}.")
            continue
        try:
            client.delete_assignment(course_id, assignment_id)
            deleted += 1
            progress(f"[assessments] Deleted {index}/{found}: {assignment_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[assessments] Failed {index}/{found}: {assignment_label} ({exc})")

    return CleanupResult(target="assessments", found=found, deleted=deleted, failed=failed)


def _cleanup_modules(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    modules = client.list_modules(course_id)
    found = len(modules)
    if dry_run:
        progress(f"[modules] Would delete {found} module(s).")
        return CleanupResult(target="modules", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, module in enumerate(modules, start=1):
        module_id = _coerce_positive_int(module.get("id"))
        module_label = str(module.get("name") or module_id or "<unknown>")
        if module_id is None:
            failed += 1
            progress(f"[modules] Failed {index}/{found}: missing module ID for {module_label}.")
            continue
        try:
            client.delete_module(course_id, module_id)
            deleted += 1
            progress(f"[modules] Deleted {index}/{found}: {module_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[modules] Failed {index}/{found}: {module_label} ({exc})")

    return CleanupResult(target="modules", found=found, deleted=deleted, failed=failed)


def _cleanup_files(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    files = client.list_files(course_id)
    found = len(files)
    if dry_run:
        progress(f"[files] Would delete {found} file(s).")
        return CleanupResult(target="files", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, file_data in enumerate(files, start=1):
        file_id = _coerce_positive_int(file_data.get("id"))
        file_label = str(file_data.get("display_name") or file_id or "<unknown>")
        if file_id is None:
            failed += 1
            progress(f"[files] Failed {index}/{found}: missing file ID for {file_label}.")
            continue
        try:
            client.delete_file(file_id)
            deleted += 1
            progress(f"[files] Deleted {index}/{found}: {file_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[files] Failed {index}/{found}: {file_label} ({exc})")

    return CleanupResult(target="files", found=found, deleted=deleted, failed=failed)


def _cleanup_discussions(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    all_topics = client.list_discussion_topics(course_id, only_announcements=False)
    discussions = [topic for topic in all_topics if not bool(topic.get("announcement"))]
    found = len(discussions)
    if dry_run:
        progress(f"[discussions] Would delete {found} discussion topic(s).")
        return CleanupResult(target="discussions", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, topic in enumerate(discussions, start=1):
        topic_id = _coerce_positive_int(topic.get("id"))
        topic_label = str(topic.get("title") or topic_id or "<unknown>")
        if topic_id is None:
            failed += 1
            progress(f"[discussions] Failed {index}/{found}: missing topic ID for {topic_label}.")
            continue
        try:
            client.delete_discussion_topic(course_id, topic_id)
            deleted += 1
            progress(f"[discussions] Deleted {index}/{found}: {topic_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[discussions] Failed {index}/{found}: {topic_label} ({exc})")

    return CleanupResult(target="discussions", found=found, deleted=deleted, failed=failed)


def _cleanup_announcements(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    topics = client.list_discussion_topics(course_id, only_announcements=True)
    found = len(topics)
    if dry_run:
        progress(f"[announcements] Would delete {found} announcement(s).")
        return CleanupResult(target="announcements", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, topic in enumerate(topics, start=1):
        topic_id = _coerce_positive_int(topic.get("id"))
        topic_label = str(topic.get("title") or topic_id or "<unknown>")
        if topic_id is None:
            failed += 1
            progress(f"[announcements] Failed {index}/{found}: missing topic ID for {topic_label}.")
            continue
        try:
            client.delete_discussion_topic(course_id, topic_id)
            deleted += 1
            progress(f"[announcements] Deleted {index}/{found}: {topic_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[announcements] Failed {index}/{found}: {topic_label} ({exc})")

    return CleanupResult(target="announcements", found=found, deleted=deleted, failed=failed)


def _cleanup_syllabus(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    course = client.get_course(course_id, include_syllabus_body=True)
    syllabus_body = course.get("syllabus_body")
    has_syllabus_content = isinstance(syllabus_body, str) and bool(syllabus_body.strip())
    found = 1 if has_syllabus_content else 0
    if dry_run:
        if has_syllabus_content:
            progress("[syllabus] Would clear syllabus body content.")
        else:
            progress("[syllabus] Syllabus body is already empty.")
        return CleanupResult(target="syllabus", found=found, deleted=0, failed=0)

    if not has_syllabus_content:
        progress("[syllabus] Nothing to delete.")
        return CleanupResult(target="syllabus", found=0, deleted=0, failed=0)

    try:
        client.update_course(course_id, {"syllabus_body": ""})
        progress("[syllabus] Cleared syllabus body content.")
        return CleanupResult(target="syllabus", found=1, deleted=1, failed=0)
    except CanvasAPIError as exc:
        progress(f"[syllabus] Failed to clear syllabus body ({exc})")
        return CleanupResult(target="syllabus", found=1, deleted=0, failed=1)


def _cleanup_calendar(
    client: CanvasClient,
    course_id: int,
    *,
    dry_run: bool,
    progress: ProgressCallback,
) -> CleanupResult:
    events = client.list_calendar_events(course_id)
    standalone_events = [event for event in events if not _is_assignment_linked_event(event)]
    found = len(standalone_events)
    if dry_run:
        progress(f"[calendar] Would delete {found} standalone calendar event(s).")
        return CleanupResult(target="calendar", found=found, deleted=0, failed=0)

    deleted = 0
    failed = 0
    for index, event in enumerate(standalone_events, start=1):
        event_id = _coerce_positive_int(event.get("id"))
        event_label = str(event.get("title") or event.get("description") or event_id or "<unknown>")
        if event_id is None:
            failed += 1
            progress(f"[calendar] Failed {index}/{found}: missing event ID for {event_label}.")
            continue
        try:
            client.delete_calendar_event(event_id)
            deleted += 1
            progress(f"[calendar] Deleted {index}/{found}: {event_label}")
        except CanvasAPIError as exc:
            failed += 1
            progress(f"[calendar] Failed {index}/{found}: {event_label} ({exc})")

    return CleanupResult(target="calendar", found=found, deleted=deleted, failed=failed)


def _coerce_positive_int(value: Any) -> int | None:
    if isinstance(value, int) and value > 0:
        return value
    if isinstance(value, str) and value.isdigit():
        parsed = int(value)
        return parsed if parsed > 0 else None
    return None


def _is_assignment_linked_event(event: dict[str, Any]) -> bool:
    if _coerce_positive_int(event.get("assignment_id")) is not None:
        return True
    return isinstance(event.get("assignment"), dict)
