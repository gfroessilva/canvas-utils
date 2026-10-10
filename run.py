"""CLI entry point for Canvas utility workflows."""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

from canvas_utils.api_client import CanvasAPIError, CanvasClient
from canvas_utils.cleanup import cleanup_course_resources, resolve_cleanup_targets
from canvas_utils.grades import export_assignment_grades
from canvas_utils.gradebook import (
    GradebookUploadLock,
    GradebookUploadJournal,
    apply_gradebook_plan,
    build_gradebook_plan,
    create_gradebook_template,
    default_gradebook_upload_journal_path,
    default_gradebook_upload_lock_path,
    describe_gradebook_change,
)
from canvas_utils.students import export_course_students
from canvas_utils.submissions import extract_submission_timestamps, resolve_optional_fields


def main(argv: list[str] | None = None) -> int:
    """Run the CLI and return an exit code."""
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "submissions":
        return _run_submissions_command(args)
    if args.command == "grades":
        return _run_grades_command(args)
    if args.command == "students":
        return _run_students_command(args)
    if args.command == "cleanup":
        return _run_cleanup_command(args)
    if args.command == "gradebook":
        return _run_gradebook_command(args)

    print(f"Error: Unsupported command '{args.command}'.", file=sys.stderr)
    return 1


def _run_submissions_command(args: argparse.Namespace) -> int:
    """Handle submissions CSV export command."""
    try:
        base_url, token, config_values = load_runtime_credentials(
            config_path=Path(args.config) if args.config else None,
            canvas_url_override=args.canvas_url,
            token_override=args.token,
        )
        assignment_id = parse_assignment_arg(args.assignment)
        optional_fields = parse_optional_fields(
            args.fields if args.fields is not None else config_values.get("submission_fields")
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    client = CanvasClient(base_url=base_url, token=token, timeout=args.timeout)

    try:
        written_files = extract_submission_timestamps(
            client=client,
            course_id=args.course,
            assignment_id=assignment_id,
            output=args.output,
            optional_fields=optional_fields,
        )
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    for output_file in written_files:
        print(f"Wrote: {output_file}")

    return 0


def _run_cleanup_command(args: argparse.Namespace) -> int:
    """Handle progressive cleanup command."""
    try:
        base_url, token, config_values = load_runtime_credentials(
            config_path=Path(args.config) if args.config else None,
            canvas_url_override=args.canvas_url,
            token_override=args.token,
        )
        targets = parse_cleanup_targets(
            args.targets if args.targets is not None else config_values.get("cleanup_targets")
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    client = CanvasClient(base_url=base_url, token=token, timeout=args.timeout)

    try:
        results = cleanup_course_resources(
            client=client,
            course_id=args.course,
            targets=targets,
            dry_run=not args.execute,
            progress_callback=print,
        )
    except (CanvasAPIError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("")
    print("Cleanup summary:")
    failed_total = 0
    for result in results:
        failed_total += result.failed
        print(
            f"  {result.target}: found={result.found}, deleted={result.deleted}, failed={result.failed}"
        )

    if not args.execute:
        print("")
        print("Dry run complete. Re-run with --execute to perform deletions.")
        return 0

    return 1 if failed_total > 0 else 0


def _run_grades_command(args: argparse.Namespace) -> int:
    """Handle grades XLSX export command."""
    try:
        base_url, token, _config_values = load_runtime_credentials(
            config_path=Path(args.config) if args.config else None,
            canvas_url_override=args.canvas_url,
            token_override=args.token,
        )
        assignment_id = parse_assignment_arg(args.assignment)
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    client = CanvasClient(base_url=base_url, token=token, timeout=args.timeout)

    try:
        written_files = export_assignment_grades(
            client=client,
            course_id=args.course,
            assignment_id=assignment_id,
            output=args.output,
            include_rubric_assessment=args.rubric_assessment,
        )
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    for output_file in written_files:
        print(f"Wrote: {output_file}")

    return 0


def _run_students_command(args: argparse.Namespace) -> int:
    """Handle students XLSX export command."""
    try:
        base_url, token, _config_values = load_runtime_credentials(
            config_path=Path(args.config) if args.config else None,
            canvas_url_override=args.canvas_url,
            token_override=args.token,
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    client = CanvasClient(base_url=base_url, token=token, timeout=args.timeout)

    try:
        written_files = export_course_students(
            client=client,
            course_id=args.course,
            output=args.output,
            active_only=args.active_only,
        )
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    for output_file in written_files:
        print(f"Wrote: {output_file}")

    return 0


def _run_gradebook_command(args: argparse.Namespace) -> int:
    """Handle controlled gradebook template, validation, and upload commands."""
    try:
        base_url, token, _config_values = load_runtime_credentials(
            config_path=Path(args.config) if args.config else None,
            canvas_url_override=args.canvas_url,
            token_override=args.token,
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    client = CanvasClient(base_url=base_url, token=token, timeout=args.timeout)

    if args.gradebook_action == "upload" and args.retry_indeterminate and not args.execute:
        print("Error: --retry-indeterminate requires --execute.", file=sys.stderr)
        return 1

    try:
        if args.gradebook_action == "template":
            output_file = create_gradebook_template(
                client=client,
                course_id=args.course,
                assignment_id=args.assignment,
                output=args.output,
            )
            print(f"Wrote: {output_file}")
            return 0
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.gradebook_action == "upload" and args.execute:
        journal_path = (
            Path(args.journal)
            if args.journal
            else default_gradebook_upload_journal_path(args.input)
        )
        try:
            with GradebookUploadLock(default_gradebook_upload_lock_path(journal_path)):
                return _run_gradebook_plan_action(
                    client=client,
                    args=args,
                    journal_path=journal_path,
                )
        except (CanvasAPIError, ValueError, OSError) as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1

    return _run_gradebook_plan_action(client=client, args=args)


def _run_gradebook_plan_action(
    *,
    client: CanvasClient,
    args: argparse.Namespace,
    journal_path: Path | None = None,
) -> int:
    """Build, display, and optionally execute a gradebook plan."""
    try:
        plan = build_gradebook_plan(
            client=client,
            course_id=args.course,
            assignment_id=args.assignment,
            input_path=args.input,
        )
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(
        f"Gradebook plan: changes={len(plan.changes)}, "
        f"unchanged_students={plan.unchanged_students}."
    )
    for change in plan.changes:
        print(f"  Student {change.user_id}: {describe_gradebook_change(change)}.")

    if args.gradebook_action == "validate":
        print("Gradebook validation succeeded.")
        return 0

    try:
        journal = GradebookUploadJournal.load(journal_path) if journal_path is not None else None
        result = apply_gradebook_plan(
            client=client,
            plan=plan,
            execute=args.execute,
            progress_callback=print,
            upload_journal=journal,
            retry_indeterminate=args.retry_indeterminate,
        )
    except (CanvasAPIError, ValueError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(
        "Gradebook summary: "
        f"planned={result.planned}, applied={result.applied}, "
        f"failed={result.failed}, conflicted={result.conflicted}, "
        f"indeterminate={result.indeterminate}, skipped={result.skipped}."
    )
    if result.journal_skipped:
        print(f"Upload journal skipped {result.journal_skipped} previously recorded change(s).")
    if result.indeterminate:
        print(
            "Indeterminate updates were recorded in the upload journal and will not be resent "
            "unless --retry-indeterminate is supplied."
        )
    if result.journal_errors:
        print(
            "One or more Canvas updates could not be recorded in the upload journal; "
            "do not rerun those rows without manually confirming Canvas."
        )
    if not args.execute:
        print("Dry run complete. Re-run with --execute to upload grades.")
        return 0
    return 1 if (result.failed or result.conflicted or result.indeterminate or result.journal_errors) else 0


def build_parser() -> argparse.ArgumentParser:
    """Build and return the argument parser."""
    parser = argparse.ArgumentParser(description="Canvas utility command-line interface.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    submissions_parser = subparsers.add_parser(
        "submissions",
        help="Extract Canvas submission timestamps to CSV.",
        description="Extract Canvas submission timestamps and export them to CSV.",
    )
    submissions_parser.add_argument(
        "--course",
        type=int,
        required=True,
        help="Canvas course ID.",
    )
    submissions_parser.add_argument(
        "--assignment",
        default="all",
        help="Assignment ID, or 'all' to process every assignment in the course (default: all).",
    )
    submissions_parser.add_argument(
        "--output",
        default=".",
        help="Output directory, or CSV file path when exporting a single assignment.",
    )
    submissions_parser.add_argument(
        "--fields",
        help=(
            "Comma-separated optional submission fields to include, such as "
            "student_name, submitted_at_excel, submitted_at_brisbane, score, "
            "grade, workflow_state, user_login_id, or user_sis_id."
        ),
    )
    _add_common_canvas_options(submissions_parser)

    cleanup_parser = subparsers.add_parser(
        "cleanup",
        help="Progressively delete selected course resources.",
        description=(
            "Progressively delete selected Canvas resources in a course "
            "(pages, assessments, modules, files, discussions, announcements, syllabus, calendar)."
        ),
    )
    cleanup_parser.add_argument(
        "--course",
        type=int,
        required=True,
        help="Canvas course ID.",
    )
    cleanup_parser.add_argument(
        "--targets",
        default="all",
        help=(
            "Comma-separated cleanup targets: pages, assessments, modules, files, "
            "discussions, announcements, syllabus, calendar, or all (default: all)."
        ),
    )
    cleanup_parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually delete resources. Without this flag, only a dry run is performed.",
    )
    _add_common_canvas_options(cleanup_parser)

    grades_parser = subparsers.add_parser(
        "grades",
        help="Export Canvas assignment marks with grader identity to XLSX.",
        description="Export Canvas assignment marks together with the resolved grader identity to XLSX.",
    )
    grades_parser.add_argument(
        "--course",
        type=int,
        required=True,
        help="Canvas course ID.",
    )
    grades_parser.add_argument(
        "--assignment",
        default="all",
        help="Assignment ID, or 'all' to process every assignment in the course (default: all).",
    )
    grades_parser.add_argument(
        "--output",
        default=".",
        help="Output directory, or XLSX file path when exporting a single assignment.",
    )
    grades_parser.add_argument(
        "--rubric-assessment",
        action="store_true",
        help="Request rubric assessments and add selected criterion descriptors to the XLSX export.",
    )
    _add_common_canvas_options(grades_parser)

    gradebook_parser = subparsers.add_parser(
        "gradebook",
        help="Create, validate, and safely upload an assignment gradebook.",
        description=(
            "Create a controlled gradebook XLSX workbook, validate it against Canvas, "
            "and upload grades only when --execute is explicitly supplied."
        ),
    )
    gradebook_actions = gradebook_parser.add_subparsers(dest="gradebook_action", required=True)

    gradebook_template_parser = gradebook_actions.add_parser(
        "template",
        help="Create a controlled XLSX template for one assignment.",
    )
    _add_gradebook_assignment_options(gradebook_template_parser)
    gradebook_template_parser.add_argument(
        "--output",
        default=".",
        help="Output directory, or the XLSX file path to create.",
    )
    _add_common_canvas_options(gradebook_template_parser)

    gradebook_validate_parser = gradebook_actions.add_parser(
        "validate",
        help="Validate an XLSX gradebook without changing Canvas.",
    )
    _add_gradebook_assignment_options(gradebook_validate_parser)
    gradebook_validate_parser.add_argument(
        "--input",
        required=True,
        help="Controlled gradebook XLSX file to validate.",
    )
    _add_common_canvas_options(gradebook_validate_parser)

    gradebook_upload_parser = gradebook_actions.add_parser(
        "upload",
        help="Upload a validated XLSX gradebook; dry run unless --execute is set.",
    )
    _add_gradebook_assignment_options(gradebook_upload_parser)
    gradebook_upload_parser.add_argument(
        "--input",
        required=True,
        help="Controlled gradebook XLSX file to upload.",
    )
    gradebook_upload_parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually upload gradebook changes. Without this flag, only a dry run is performed.",
    )
    gradebook_upload_parser.add_argument(
        "--journal",
        help=(
            "Upload journal path. Defaults to INPUT.xlsx.upload-journal.json when --execute is used."
        ),
    )
    gradebook_upload_parser.add_argument(
        "--retry-indeterminate",
        action="store_true",
        help=(
            "Retry changes whose prior Canvas request could not be confirmed. "
            "This can duplicate submission comments."
        ),
    )
    _add_common_canvas_options(gradebook_upload_parser)

    students_parser = subparsers.add_parser(
        "students",
        help="Export course student roster to XLSX.",
        description="Export Canvas course students to XLSX with student identity columns.",
    )
    students_parser.add_argument(
        "--course",
        type=int,
        required=True,
        help="Canvas course ID.",
    )
    students_parser.add_argument(
        "--output",
        default=".",
        help="Output directory, or XLSX file path for the exported student roster.",
    )
    students_parser.add_argument(
        "--active-only",
        action="store_true",
        help="Export only active student enrollments. Default behavior includes all student enrollments.",
    )
    _add_common_canvas_options(students_parser)

    return parser


def parse_assignment_arg(raw_assignment: str) -> int | None:
    """Parse assignment CLI value into assignment ID or None for all."""
    lowered = raw_assignment.strip().lower()
    if lowered == "all":
        return None
    if lowered.isdigit() and int(lowered) > 0:
        return int(lowered)
    raise ValueError("--assignment must be a positive integer or 'all'.")


def load_runtime_credentials(
    config_path: Path | None,
    canvas_url_override: str | None,
    token_override: str | None,
) -> tuple[str, str, dict[str, Any]]:
    """Load credentials from CLI overrides, then environment, then config file."""
    config_values = _load_config_values(config_path)
    base_url = canvas_url_override or os.environ.get("CANVAS_BASE_URL")
    token = token_override or os.environ.get("CANVAS_API_TOKEN")

    if base_url and token:
        return base_url, token, config_values

    base_url = base_url or config_values.get("canvas_base_url") or config_values.get("base_url")
    token = token or config_values.get("canvas_api_token") or config_values.get("token")

    if not base_url:
        raise ValueError(
            "Canvas base URL not provided. Set --canvas-url, CANVAS_BASE_URL, or config field canvas_base_url."
        )
    if not token:
        raise ValueError(
            "Canvas API token not provided. Set --token, CANVAS_API_TOKEN, or config field canvas_api_token."
        )

    return str(base_url), str(token), config_values


def parse_optional_fields(raw_fields: object) -> tuple[str, ...]:
    """Parse CLI/config optional field selections into normalized field names."""
    if raw_fields is None:
        return tuple()

    if isinstance(raw_fields, str):
        parts = [part.strip() for part in raw_fields.split(",")]
        return resolve_optional_fields(parts)

    if isinstance(raw_fields, list):
        parts = [str(part) for part in raw_fields]
        return resolve_optional_fields(parts)

    raise ValueError("--fields or submission_fields must be a comma-separated string or list of strings.")


def parse_cleanup_targets(raw_targets: object) -> tuple[str, ...]:
    """Parse CLI/config cleanup target selection."""
    if raw_targets is None:
        return resolve_cleanup_targets(["all"])

    if isinstance(raw_targets, str):
        parts = [part.strip() for part in raw_targets.split(",")]
        return resolve_cleanup_targets(parts)

    if isinstance(raw_targets, list):
        parts = [str(part) for part in raw_targets]
        return resolve_cleanup_targets(parts)

    raise ValueError("--targets or cleanup_targets must be a comma-separated string or list of strings.")


def _load_config_values(config_path: Path | None) -> dict[str, Any]:
    if config_path is None or not config_path.exists():
        return {}
    try:
        raw_data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Config file is not valid JSON: {config_path} ({exc})") from exc

    if not isinstance(raw_data, dict):
        raise ValueError(f"Config file must contain a JSON object: {config_path}")

    return raw_data


def _add_common_canvas_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--canvas-url",
        help="Canvas base URL. Defaults to CANVAS_BASE_URL, then config file.",
    )
    parser.add_argument(
        "--token",
        help="Canvas API token. Defaults to CANVAS_API_TOKEN, then config file.",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Optional config JSON path for credentials fallback (default: ./config.json).",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="Per-request timeout in seconds.",
    )


def _add_gradebook_assignment_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--course",
        type=_parse_positive_int,
        required=True,
        help="Canvas course ID.",
    )
    parser.add_argument(
        "--assignment",
        type=_parse_positive_int,
        required=True,
        help="Canvas assignment ID.",
    )


def _parse_positive_int(raw_value: str) -> int:
    try:
        value = int(raw_value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a positive integer.") from exc
    if value <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer.")
    return value


if __name__ == "__main__":
    raise SystemExit(main())
