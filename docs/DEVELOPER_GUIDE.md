# canvas-utils developer guide

This guide describes the implementation, module boundaries, API behavior, tests, and the conventions to follow when extending `canvas-utils`.

## Contents

- [Architecture](#architecture)
- [Repository layout](#repository-layout)
- [Runtime flow](#runtime-flow)
- [Module reference](#module-reference)
- [Canvas API client](#canvas-api-client)
- [Data and output conventions](#data-and-output-conventions)
- [Testing](#testing)
- [Extending the project](#extending-the-project)
- [Current scope](#current-scope)

## Architecture

The project is a lightweight Python CLI rather than an installable package or web service.

```text
run.py
  -> parse command-line arguments
  -> load credentials and command-specific settings
  -> construct CanvasClient
  -> call an export, gradebook, or cleanup workflow
  -> write local files or report cleanup progress
```

The implementation deliberately separates:

- `run.py`, which owns argument parsing, configuration resolution, user-facing error handling, and process exit codes.
- `canvas_utils.api_client`, which owns authenticated HTTP requests, pagination, JSON validation, and Canvas-specific error messages.
- Feature modules, which map Canvas payloads to rows, workbooks, or cleanup actions without knowing about command-line parsing.

There is no persistent database, background worker, or state file. Every command obtains its data from Canvas at runtime and writes its result to the requested output location.

## Repository layout

| Path | Responsibility |
| --- | --- |
| `run.py` | CLI entry point and command dispatch. |
| `canvas_utils/__init__.py` | Package metadata and version. |
| `canvas_utils/api_client.py` | Canvas HTTP client, pagination, response validation, and error formatting. |
| `canvas_utils/submissions.py` | Submission CSV extraction and optional-field mapping. |
| `canvas_utils/grades.py` | Assignment-grade XLSX export and marker-statistics formulas. |
| `canvas_utils/gradebook.py` | Controlled gradebook template, validation, conflict detection, and upload workflow. |
| `canvas_utils/rubrics.py` | Shared Canvas rubric extraction, identifier normalization, and rating lookup helpers. |
| `canvas_utils/students.py` | Student-roster XLSX export and group-membership mapping. |
| `canvas_utils/cleanup.py` | Dry-run and progressive destructive cleanup workflows. |
| `canvas_utils/cohort.py` | Reserved placeholder for future cohort analytics; it has no implementation yet. |
| `tests/` | Unit tests using fake Canvas clients and local temporary output files. |
| `config.json.template` | Safe configuration-file shape without a real credential. |
| `requirements.txt` | Runtime dependency list. |
| `docs/USER_GUIDE.md` | User-facing installation, commands, outputs, and safety documentation. |

## Runtime flow

`main()` builds an `argparse` parser and dispatches to one handler per subcommand:

| CLI command | Handler | Workflow |
| --- | --- | --- |
| `submissions` | `_run_submissions_command` | `extract_submission_timestamps` |
| `grades` | `_run_grades_command` | `export_assignment_grades` |
| `gradebook template` | `_run_gradebook_command` | `create_gradebook_template` |
| `gradebook validate` | `_run_gradebook_command` | `build_gradebook_plan` |
| `gradebook upload` | `_run_gradebook_command` | `build_gradebook_plan` and `apply_gradebook_plan` |
| `students` | `_run_students_command` | `export_course_students` |
| `cleanup` | `_run_cleanup_command` | `cleanup_course_resources` |

Each handler:

1. Calls `load_runtime_credentials`.
2. Parses and validates command-specific values.
3. Instantiates `CanvasClient` with the configured timeout.
4. Calls the feature module.
5. Catches expected `CanvasAPIError`, `ValueError`, and relevant filesystem errors, printing a concise error to standard error.

The export handlers print one `Wrote:` line per generated file and return exit code `0` on success. Cleanup returns `0` for a dry run, and returns `1` after execution if any individual resource deletion failed.

## Module reference

### `canvas_utils.api_client`

`CanvasClient` is the common HTTP abstraction. It accepts `base_url`, `token`, and `timeout`, and exposes resource-oriented methods rather than leaking URL construction into feature modules.

Important public methods include:

```python
client.list_assignments(course_id)
client.get_assignment(course_id, assignment_id, include_rubric=False)
client.list_submissions(
    course_id,
    assignment_id,
    include_user=True,
    include_rubric_assessment=False,
)
client.grade_submission(
    course_id,
    assignment_id,
    user_id,
    posted_grade=None,
    submission_comment=None,
    submission_comment_attempt=None,
    rubric_assessment=None,
)
client.get_submission(
    course_id,
    assignment_id,
    user_id,
    include_rubric_assessment=False,
)
client.list_course_students(course_id, active_only=False)
client.list_course_groups(course_id)
client.get_user(user_id)
client.get_course_user(course_id, user_id)
```

It also provides the list, delete, and update methods used by cleanup:

```python
client.list_pages(course_id)
client.delete_page(course_id, page_url)
client.delete_assignment(course_id, assignment_id)
client.list_modules(course_id)
client.delete_module(course_id, module_id)
client.list_files(course_id)
client.delete_file(file_id)
client.list_discussion_topics(course_id, only_announcements=False)
client.delete_discussion_topic(course_id, topic_id)
client.get_course(course_id, include_syllabus_body=False)
client.update_course(course_id, updates)
client.list_calendar_events(course_id)
client.delete_calendar_event(event_id)
```

`CanvasAPIError` is the domain exception for HTTP errors and unexpected Canvas response shapes. It carries an optional `status_code`, which grade export uses to distinguish a 404 user lookup from other failures. `CanvasTransportError` is its subclass for network or malformed-response failures and identifies whether a mutation may already have reached Canvas.

### `canvas_utils.submissions`

The main entry point is:

```python
extract_submission_timestamps(
    client,
    course_id,
    assignment_id,
    output=".",
    optional_fields=None,
)
```

`assignment_id=None` means all assignments. The workflow requests assignments first, then gets submissions for each assignment with user data included. `resolve_optional_fields` validates supported field names, removes duplicates, and preserves input order. `build_csv_headers` maps those normalized names to CSV columns.

Keep field declarations centralized in `OPTIONAL_FIELD_SPECS`. Adding a submission export field normally requires one entry with the external field name, output header, and extractor function, plus corresponding tests and user documentation.

### `canvas_utils.grades`

The main entry point is:

```python
export_assignment_grades(
    client,
    course_id,
    assignment_id,
    output=".",
    include_rubric_assessment=False,
)
```

It writes a `Grades` sheet and a `Marker Statistics` sheet for each selected assignment. Grade records use the top-level Canvas submission values `score`, `grade`, and `entered_grade`.

When `include_rubric_assessment=True`, the workflow requests `rubric_assessment` with each submission and requests the assignment rubric. It writes a `Rubric Assessments` sheet with a row per criterion assessment. The assessment's `rating_id` is matched to the rubric rating ID so the rating description can be exported as `Selected_Descriptor`. Canvas assessment keys such as `criterion_123` are normalized to match rubric criterion IDs such as `_123`. A directly supplied assessment descriptor is used only if the rating cannot be resolved.

Positive grader IDs are resolved through `get_user`, then `get_course_user` after a global 404. Negative IDs become `automatic`; missing IDs become `ungraded`; IDs that cannot be resolved from either endpoint become `unresolvable`. Grader lookups are cached for the lifetime of an export operation.

`build_marker_statistics` groups rows by the complete grader identity tuple. Formula helpers build Excel formulas that calculate count, scored count, unscored count, average, minimum, and maximum from the `Grades` sheet. Preserve the relationship between `GRADE_HEADERS`, `MARKER_STATS_HEADERS`, and formula helpers when adding or reordering columns.

### `canvas_utils.gradebook`

The gradebook workflow is deliberately separate from the read-only `grades` export. Its public operations are:

```python
create_gradebook_template(client, course_id, assignment_id, output=".")
build_gradebook_plan(client, course_id, assignment_id, input_path)
apply_gradebook_plan(client, plan, execute=False, progress_callback=None, upload_journal=None)
```

`create_gradebook_template` retrieves the assignment rubric and submissions with users and rubric assessments included. It writes a versioned workbook with visible `Grades`, `Rubric`, `Rubric Catalog`, `Canvas Mapping`, and `Instructions` sheets, plus hidden metadata, snapshot, and descriptor-list sheets. The `Rubric` sheet is normalized to one row per student and criterion; descriptor dropdown labels resolve to immutable Canvas rating IDs.

`build_gradebook_plan` is the no-write preflight boundary. It requires the generated metadata and snapshot sheets, including a fresh per-template UUID, verifies the course and assignment IDs, checks a stable rubric fingerprint, resolves explicit `Canvas Mapping` entries, rejects formulas in mapped inputs, validates assignment grading modes, and validates every requested student/rubric change against live submission data before returning a plan. `Canvas_User_ID` is the only student matching key. `Grade_Action` accepts `KEEP` or `SET`; `Comment_Action` accepts `KEEP` or `ADD`; rubric `Action` accepts `KEEP` or `SET`. `not_graded` assignments reject grade changes; pass/fail values are constrained to Canvas's valid endpoints and textual states.

`apply_gradebook_plan` makes no requests in dry-run mode. During execution it re-fetches each affected grade, rubric, or comment submission immediately before sending one `grade_submission` request per changed student. A mismatch becomes a reported conflict rather than an overwrite; comment requests also target the validated submission attempt. Canvas does not expose a conditional version field for the final write, so this only narrows the read/write race. Execution requires a `GradebookUploadJournal`; the CLI holds its paired local advisory lock from plan creation through the final journal update. Before each mutation, the journal writes a durable `pending` request hash; it is updated to `applied` or `indeterminate` only after Canvas responds. Later executions skip all applied, pending, and indeterminate hashes unless `--retry-indeterminate` is explicitly selected. The per-template UUID prevents a historical journal entry from suppressing the same intended payload in a newly generated workbook. Transport and server failures that may have occurred after Canvas received the request are indeterminate, not ordinary failures. Do not add automatic mutation retries for submission comments: Canvas comments are additive and an uncertain retry can duplicate student feedback.

The initial scope rejects anonymous, moderated, and group assignments. It intentionally does not implement clearing grades, rubric selections, or comments, because Canvas's write semantics for those operations need their own explicit design.

### `canvas_utils.rubrics`

Shared helpers centralize Canvas's inconsistent criterion key forms, such as rubric ID `_123` and assessment key `criterion_123`. Both grade export and gradebook upload use these helpers to resolve criteria and ratings rather than duplicating identifier logic.

### `canvas_utils.students`

The main entry point is:

```python
export_course_students(client, course_id, output=".", active_only=False)
```

It retrieves student enrollments and course groups. `_build_group_memberships` maps Canvas user IDs to a case-insensitively sorted, semicolon-separated group string. The workbook writer uses `STUDENT_HEADERS` as the source of truth and sizes columns from written values.

The current output deliberately maps:

| Output column | Canvas field |
| --- | --- |
| `Student` | `name` |
| `Email` | `email`, with `login_id` fallback |
| `Student Number` | `integration_id` |
| `ID` | numeric `sis_user_id` |
| `Groups` | group memberships found by user ID |

### `canvas_utils.cleanup`

The public workflow is:

```python
cleanup_course_resources(
    client,
    course_id,
    targets,
    dry_run=True,
    progress_callback=None,
)
```

`resolve_cleanup_targets` normalizes a sequence of names and expands `all` to `ALL_CLEANUP_TARGETS`. Each handler returns an immutable `CleanupResult(target, found, deleted, failed)`.

The handlers deliberately catch `CanvasAPIError` around individual mutation requests and continue to the next resource. This allows a useful end-of-run summary even when Canvas rejects a specific deletion. Do not make cleanup fail silently or convert failed deletions into successful counts.

`calendar` excludes assignment-linked events. `syllabus` clears only a nonempty `syllabus_body`. The rest of the targets list and delete their corresponding Canvas resources.

### `run.py`

`run.py` is the executable boundary. It should remain responsible for:

- `argparse` parser construction and help text.
- Parsing assignment, optional-field, and cleanup-target inputs.
- Credential resolution and configuration-file parsing.
- Instantiating `CanvasClient`.
- Rendering command output and exit status.

Feature modules should not call `sys.exit`, parse environment variables, or print CLI-specific messages. Keeping that boundary makes their logic easy to unit-test with fake clients.

## Canvas API client

All Canvas paths use `/api/v1`. List methods request `per_page=100` and `get_paginated_json` follows the Canvas `Link` header while a `rel="next"` URL is present. Later pages use Canvas's supplied next URL without reapplying the original query parameters.

The currently used endpoint families are:

| Feature | Endpoint family |
| --- | --- |
| Assignment discovery and rubric lookup | `/courses/{course_id}/assignments` and `/courses/{course_id}/assignments/{assignment_id}` |
| Submission exports | `/courses/{course_id}/assignments/{assignment_id}/submissions` |
| Gradebook pre-write check and upload | `GET` and `PUT /courses/{course_id}/assignments/{assignment_id}/submissions/{user_id}` |
| Student roster | `/courses/{course_id}/users` with a student-enrollment filter |
| Group membership | `/courses/{course_id}/groups` with `users` included |
| Grader resolution | `/users/{user_id}` and `/courses/{course_id}/users/{user_id}` |
| Pages | `/courses/{course_id}/pages` |
| Modules | `/courses/{course_id}/modules` |
| Files | `/courses/{course_id}/files` |
| Discussions and announcements | `/courses/{course_id}/discussion_topics` |
| Course and syllabus | `/courses/{course_id}` |
| Calendar events | `/calendar_events` scoped by `context_codes[]=course_{course_id}` |

`request_json` uses the standard library HTTP client, parses JSON responses, and adds a Canvas authorization header. HTTP errors become `CanvasAPIError` with contextual messages for common 401 and 404 cases. Do not catch broad exceptions around this client: unexpected failures should remain visible rather than being mistaken for successful exports or cleanup.

## Data and output conventions

- Export functions return `list[pathlib.Path]`, even when they write one file.
- `assignment_id=None` consistently means "all assignments."
- A single-assignment export accepts a file path only when it has the expected extension. All-assignment mode requires a directory.
- Filenames use a sanitized, lowercase assignment name and the local date.
- Output writers create missing parent directories.
- CSV files are UTF-8 and include a header even when Canvas returns no submissions.
- XLSX writers use `openpyxl`; grade workbooks request full recalculation when opened so formula-based statistics update in Excel.
- Mapping code tolerates missing or malformed optional Canvas fields by writing `None` instead of raising, where that behavior is already encoded by a coercion helper.

When changing a schema, update the associated header constant, mapping logic, tests, and `docs/USER_GUIDE.md` together.

## Testing

The repository uses the Python standard-library `unittest` runner. Run the full suite from the repository root:

```powershell
python -m unittest discover -s tests -v
```

On Linux or macOS:

```bash
python3 -m unittest discover -s tests -v
```

Tests use fake clients rather than real Canvas credentials or network calls. They cover:

- Argument parsing and validation.
- Submission headers, optional fields, CSV contents, and filenames.
- Grade rows, grader resolution behavior, marker-statistics formulas, rubric descriptor resolution, and rubric request parameters.
- Gradebook template sheets, descriptor dropdowns, mapping resolution, conflict detection, dry runs, and per-student upload failure reporting.
- Student fields, explicit XLSX output paths, active-only forwarding, and group membership aggregation.
- Cleanup dry-run, selected-target execution, syllabus handling, and calendar-event filtering.
- API-client query parameters and contextual error formatting.

Add focused tests for every behavior change. For output changes, open the generated CSV or workbook in the test and assert the exact header and representative row values.

## Extending the project

To add a new CLI workflow:

1. Add a focused module under `canvas_utils/` with a public operation that accepts a `CanvasClient`.
2. Add only the client methods needed for its Canvas endpoints to `CanvasClient`.
3. Add a subparser, handler, and dispatch branch in `run.py`.
4. Keep input validation in `run.py` and data transformation in the feature module.
5. Test the feature with a fake client and temporary output directory when files are produced.
6. Update `README.md`, `docs/USER_GUIDE.md`, and this guide when interfaces, behavior, or project structure change.

Use existing patterns for output paths, positive integer coercion, and Canvas API errors instead of introducing inconsistent behavior. Treat destructive actions conservatively: dry-run by default, report every attempted mutation, and surface failures in the result.

## Current scope

- The project is an executable repository, not a published Python package with a packaging manifest.
- `canvas_utils.cohort` is intentionally a placeholder and does not expose cohort analytics.
- Grade export is assignment-level by default. `--rubric-assessment` adds the returned per-criterion assessments, selected rating descriptors, points, and comments.
- Gradebook upload is the submission-write workflow and requires `--execute`; cleanup mutations also require `--execute`.
- The CLI uses Canvas data as returned to the current token. Hidden users, restricted submissions, and inaccessible resources can affect what appears in an export.
