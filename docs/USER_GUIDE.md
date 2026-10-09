# canvas-utils user guide

`canvas-utils` is a command-line tool for exporting Canvas LMS course data, progressively removing selected course resources, and safely uploading assignment grades from controlled XLSX workbooks. It uses the Canvas REST API directly and writes CSV or XLSX files locally.

## Contents

- [Capabilities](#capabilities)
- [Requirements and installation](#requirements-and-installation)
- [Authentication and configuration](#authentication-and-configuration)
- [Common command behavior](#common-command-behavior)
- [Command reference](#command-reference)
- [Output reference](#output-reference)
- [Gradebook safety](#gradebook-safety)
- [Cleanup safety](#cleanup-safety)
- [Canvas access](#canvas-access)
- [Troubleshooting](#troubleshooting)

## Capabilities

| Command | Purpose | Output or effect |
| --- | --- | --- |
| `submissions` | Exports submission timestamps, status, and selected metadata. | One CSV per assignment. |
| `grades` | Exports assignment-level grades and resolved grader identity. | One XLSX workbook per assignment. |
| `gradebook` | Creates, validates, and uploads controlled gradebook workbooks. | One XLSX template, validation output, or guarded submission updates. |
| `students` | Exports the student roster and group memberships. | One XLSX workbook per course. |
| `cleanup` | Lists or deletes selected course resources. | Progress messages and a per-target summary. |

The export commands only read Canvas data. `gradebook upload` and `cleanup --execute` change Canvas data; gradebook upload is intentionally protected by a dry-run default.

## Requirements and installation

- Python 3.10 or later
- A Canvas base URL, such as `https://canvas.example.edu`
- A Canvas API token with access appropriate to the command being run
- Internet access to the Canvas instance

The only third-party dependency is `openpyxl`, which creates XLSX exports.

### PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Linux or macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
```

Run the CLI from the repository root:

```text
python run.py <subcommand> [options]
```

Use `python3` instead of `python` where that is how Python 3 is installed.

## Authentication and configuration

Each setting is resolved independently in this order:

1. Command-line option
2. Environment variable
3. JSON configuration file

The common options are:

| Option | Environment variable | Description |
| --- | --- | --- |
| `--canvas-url` | `CANVAS_BASE_URL` | Canvas instance base URL. |
| `--token` | `CANVAS_API_TOKEN` | Canvas API token. |
| `--config` | N/A | Path to a JSON configuration file. Defaults to `config.json`. |
| `--timeout` | N/A | HTTP request timeout in seconds. Defaults to `30`. |

The default `config.json` is optional. If it does not exist, environment variables or CLI options can provide the credentials instead.

Create a local configuration file from the template:

### PowerShell

```powershell
Copy-Item .\config.json.template .\config.json
```

### Linux or macOS

```bash
cp config.json.template config.json
```

Use the canonical keys shown in the template:

```json
{
  "canvas_base_url": "https://canvas.example.edu",
  "canvas_api_token": "REPLACE_WITH_TOKEN",
  "submission_fields": ["student_name", "submitted_at_brisbane"]
}
```

For compatibility, `base_url` and `token` are also accepted in a configuration file. `submission_fields` accepts either a comma-separated string or a JSON list.

`config.json` is ignored by Git. Do not put a real token in a tracked file. Environment variables or the ignored configuration file are generally safer than passing `--token` on a shell command line, which can be retained in shell history.

## Common command behavior

Every subcommand requires `--course COURSE_ID`. Course and assignment values are Canvas numeric IDs, not course codes or assignment names.

`submissions` and `grades` also support:

| Option | Meaning |
| --- | --- |
| `--assignment ID` | Process one assignment. |
| `--assignment all` | Process every assignment returned by Canvas for the course. This is the default. |
| `--output PATH` | A directory, or a file path when processing one assignment. |

When `--assignment all` is used, the output must be a directory. The command writes a separate file for every returned assignment. A `.csv` or `.xlsx` file path is only valid when a single assignment is selected.

The date in generated filenames is the local date of the machine running the command. Assignment names are converted to lowercase, filename-safe text.

`gradebook` always operates on one assignment and requires both `--course COURSE_ID` and `--assignment ASSIGNMENT_ID`. Its `template`, `validate`, and `upload` actions share the common Canvas credential options.

## Command reference

### `submissions`

Exports one CSV row for each submission record returned by Canvas. The default fields focus on submission attempts and late status. Use `--fields` to append supported metadata.

```text
python run.py submissions --course COURSE_ID [--assignment ID|all] [--output PATH] [--fields FIELD[,FIELD...]]
```

Examples:

```powershell
python .\run.py submissions --course 12345 --assignment 67890 --output .\exports
python .\run.py submissions --course 12345 --assignment all --output .\exports
python .\run.py submissions --course 12345 --assignment 67890 --fields student_name,score,workflow_state
```

The optional field names below are case-insensitive. Empty values and duplicate field names are ignored; requested fields retain their first-seen order.

| Field name | CSV column | Value |
| --- | --- | --- |
| `student_name` | `Student_Name` | Canvas submission user's name. |
| `submitted_at_excel` | `Submitted_At_Excel` | Submission time normalized to UTC as `YYYY-MM-DD HH:MM:SS`. |
| `submitted_at_brisbane` | `Submitted_At_Brisbane` | Submission time in Brisbane time (UTC+10) as `YYYY-MM-DD HH:MM:SS`. |
| `score` | `Score` | Canvas assignment score. |
| `grade` | `Grade` | Canvas assignment grade. |
| `entered_grade` | `Entered_Grade` | Grade entered by the grader. |
| `workflow_state` | `Workflow_State` | Canvas submission workflow state. |
| `missing` | `Missing` | Canvas missing flag. |
| `excused` | `Excused` | Canvas excused flag. |
| `late_policy_status` | `Late_Policy_Status` | Canvas late-policy status. |
| `posted_at` | `Posted_At` | Grade-posting timestamp returned by Canvas. |
| `seconds_late` | `Seconds_Late` | Canvas late duration in seconds. |
| `user_login_id` | `Student_Login_ID` | Canvas submission user's login ID. |
| `user_sis_id` | `Student_SIS_ID` | Canvas submission user's SIS user ID. |
| `user_short_name` | `Student_Short_Name` | Canvas submission user's short name. |

An unsupported field fails the command before Canvas requests are made and lists the supported field names.

### `grades`

Exports one XLSX workbook per assignment. The workbook has a `Grades` sheet with assignment, student, score, grade, workflow, and grader information, plus a `Marker Statistics` sheet with formulas grouped by grader. With `--rubric-assessment`, it also has a `Rubric Assessments` sheet with one row per returned criterion assessment.

```text
python run.py grades --course COURSE_ID [--assignment ID|all] [--output PATH] [--rubric-assessment]
```

Examples:

```powershell
python .\run.py grades --course 12345 --assignment 67890 --output .\exports
python .\run.py grades --course 12345 --assignment all --output .\exports
python .\run.py grades --course 12345 --assignment 67890 --output .\exports\assessment-1-grades.xlsx
python .\run.py grades --course 12345 --assignment 67890 --rubric-assessment --output .\exports
```

By default, the export contains Canvas's top-level assignment submission fields: `score`, `grade`, and `entered_grade`. For an assignment assessed with a rubric, these values represent Canvas's assignment-level result.

`--rubric-assessment` requests `rubric_assessment` with each submission and retrieves the assignment rubric. It matches the assessment's `rating_id` to the rubric rating and writes that rating's description to `Selected_Descriptor`. This makes the marker's selected performance descriptor visible alongside the criterion, points, comments, student, and grader. A criterion assessed with free-form points but no selected rating has an empty descriptor; a descriptor supplied directly by Canvas is used as a fallback when a rating cannot be matched.

For a positive `grader_id`, the command first attempts a global Canvas user lookup and then falls back to a course-scoped lookup when the global lookup returns 404. Negative grader IDs are identified as automatic grading. Missing grader IDs are identified as ungraded. If neither user lookup can resolve a positive grader ID, the row is retained and marked as unresolvable.

### `gradebook`

Creates a controlled XLSX workbook for one assignment, validates changes against live Canvas data, and uploads only when explicitly requested. Generate a template before entering grades; `gradebook upload` requires the template's metadata and baseline sheets even when the visible data columns are renamed.

```text
python run.py gradebook template --course COURSE_ID --assignment ASSIGNMENT_ID [--output PATH]
python run.py gradebook validate --course COURSE_ID --assignment ASSIGNMENT_ID --input WORKBOOK.xlsx
python run.py gradebook upload --course COURSE_ID --assignment ASSIGNMENT_ID --input WORKBOOK.xlsx [--execute] [--journal PATH] [--retry-indeterminate]
```

Examples:

```powershell
# Create a workbook with every submission recipient and rubric criterion.
python .\run.py gradebook template --course 12345 --assignment 67890 --output .\exports\assessment-1-gradebook.xlsx

# Check identifiers, rubric selections, values, and Canvas conflicts. No Canvas data changes.
python .\run.py gradebook validate --course 12345 --assignment 67890 --input .\exports\assessment-1-gradebook.xlsx

# Review the planned requests, then explicitly upload the same workbook.
python .\run.py gradebook upload --course 12345 --assignment 67890 --input .\exports\assessment-1-gradebook.xlsx
python .\run.py gradebook upload --course 12345 --assignment 67890 --input .\exports\assessment-1-gradebook.xlsx --execute
```

The template's visible sheets are:

| Sheet | Purpose |
| --- | --- |
| `Grades` | One row per student. Set `Grade_Action` to `SET` with `New_Grade`, or set `Comment_Action` to `ADD` with `New_Submission_Comment`. |
| `Rubric` | One row per student and criterion. Set `Action` to `SET`, choose `Selected_Descriptor` from its dropdown or provide `Selected_Rating_ID`, then optionally provide `Points` and `Criterion_Comment`. |
| `Rubric Catalog` | Reference for the Canvas criterion IDs, rating IDs, descriptors, and default points. |
| `Canvas Mapping` | Maps the logical gradebook fields to data-sheet column names. Update it if data columns or sheet names are renamed. |
| `Instructions` | Workbook editing and safety rules. |

`Canvas_User_ID` is the required student identifier. Student names and login IDs are audit fields only; they are never used to match students. `KEEP` means no change. A row with an editable value but a `KEEP` action fails validation rather than being silently ignored.

`New_Submission_Comment` is intentionally additive: Canvas creates a new submission comment and does not edit previous feedback. Re-running an executed workbook with `Comment_Action` set to `ADD` creates another comment.

The uploader validates every workbook row before its first write. It checks the course and assignment IDs, rubric fingerprint, Canvas user IDs, descriptor-to-rating mapping, numeric rubric points, duplicate rubric rows, assignment grading type, and template baselines. A requested grade, rubric, or comment change fails if the corresponding Canvas value changed after the template was created. Immediately before each write, it rechecks that student's live submission again; comments target the validated submission attempt. Canvas does not provide a conditional write, so a narrow read/write race remains possible.

Gradebook upload currently rejects anonymous, moderated, and group assignments so that unsupported Canvas grading behavior cannot be changed accidentally. It does not clear existing grades, rubric selections, or comments; create a new template and use Canvas directly if clearing is required.

### `students`

Exports a course roster to one XLSX workbook. It includes all student enrollments returned by Canvas by default, including non-active enrollments where Canvas permits them. Add `--active-only` to request only active student enrollments.

```text
python run.py students --course COURSE_ID [--output PATH] [--active-only]
```

Examples:

```powershell
python .\run.py students --course 12345 --output .\exports
python .\run.py students --course 12345 --output .\exports\roster.xlsx
python .\run.py students --course 12345 --output .\exports --active-only
```

The command also requests course groups and writes every matching group name for each student. Group names are sorted case-insensitively and joined with `; `.

### `cleanup`

Lists or progressively removes selected resources. Without `--execute`, it is a dry run: Canvas is queried, but no delete or update request is sent.

```text
python run.py cleanup --course COURSE_ID [--targets TARGET[,TARGET...]] [--execute]
```

Examples:

```powershell
# Review only; no data is changed.
python .\run.py cleanup --course 12345 --targets pages,assessments,modules

# Perform the selected deletions.
python .\run.py cleanup --course 12345 --targets pages,modules --execute

# Perform every supported cleanup target.
python .\run.py cleanup --course 12345 --targets all --execute
```

Supported targets are `pages`, `assessments`, `modules`, `files`, `discussions`, `announcements`, `syllabus`, `calendar`, and `all`. Target names are case-insensitive; duplicate names are ignored. `all` selects the complete supported target list.

| Target | Canvas resources affected |
| --- | --- |
| `pages` | All course pages. |
| `assessments` | All course assignments. |
| `modules` | All course modules. |
| `files` | All files owned by the course. |
| `discussions` | Discussion topics that are not announcements. |
| `announcements` | Announcement discussion topics. |
| `syllabus` | The course syllabus body only. |
| `calendar` | Standalone course calendar events; assignment-linked events are skipped. |

Each target is processed progressively. A failure for one resource is reported and does not prevent later resources in the same target from being attempted. The command exits with a nonzero status after `--execute` if any resource failed.

## Output reference

### Generated filenames

| Command | Generated filename |
| --- | --- |
| `submissions` | `canvas_submissions_<assignment-id>_<assignment-name>_<YYYY-MM-DD>.csv` |
| `grades` | `canvas_grades_<assignment-id>_<assignment-name>_<YYYY-MM-DD>.xlsx` |
| `gradebook template` | `canvas_gradebook_<assignment-id>_<assignment-name>_<YYYY-MM-DD>.xlsx` |
| `students` | `canvas_students_<course-id>_<YYYY-MM-DD>.xlsx` |

When a single-assignment `--output` path ends in the expected extension, the supplied filename is used instead.

### Submission CSV

Every CSV begins with these columns:

| Column | Meaning |
| --- | --- |
| `Assignment_ID` | Canvas assignment ID from the submission, or the requested assignment ID when unavailable. |
| `Student_ID` | Canvas user ID from the submission. |
| `Attempt_Number` | Canvas submission attempt number. |
| `Submitted_At` | Raw Canvas submission timestamp. |
| `Canvas_Late_Flag` | Boolean value based on Canvas's `late` field. |

Selected optional fields are appended after the base columns in the order requested.

### Grades workbook

The `Grades` worksheet has the following columns:

| Column | Meaning |
| --- | --- |
| `Assignment_ID` | Canvas assignment ID. |
| `Assignment_Name` | Assignment name returned by Canvas. |
| `Student_ID` | Canvas student user ID. |
| `Student_Name` | Student name returned with the submission. |
| `Student_Login_ID` | Student login ID returned with the submission. |
| `Score` | Assignment-level numeric score. |
| `Grade` | Assignment-level Canvas grade. |
| `Entered_Grade` | Assignment-level grade entered by the grader. |
| `Workflow_State` | Canvas submission workflow state. |
| `Grader_ID` | Canvas grader ID when provided. |
| `Grader_Name` | Resolved grader name or a status label. |
| `Grader_Login_ID` | Resolved grader login ID. |
| `Grader_SIS_ID` | Resolved grader SIS user ID. |
| `Grader_Type` | `human`, `automatic`, `ungraded`, or `unresolvable`. |

The `Marker Statistics` worksheet groups rows by all grader identity columns and contains `Submission_Count`, `Scored_Submission_Count`, `Unscored_Submission_Count`, `Average_Score`, `Minimum_Score`, and `Maximum_Score`. These are Excel formulas based on the `Grades` worksheet. The workbook is marked for recalculation when opened in Excel.

When `--rubric-assessment` is supplied, the workbook also has a `Rubric Assessments` worksheet. It contains one row for each criterion assessment returned by Canvas; assignments without a rubric or without returned assessments produce an empty sheet with headers.

| Column | Meaning |
| --- | --- |
| `Assignment_ID` | Canvas assignment ID. |
| `Assignment_Name` | Assignment name returned by Canvas. |
| `Student_ID` | Canvas student user ID. |
| `Student_Name` | Student name returned with the submission. |
| `Student_Login_ID` | Student login ID returned with the submission. |
| `Grader_ID` | Canvas grader ID when provided. |
| `Grader_Name` | Resolved grader name or a status label. |
| `Grader_Login_ID` | Resolved grader login ID. |
| `Grader_SIS_ID` | Resolved grader SIS user ID. |
| `Grader_Type` | `human`, `automatic`, `ungraded`, or `unresolvable`. |
| `Criterion_ID` | Canvas rubric criterion ID. |
| `Criterion_Description` | Rubric criterion description. |
| `Rating_ID` | Canvas ID of the selected rubric rating. |
| `Selected_Descriptor` | Description of the selected rating. |
| `Points` | Points awarded for the criterion. |
| `Comments` | Criterion-level marker comments. |

### Gradebook workbook

The controlled gradebook has `Instructions`, `Grades`, `Rubric`, `Rubric Catalog`, and `Canvas Mapping` sheets. It also contains hidden `Canvas Metadata`, `Canvas Snapshot`, and `Rubric Lists` sheets that support assignment verification, changed-value detection, and descriptor dropdowns. Do not remove or alter the hidden sheets.

`Canvas Mapping` has `Field`, `Sheet`, and `Column` columns. It maps standard logical fields such as `grades.new_grade`, `grades.submission_comment`, `rubric.selected_descriptor`, and `rubric.criterion_comment` to the visible column headers. This supports renamed columns without guessing at headers. It does not turn an arbitrary workbook into a gradebook: the generated metadata and baseline sheets remain required.

### Students workbook

The `Students` worksheet has these columns:

| Column | Canvas source or behavior |
| --- | --- |
| `Student` | User `name`. |
| `Email` | User `email`, with `login_id` as a fallback. |
| `Student Number` | User `integration_id`. |
| `ID` | Numeric user `sis_user_id`. |
| `Groups` | Matching course-group names, separated by `; `. |

## Gradebook safety

Always use `gradebook validate` or the default dry-run `gradebook upload` before adding `--execute`. The dry run prints each proposed student update without sending a Canvas mutation. The executed upload sends one Canvas grade-or-comment request per changed student and reports individual failures; Canvas cannot make the whole workbook transactional.

An executed upload creates a sidecar journal named `<workbook>.upload-journal.json` unless `--journal PATH` selects another location. Before each Canvas write, it records a durable pending hash without storing feedback text, then marks that hash applied or indeterminate after Canvas responds. A later execution skips the same pending, applied, or indeterminate request, preventing accidental duplicate comments after an interruption. The journal is scoped to each generated template, so a fresh template can intentionally submit the same value later. Use `--retry-indeterminate --execute` only after checking Canvas: it deliberately overrides the pending/indeterminate protection and can add a duplicate submission comment. Executed uploads also take a local lock beside the journal, so do not start another upload that uses the same journal while one is running.

Only enter literal values in editable gradebook cells. Formulas in mapped input fields are rejected. The `Instructions` sheet shows the assignment's grading type and points possible. `not_graded` assignments cannot accept `New_Grade`; pass/fail assignments accept only `pass`, `complete`, `fail`, `incomplete`, `0%`, `100%`, zero points, or the assignment's points possible. Select descriptors from the supplied dropdown whenever possible, and use a rating ID only when a compatible external mapping requires it. A selected descriptor supplies its Canvas default points unless `Points` provides an explicit numeric override. Free-form rubric scoring requires numeric `Points`.

Do not use a template after the assignment rubric changes. The workbook's rubric fingerprint causes validation to fail so the current rubric and descriptor choices can be regenerated.

## Cleanup safety

`cleanup --execute` is destructive and has no undo or rollback. A partial failure can leave the course partly cleaned. Always run the same target selection without `--execute` first, review the counts, and confirm that the course ID is correct.

The command does not delete assignment-linked calendar events through the `calendar` target; delete the related assignment through `assessments` if that is intended. Clearing `syllabus` only empties the editable syllabus body. Canvas's generated Course Summary can still contain dated items until the related assignments, discussions, or calendar events are removed.

## Canvas access

Canvas decides what a token can see and modify according to the token owner's role and course permissions. Export commands generally require permission to read the course, assignments, submissions, users, and groups involved. Grade exports may need permission to view submission grades, grader identities, rubrics, and rubric assessments. Gradebook template and validation require the same assignment and submission access; upload also requires permission to grade or comment on every targeted submission. Cleanup requires the corresponding delete or course-update permission for every selected target.

The client follows Canvas pagination links, so an export or cleanup includes all pages returned by the API, rather than only the first 100 records.

## Troubleshooting

| Symptom | Likely cause and action |
| --- | --- |
| `Canvas base URL not provided` | Set `--canvas-url`, `CANVAS_BASE_URL`, or `canvas_base_url` in the configuration file. |
| `Canvas API token not provided` | Set `--token`, `CANVAS_API_TOKEN`, or `canvas_api_token` in the configuration file. |
| `401 Unauthorized` | Verify the token, Canvas base URL, and token permissions. |
| `404` for an assignment or submissions | Verify the numeric course and assignment IDs and confirm that the token can access submissions for that assignment. |
| `404` while resolving a grader | The user may be deleted, hidden from the token, or only available in the course context. Grade data is still exported as `unresolvable` if both lookups return 404. |
| `--assignment must be a positive integer or 'all'` | Pass a numeric Canvas assignment ID or `all`. |
| Output-path error with `--assignment all` | Pass a directory rather than a `.csv` or `.xlsx` file path. |
| `Rubric Assessments` has no data rows | The assignment may not have a rubric, no criterion assessments may have been returned, or the token may not be able to view rubric assessments. |
| `Gradebook validation failed` | Correct the listed workbook action, identifier, descriptor, mapping, or Canvas conflict before retrying. |
| Rubric changed since template creation | Generate a fresh `gradebook template`, then re-enter or transfer the intended changes. |
| Another gradebook upload is already using lock | Wait for the other upload using this workbook's journal to finish. |
| Unsupported field or target | Use the field and target names listed in this guide. |

Run `python run.py --help` or `python run.py <subcommand> --help` to see the CLI's current arguments.
