# canvas-utils

canvas-utils is a Python utility package for automating Canvas LMS course administration workflows.

## Documentation

- [User guide](docs/USER_GUIDE.md) - installation, authentication, complete command reference, output schemas, and cleanup safety.
- [Developer guide](docs/DEVELOPER_GUIDE.md) - architecture, Canvas API integration, module contracts, testing, and extension guidance.

## Table of Contents

- [Documentation](#documentation)
- [Requirements](#requirements)
- [Setup](#setup)
- [Credentials](#credentials)
- [Command Overview](#command-overview)
- [Subcommands](#subcommands)
  - [submissions](#submissions)
  - [grades](#grades)
  - [gradebook](#gradebook)
  - [students](#students)
  - [cleanup](#cleanup)
  - [quiz](#quiz)
- [Outputs](#outputs)
  - [Submissions Output](#submissions-output)
  - [Grades Output](#grades-output)
  - [Gradebook Output](#gradebook-output)
  - [Students Output](#students-output)
- [Examples](#examples)

## Requirements

- Python 3.10+
- Canvas base URL (for example https://canvas.example.edu)
- Canvas API token with permission to read course data; gradebook upload also requires grading permission

## Setup

<details open>
<summary>PowerShell</summary>

```powershell
python -m pip install -r requirements.txt
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 -m pip install -r requirements.txt
```

</details>

## Credentials

Credentials are resolved in this order:

1. CLI arguments (--canvas-url, --token)
2. Environment variables (CANVAS_BASE_URL, CANVAS_API_TOKEN)
3. Config file (default: ./config.json, override with --config)

Create config from template:

<details open>
<summary>PowerShell</summary>

```powershell
Copy-Item .\config.json.template .\config.json
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
cp ./config.json.template ./config.json
```

</details>

Current template keys:

- canvas_base_url
- canvas_api_token
- submission_fields (used by submissions)

## Command Overview

All commands are invoked through:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py <subcommand> [options]
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py <subcommand> [options]
```

</details>

Common options available on every subcommand:

- --canvas-url
- --token
- --config
- --timeout

## Subcommands

### submissions

Export submission timestamp and status data to CSV.

Options:

- --course (required)
- --assignment (default: all)
- --output (default: current directory)
- --fields (optional comma-separated extra fields)

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py submissions --course 12345 --assignment 67890 --output .\exports
python .\run.py submissions --course 12345 --assignment all --output .\exports
python .\run.py submissions --course 12345 --assignment 67890 --fields score,workflow_state,user_login_id
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py submissions --course 12345 --assignment 67890 --output ./exports
python3 ./run.py submissions --course 12345 --assignment all --output ./exports
python3 ./run.py submissions --course 12345 --assignment 67890 --fields score,workflow_state,user_login_id
```

</details>

### grades

Export assignment grades and grader identity to XLSX.

The `Grades` worksheet contains Canvas's assignment-level `score`, `grade`, and `entered_grade` values. Add `--rubric-assessment` to request each submission's `rubric_assessment` data and write a per-criterion `Rubric Assessments` worksheet with the selected descriptor.

Options:

- --course (required)
- --assignment (default: all)
- --output (default: current directory)
- --rubric-assessment (optional; includes criterion ratings and comments in a Rubric Assessments worksheet)

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py grades --course 12345 --assignment 67890 --output .\exports
python .\run.py grades --course 12345 --assignment all --output .\exports
python .\run.py grades --course 12345 --assignment 67890 --output .\exports --rubric-assessment
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py grades --course 12345 --assignment 67890 --output ./exports
python3 ./run.py grades --course 12345 --assignment all --output ./exports
python3 ./run.py grades --course 12345 --assignment 67890 --output ./exports --rubric-assessment
```

</details>

### gradebook

Create a controlled XLSX template for one assignment, validate its planned changes against Canvas, then upload only with explicit confirmation. The template supports assignment grades, rubric descriptor selection and points, criterion comments, and new submission comments.

Options:

- template: --course (required), --assignment (required), --output (optional)
- validate: --course (required), --assignment (required), --input (required)
- upload: --course (required), --assignment (required), --input (required), --execute (optional), --journal (optional), --retry-indeterminate (optional)

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py gradebook template --course 12345 --assignment 67890 --output .\exports\gradebook.xlsx
python .\run.py gradebook validate --course 12345 --assignment 67890 --input .\exports\gradebook.xlsx
python .\run.py gradebook upload --course 12345 --assignment 67890 --input .\exports\gradebook.xlsx
python .\run.py gradebook upload --course 12345 --assignment 67890 --input .\exports\gradebook.xlsx --execute
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py gradebook template --course 12345 --assignment 67890 --output ./exports/gradebook.xlsx
python3 ./run.py gradebook validate --course 12345 --assignment 67890 --input ./exports/gradebook.xlsx
python3 ./run.py gradebook upload --course 12345 --assignment 67890 --input ./exports/gradebook.xlsx
python3 ./run.py gradebook upload --course 12345 --assignment 67890 --input ./exports/gradebook.xlsx --execute
```

</details>

The template uses Canvas user IDs, explicit `SET`/`KEEP`/`ADD` actions, rubric descriptor dropdowns, and hidden metadata/snapshot sheets. `Canvas Mapping` lets renamed data columns be mapped explicitly. Upload validates the workbook and live Canvas state before a write, then rechecks each grade, rubric, or comment change immediately before sending it; comments target the validated submission attempt. Without `--execute`, it only reports the proposed per-student changes. Submission comments are additive.

Executed uploads create a `<workbook>.upload-journal.json` sidecar and local lock by default. The journal records a pending request before Canvas is contacted, then prevents that pending, successfully applied, or indeterminate request from being resent accidentally. Each generated template has a distinct journal identity, so a fresh template can still submit the same value later. `--retry-indeterminate --execute` explicitly overrides pending/indeterminate protection and can duplicate a submission comment.

### students

Export course student roster to XLSX.

The export includes all student enrollments returned by Canvas and the groups each student belongs to.

Options:

- --course (required)
- --output (default: current directory)
- --active-only (optional; limits export to active student enrollments)

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py students --course 12345 --output .\exports
python .\run.py students --course 12345 --output .\exports\roster.xlsx
python .\run.py students --course 12345 --output .\exports --active-only
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py students --course 12345 --output ./exports
python3 ./run.py students --course 12345 --output ./exports/roster.xlsx
python3 ./run.py students --course 12345 --output ./exports --active-only
```

</details>

### cleanup

Delete selected course resources progressively. By default, cleanup performs a dry run unless --execute is supplied.

Options:

- --course (required)
- --targets (default: all)
- --execute (optional flag)

Supported cleanup targets:

- pages
- assessments
- modules
- files
- discussions
- announcements
- syllabus
- calendar
- all

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py cleanup --course 12345 --targets pages,assessments,modules
python .\run.py cleanup --course 12345 --targets all --execute
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py cleanup --course 12345 --targets pages,assessments,modules
python3 ./run.py cleanup --course 12345 --targets all --execute
```

</details>

Note: the Canvas syllabus Course Summary may still contain dated course items unless related resources are also removed (for example assessments, discussions, calendar).

### quiz

Author, validate, and programmatically create Canvas quizzes from Excel or JSON specifications.

Actions:

- template: `--output` (default: `quiz_template.xlsx`)
- create: `--course` (required), `--input` (required; `.xlsx` or `.json`), `--publish` (optional; defaults to draft), `--dry-run` (optional)

Supported Question Types:

- `multiple_choice`: Single correct answer from multiple choices
- `true_false`: True or False question
- `multiple_answers`: Multiple correct answers (checkboxes)
- `short_answer`: Text input matching
- `essay`: Freeform text response

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
# 1. Generate an Excel or JSON quiz template
python .\run.py quiz template --output .\exports\quiz_template.xlsx
python .\run.py quiz template --output .\exports\quiz_template.json

# 2. Validate the quiz format without creating it in Canvas (dry-run)
python .\run.py quiz create --course 12345 --input .\exports\quiz_template.xlsx --dry-run
python .\run.py quiz create --course 12345 --input .\exports\quiz_template.json --dry-run

# 3. Create the quiz in Canvas
python .\run.py quiz create --course 12345 --input .\exports\quiz_template.xlsx

# 4. Create and publish immediately
python .\run.py quiz create --course 12345 --input .\exports\quiz_template.xlsx --publish
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
# 1. Generate an Excel or JSON quiz template
python3 ./run.py quiz template --output ./exports/quiz_template.xlsx
python3 ./run.py quiz template --output ./exports/quiz_template.json

# 2. Validate the quiz format without creating it in Canvas (dry-run)
python3 ./run.py quiz create --course 12345 --input ./exports/quiz_template.xlsx --dry-run
python3 ./run.py quiz create --course 12345 --input ./exports/quiz_template.json --dry-run

# 3. Create the quiz in Canvas
python3 ./run.py quiz create --course 12345 --input ./exports/quiz_template.xlsx

# 4. Create and publish immediately
python3 ./run.py quiz create --course 12345 --input ./exports/quiz_template.xlsx --publish
```

</details>

## Outputs

### Submissions Output

Base CSV columns:

- Assignment_ID
- Student_ID
- Attempt_Number
- Submitted_At
- Canvas_Late_Flag

Optional fields supported through --fields or submission_fields in config:

- student_name
- submitted_at_excel
- submitted_at_brisbane
- score
- grade
- entered_grade
- workflow_state
- missing
- excused
- late_policy_status
- posted_at
- seconds_late
- user_login_id
- user_sis_id
- user_short_name

Filename pattern:

- canvas_submissions_[assignment_id]_[assignment_name]_[export_date].csv

### Grades Output

XLSX worksheets:

- Grades
- Marker Statistics
- Rubric Assessments (optional; generated when --rubric-assessment is supplied)

Grades columns:

- Assignment_ID
- Assignment_Name
- Student_ID
- Student_Name
- Student_Login_ID
- Score
- Grade
- Entered_Grade
- Workflow_State
- Grader_ID
- Grader_Name
- Grader_Login_ID
- Grader_SIS_ID
- Grader_Type

Rubric Assessments columns:

- Assignment_ID
- Assignment_Name
- Student_ID
- Student_Name
- Student_Login_ID
- Grader_ID
- Grader_Name
- Grader_Login_ID
- Grader_SIS_ID
- Grader_Type
- Criterion_ID
- Criterion_Description
- Rating_ID
- Selected_Descriptor
- Points
- Comments

Marker Statistics summary columns:

- Submission_Count
- Scored_Submission_Count
- Unscored_Submission_Count
- Average_Score
- Minimum_Score
- Maximum_Score

Notes:

- Grader_Type values include human, automatic, ungraded, and unresolvable.
- Marker statistics are written as formulas and recalculate in Excel.
- Grade values on the `Grades` worksheet are assignment totals.
- The optional `Rubric Assessments` worksheet has one row per assessment returned by Canvas and includes criterion, selected descriptor, points, and comments.

Filename pattern:

- canvas_grades_[assignment_id]_[assignment_name]_[export_date].xlsx

### Gradebook Output

The generated gradebook contains `Grades`, `Rubric`, `Rubric Catalog`, `Canvas Mapping`, and `Instructions` sheets. `Grades` accepts an overall mark and an additive student comment. `Rubric` has one row per student and criterion, with a descriptor dropdown, optional points override, and criterion comment. Generate the workbook with `gradebook template`; do not use a plain `grades` export as an upload input.

The template also contains hidden metadata and baseline sheets. Upload rejects a workbook when its course, assignment, rubric, selected descriptor, or baseline grade/rubric assessment no longer matches Canvas. It deliberately does not clear existing grades or comments.

Filename pattern:

- canvas_gradebook_[assignment_id]_[assignment_name]_[export_date].xlsx

### Students Output

XLSX worksheet:

- Students

Students columns:

- Student
- Email
- Student Number
- ID
- Groups

Student Number contains the Canvas SIS user ID; ID contains the Canvas user ID. Email uses Canvas email when available, otherwise login_id fallback. Multiple group memberships are semicolon-separated.

Filename pattern:

- canvas_students_[course_id]_[export_date].xlsx

## Examples

Using environment variables:

<details open>
<summary>PowerShell</summary>

```powershell
$env:CANVAS_BASE_URL="https://canvas.example.edu"
$env:CANVAS_API_TOKEN="YOUR_TOKEN"
python .\run.py students --course 12345 --output .\exports
python .\run.py submissions --course 12345 --assignment all --output .\exports
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
export CANVAS_BASE_URL="https://canvas.example.edu"
export CANVAS_API_TOKEN="YOUR_TOKEN"
python3 ./run.py students --course 12345 --output ./exports
python3 ./run.py submissions --course 12345 --assignment all --output ./exports
```

</details>
