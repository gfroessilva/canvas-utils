# canvas-utils

canvas-utils is a Python utility package for automating Canvas LMS course administration workflows.

## Table of Contents

- [Requirements](#requirements)
- [Setup](#setup)
- [Credentials](#credentials)
- [Command Overview](#command-overview)
- [Subcommands](#subcommands)
  - [submissions](#submissions)
  - [grades](#grades)
  - [students](#students)
  - [cleanup](#cleanup)
- [Outputs](#outputs)
  - [Submissions Output](#submissions-output)
  - [Grades Output](#grades-output)
  - [Students Output](#students-output)
- [Examples](#examples)

## Requirements

- Python 3.10+
- Canvas base URL (for example https://canvas.example.edu)
- Canvas API token with permission to read course data

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

Options:

- --course (required)
- --assignment (default: all)
- --output (default: current directory)

Examples:

<details open>
<summary>PowerShell</summary>

```powershell
python .\run.py grades --course 12345 --assignment 67890 --output .\exports
python .\run.py grades --course 12345 --assignment all --output .\exports
```

</details>

<details>
<summary>Linux/macOS (bash)</summary>

```bash
python3 ./run.py grades --course 12345 --assignment 67890 --output ./exports
python3 ./run.py grades --course 12345 --assignment all --output ./exports
```

</details>

### students

Export course student roster to XLSX.

The export includes all student enrollments returned by Canvas.

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

Filename pattern:

- canvas_grades_[assignment_id]_[assignment_name]_[export_date].xlsx

### Students Output

XLSX worksheet:

- Students

Students columns:

- Student
- Email
- Student Number
- ID

Email uses Canvas email when available, otherwise login_id fallback.

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
