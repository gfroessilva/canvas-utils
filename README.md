# canvas-utils

`canvas-utils` is a Python utility package for automating academic administration tasks in Canvas LMS.

This initial version implements **Submission Timestamp Extractor**, which exports submission metadata to CSV.

It also includes a **Grades Exporter**, which exports assignment marks to XLSX together with grader identity where Canvas exposes it.

## Requirements

- Python 3.10+
- Canvas API token with permission to read course assignments and submissions
- Canvas base URL (for example `https://canvas.example.edu`)

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Credentials

Credentials are resolved in this order:

1. CLI arguments (`--canvas-url`, `--token`)
2. Environment variables (`CANVAS_BASE_URL`, `CANVAS_API_TOKEN`) **(default/recommended)**
3. `config.json` (or another file via `--config`)

You can start from:

```powershell
Copy-Item .\config.json.template .\config.json
```

## Usage

### Submissions extractor

Single assignment:

```powershell
python .\run.py submissions --course 12345 --assignment 67890 --output .\exports
```

All assignments in a course:

```powershell
python .\run.py submissions --course 12345 --assignment all --output .\exports
```

Using environment variables:

```powershell
$env:CANVAS_BASE_URL="https://canvas.example.edu"
$env:CANVAS_API_TOKEN="YOUR_TOKEN"
python .\run.py submissions --course 12345 --assignment all --output .\exports
```

Including optional fields:

```powershell
python .\run.py submissions --course 12345 --assignment 67890 --fields score,workflow_state,user_login_id
```

### Course cleanup (progressive deletion)

Dry run (recommended first):

```powershell
python .\run.py cleanup --course 12345 --targets pages,assessments,modules
```

Execute deletion:

```powershell
python .\run.py cleanup --course 12345 --targets all --execute
```

### Grades exporter

Single assignment:

```powershell
python .\run.py grades --course 12345 --assignment 67890 --output .\exports
```

All assignments in a course:

```powershell
python .\run.py grades --course 12345 --assignment all --output .\exports
```

Cleanup targets:

- `pages`
- `assessments`
- `modules`
- `files`
- `discussions`
- `announcements`
- `syllabus` (clears syllabus body content)
- `calendar` (deletes standalone calendar events in the course context)
- `all`

Note: the Syllabus **Course Summary** includes dated course items (for example assignments). To fully clear it, use the relevant cleanup targets too (such as `assessments`, `discussions`, and `calendar`).

## Output

The extractor writes CSV file(s) with the schema:

- `Assignment_ID`
- `Student_ID`
- `Attempt_Number`
- `Submitted_At`
- `Canvas_Late_Flag`

Optional fields can be added with `--fields` or `submission_fields` in `config.json`. Supported values include `student_name`, `submitted_at_excel`, `submitted_at_brisbane`, `score`, `grade`, `entered_grade`, `workflow_state`, `missing`, `excused`, `late_policy_status`, `posted_at`, `seconds_late`, `user_login_id`, `user_sis_id`, and `user_short_name`.

Filename format:

- `canvas_submissions_[assignment_id]_[assignment_name]_[export_date].csv`

The grades exporter writes XLSX file(s) with:

- a `Grades` worksheet with the schema:

- `Assignment_ID`
- `Assignment_Name`
- `Student_ID`
- `Student_Name`
- `Student_Login_ID`
- `Score`
- `Grade`
- `Entered_Grade`
- `Workflow_State`
- `Grader_ID`
- `Grader_Name`
- `Grader_Login_ID`
- `Grader_SIS_ID`
- `Grader_Type`

- a `Marker Statistics` worksheet with per-marker summary rows:
  - `Submission_Count`
  - `Scored_Submission_Count`
  - `Unscored_Submission_Count`
  - `Average_Score`
  - `Minimum_Score`
  - `Maximum_Score`

Notes:

- `Grader_Type` is `human`, `automatic`, or `ungraded`.
- The exporter first resolves graders through the global Users API, then falls back to a course-scoped user lookup before marking a grader as `unresolvable`.
- Canvas may hide grader identity for anonymous grading.
- Auto-graded submissions can expose a non-human grader identifier.
- Marker statistics are written as Excel formulas so they recalculate if you edit the `Grades` sheet manually.
- `Average_Score` is rounded to 2 decimal places when needed.

Filename format:

- `canvas_grades_[assignment_id]_[assignment_name]_[export_date].xlsx`
