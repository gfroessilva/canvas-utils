# canvas-utils

`canvas-utils` is a Python utility package for automating academic administration tasks in Canvas LMS.

This initial version implements **Submission Timestamp Extractor**, which exports submission metadata to CSV.

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

- `canvas_submissions_[assignment_id]_[assignment_name].csv`
