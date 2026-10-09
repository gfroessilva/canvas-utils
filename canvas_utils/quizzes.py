"""Data structures, parsers, and deployment workflows for Canvas Quizzes."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from canvas_utils.api_client import CanvasAPIError, CanvasClient

SUPPORTED_QUESTION_TYPES = {
    "multiple_choice_question": "multiple_choice_question",
    "multiple_choice": "multiple_choice_question",
    "true_false_question": "true_false_question",
    "true_false": "true_false_question",
    "multiple_answers_question": "multiple_answers_question",
    "multiple_answers": "multiple_answers_question",
    "short_answer_question": "short_answer_question",
    "short_answer": "short_answer_question",
    "essay_question": "essay_question",
    "essay": "essay_question",
    "numerical_question": "numerical_question",
    "numerical": "numerical_question",
}

VALID_QUIZ_TYPES = {
    "practice_quiz",
    "assignment",
    "graded_survey",
    "survey",
}


class QuizValidationError(ValueError):
    """Raised when quiz specification or questions fail validation."""


@dataclass
class AnswerSpec:
    """Specification for an individual question answer choice."""

    text: str
    weight: float = 0.0
    comments: str | None = None
    numerical_answer_type: str | None = None
    exact_value: float | None = None
    margin: float | None = None

    def to_api_payload(self) -> dict[str, Any]:
        """Convert answer spec into Canvas question answer payload."""
        payload: dict[str, Any] = {
            "answer_text": self.text,
            "answer_weight": self.weight,
        }
        if self.comments:
            payload["answer_comment"] = self.comments
        if self.numerical_answer_type:
            payload["numerical_answer_type"] = self.numerical_answer_type
            if self.exact_value is not None:
                payload["answer_exact"] = self.exact_value
            if self.margin is not None:
                payload["answer_error_margin"] = self.margin
        return payload


@dataclass
class QuestionSpec:
    """Specification for an individual quiz question."""

    name: str
    text: str
    question_type: str
    points_possible: float = 1.0
    answers: list[AnswerSpec] = field(default_factory=list)
    general_feedback: str | None = None
    correct_comments: str | None = None
    incorrect_comments: str | None = None

    def validate(self, index: int = 1) -> None:
        """Validate question fields."""
        prefix = f"Question #{index} ('{self.name}')"
        if not self.name.strip():
            raise QuizValidationError(f"{prefix}: Question name cannot be blank.")
        if not self.text.strip():
            raise QuizValidationError(f"{prefix}: Question text cannot be blank.")

        norm_type = SUPPORTED_QUESTION_TYPES.get(self.question_type.lower())
        if not norm_type:
            raise QuizValidationError(
                f"{prefix}: Unsupported question type '{self.question_type}'. "
                f"Supported types: {', '.join(sorted(set(SUPPORTED_QUESTION_TYPES.values())))}"
            )

        if self.points_possible < 0:
            raise QuizValidationError(f"{prefix}: Points possible must be non-negative (got {self.points_possible}).")

        if norm_type == "multiple_choice_question":
            if len(self.answers) < 2:
                raise QuizValidationError(f"{prefix}: Multiple choice questions must have at least 2 answer options.")
            correct_answers = [ans for ans in self.answers if ans.weight > 0]
            if len(correct_answers) != 1:
                raise QuizValidationError(
                    f"{prefix}: Multiple choice question must have exactly 1 correct answer (weight > 0), "
                    f"found {len(correct_answers)}."
                )

        elif norm_type == "true_false_question":
            if len(self.answers) != 2:
                raise QuizValidationError(f"{prefix}: True/False questions must have exactly 2 options.")
            correct_answers = [ans for ans in self.answers if ans.weight > 0]
            if len(correct_answers) != 1:
                raise QuizValidationError(f"{prefix}: True/False questions must have exactly 1 correct option.")

        elif norm_type == "multiple_answers_question":
            if len(self.answers) < 2:
                raise QuizValidationError(f"{prefix}: Multiple answers questions must have at least 2 options.")
            correct_answers = [ans for ans in self.answers if ans.weight > 0]
            if not correct_answers:
                raise QuizValidationError(f"{prefix}: Multiple answers question must have at least 1 correct answer.")

        elif norm_type == "short_answer_question":
            if not self.answers:
                raise QuizValidationError(f"{prefix}: Short answer questions must have at least one acceptable answer.")

    def to_api_payload(self) -> dict[str, Any]:
        """Convert question spec into Canvas question payload."""
        norm_type = SUPPORTED_QUESTION_TYPES[self.question_type.lower()]
        payload: dict[str, Any] = {
            "question_name": self.name,
            "question_text": self.text,
            "question_type": norm_type,
            "points_possible": self.points_possible,
            "answers": [ans.to_api_payload() for ans in self.answers],
        }
        if self.general_feedback:
            payload["neutral_comments"] = self.general_feedback
        if self.correct_comments:
            payload["correct_comments"] = self.correct_comments
        if self.incorrect_comments:
            payload["incorrect_comments"] = self.incorrect_comments
        return payload


@dataclass
class QuizSpec:
    """Specification for an entire Canvas Quiz."""

    title: str
    description: str = ""
    quiz_type: str = "assignment"  # or 'practice_quiz'
    time_limit: int | None = None  # in minutes
    shuffle_answers: bool = True
    allowed_attempts: int = 1  # -1 for unlimited
    scoring_policy: str = "keep_highest"  # 'keep_highest', 'keep_latest'
    show_correct_answers: bool = True
    published: bool = False
    due_at: str | None = None
    lock_at: str | None = None
    unlock_at: str | None = None
    questions: list[QuestionSpec] = field(default_factory=list)

    def validate(self) -> None:
        """Validate quiz and all underlying questions."""
        if not self.title.strip():
            raise QuizValidationError("Quiz title cannot be blank.")
        if self.quiz_type not in VALID_QUIZ_TYPES:
            raise QuizValidationError(
                f"Invalid quiz_type '{self.quiz_type}'. Must be one of: {', '.join(sorted(VALID_QUIZ_TYPES))}"
            )
        if not self.questions:
            raise QuizValidationError("Quiz must contain at least one question.")

        for idx, question in enumerate(self.questions, start=1):
            question.validate(index=idx)

    @property
    def total_points(self) -> float:
        """Calculate total points across all questions."""
        return sum(q.points_possible for q in self.questions)

    def to_api_payload(self) -> dict[str, Any]:
        """Convert quiz metadata to Canvas quiz creation payload."""
        payload: dict[str, Any] = {
            "title": self.title,
            "description": self.description,
            "quiz_type": self.quiz_type,
            "shuffle_answers": self.shuffle_answers,
            "allowed_attempts": self.allowed_attempts,
            "scoring_policy": self.scoring_policy,
            "published": self.published,
        }
        if self.time_limit is not None:
            payload["time_limit"] = self.time_limit
        if self.show_correct_answers is not None:
            payload["show_correct_answers"] = self.show_correct_answers
        if self.due_at:
            payload["due_at"] = self.due_at
        if self.lock_at:
            payload["lock_at"] = self.lock_at
        if self.unlock_at:
            payload["unlock_at"] = self.unlock_at
        return payload


def load_quiz_from_dict(data: dict[str, Any]) -> QuizSpec:
    """Parse and validate a dictionary into a QuizSpec."""
    questions_data = data.get("questions") or []
    questions: list[QuestionSpec] = []

    for q_dict in questions_data:
        answers_data = q_dict.get("answers") or []
        answers = [
            AnswerSpec(
                text=str(a.get("text", "")),
                weight=float(a.get("weight", 0.0)),
                comments=a.get("comments"),
                numerical_answer_type=a.get("numerical_answer_type"),
                exact_value=float(a["exact_value"]) if "exact_value" in a and a["exact_value"] is not None else None,
                margin=float(a["margin"]) if "margin" in a and a["margin"] is not None else None,
            )
            for a in answers_data
        ]
        questions.append(
            QuestionSpec(
                name=str(q_dict.get("name", "Question")),
                text=str(q_dict.get("text", "")),
                question_type=str(q_dict.get("type") or q_dict.get("question_type", "multiple_choice_question")),
                points_possible=float(q_dict.get("points_possible", 1.0)),
                answers=answers,
                general_feedback=q_dict.get("general_feedback"),
                correct_comments=q_dict.get("correct_comments"),
                incorrect_comments=q_dict.get("incorrect_comments"),
            )
        )

    quiz = QuizSpec(
        title=str(data.get("title", "")),
        description=str(data.get("description", "")),
        quiz_type=str(data.get("quiz_type", "assignment")),
        time_limit=int(data["time_limit"]) if data.get("time_limit") is not None else None,
        shuffle_answers=bool(data.get("shuffle_answers", True)),
        allowed_attempts=int(data.get("allowed_attempts", 1)),
        scoring_policy=str(data.get("scoring_policy", "keep_highest")),
        show_correct_answers=bool(data.get("show_correct_answers", True)),
        published=bool(data.get("published", False)),
        due_at=data.get("due_at"),
        lock_at=data.get("lock_at"),
        unlock_at=data.get("unlock_at"),
        questions=questions,
    )
    quiz.validate()
    return quiz


def load_quiz_from_json(path: str | Path) -> QuizSpec:
    """Load and validate a quiz specification from a JSON file."""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"Quiz JSON file not found: {p}")
    with p.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise QuizValidationError("Root of quiz JSON must be a JSON object/dictionary.")
    return load_quiz_from_dict(data)


def load_quiz_from_excel(path: str | Path) -> QuizSpec:
    """Load and validate a quiz specification from an Excel workbook (.xlsx)."""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"Quiz Excel file not found: {p}")

    wb = load_workbook(p, data_only=True)

    # 1. Parse Settings sheet
    settings: dict[str, Any] = {}
    if "Settings" in wb.sheetnames:
        settings_sheet = wb["Settings"]
        for row in settings_sheet.iter_rows(min_row=1, values_only=True):
            if not row or not row[0]:
                continue
            key = str(row[0]).strip().lower()
            val = row[1] if len(row) > 1 else None
            settings[key] = val

    title = str(settings.get("title") or p.stem)
    description = str(settings.get("description") or "")
    quiz_type = str(settings.get("quiz_type") or "assignment")
    time_limit = int(settings["time_limit"]) if settings.get("time_limit") is not None and str(settings["time_limit"]).strip() != "" else None
    
    shuffle_val = settings.get("shuffle_answers")
    shuffle_answers = True if shuffle_val is None else (str(shuffle_val).strip().upper() not in {"FALSE", "0", "NO"})

    attempts_val = settings.get("allowed_attempts")
    allowed_attempts = int(attempts_val) if attempts_val is not None and str(attempts_val).strip() != "" else 1

    scoring_policy = str(settings.get("scoring_policy") or "keep_highest")
    
    show_ans_val = settings.get("show_correct_answers")
    show_correct_answers = True if show_ans_val is None else (str(show_ans_val).strip().upper() not in {"FALSE", "0", "NO"})

    pub_val = settings.get("published")
    published = False if pub_val is None else (str(pub_val).strip().upper() in {"TRUE", "1", "YES"})

    # 2. Parse Questions sheet
    if "Questions" not in wb.sheetnames:
        raise QuizValidationError("Excel workbook must contain a 'Questions' worksheet.")

    questions_sheet = wb["Questions"]
    rows = list(questions_sheet.iter_rows(values_only=True))
    if not rows:
        raise QuizValidationError("The 'Questions' worksheet is empty.")

    header = [str(col).strip().lower() if col is not None else "" for col in rows[0]]
    header_idx = {name: i for i, name in enumerate(header) if name}

    def get_cell(row_vals: tuple[Any, ...], col_name: str) -> Any:
        idx = header_idx.get(col_name.lower())
        if idx is not None and idx < len(row_vals):
            return row_vals[idx]
        return None

    questions: list[QuestionSpec] = []
    for r_idx, row_vals in enumerate(rows[1:], start=2):
        if not any(row_vals):
            continue

        q_name = get_cell(row_vals, "question_name") or get_cell(row_vals, "name") or f"Question {r_idx - 1}"
        q_text = get_cell(row_vals, "question_text") or get_cell(row_vals, "text") or ""
        q_type_raw = str(get_cell(row_vals, "question_type") or get_cell(row_vals, "type") or "multiple_choice").strip()
        points_raw = get_cell(row_vals, "points") or get_cell(row_vals, "points_possible") or 1.0
        try:
            points = float(points_raw)
        except (ValueError, TypeError):
            points = 1.0

        general_feedback = get_cell(row_vals, "general_feedback") or get_cell(row_vals, "feedback")
        correct_answer_raw = str(get_cell(row_vals, "correct_answer") or "").strip()

        # Gather choice columns: Choice_A through Choice_Z or Option 1...
        choices: list[tuple[str, str]] = []  # (label, text)
        for letter in "ABCDEFGH":
            val = get_cell(row_vals, f"choice_{letter}") or get_cell(row_vals, f"option_{letter}")
            if val is not None and str(val).strip():
                choices.append((letter, str(val).strip()))

        answers: list[AnswerSpec] = []
        norm_type = SUPPORTED_QUESTION_TYPES.get(q_type_raw.lower(), "multiple_choice_question")

        if norm_type == "true_false_question":
            is_true_correct = correct_answer_raw.upper() in {"TRUE", "T", "YES", "1"}
            answers = [
                AnswerSpec(text="True", weight=100.0 if is_true_correct else 0.0),
                AnswerSpec(text="False", weight=0.0 if is_true_correct else 100.0),
            ]
        elif norm_type == "multiple_choice_question":
            target_key = correct_answer_raw.upper()
            for letter, c_text in choices:
                is_correct = (letter == target_key or c_text.strip().lower() == correct_answer_raw.lower())
                answers.append(AnswerSpec(text=c_text, weight=100.0 if is_correct else 0.0))
        elif norm_type == "multiple_answers_question":
            correct_keys = {k.strip().upper() for k in correct_answer_raw.replace(";", ",").split(",") if k.strip()}
            num_correct = len(correct_keys)
            weight_per_correct = (100.0 / num_correct) if num_correct > 0 else 0.0
            for letter, c_text in choices:
                is_correct = (letter in correct_keys or c_text.strip().lower() in {k.lower() for k in correct_keys})
                answers.append(AnswerSpec(text=c_text, weight=weight_per_correct if is_correct else 0.0))
        elif norm_type == "short_answer_question":
            acceptable = [ans.strip() for ans in correct_answer_raw.replace(";", ",").split(",") if ans.strip()]
            for acc in acceptable:
                answers.append(AnswerSpec(text=acc, weight=100.0))
        elif norm_type == "essay_question":
            answers = []

        questions.append(
            QuestionSpec(
                name=str(q_name),
                text=str(q_text),
                question_type=norm_type,
                points_possible=points,
                answers=answers,
                general_feedback=str(general_feedback) if general_feedback else None,
            )
        )

    quiz = QuizSpec(
        title=title,
        description=description,
        quiz_type=quiz_type,
        time_limit=time_limit,
        shuffle_answers=shuffle_answers,
        allowed_attempts=allowed_attempts,
        scoring_policy=scoring_policy,
        show_correct_answers=show_correct_answers,
        published=published,
        questions=questions,
    )
    quiz.validate()
    return quiz


def create_quiz_template_excel(output_path: str | Path) -> Path:
    """Generate a starter Excel workbook template for authoring quizzes."""
    wb = Workbook()

    # Sheet 1: Settings
    ws_settings = wb.active
    ws_settings.title = "Settings"

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")

    ws_settings.append(["Setting", "Value", "Notes"])
    for cell in ws_settings[1]:
        cell.font = header_font
        cell.fill = header_fill

    settings_data = [
        ("title", "Sample Module Quiz", "Name of the quiz shown in Canvas"),
        ("description", "<p>Please answer the following questions.</p>", "HTML or plain text quiz instructions"),
        ("quiz_type", "assignment", "assignment (graded) or practice_quiz (self-evaluation)"),
        ("time_limit", "", "Time limit in minutes (leave empty for untimed)"),
        ("shuffle_answers", "TRUE", "TRUE to randomize answer choices, FALSE to keep order"),
        ("allowed_attempts", "1", "Number of attempts, or -1 for unlimited attempts"),
        ("scoring_policy", "keep_highest", "keep_highest or keep_latest"),
        ("show_correct_answers", "TRUE", "Whether students see correct answers after submission"),
        ("published", "FALSE", "Set to TRUE to publish immediately upon creation"),
    ]
    for row in settings_data:
        ws_settings.append(row)

    # Sheet 2: Questions
    ws_q = wb.create_sheet(title="Questions")
    q_headers = [
        "Question_Name",
        "Question_Type",
        "Points",
        "Question_Text",
        "Choice_A",
        "Choice_B",
        "Choice_C",
        "Choice_D",
        "Correct_Answer",
        "General_Feedback",
    ]
    ws_q.append(q_headers)
    for cell in ws_q[1]:
        cell.font = header_font
        cell.fill = header_fill

    sample_questions = [
        (
            "Q1",
            "multiple_choice",
            1.0,
            "What is the capital of France?",
            "Berlin",
            "Madrid",
            "Paris",
            "Rome",
            "C",
            "Paris has been the capital since 508 AD.",
        ),
        (
            "Q2",
            "true_false",
            1.0,
            "Python supports both object-oriented and functional programming paradigms.",
            "",
            "",
            "",
            "",
            "True",
            "Python is multi-paradigm.",
        ),
        (
            "Q3",
            "multiple_answers",
            2.0,
            "Which of the following are prime numbers?",
            "2",
            "4",
            "5",
            "9",
            "A, C",
            "2 and 5 are primes. 4 and 9 are composite.",
        ),
        (
            "Q4",
            "short_answer",
            1.0,
            "What keyword defines a function in Python?",
            "",
            "",
            "",
            "",
            "def",
            "The 'def' keyword introduces a function definition.",
        ),
        (
            "Q5",
            "essay",
            5.0,
            "Explain the difference between overfitting and underfitting in machine learning.",
            "",
            "",
            "",
            "",
            "",
            "Grade manually according to the rubric.",
        ),
    ]
    for row in sample_questions:
        ws_q.append(row)

    # Adjust column widths
    for sheet in (ws_settings, ws_q):
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 45)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def create_quiz_template_json(output_path: str | Path) -> Path:
    """Generate a starter JSON template for authoring quizzes."""
    template_data: dict[str, Any] = {
        "title": "Sample Module Quiz",
        "description": "<p>Please answer the following questions to test your knowledge.</p>",
        "quiz_type": "assignment",
        "time_limit": None,
        "shuffle_answers": True,
        "allowed_attempts": 1,
        "scoring_policy": "keep_highest",
        "show_correct_answers": True,
        "published": False,
        "questions": [
            {
                "name": "Q1: Capital City",
                "text": "<p>What is the capital of France?</p>",
                "type": "multiple_choice_question",
                "points_possible": 1.0,
                "general_feedback": "Paris has been the capital since 508 AD.",
                "answers": [
                    {"text": "Berlin", "weight": 0, "comments": "Incorrect. Capital of Germany."},
                    {"text": "Madrid", "weight": 0, "comments": "Incorrect. Capital of Spain."},
                    {"text": "Paris", "weight": 100, "comments": "Correct! Paris is the capital of France."},
                    {"text": "Rome", "weight": 0, "comments": "Incorrect. Capital of Italy."}
                ]
            },
            {
                "name": "Q2: Programming Paradigms",
                "text": "<p>Python supports both object-oriented and functional programming paradigms.</p>",
                "type": "true_false_question",
                "points_possible": 1.0,
                "general_feedback": "Python is multi-paradigm.",
                "answers": [
                    {"text": "True", "weight": 100, "comments": "Correct. Python supports multiple paradigms."},
                    {"text": "False", "weight": 0, "comments": "Incorrect. Python is multi-paradigm."}
                ]
            },
            {
                "name": "Q3: Prime Numbers",
                "text": "<p>Which of the following numbers are prime?</p>",
                "type": "multiple_answers_question",
                "points_possible": 2.0,
                "general_feedback": "Prime numbers are only divisible by 1 and themselves.",
                "answers": [
                    {"text": "2", "weight": 50, "comments": "Correct. 2 is the smallest prime."},
                    {"text": "4", "weight": 0, "comments": "Incorrect. 4 is divisible by 2."},
                    {"text": "5", "weight": 50, "comments": "Correct. 5 is a prime."},
                    {"text": "9", "weight": 0, "comments": "Incorrect. 9 is divisible by 3."}
                ]
            },
            {
                "name": "Q4: Function Keyword",
                "text": "<p>What keyword defines a function in Python?</p>",
                "type": "short_answer_question",
                "points_possible": 1.0,
                "general_feedback": "The 'def' keyword introduces a function definition.",
                "answers": [
                    {"text": "def", "weight": 100}
                ]
            },
            {
                "name": "Q5: Machine Learning Concepts",
                "text": "<p>Explain the difference between overfitting and underfitting in machine learning.</p>",
                "type": "essay_question",
                "points_possible": 5.0,
                "general_feedback": "Assessed manually against the grading rubric.",
                "answers": []
            }
        ]
    }
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        json.dump(template_data, f, indent=2)
    return out


def create_quiz_template(output_path: str | Path) -> Path:
    """Generate a starter template (JSON or Excel based on file extension)."""
    p = Path(output_path)
    if p.suffix.lower() == ".json":
        return create_quiz_template_json(p)
    if p.suffix.lower() in {".xlsx", ".xlsm"}:
        return create_quiz_template_excel(p)
    # Default to xlsx if no recognized extension
    if not p.suffix:
        return create_quiz_template_excel(p.with_suffix(".xlsx"))
    raise ValueError(f"Unsupported template format '{p.suffix}'. Use .xlsx or .json.")


def deploy_quiz(
    client: CanvasClient,
    course_id: int,
    quiz_spec: QuizSpec,
    *,
    publish: bool = False,
) -> dict[str, Any]:
    """Create a quiz and all its questions in Canvas via the Classic Quizzes REST API."""
    quiz_spec.validate()

    quiz_payload = quiz_spec.to_api_payload()
    if publish:
        quiz_payload["published"] = True

    # 1. Create quiz container
    created_quiz = client.create_quiz(course_id, quiz_payload)
    quiz_id = created_quiz.get("id")
    if not quiz_id:
        raise CanvasAPIError("Failed to obtain quiz ID from Canvas response.")

    created_questions: list[dict[str, Any]] = []

    # 2. Upload questions sequentially
    for q_idx, q_spec in enumerate(quiz_spec.questions, start=1):
        q_payload = q_spec.to_api_payload()
        created_q = client.create_quiz_question(course_id, quiz_id, q_payload)
        created_questions.append(created_q)

    return {
        "quiz": created_quiz,
        "questions": created_questions,
        "quiz_id": quiz_id,
        "html_url": created_quiz.get("html_url"),
        "total_points": quiz_spec.total_points,
        "question_count": len(created_questions),
    }
