"""Utilities for Canvas LMS automation workflows."""

from canvas_utils.quizzes import (
    AnswerSpec,
    QuestionSpec,
    QuizSpec,
    QuizValidationError,
    create_quiz_template,
    create_quiz_template_excel,
    create_quiz_template_json,
    deploy_quiz,
    load_quiz_from_dict,
    load_quiz_from_excel,
    load_quiz_from_json,
)

__all__ = [
    "__version__",
    "AnswerSpec",
    "QuestionSpec",
    "QuizSpec",
    "QuizValidationError",
    "create_quiz_template",
    "create_quiz_template_excel",
    "create_quiz_template_json",
    "deploy_quiz",
    "load_quiz_from_dict",
    "load_quiz_from_excel",
    "load_quiz_from_json",
]

__version__ = "0.1.0"
