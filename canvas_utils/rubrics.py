"""Shared helpers for working with Canvas assignment rubrics."""

from __future__ import annotations

from typing import Any


def extract_rubric_criteria(assignment: dict[str, Any]) -> list[dict[str, Any]]:
    """Return rubric criteria from a Canvas assignment payload."""
    rubric = assignment.get("rubric")
    if isinstance(rubric, list):
        return [criterion for criterion in rubric if isinstance(criterion, dict)]
    if not isinstance(rubric, dict):
        return []
    if "id" in rubric or "ratings" in rubric:
        return [rubric]
    return [criterion for criterion in rubric.values() if isinstance(criterion, dict)]


def index_rubric_criteria(rubric_criteria: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Index criteria by the identifier variants Canvas uses in assessments."""
    criteria_by_id: dict[str, dict[str, Any]] = {}
    for criterion in rubric_criteria:
        criterion_id = identifier_or_none(criterion.get("id"))
        if criterion_id is None:
            continue
        for key in criterion_id_aliases(criterion_id):
            criteria_by_id.setdefault(key, criterion)
    return criteria_by_id


def criterion_id_aliases(criterion_id: str) -> tuple[str, ...]:
    """Return the equivalent criterion key forms used by Canvas."""
    aliases = [criterion_id]
    if criterion_id.startswith("criterion_"):
        suffix = criterion_id.removeprefix("criterion_")
        if suffix:
            aliases.append(f"_{suffix}")
    elif criterion_id.startswith("_") and len(criterion_id) > 1:
        aliases.append(f"criterion_{criterion_id[1:]}")
    return tuple(dict.fromkeys(aliases))


def find_rubric_rating(
    criterion: dict[str, Any] | None,
    rating_id: str | None,
) -> dict[str, Any] | None:
    """Find a rating in a criterion by its Canvas rating identifier."""
    if criterion is None or rating_id is None:
        return None

    for rating in rubric_ratings(criterion):
        if identifier_or_none(rating.get("id")) == rating_id:
            return rating
    return None


def rubric_ratings(criterion: dict[str, Any]) -> list[dict[str, Any]]:
    """Return well-formed rating records from a Canvas rubric criterion."""
    ratings = criterion.get("ratings")
    if isinstance(ratings, list):
        return [rating for rating in ratings if isinstance(rating, dict)]
    if isinstance(ratings, dict):
        if "id" in ratings:
            return [ratings]
        return [rating for rating in ratings.values() if isinstance(rating, dict)]
    return []


def rubric_criterion_description(criterion: dict[str, Any] | None) -> str | None:
    """Return a criterion's short or long description."""
    if criterion is None:
        return None
    return string_or_none(criterion.get("description")) or string_or_none(criterion.get("long_description"))


def rubric_rating_description(rating: dict[str, Any] | None) -> str | None:
    """Return a rating's short or long descriptor."""
    if rating is None:
        return None
    return string_or_none(rating.get("description")) or string_or_none(rating.get("long_description"))


def identifier_or_none(value: object) -> str | None:
    """Normalize a Canvas identifier to a nonempty string."""
    if isinstance(value, str) and value.strip():
        return value.strip()
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return None


def string_or_none(value: object) -> str | None:
    """Return a nonempty string value, otherwise ``None``."""
    if isinstance(value, str) and value.strip():
        return value
    return None
