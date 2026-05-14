from __future__ import annotations

from datetime import date
import re


def format_date(value: date | str) -> str:
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def format_percent(value: float | None, decimals: int = 1) -> str:
    if value is None:
        return "N/A"
    return f"{value:.{decimals}f}%"


def format_dose(value: float, unit: str) -> str:
    return f"{value:g} {unit}"


def format_concentration(value: float, unit: str) -> str:
    return f"{value:g} {unit}"


def format_classification(value: str) -> str:
    return value.strip().upper()


def format_metric_name(value: str) -> str:
    return value.replace("_", " ").title()


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", value.strip().lower()).strip("-")
    return slug or "item"
