from __future__ import annotations


def should_show_high_risk_notice(risk_score: int) -> bool:
    return risk_score > 70


def should_show_restrictions(approval_status: str) -> bool:
    return approval_status.strip().lower() == "conditional"


def is_german_country(country: str) -> bool:
    return country.strip().upper() == "DE"
