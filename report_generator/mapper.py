from __future__ import annotations

from pathlib import Path
from statistics import mean

from .formatting import (
    format_classification,
    format_concentration,
    format_date,
    format_dose,
    format_metric_name,
    format_percent,
)
from .models import Application, Product, ReportInput
from .rules import is_german_country, should_show_high_risk_notice, should_show_restrictions


def calculate_product_average_efficacy(applications: list[Application]) -> float | None:
    if not applications:
        return None
    return mean(application.efficacy_percent for application in applications)


def calculate_overall_average_efficacy(products: list[Product]) -> float | None:
    product_averages = [
        average
        for average in (calculate_product_average_efficacy(product.applications) for product in products)
        if average is not None
    ]
    if not product_averages:
        return None
    return mean(product_averages)


def map_report(report: ReportInput, chart_paths: dict[str, Path | None]) -> dict:
    metadata = report.report_metadata
    products = [_map_product(product, chart_paths.get(product.product_id)) for product in report.products]
    conditional_count = sum(1 for product in report.products if should_show_restrictions(product.approval_status))
    high_risk_count = sum(1 for product in report.products if should_show_high_risk_notice(product.risk_score))

    return {
        "metadata": {
            "report_id": metadata.report_id,
            "report_title": metadata.report_title,
            "issuing_authority": metadata.issuing_authority,
            "issued_on": format_date(metadata.issued_on),
            "language_locale": metadata.language_locale,
            "version": metadata.version,
            "classification": format_classification(metadata.classification),
        },
        "summary": {
            "product_count": len(report.products),
            "conditional_count": conditional_count,
            "high_risk_count": high_risk_count,
            "overall_average_efficacy": format_percent(calculate_overall_average_efficacy(report.products)),
            "summary_text": _format_summary_text(report.products, conditional_count),
        },
        "products": products,
        "footer": {
            "classification": format_classification(report.footer.classification),
            "disclaimer": report.footer.disclaimer,
        },
    }


def _map_product(product: Product, chart_path: Path | None) -> dict:
    is_german = is_german_country(product.country)
    show_restrictions = should_show_restrictions(product.approval_status)
    restrictions = product.restrictions or ["No restrictions provided for this conditional approval."]

    return {
        "product_id": product.product_id,
        "name": product.product_name,
        "active_substance": product.active_substance,
        "concentration": format_concentration(
            product.active_substance_concentration.value,
            product.active_substance_concentration.unit,
        ),
        "formulation": f"{product.formulation_type} - {product.formulation_type_long}",
        "formulation_display": f"{product.formulation_type_long} ({product.formulation_type})",
        "manufacturer": product.manufacturer,
        "country": product.country,
        "approval_status": product.approval_status.value,
        "approval_line": _format_approval_line(product, is_german),
        "approval_summary": _format_approval_summary(product, is_german),
        "risk_score": product.risk_score,
        "show_high_risk_notice": should_show_high_risk_notice(product.risk_score),
        "high_risk_notice": _format_high_risk_notice(product, is_german),
        "show_restrictions": show_restrictions,
        "restrictions_heading": "Auflagen und Beschrankungen" if is_german else "Restrictions",
        "restrictions": restrictions,
        "application_count": len(product.applications),
        "average_efficacy": format_percent(calculate_product_average_efficacy(product.applications)),
        "registration_summary": _format_registration_summary(product),
        "applications": _map_applications(product.applications),
        "toxicity_rows": _map_toxicity_rows(product),
        "risk_component_rows": _map_risk_component_rows(product),
        "environmental_notes": product.environmental_notes or ["No environmental notes provided."],
        "references": [
            {"source": reference.source, "id": reference.id}
            for reference in product.references
        ],
        "chart_path": str(chart_path) if chart_path else "",
        "chart_note": "" if chart_path else "No efficacy data available.",
        "efficacy_chart": "",
    }


def _format_summary_text(products: list[Product], conditional_count: int) -> str:
    product_label = "product" if len(products) == 1 else "products"
    return (
        f"This report covers {len(products)} {product_label}; {conditional_count} of them under conditional approval. "
        f"Average efficacy across all applications: {format_percent(calculate_overall_average_efficacy(products))}."
    )


def _format_approval_line(product: Product, is_german: bool) -> str:
    valid_from = format_date(product.approval_period.valid_from)
    valid_to = format_date(product.approval_period.valid_to)
    if is_german:
        return (
            f"Germany: Zugelassen gem. PflSchG - Status: {product.approval_status.value}. "
            f"Valid from {valid_from} to {valid_to}."
        )
    return (
        f"Country: {product.country}. Status: {product.approval_status.value}. "
        f"Valid from {valid_from} to {valid_to}."
    )


def _format_approval_summary(product: Product, is_german: bool) -> str:
    valid_from = format_date(product.approval_period.valid_from)
    valid_to = format_date(product.approval_period.valid_to)
    if is_german:
        return (
            f"Germany: Zugelassen gem. PflSchG - Status: {product.approval_status.value}. "
            f"Valid from {valid_from} to {valid_to}."
        )
    return (
        f"Country: {product.country}. Approval status: {product.approval_status.value}. "
        f"Valid from {valid_from} to {valid_to}."
    )


def _format_high_risk_notice(product: Product, is_german: bool) -> str:
    return (
        f"This product carries a composite risk score of {product.risk_score} out of 100 and requires elevated "
        "regulatory scrutiny. Consult the Restrictions section before authorising any application."
    )


def _format_registration_summary(product: Product) -> str:
    application_count = len(product.applications)
    crop_label = "crop" if application_count == 1 else "crops"
    return (
        f"This product is registered for {application_count} {crop_label} with an average efficacy of "
        f"{format_percent(calculate_product_average_efficacy(product.applications))}."
    )


def _map_applications(applications: list[Application]) -> list[dict]:
    if not applications:
        return [
            {
                "crop": "No applications provided",
                "target_pest": "N/A",
                "bbch": "N/A",
                "dose": "N/A",
                "max_applications": "N/A",
                "efficacy": "N/A",
                "phi_days": "N/A",
            }
        ]

    return [
        {
            "crop": application.crop,
            "target_pest": application.target_pest,
            "bbch": application.growth_stage_bbch,
            "dose": format_dose(application.dose, application.dose_unit),
            "max_applications": application.max_applications_per_season,
            "efficacy": format_percent(application.efficacy_percent, decimals=0),
            "phi_days": application.preharvest_interval_days,
        }
        for application in applications
    ]


def _map_toxicity_rows(product: Product) -> list[dict]:
    indicators = product.toxicity_indicators.model_dump()
    return [
        {"indicator": format_metric_name(name), "value": f"{value:g}"}
        for name, value in indicators.items()
    ]


def _map_risk_component_rows(product: Product) -> list[dict]:
    components = product.risk_components.model_dump()
    return [
        {"component": format_metric_name(name), "score": score}
        for name, score in components.items()
    ]
