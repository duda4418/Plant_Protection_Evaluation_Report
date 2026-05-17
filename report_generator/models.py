from __future__ import annotations

from datetime import date
from enum import StrEnum

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class ApprovalStatus(StrEnum):
    APPROVED = "approved"
    CONDITIONAL = "conditional"
    REJECTED = "rejected"
    EXPIRED = "expired"


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ReportMetadata(StrictModel):
    report_id: str = Field(min_length=1)
    report_title: str = Field(min_length=1)
    issuing_authority: str = Field(min_length=1)
    issued_on: date
    language_locale: str = Field(min_length=1)
    version: str = Field(min_length=1)
    classification: str = Field(min_length=1)


class Concentration(StrictModel):
    value: float = Field(gt=0)
    unit: str = Field(min_length=1)


class ApprovalPeriod(StrictModel):
    valid_from: date
    valid_to: date

    @model_validator(mode="after")
    def check_date_order(self) -> "ApprovalPeriod":
        if self.valid_from > self.valid_to:
            raise ValueError(
                f"valid_from ({self.valid_from}) must not be after valid_to ({self.valid_to})"
            )
        return self


class RiskComponents(StrictModel):
    human_health: int = Field(ge=0, le=100)
    ecotoxicology: int = Field(ge=0, le=100)
    environmental_fate: int = Field(ge=0, le=100)
    residue_in_food: int = Field(ge=0, le=100)


class Application(StrictModel):
    crop: str = Field(min_length=1)
    target_pest: str = Field(min_length=1)
    growth_stage_bbch: str = Field(min_length=1)
    dose: float = Field(gt=0)
    dose_unit: str = Field(min_length=1)
    max_applications_per_season: int = Field(ge=0)
    efficacy_percent: int = Field(ge=0, le=100)
    preharvest_interval_days: int = Field(ge=0)


class ToxicityIndicators(StrictModel):
    ld50_oral_rat_mg_per_kg: float = Field(gt=0)
    ld50_dermal_rat_mg_per_kg: float = Field(gt=0)
    fish_lc50_96h_mg_per_l: float = Field(gt=0)
    bee_ld50_oral_microg_per_bee: float = Field(gt=0)
    earthworm_lc50_14d_mg_per_kg: float = Field(gt=0)


class Reference(StrictModel):
    source: str = Field(min_length=1)
    id: str = Field(min_length=1)


class Product(StrictModel):
    product_id: str = Field(min_length=1)
    product_name: str = Field(min_length=1)
    active_substance: str = Field(min_length=1)
    active_substance_concentration: Concentration
    formulation_type: str = Field(min_length=1)
    formulation_type_long: str = Field(min_length=1)
    manufacturer: str = Field(min_length=1)
    country: str = Field(min_length=1)
    approval_status: ApprovalStatus
    approval_period: ApprovalPeriod
    risk_score: int = Field(ge=0, le=100)
    risk_components: RiskComponents
    applications: list[Application] = Field(default_factory=list)
    toxicity_indicators: ToxicityIndicators
    restrictions: list[str] = Field(default_factory=list)
    environmental_notes: list[str] = Field(default_factory=list)
    references: list[Reference] = Field(default_factory=list)

    @field_validator("restrictions", "environmental_notes", mode="before")
    @classmethod
    def validate_string_list_items(cls, v: object) -> object:
        if isinstance(v, list):
            for i, item in enumerate(v):
                if not isinstance(item, str) or not str(item).strip():
                    raise ValueError(f"item {i} must be a non-empty string")
        return v

    @field_validator("country", mode="before")
    @classmethod
    def normalize_country(cls, value: str) -> str:
        normalized = str(value).strip().upper()
        if not normalized:
            raise ValueError("country must not be empty")
        return normalized

    @field_validator("approval_status", mode="before")
    @classmethod
    def normalize_approval_status(cls, value: str) -> str:
        return str(value).strip().lower()


class Footer(StrictModel):
    classification: str = Field(min_length=1)
    disclaimer: str = Field(min_length=1)


class ReportInput(StrictModel):
    report_metadata: ReportMetadata
    products: list[Product] = Field(min_length=1)
    footer: Footer

    @model_validator(mode="after")
    def check_unique_product_ids(self) -> "ReportInput":
        ids = [p.product_id for p in self.products]
        seen: set[str] = set()
        duplicates = {pid for pid in ids if pid in seen or seen.add(pid)}  # type: ignore[func-returns-value]
        if duplicates:
            raise ValueError(
                f"product_id values must be unique; duplicates found: {', '.join(sorted(duplicates))}"
            )
        return self
