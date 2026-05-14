from __future__ import annotations

import argparse
import json
import sys
from json import JSONDecodeError
from pathlib import Path

from pydantic import ValidationError

from .charts import generate_efficacy_chart
from .mapper import map_report
from .models import ReportInput
from .renderer import RenderError, render_report
from .template_builder import build_default_template


class CliError(RuntimeError):
    """Raised when the command-line workflow cannot continue."""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a regulatory DOCX report from JSON input.")
    parser.add_argument("--input", type=Path, default=Path("data/input_data.json"), help="Path to input JSON.")
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("templates/report_template.docx"),
        help="Path to the DOCX template.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("generated/report.docx"),
        help="Path for the generated DOCX report.",
    )
    parser.add_argument(
        "--charts-dir",
        type=Path,
        default=Path("build/charts"),
        help="Directory for generated chart images.",
    )
    parser.add_argument(
        "--rebuild-template",
        action="store_true",
        help="Recreate the default DOCX template before rendering.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        run(args)
    except (CliError, RenderError, ValidationError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    return 0


def run(args: argparse.Namespace) -> None:
    if args.rebuild_template or not args.template.exists():
        build_default_template(args.template)

    raw_data = load_json(args.input)
    report = ReportInput.model_validate(raw_data)

    chart_paths = {
        product.product_id: generate_efficacy_chart(product, args.charts_dir)
        for product in report.products
    }
    context = map_report(report, chart_paths)
    render_report(context, args.template, args.output)

    print(f"Generated report: {args.output}")
    print(f"Generated template: {args.template}")


def load_json(path: Path) -> dict:
    if not path.exists():
        raise CliError(f"Input JSON not found: {path}")

    try:
        with path.open("r", encoding="utf-8") as input_file:
            return json.load(input_file)
    except JSONDecodeError as error:
        raise CliError(
            f"Invalid JSON in {path}: {error.msg} at line {error.lineno}, column {error.colno}"
        ) from error
