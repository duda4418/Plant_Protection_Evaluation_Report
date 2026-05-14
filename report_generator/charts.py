from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt

from .formatting import slugify
from .models import Product


def generate_efficacy_chart(product: Product, output_dir: Path) -> Path | None:
    if not product.applications:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    crops = [application.crop for application in product.applications]
    values = [application.efficacy_percent for application in product.applications]

    figure_height = max(2.6, 1.2 + 0.45 * len(crops))
    figure, axis = plt.subplots(figsize=(7.2, figure_height))
    bars = axis.barh(crops, values, color="#2f6f73")

    axis.set_xlim(0, 100)
    axis.set_xlabel("Efficacy (%)")
    axis.set_title(f"Efficacy by crop - {product.product_name}")
    axis.grid(axis="x", linestyle="--", alpha=0.35)
    axis.invert_yaxis()

    for bar, value in zip(bars, values, strict=True):
        axis.text(
            min(value + 1.2, 98),
            bar.get_y() + bar.get_height() / 2,
            f"{value}%",
            va="center",
            fontsize=9,
        )

    figure.tight_layout()
    chart_path = output_dir / f"{slugify(product.product_id)}-efficacy.png"
    figure.savefig(chart_path, dpi=180, bbox_inches="tight")
    plt.close(figure)
    return chart_path
