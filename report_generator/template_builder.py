from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .template_styles import (
    BRAND_NAVY,
    BRAND_ORANGE,
    add_accent_bar,
    configure_document_theme,
    set_paragraph_bottom_border,
    style_key_value_table,
    style_run,
    style_table_body,
    style_table_header,
    style_warning_cell,
)


def build_default_template(output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = Document()
    configure_document_theme(document)
    _add_header_footer(document)
    _add_title_page(document)
    _add_summary_section(document)
    _add_products_section(document)
    document.save(output_path)


# Document shell
def _add_header_footer(document: Document) -> None:
    section = document.sections[0]
    header = section.header.paragraphs[0]
    header.clear()
    header_run = header.add_run("{{ metadata.issuing_authority }}")
    style_run(header_run, color=BRAND_NAVY, size=8, bold=True)
    separator_run = header.add_run(" | {{ metadata.report_id }}")
    style_run(separator_run, color=BRAND_ORANGE, size=8)
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_bottom_border(header, BRAND_ORANGE)

    footer = section.footer.paragraphs[0]
    footer.clear()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    classification = footer.add_run("{{ footer.classification }}")
    style_run(classification, color=BRAND_ORANGE, size=7, bold=True)
    footer_text = footer.add_run(" | Page ")
    style_run(footer_text, color=BRAND_NAVY, size=7)
    _add_field(footer, "PAGE")
    page_text = footer.add_run(" of ")
    style_run(page_text, color=BRAND_NAVY, size=7)
    _add_field(footer, "NUMPAGES")
    disclaimer = footer.add_run(" | {{ footer.disclaimer }}")
    style_run(disclaimer, color=BRAND_NAVY, size=7)


def _add_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instruction

    fld_char_separate = OxmlElement("w:fldChar")
    fld_char_separate.set(qn("w:fldCharType"), "separate")

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_char_begin)
    run._r.append(instr_text)
    run._r.append(fld_char_separate)
    run._r.append(fld_char_end)


# Front matter
def _add_title_page(document: Document) -> None:
    add_accent_bar(document, BRAND_ORANGE)

    eyebrow = document.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    eyebrow_run = eyebrow.add_run("DYNAMIC REGULATORY REPORT GENERATOR")
    style_run(eyebrow_run, color=BRAND_ORANGE, size=10, bold=True)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(18)
    title.paragraph_format.space_after = Pt(8)
    title.add_run("{{ metadata.report_title }}")

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(24)
    subtitle_run = subtitle.add_run("{{ metadata.issuing_authority }}")
    style_run(subtitle_run, color=BRAND_NAVY, size=12, bold=True)

    details = document.add_table(rows=5, cols=2)
    details.style = "Table Grid"
    detail_rows = [
        ("Report ID", "{{ metadata.report_id }}"),
        ("Issued on", "{{ metadata.issued_on }}"),
        ("Version", "{{ metadata.version }}"),
        ("Locale", "{{ metadata.language_locale }}"),
        ("Classification", "{{ metadata.classification }}"),
    ]
    for row, (label, value) in zip(details.rows, detail_rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value
    style_key_value_table(details)

    document.add_page_break()


def _add_summary_section(document: Document) -> None:
    document.add_heading("Executive Summary", level=1)
    summary = document.add_table(rows=4, cols=2)
    summary.style = "Table Grid"
    summary_rows = [
        ("Products assessed", "{{ summary.product_count }}"),
        ("Conditional approvals", "{{ summary.conditional_count }}"),
        ("High-risk products", "{{ summary.high_risk_count }}"),
        ("Overall average efficacy", "{{ summary.overall_average_efficacy }}"),
    ]
    for row, (label, value) in zip(summary.rows, summary_rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value
    style_key_value_table(summary)

    summary_note = document.add_paragraph(
        "This report was generated from structured JSON input after schema validation, business-rule mapping, chart generation, and DOCX template rendering."
    )
    summary_note.paragraph_format.space_before = Pt(12)
    summary_note.paragraph_format.space_after = Pt(12)


# Repeating product section
def _add_products_section(document: Document) -> None:
    document.add_paragraph("{% for product in products %}")
    document.add_section(WD_SECTION.NEW_PAGE)
    add_accent_bar(document, BRAND_ORANGE)
    document.add_heading("{{ product.name }}", level=1)
    _add_product_overview(document)
    _add_high_risk_notice(document)
    _add_product_data_sections(document)
    _add_chart_section(document)
    _add_restrictions_section(document)
    _add_environmental_notes_section(document)
    _add_references_section(document)
    document.add_paragraph("{% endfor %}")


# Product overview
def _add_product_overview(document: Document) -> None:
    product_facts = document.add_table(rows=8, cols=2)
    product_facts.style = "Table Grid"
    fact_rows = [
        ("Product ID", "{{ product.product_id }}"),
        ("Active substance", "{{ product.active_substance }} ({{ product.concentration }})"),
        ("Formulation", "{{ product.formulation }}"),
        ("Manufacturer", "{{ product.manufacturer }}"),
        ("Approval", "{{ product.approval_line }}"),
        ("Risk score", "{{ product.risk_score }}"),
        ("Applications assessed", "{{ product.application_count }}"),
        ("Average efficacy", "{{ product.average_efficacy }}"),
    ]
    for row, (label, value) in zip(product_facts.rows, fact_rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value
    style_key_value_table(product_facts)


def _add_high_risk_notice(document: Document) -> None:
    document.add_paragraph("{% if product.show_high_risk_notice %}")
    notice_table = document.add_table(rows=1, cols=1)
    notice_table.style = "Table Grid"
    notice_cell = notice_table.rows[0].cells[0]
    notice_cell.text = "{{ product.high_risk_notice }}"
    style_warning_cell(notice_cell)
    document.add_paragraph("{% endif %}")


# Product data blocks
def _add_product_data_sections(document: Document) -> None:
    document.add_heading("Applications", level=2)
    _add_applications_table(document)

    document.add_heading("Risk Components", level=2)
    _add_risk_table(document)

    document.add_heading("Toxicity Indicators", level=2)
    _add_toxicity_table(document)


def _add_chart_section(document: Document) -> None:
    document.add_heading("Efficacy Chart", level=2)
    chart_paragraph = document.add_paragraph()
    chart_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    chart_paragraph.add_run("{{ product.efficacy_chart }}")
    document.add_paragraph("{{ product.chart_note }}")


def _add_restrictions_section(document: Document) -> None:
    document.add_paragraph("{% if product.show_restrictions %}")
    document.add_heading("{{ product.restrictions_heading }}", level=2)
    document.add_paragraph("{% for restriction in product.restrictions %}")
    document.add_paragraph("- {{ restriction }}")
    document.add_paragraph("{% endfor %}")
    document.add_paragraph("{% endif %}")


def _add_environmental_notes_section(document: Document) -> None:
    document.add_heading("Environmental Notes", level=2)
    document.add_paragraph("{% for note in product.environmental_notes %}")
    document.add_paragraph("- {{ note }}")
    document.add_paragraph("{% endfor %}")


def _add_references_section(document: Document) -> None:
    document.add_heading("References", level=2)
    document.add_paragraph("{% if product.references %}")
    document.add_paragraph("{% for reference in product.references %}")
    document.add_paragraph("- {{ reference.source }}: {{ reference.id }}")
    document.add_paragraph("{% endfor %}")
    document.add_paragraph("{% else %}")
    document.add_paragraph("No references provided.")
    document.add_paragraph("{% endif %}")


# Reusable tables
def _add_applications_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=7)
    table.style = "Table Grid"
    headers = ["Crop", "Target pest", "BBCH", "Dose", "Max apps", "Efficacy", "PHI days"]
    for cell, header in zip(table.rows[0].cells, headers, strict=True):
        cell.text = header
    style_table_header(table.rows[0], fill=BRAND_NAVY)
    table.rows[1].cells[0].text = "{%tr for app in product.applications %}"
    values = [
        "{{ app.crop }}",
        "{{ app.target_pest }}",
        "{{ app.bbch }}",
        "{{ app.dose }}",
        "{{ app.max_applications }}",
        "{{ app.efficacy }}",
        "{{ app.phi_days }}",
    ]
    for cell, value in zip(table.rows[2].cells, values, strict=True):
        cell.text = value
    table.rows[3].cells[0].text = "{%tr endfor %}"
    style_table_body(table)


def _add_risk_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Component"
    table.rows[0].cells[1].text = "Score"
    style_table_header(table.rows[0], fill=BRAND_ORANGE)
    table.rows[1].cells[0].text = "{%tr for row in product.risk_component_rows %}"
    table.rows[2].cells[0].text = "{{ row.component }}"
    table.rows[2].cells[1].text = "{{ row.score }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
    style_table_body(table)


def _add_toxicity_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Indicator"
    table.rows[0].cells[1].text = "Value"
    style_table_header(table.rows[0], fill=BRAND_ORANGE)
    table.rows[1].cells[0].text = "{%tr for row in product.toxicity_rows %}"
    table.rows[2].cells[0].text = "{{ row.indicator }}"
    table.rows[2].cells[1].text = "{{ row.value }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
    style_table_body(table)
