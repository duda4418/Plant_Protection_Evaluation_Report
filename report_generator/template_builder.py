from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


def build_default_template(output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = Document()
    _configure_document(document)
    _add_header_footer(document)
    _add_title_page(document)
    _add_summary_section(document)
    _add_product_loop(document)
    document.save(output_path)


def _configure_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.size = Pt(26)
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.size = Pt(18)
    styles["Heading 2"].font.name = "Aptos Display"
    styles["Heading 2"].font.size = Pt(13)


def _add_header_footer(document: Document) -> None:
    section = document.sections[0]
    header = section.header.paragraphs[0]
    header.text = "{{ metadata.issuing_authority }} | {{ metadata.report_id }}"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("{{ footer.classification }} | Page ")
    _add_field(footer, "PAGE")
    footer.add_run(" of ")
    _add_field(footer, "NUMPAGES")
    footer.add_run(" | {{ footer.disclaimer }}")


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


def _add_title_page(document: Document) -> None:
    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("{{ metadata.report_title }}")

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("{{ metadata.issuing_authority }}").bold = True

    details = document.add_table(rows=5, cols=2)
    details.style = "Table Grid"
    rows = [
        ("Report ID", "{{ metadata.report_id }}"),
        ("Issued on", "{{ metadata.issued_on }}"),
        ("Version", "{{ metadata.version }}"),
        ("Locale", "{{ metadata.language_locale }}"),
        ("Classification", "{{ metadata.classification }}"),
    ]
    for row, (label, value) in zip(details.rows, rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value

    document.add_page_break()


def _add_summary_section(document: Document) -> None:
    document.add_heading("Executive Summary", level=1)
    summary = document.add_table(rows=4, cols=2)
    summary.style = "Table Grid"
    rows = [
        ("Products assessed", "{{ summary.product_count }}"),
        ("Conditional approvals", "{{ summary.conditional_count }}"),
        ("High-risk products", "{{ summary.high_risk_count }}"),
        ("Overall average efficacy", "{{ summary.overall_average_efficacy }}"),
    ]
    for row, (label, value) in zip(summary.rows, rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value

    document.add_paragraph(
        "This report was generated from structured JSON input after schema validation, business-rule mapping, chart generation, and DOCX template rendering."
    )


def _add_product_loop(document: Document) -> None:
    document.add_paragraph("{% for product in products %}")
    document.add_section(WD_SECTION.NEW_PAGE)
    document.add_heading("{{ product.name }}", level=1)
    document.add_paragraph("Product ID: {{ product.product_id }}")
    document.add_paragraph("Active substance: {{ product.active_substance }} ({{ product.concentration }})")
    document.add_paragraph("Formulation: {{ product.formulation }}")
    document.add_paragraph("Manufacturer: {{ product.manufacturer }}")
    document.add_paragraph("{{ product.approval_line }}")
    document.add_paragraph("Risk score: {{ product.risk_score }}")
    document.add_paragraph("Applications assessed: {{ product.application_count }}")
    document.add_paragraph("Average efficacy: {{ product.average_efficacy }}")

    document.add_paragraph("{% if product.show_high_risk_notice %}")
    notice = document.add_paragraph()
    notice_run = notice.add_run("{{ product.high_risk_notice }}")
    notice_run.bold = True
    notice_run.font.color.rgb = RGBColor(156, 43, 43)
    document.add_paragraph("{% endif %}")

    document.add_heading("Applications", level=2)
    _add_applications_table(document)

    document.add_heading("Risk Components", level=2)
    _add_risk_table(document)

    document.add_heading("Toxicity Indicators", level=2)
    _add_toxicity_table(document)

    document.add_heading("Efficacy Chart", level=2)
    chart_paragraph = document.add_paragraph()
    chart_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    chart_paragraph.add_run("{{ product.efficacy_chart }}")
    document.add_paragraph("{{ product.chart_note }}")

    document.add_paragraph("{% if product.show_restrictions %}")
    document.add_heading("{{ product.restrictions_heading }}", level=2)
    document.add_paragraph("{% for restriction in product.restrictions %}")
    document.add_paragraph("- {{ restriction }}")
    document.add_paragraph("{% endfor %}")
    document.add_paragraph("{% endif %}")

    document.add_heading("Environmental Notes", level=2)
    document.add_paragraph("{% for note in product.environmental_notes %}")
    document.add_paragraph("- {{ note }}")
    document.add_paragraph("{% endfor %}")

    document.add_heading("References", level=2)
    document.add_paragraph("{% if product.references %}")
    document.add_paragraph("{% for reference in product.references %}")
    document.add_paragraph("- {{ reference.source }}: {{ reference.id }}")
    document.add_paragraph("{% endfor %}")
    document.add_paragraph("{% else %}")
    document.add_paragraph("No references provided.")
    document.add_paragraph("{% endif %}")
    document.add_paragraph("{% endfor %}")


def _add_applications_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=7)
    table.style = "Table Grid"
    headers = ["Crop", "Target pest", "BBCH", "Dose", "Max apps", "Efficacy", "PHI days"]
    for cell, header in zip(table.rows[0].cells, headers, strict=True):
        cell.text = header
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


def _add_risk_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Component"
    table.rows[0].cells[1].text = "Score"
    table.rows[1].cells[0].text = "{%tr for row in product.risk_component_rows %}"
    table.rows[2].cells[0].text = "{{ row.component }}"
    table.rows[2].cells[1].text = "{{ row.score }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"


def _add_toxicity_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Indicator"
    table.rows[0].cells[1].text = "Value"
    table.rows[1].cells[0].text = "{%tr for row in product.toxicity_rows %}"
    table.rows[2].cells[0].text = "{{ row.indicator }}"
    table.rows[2].cells[1].text = "{{ row.value }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
