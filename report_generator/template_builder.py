from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BRAND_ORANGE = "F07D00"
BRAND_NAVY = "1D3A5E"
LIGHT_GREY = "F5F5F5"
MID_GREY = "D9DEE7"
PALE_ORANGE = "FFF2E6"
PALE_RED = "FCE4E4"
DARK_RED = "9C2B2B"


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
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = "Arial"
    styles["Title"].font.size = Pt(28)
    styles["Title"].font.bold = True
    styles["Title"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 1"].font.name = "Arial"
    styles["Heading 1"].font.size = Pt(18)
    styles["Heading 1"].font.bold = True
    styles["Heading 1"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 2"].font.name = "Arial"
    styles["Heading 2"].font.size = Pt(13)
    styles["Heading 2"].font.bold = True
    styles["Heading 2"].font.color.rgb = RGBColor.from_string(BRAND_ORANGE)


def _add_header_footer(document: Document) -> None:
    section = document.sections[0]
    header = section.header.paragraphs[0]
    header.clear()
    header_run = header.add_run("{{ metadata.issuing_authority }}")
    header_run.bold = True
    header_run.font.name = "Arial"
    header_run.font.size = Pt(8)
    header_run.font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    separator_run = header.add_run(" | {{ metadata.report_id }}")
    separator_run.font.name = "Arial"
    separator_run.font.size = Pt(8)
    separator_run.font.color.rgb = RGBColor.from_string(BRAND_ORANGE)
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_paragraph_bottom_border(header, BRAND_ORANGE)

    footer = section.footer.paragraphs[0]
    footer.clear()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    classification = footer.add_run("{{ footer.classification }}")
    classification.bold = True
    classification.font.name = "Arial"
    classification.font.size = Pt(7)
    classification.font.color.rgb = RGBColor.from_string(BRAND_ORANGE)
    footer_text = footer.add_run(" | Page ")
    footer_text.font.name = "Arial"
    footer_text.font.size = Pt(7)
    footer_text.font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    _add_field(footer, "PAGE")
    page_text = footer.add_run(" of ")
    page_text.font.name = "Arial"
    page_text.font.size = Pt(7)
    page_text.font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    _add_field(footer, "NUMPAGES")
    disclaimer = footer.add_run(" | {{ footer.disclaimer }}")
    disclaimer.font.name = "Arial"
    disclaimer.font.size = Pt(7)
    disclaimer.font.color.rgb = RGBColor.from_string(BRAND_NAVY)


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
    _add_accent_bar(document, BRAND_ORANGE)

    eyebrow = document.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    eyebrow_run = eyebrow.add_run("DYNAMIC REGULATORY REPORT GENERATOR")
    eyebrow_run.bold = True
    eyebrow_run.font.name = "Arial"
    eyebrow_run.font.size = Pt(10)
    eyebrow_run.font.color.rgb = RGBColor.from_string(BRAND_ORANGE)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(18)
    title.paragraph_format.space_after = Pt(8)
    title.add_run("{{ metadata.report_title }}")

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(24)
    subtitle_run = subtitle.add_run("{{ metadata.issuing_authority }}")
    subtitle_run.bold = True
    subtitle_run.font.name = "Arial"
    subtitle_run.font.size = Pt(12)
    subtitle_run.font.color.rgb = RGBColor.from_string(BRAND_NAVY)

    details = document.add_table(rows=5, cols=2)
    details.style = "Table Grid"
    details.alignment = WD_TABLE_ALIGNMENT.CENTER
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
    _style_key_value_table(details)

    document.add_page_break()


def _add_summary_section(document: Document) -> None:
    document.add_heading("Executive Summary", level=1)
    summary = document.add_table(rows=4, cols=2)
    summary.style = "Table Grid"
    summary.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows = [
        ("Products assessed", "{{ summary.product_count }}"),
        ("Conditional approvals", "{{ summary.conditional_count }}"),
        ("High-risk products", "{{ summary.high_risk_count }}"),
        ("Overall average efficacy", "{{ summary.overall_average_efficacy }}"),
    ]
    for row, (label, value) in zip(summary.rows, rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value
    _style_key_value_table(summary)

    summary_note = document.add_paragraph(
        "This report was generated from structured JSON input after schema validation, business-rule mapping, chart generation, and DOCX template rendering."
    )
    summary_note.paragraph_format.space_before = Pt(12)
    summary_note.paragraph_format.space_after = Pt(12)


def _add_product_loop(document: Document) -> None:
    document.add_paragraph("{% for product in products %}")
    document.add_section(WD_SECTION.NEW_PAGE)
    _add_accent_bar(document, BRAND_ORANGE)
    document.add_heading("{{ product.name }}", level=1)

    product_facts = document.add_table(rows=8, cols=2)
    product_facts.style = "Table Grid"
    product_facts_rows = [
        ("Product ID", "{{ product.product_id }}"),
        ("Active substance", "{{ product.active_substance }} ({{ product.concentration }})"),
        ("Formulation", "{{ product.formulation }}"),
        ("Manufacturer", "{{ product.manufacturer }}"),
        ("Approval", "{{ product.approval_line }}"),
        ("Risk score", "{{ product.risk_score }}"),
        ("Applications assessed", "{{ product.application_count }}"),
        ("Average efficacy", "{{ product.average_efficacy }}"),
    ]
    for row, (label, value) in zip(product_facts.rows, product_facts_rows, strict=True):
        row.cells[0].text = label
        row.cells[1].text = value
    _style_key_value_table(product_facts)

    document.add_paragraph("{% if product.show_high_risk_notice %}")
    notice_table = document.add_table(rows=1, cols=1)
    notice_table.style = "Table Grid"
    notice_cell = notice_table.rows[0].cells[0]
    notice_cell.text = "{{ product.high_risk_notice }}"
    _shade_cell(notice_cell, PALE_RED)
    _set_cell_border(notice_cell, DARK_RED)
    _style_cell_text(notice_cell, color=DARK_RED, bold=True)
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
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Crop", "Target pest", "BBCH", "Dose", "Max apps", "Efficacy", "PHI days"]
    for cell, header in zip(table.rows[0].cells, headers, strict=True):
        cell.text = header
    _style_table_header(table.rows[0], fill=BRAND_NAVY)
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
    _style_table_body(table)


def _add_risk_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0].cells[0].text = "Component"
    table.rows[0].cells[1].text = "Score"
    _style_table_header(table.rows[0], fill=BRAND_ORANGE)
    table.rows[1].cells[0].text = "{%tr for row in product.risk_component_rows %}"
    table.rows[2].cells[0].text = "{{ row.component }}"
    table.rows[2].cells[1].text = "{{ row.score }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
    _style_table_body(table)


def _add_toxicity_table(document: Document) -> None:
    table = document.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0].cells[0].text = "Indicator"
    table.rows[0].cells[1].text = "Value"
    _style_table_header(table.rows[0], fill=BRAND_ORANGE)
    table.rows[1].cells[0].text = "{%tr for row in product.toxicity_rows %}"
    table.rows[2].cells[0].text = "{{ row.indicator }}"
    table.rows[2].cells[1].text = "{{ row.value }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
    _style_table_body(table)


def _add_accent_bar(document: Document, fill: str) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.text = ""
    _shade_cell(cell, fill)
    _set_cell_border(cell, fill)
    table.rows[0].height = Pt(8)


def _style_key_value_table(table) -> None:
    table.autofit = True
    for row in table.rows:
        row.cells[0].width = Inches(1.9)
        for index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_border(cell, MID_GREY)
            _style_cell_text(cell, color=BRAND_NAVY, bold=index == 0)
            _shade_cell(cell, PALE_ORANGE if index == 0 else LIGHT_GREY)


def _style_table_header(row, fill: str) -> None:
    for cell in row.cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _shade_cell(cell, fill)
        _set_cell_border(cell, fill)
        _style_cell_text(cell, color="FFFFFF", bold=True)


def _style_table_body(table) -> None:
    for row in table.rows[1:]:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_border(cell, MID_GREY)
            _style_cell_text(cell, color=BRAND_NAVY, bold=False)


def _style_cell_text(cell, color: str, bold: bool) -> None:
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_after = Pt(2)
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string(color)
            run.bold = bold


def _shade_cell(cell, fill: str) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = cell_properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        cell_properties.append(shading)
    shading.set(qn("w:fill"), fill)


def _set_cell_border(cell, color: str) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    borders = cell_properties.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        cell_properties.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def _set_paragraph_bottom_border(paragraph, color: str) -> None:
    paragraph_properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    borders.append(bottom)
    paragraph_properties.append(borders)
