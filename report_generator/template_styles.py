from __future__ import annotations

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
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
BODY_FONT = "Arial"


def configure_document_theme(document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = document.styles
    styles["Normal"].font.name = BODY_FONT
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = BODY_FONT
    styles["Title"].font.size = Pt(28)
    styles["Title"].font.bold = True
    styles["Title"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 1"].font.name = BODY_FONT
    styles["Heading 1"].font.size = Pt(18)
    styles["Heading 1"].font.bold = True
    styles["Heading 1"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 2"].font.name = BODY_FONT
    styles["Heading 2"].font.size = Pt(13)
    styles["Heading 2"].font.bold = True
    styles["Heading 2"].font.color.rgb = RGBColor.from_string(BRAND_ORANGE)


def style_run(run, *, color: str, size: int, bold: bool = False) -> None:
    run.font.name = BODY_FONT
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = bold


def add_accent_bar(document, fill: str) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.text = ""
    shade_cell(cell, fill)
    set_cell_border(cell, fill)
    table.rows[0].height = Pt(8)


def style_key_value_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for row in table.rows:
        row.cells[0].width = Inches(1.9)
        for index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell, MID_GREY)
            style_cell_text(cell, color=BRAND_NAVY, bold=index == 0)
            shade_cell(cell, PALE_ORANGE if index == 0 else LIGHT_GREY)


def style_table_header(row, fill: str) -> None:
    for cell in row.cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade_cell(cell, fill)
        set_cell_border(cell, fill)
        style_cell_text(cell, color="FFFFFF", bold=True)


def style_table_body(table) -> None:
    for row in table.rows[1:]:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell, MID_GREY)
            style_cell_text(cell, color=BRAND_NAVY, bold=False)


def style_warning_cell(cell) -> None:
    shade_cell(cell, PALE_RED)
    set_cell_border(cell, DARK_RED)
    style_cell_text(cell, color=DARK_RED, bold=True)


def style_cell_text(cell, *, color: str, bold: bool) -> None:
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_after = Pt(2)
        for run in paragraph.runs:
            style_run(run, color=color, size=9, bold=bold)


def shade_cell(cell, fill: str) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = cell_properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        cell_properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_border(cell, color: str) -> None:
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


def set_paragraph_bottom_border(paragraph, color: str) -> None:
    paragraph_properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    borders.append(bottom)
    paragraph_properties.append(borders)
