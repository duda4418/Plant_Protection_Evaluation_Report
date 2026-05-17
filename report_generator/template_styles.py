from __future__ import annotations

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BRAND_ORANGE = "F07D00"
BRAND_NAVY = "17365D"
TEXT_DARK = "1A1A1A"
TEXT_MUTED = "666666"
LIGHT_GREY = "F3F3F3"
MID_GREY = "D9D9D9"
HEADER_GREY = "EFEFEF"
PALE_ORANGE = "FFF7EE"
PALE_RED = "FFF7EE"
DARK_RED = "F07D00"
BODY_FONT = "Arial"


def configure_document_theme(document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.header_distance = Inches(0.24)
    section.footer_distance = Inches(0.28)

    styles = document.styles
    styles["Normal"].font.name = BODY_FONT
    styles["Normal"].font.size = Pt(8)
    styles["Normal"].font.color.rgb = RGBColor.from_string(TEXT_DARK)
    styles["Title"].font.name = BODY_FONT
    styles["Title"].font.size = Pt(22)
    styles["Title"].font.bold = True
    styles["Title"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 1"].font.name = BODY_FONT
    styles["Heading 1"].font.size = Pt(14)
    styles["Heading 1"].font.bold = True
    styles["Heading 1"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)
    styles["Heading 2"].font.name = BODY_FONT
    styles["Heading 2"].font.size = Pt(10)
    styles["Heading 2"].font.bold = True
    styles["Heading 2"].font.color.rgb = RGBColor.from_string(BRAND_NAVY)


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
    table.rows[0].height = Pt(4)


def style_summary_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    left_cell = table.rows[0].cells[0]
    body_cell = table.rows[0].cells[1]
    left_cell.width = Inches(0.05)
    body_cell.width = Inches(6.2)
    shade_cell(left_cell, BRAND_NAVY)
    set_cell_border(left_cell, BRAND_NAVY)
    shade_cell(body_cell, LIGHT_GREY)
    set_cell_border(body_cell, LIGHT_GREY)
    set_cell_margins(body_cell, top=120, start=140, bottom=120, end=120)
    style_cell_text(body_cell, color=TEXT_DARK, bold=False, size=10)


def style_key_value_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for row in table.rows:
        row.cells[0].width = Inches(1.9)
        for index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell, MID_GREY)
            set_cell_margins(cell)
            style_cell_text(cell, color=TEXT_DARK, bold=index == 0, size=10)
            shade_cell(cell, HEADER_GREY if index == 0 else "FFFFFF")


def style_table_header(row, fill: str) -> None:
    for cell in row.cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade_cell(cell, fill)
        set_cell_border(cell, fill)
        set_cell_margins(cell)
        style_cell_text(cell, color="FFFFFF", bold=True, size=7)


def style_table_body(table) -> None:
    for row_index, row in enumerate(table.rows[1:], start=1):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell, MID_GREY)
            set_cell_margins(cell)
            shade_cell(cell, "FFFFFF")
            style_cell_text(cell, color=TEXT_DARK, bold=False, size=7)


def style_metric_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_border(cell, MID_GREY)
            set_cell_margins(cell, top=65, start=90, bottom=65, end=90)
            shade_cell(cell, LIGHT_GREY if index == 0 else "FFFFFF")
            style_cell_text(cell, color=TEXT_DARK, bold=index == 0, size=7)


def style_warning_cell(cell) -> None:
    shade_cell(cell, PALE_RED)
    set_cell_border(cell, BRAND_ORANGE)
    set_cell_margins(cell, top=110, start=120, bottom=110, end=120)
    style_cell_text(cell, color=TEXT_DARK, bold=False, size=10)


def style_cell_text(cell, *, color: str, bold: bool, size: int = 8, italic: bool = False) -> None:
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1)
        for run in paragraph.runs:
            style_run(run, color=color, size=size, bold=bold)
            run.italic = italic


def shade_cell(cell, fill: str) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = cell_properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        cell_properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_margins(cell, *, top: int = 70, start: int = 90, bottom: int = 70, end: int = 90) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    margins = cell_properties.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        cell_properties.append(margins)

    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        element = margins.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            margins.append(element)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")


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
