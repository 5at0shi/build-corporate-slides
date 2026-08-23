from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

from .theme import PALETTE
from .typography import set_run


def _set_border(cell, color="D7DEE6", width="6350"):
    properties = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        existing = properties.find(qn(tag))
        if existing is not None:
            properties.remove(existing)
        line = OxmlElement(tag)
        line.set("w", width)
        fill = OxmlElement("a:solidFill")
        value = OxmlElement("a:srgbClr")
        value.set("val", color)
        fill.append(value)
        line.append(fill)
        properties.append(line)


def add_data_table(slide, region, columns, rows, *, highlight_key="_highlight"):
    typography = slide._slidekit_typography
    shape = slide.shapes.add_table(len(rows) + 1, len(columns),
                                   region.x, region.y, region.w, region.h)
    table = shape.table
    total_weight = sum(column.get("weight", 1) for column in columns)
    for index, column in enumerate(columns):
        table.columns[index].width = int(
            region.w * column.get("weight", 1) / total_weight)

    for row_index in range(len(rows) + 1):
        table.rows[row_index].height = int(region.h / (len(rows) + 1))
        row_data = None if row_index == 0 else rows[row_index - 1]
        highlighted = bool(row_data and row_data.get(highlight_key))
        for col_index, column in enumerate(columns):
            cell = table.cell(row_index, col_index)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.06)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if row_index == 0:
                cell.fill.fore_color.rgb = PALETTE.surface_subtle
                value = column.get("label", column.get("key", ""))
            else:
                cell.fill.fore_color.rgb = (
                    PALETTE.surface_brand_soft if highlighted else PALETTE.surface_base)
                value = str(row_data.get(column["key"], ""))
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = (PP_ALIGN.CENTER if column.get("align") == "center"
                                   else PP_ALIGN.LEFT)
            run = paragraph.add_run(); run.text = value
            set_run(run,
                    size=typography.small if row_index else typography.small,
                    color=(PALETTE.blue if highlighted and row_index else
                           PALETTE.text_primary),
                    bold=(row_index == 0 or highlighted),
                    font=typography.body_font)
            _set_border(cell)
    return shape
