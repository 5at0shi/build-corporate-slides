from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from .theme import PALETTE, TYPE


def style_text_frame(text_frame, *, margin=0, vertical_anchor=MSO_ANCHOR.TOP):
    text_frame.clear()
    text_frame.margin_left = margin
    text_frame.margin_right = margin
    text_frame.margin_top = margin
    text_frame.margin_bottom = margin
    text_frame.vertical_anchor = vertical_anchor
    text_frame.word_wrap = True
    return text_frame


def set_run(run, *, size=TYPE.body, color=PALETTE.ink, bold=False,
            font=TYPE.body_font):
    run.font.name = font
    properties = run._r.get_or_add_rPr()
    east_asian = properties.find(qn("a:ea"))
    if east_asian is None:
        east_asian = OxmlElement("a:ea")
        properties.append(east_asian)
    east_asian.set("typeface", font)
    run.font.size = size
    run.font.color.rgb = color
    run.font.bold = bold
    return run


def add_textbox(slide, x, y, w, h, text, *, size=TYPE.body,
                color=PALETTE.ink, bold=False, font=TYPE.body_font,
                align=PP_ALIGN.LEFT, margin=0, line_spacing=1.08):
    shape = slide.shapes.add_textbox(x, y, w, h)
    tf = style_text_frame(shape.text_frame, margin=margin)
    paragraph = tf.paragraphs[0]
    paragraph.alignment = align
    paragraph.line_spacing = line_spacing
    paragraph.space_after = Pt(0)
    set_run(paragraph.add_run(), size=size, color=color, bold=bold, font=font)
    paragraph.runs[0].text = text
    return shape


def add_rich_textbox(slide, x, y, w, h, segments, *, align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(x, y, w, h)
    tf = style_text_frame(shape.text_frame)
    paragraph = tf.paragraphs[0]
    paragraph.alignment = align
    for text, kwargs in segments:
        run = paragraph.add_run()
        run.text = text
        set_run(run, **kwargs)
    return shape


def add_paragraph_textbox(slide, x, y, w, h, paragraphs, *,
                          margin=0, vertical_anchor=MSO_ANCHOR.TOP):
    """複数の箇条書き・見出し・本文を一つの編集可能なtextboxへまとめる。

    paragraphsは {"segments": [(text, run_kwargs)], "space_after": 6,
    "level": 0, "bullet": False} の配列。意味上ひとまとまりの文章を、
    見た目上の行ごとに別shapeへ分割しないために使う。
    """
    shape = slide.shapes.add_textbox(x, y, w, h)
    tf = style_text_frame(shape.text_frame, margin=margin,
                          vertical_anchor=vertical_anchor)
    for index, spec in enumerate(paragraphs):
        paragraph = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        paragraph.alignment = spec.get("align", PP_ALIGN.LEFT)
        paragraph.level = spec.get("level", 0)
        paragraph.line_spacing = spec.get("line_spacing", 1.08)
        paragraph.space_before = Pt(spec.get("space_before", 0))
        paragraph.space_after = Pt(spec.get("space_after", 0))
        if spec.get("bullet"):
            paragraph.text = "• "
        for value, kwargs in spec.get("segments", []):
            run = paragraph.add_run()
            run.text = value
            set_run(run, **kwargs)
    return shape
