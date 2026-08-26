from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from .atoms import add_hairline
from .icons import add_icon
from .textmetrics import (adaptive_gap_pt, estimate_item_list_height_pt,
                          estimate_paragraph_height_pt)
from .theme import PALETTE, TYPE


def _type_for(slide):
    slide_typography = getattr(slide, "_slidekit_typography", None)
    if slide_typography is not None:
        return slide_typography
    presentation = slide.part.package.presentation_part.presentation
    return getattr(presentation, "_slidekit_typography", TYPE)


def style_text_frame(text_frame, *, margin=0, margin_x=None, margin_y=None,
                     vertical_anchor=MSO_ANCHOR.TOP):
    """textboxの余白・折り返し・垂直配置を初期化する。

    margin_x/margin_yを指定すると左右/上下で異なる余白にできる（未指定
    の側はmargin、既定0を使う）。
    """
    text_frame.clear()
    text_frame.margin_left = text_frame.margin_right = (
        margin_x if margin_x is not None else margin)
    text_frame.margin_top = text_frame.margin_bottom = (
        margin_y if margin_y is not None else margin)
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


def _normalize_list_item(item):
    if isinstance(item, str):
        return item, None
    return item.get("title", ""), item.get("body")


def _list_item_segments(marker, index, title, body, bullet_char, typography):
    """1項目分のsegments（(text, run_kwargs)の配列）を組み立てる。"""
    segments = []
    if marker == "number":
        segments.append((f"{index + 1:02}   ", {
            "size": typography.small, "color": PALETTE.blue,
            "bold": True, "font": typography.body_font,
        }))
        title_text, title_bold, indent = title, True, "      "
    elif marker == "bullet":
        title_text, title_bold, indent = f"{bullet_char}  {title}", True, "    "
    else:
        # "icon" / "none": 記号なしの本文（iconは呼び出し側が別図形で置く）
        title_text, title_bold, indent = title, False, "  "
    segments.append((title_text, {
        "size": typography.body, "color": PALETTE.text_primary,
        "bold": title_bold, "font": typography.body_font,
    }))
    if body:
        segments.append((f"\n{indent}{body}", {
            "size": typography.small, "color": PALETTE.text_secondary,
            "font": typography.body_font,
        }))
    return segments


def add_text_list(slide, x, y, w, h, items, *, marker="bullet", bullet_char="・",
                  icon=None, icon_color=None, icon_size=None, text_gap=Inches(0.16),
                  divider=False, max_row_h=Inches(0.9),
                  vertical_anchor=MSO_ANCHOR.TOP, gap=3, adaptive=True):
    """複数項目のテキストを描く統合atom。

    旧add_item_list / add_icon_list / add_numbered_rowを1つの実装へ統合
    したもの。itemsは文字列、または{"title": ..., "body": ...}の配列。

    marker: "bullet"（既定、bullet_charでの箇条書き）/ "number"（01, 02...
      を自動採番）/ "icon"（iconで指定したアイコンを行ごとに独立した図形
      で置く）/ "none"（記号なし）。

    divider=False（既定）は全項目を1つのtextboxへまとめる。PowerPoint上
    での編集性を優先し、項目ごとにshapeへ分割しない。divider=Trueは各
    項目を独立した行として描き、行の間に実際の文字高さへ合わせた安全な
    罫線を引く（numbered_listのような単列の項目向け。安全な余白がない
    行は罫線を省略する）。項目数が少ないときはmax_row_hを上限に行の
    高さを詰め、領域内で上下中央へ配置する（上詰めのまま下部だけ余ると
    間延びして見えるため）。
    """
    typography = _type_for(slide)
    normalized = [_normalize_list_item(item) for item in items]

    if divider:
        row_h = Inches(0.9) if not items else min(max_row_h, h // len(items))
        block_h = row_h * len(items)
        cursor = y + max(0, (h - block_h) // 2)
        text_w_pt = w / 12700
        indent = "      " if marker == "number" else "    "
        for index, (title, body) in enumerate(normalized):
            segments = _list_item_segments(marker, index, title, body, bullet_char, typography)
            add_paragraph_textbox(slide, x, cursor, w, Inches(0.76), [
                {"segments": segments, "space_after": 0}
            ])
            # 罫線は固定オフセットではなく実際の文字高さに合わせる。行が
            # 短い場合、固定オフセットのままだと罫線が自分の行から離れ、
            # 次の行の文字に近づきすぎて見える（区切り線が次行に属して
            # 見える）ため。row_hに対して安全な余白がない場合は、中途半端
            # な位置に引いて文字に被せるより罫線を省略する。
            content_pt = estimate_paragraph_height_pt(title, typography.body.pt, text_w_pt,
                                                       line_spacing=1.08)
            if body:
                content_pt += estimate_paragraph_height_pt(
                    indent + body, typography.small.pt, text_w_pt, line_spacing=1.08)
            content_h = Inches(content_pt / 72)
            offset = content_h + Inches(0.14)
            safe_offset = row_h - Inches(0.08)
            if safe_offset >= content_h + Inches(0.04):
                add_hairline(slide, x, cursor + min(offset, safe_offset), w)
            cursor += row_h
        return None

    if marker == "icon":
        names = icon if isinstance(icon, list) else [icon or "check"] * len(items)
        size = icon_size if icon_size is not None else Pt(typography.body.pt * 2.0)
        color = icon_color or PALETTE.blue
        text_x = x + size + text_gap
        text_w = w - size - text_gap
        text_w_pt = text_w / 12700
        heights_pt = [estimate_paragraph_height_pt(title, typography.body.pt, text_w_pt,
                                                    line_spacing=1.15)
                     for title, _ in normalized]
        gap_pt = adaptive_gap_pt(sum(heights_pt), len(items), h / 12700, base_gap=gap)
        paragraphs, cursor = [], y
        for index, ((title, body), name, item_h_pt) in enumerate(
                zip(normalized, names, heights_pt)):
            add_icon(slide, x, cursor, size, name, color=color)
            segments = _list_item_segments("icon", index, title, body, bullet_char, typography)
            paragraphs.append({
                "segments": segments,
                "line_spacing": 1.15,
                "space_after": gap_pt if index < len(items) - 1 else 0,
            })
            cursor += Inches(item_h_pt / 72) + Pt(gap_pt)
        return add_paragraph_textbox(slide, text_x, y, text_w, h, paragraphs)

    body_gap = gap
    if adaptive and items:
        content_pt = estimate_item_list_height_pt(typography, items, w / 12700, body_gap=0)
        body_gap = int(adaptive_gap_pt(content_pt, len(items), h / 12700, base_gap=gap))
    paragraphs = []
    for index, (title, body) in enumerate(normalized):
        segments = _list_item_segments(marker, index, title, body, bullet_char, typography)
        paragraphs.append({
            "segments": segments,
            "space_after": body_gap if index < len(items) - 1 else 0,
        })
    return add_paragraph_textbox(slide, x, y, w, h, paragraphs,
                                 vertical_anchor=vertical_anchor)
