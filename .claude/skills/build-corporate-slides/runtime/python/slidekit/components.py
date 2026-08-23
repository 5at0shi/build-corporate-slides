from datetime import date

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from .textmetrics import adaptive_gap_pt, estimate_item_list_height_pt
from .theme import LAYOUT, PALETTE, TYPE
from .typography import add_paragraph_textbox, add_textbox, set_run


def _type_for(slide):
    slide_typography = getattr(slide, "_slidekit_typography", None)
    if slide_typography is not None:
        return slide_typography
    presentation = slide.part.package.presentation_part.presentation
    return getattr(presentation, "_slidekit_typography", TYPE)


def _flat(shape, *, rounding=None):
    shape.shadow.inherit = False
    effect_list = shape._element.spPr.find(qn("a:effectLst"))
    if effect_list is not None:
        shape._element.spPr.remove(effect_list)
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    if rounding is not None and len(shape.adjustments):
        shape.adjustments[0] = rounding
    return shape


def add_hairline(slide, x, y, w, *, color=PALETTE.grey_300, width=0.75):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Pt(width))
    _flat(line)
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()
    return line


def add_slide_title(slide, title, *, kicker=None, page=None):
    typography = _type_for(slide)
    if kicker:
        add_textbox(slide, LAYOUT.margin_x, Inches(0.34), Inches(5.5), Inches(0.2),
                    kicker.upper(), size=typography.small, color=PALETTE.blue, bold=True)
    add_textbox(slide, LAYOUT.margin_x, Inches(0.62), Inches(11.7), Inches(0.52),
                title, size=typography.title, color=PALETTE.ink, bold=True,
                font=typography.headline_font)
    add_hairline(slide, LAYOUT.margin_x, Inches(1.25),
                 LAYOUT.slide_width - 2 * LAYOUT.margin_x,
                 color=PALETTE.line_brand, width=1.4)
    if page is not None:
        add_textbox(slide, Inches(12.1), Inches(0.37), Inches(0.55), Inches(0.2),
                    f"{page:02}", size=typography.small, color=PALETTE.grey_500)


def add_cover(slide, title, *, subtitle=None, department="", created=None,
              eyebrow=None, logo_path=None, classification="",
              brand_side="right", brand_shape="diagonal", brand_width=None):
    """表紙を描く。brand_widthでブランド面の幅を調整できる（既定は3.4in/2.84in）。

    左配置時はcontent_xをbrand_widthに応じて自動的に押し出す。右配置時に
    大きく広げる場合は、本文が8.5in幅に収まっているか確認する。
    """
    typography = _type_for(slide)
    created = created or date.today().isoformat()
    if brand_side not in {"left", "right", "none"}:
        raise ValueError("brand_sideは 'left'、'right'、'none' を指定してください")
    if brand_shape not in {"diagonal", "curve", "straight"}:
        raise ValueError("brand_shapeは 'diagonal'、'curve'、'straight' を指定してください")

    default_widths = {"curve": Inches(3.4), "straight": Inches(2.84),
                      "diagonal": Inches(3.4)}
    width = brand_width or default_widths[brand_shape]

    if brand_side != "none":
        brand_shapes = []
        if brand_shape == "curve":
            inner_w = int(width * (1.25 / 3.4))
            if brand_side == "left":
                brand_shapes.append(slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, 0, 0, inner_w, LAYOUT.slide_height))
                brand_shapes.append(slide.shapes.add_shape(
                    MSO_SHAPE.OVAL, 0, 0, width, LAYOUT.slide_height))
            else:
                brand_shapes.append(slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, LAYOUT.slide_width - inner_w, 0,
                    inner_w, LAYOUT.slide_height))
                brand_shapes.append(slide.shapes.add_shape(
                    MSO_SHAPE.OVAL, LAYOUT.slide_width - width, 0,
                    width, LAYOUT.slide_height))
        elif brand_shape == "straight":
            x = 0 if brand_side == "left" else LAYOUT.slide_width - width
            brand_shapes.append(slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, x, 0, width, LAYOUT.slide_height))
        else:
            x = 0 if brand_side == "left" else LAYOUT.slide_width - width
            brand_shapes.append(slide.shapes.add_shape(
                MSO_SHAPE.PARALLELOGRAM, x, 0, width, LAYOUT.slide_height))
        for brand in brand_shapes:
            _flat(brand)
            brand.fill.solid()
            brand.fill.fore_color.rgb = PALETTE.surface_brand_soft
            brand.line.fill.background()

    if brand_side == "left":
        content_x = max(Inches(3.25), width + Inches(0.35))
        content_w = LAYOUT.slide_width - LAYOUT.margin_x - content_x
    else:
        content_x = LAYOUT.margin_x
        content_w = Inches(8.5)
    if eyebrow:
        add_textbox(slide, content_x, Inches(2.08), content_w, Inches(0.28),
                    eyebrow.upper(), size=typography.small, color=PALETTE.blue, bold=True)
    add_textbox(slide, content_x, Inches(2.48), content_w, Inches(1.3),
                title, size=typography.title, color=PALETTE.text_primary, bold=True,
                font=typography.headline_font, line_spacing=0.98)
    if subtitle:
        add_textbox(slide, content_x, Inches(4.03), content_w, Inches(0.62),
                    subtitle, size=typography.section, color=PALETTE.text_secondary)

    meta_x, meta_w = Inches(9.2), Inches(3.45)
    if classification:
        label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Inches(11.42), Inches(0.32),
                                       Inches(1.22), Inches(0.3))
        _flat(label, rounding=0.05)
        label.fill.background(); label.line.color.rgb = PALETTE.line_neutral
        label.line.width = Pt(0.7)
        tf = label.text_frame
        tf.clear(); tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        run = p.add_run(); run.text = classification
        set_run(run, size=typography.small, color=PALETTE.text_secondary,
                bold=True, font=typography.body_font)
    if department:
        add_textbox(slide, meta_x, Inches(0.72), meta_w, Inches(0.24), department,
                    size=typography.small, color=PALETTE.text_secondary,
                    bold=True, align=PP_ALIGN.RIGHT)
    add_textbox(slide, meta_x, Inches(1.0), meta_w, Inches(0.24), created,
                size=typography.small, color=PALETTE.text_secondary,
                align=PP_ALIGN.RIGHT)
    if logo_path:
        logo_x = Inches(0.55) if brand_side == "left" else Inches(11.6)
        slide.shapes.add_picture(str(logo_path), logo_x, Inches(5.7),
                                 width=Inches(0.55))


def add_section_divider(slide, title, *, kicker=None, subtitle=None, page=None):
    """章扉。通常ページの小見出しヘッダーは使わず、単独で成立させる。"""
    typography = _type_for(slide)
    center_y = Inches(3.2)
    marker = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    LAYOUT.margin_x, center_y,
                                    Inches(0.09), Inches(0.6))
    _flat(marker, rounding=0.16)
    marker.fill.solid(); marker.fill.fore_color.rgb = PALETTE.line_brand
    marker.line.fill.background()
    text_x = LAYOUT.margin_x + Inches(0.3)
    title_y = center_y
    if kicker:
        add_textbox(slide, text_x, center_y - Inches(0.02), Inches(10),
                    Inches(0.3), kicker.upper(), size=typography.small,
                    color=PALETTE.blue, bold=True)
        title_y = center_y + Inches(0.36)
    add_textbox(slide, text_x, title_y, Inches(11), Inches(1.2), title,
                size=Pt(34), color=PALETTE.text_primary, bold=True,
                font=typography.headline_font, line_spacing=1.02)
    if subtitle:
        add_textbox(slide, text_x, title_y + Inches(1.05), Inches(9.5),
                    Inches(0.6), subtitle, size=typography.section,
                    color=PALETTE.text_secondary)
    if page is not None:
        add_textbox(slide, Inches(12.1), Inches(0.37), Inches(0.55),
                    Inches(0.2), f"{page:02}", size=typography.small,
                    color=PALETTE.grey_500)


def add_section_lead(slide, x, y, w, text, *, color=PALETTE.line_brand,
                     size=None):
    typography = _type_for(slide)
    size = size or typography.section
    marker = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y,
                                    Inches(0.06), Inches(0.38))
    _flat(marker, rounding=0.16)
    marker.fill.solid(); marker.fill.fore_color.rgb = color
    marker.line.fill.background()
    label = add_textbox(slide, x + Inches(0.16), y - Inches(0.01),
                        w - Inches(0.16), Inches(0.42), text,
                        size=size, color=PALETTE.text_primary, bold=True)
    return marker, label


def add_background_zone(slide, x, y, w, h, *, tone="brand-soft",
                        rounded=False):
    tones = {
        "brand-soft": PALETTE.surface_brand_soft,
        "neutral": PALETTE.surface_subtle,
        "teal-soft": PALETTE.surface_teal_soft,
    }
    if tone not in tones:
        raise ValueError(f"未定義のtoneです: {tone}")
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    zone = slide.shapes.add_shape(shape_type, x, y, w, h)
    _flat(zone, rounding=0.018 if rounded else None)
    zone.fill.solid(); zone.fill.fore_color.rgb = tones[tone]
    zone.line.fill.background()
    return zone


def add_card(slide, x, y, w, h, *, fill=PALETTE.surface_base,
             line=PALETTE.line_neutral, elevated=False):
    if elevated:
        shadow = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        x + Pt(1.5), y + Pt(2), w, h)
        _flat(shadow, rounding=0.018)
        shadow.fill.solid(); shadow.fill.fore_color.rgb = PALETTE.surface_subtle
        shadow.line.fill.background()
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    _flat(card, rounding=0.018)
    card.fill.solid(); card.fill.fore_color.rgb = fill
    card.line.color.rgb = line; card.line.width = Pt(0.7)
    return card


def add_focus_panel(slide, x, y, w, h, *, tone="solid"):
    if tone == "solid":
        fill, line = PALETTE.focus_primary, PALETTE.focus_primary
    elif tone == "brand":
        fill, line = PALETTE.surface_brand_soft, PALETTE.line_brand
    else:
        raise ValueError("toneは 'solid' または 'brand' を指定してください")
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    _flat(panel, rounding=0.018)
    panel.fill.solid(); panel.fill.fore_color.rgb = fill
    panel.line.color.rgb = line; panel.line.width = Pt(1.0)
    return panel


def add_key_message(slide, x, y, w, text, *, style="editorial"):
    """ページの結論・導入文を示す。

    styleは editorial/subtle/solid/card に加え、装飾なしの plain を選べる。
    plain は、ページタイトル直下の罫線に隣接して短い線を重ねたくない場合
    （例: ページ冒頭の導入文）に使う。
    """
    typography = _type_for(slide)
    heights = {"editorial": Inches(0.72), "subtle": Inches(0.7),
               "solid": Inches(0.78), "card": Inches(0.82),
               "plain": Inches(0.5)}
    h = heights.get(style, heights["editorial"])
    if style == "plain":
        return add_textbox(slide, x, y, w, h, text, size=typography.section,
                           bold=True, font=typography.headline_font)
    if style == "editorial":
        add_hairline(slide, x, y, Inches(0.38), color=PALETTE.blue, width=2)
        return add_textbox(slide, x, y + Inches(0.17), w, h - Inches(0.17), text,
                           size=typography.section, bold=True,
                           font=typography.headline_font)
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    _flat(box, rounding=0.018)
    box.line.width = Pt(0.7)
    if style == "solid":
        box.fill.solid(); box.fill.fore_color.rgb = PALETTE.navy
        box.line.fill.background(); color = PALETTE.white
    else:
        box.fill.solid(); box.fill.fore_color.rgb = PALETTE.grey_100
        box.line.color.rgb = PALETTE.grey_300; color = PALETTE.ink
    tf = box.text_frame
    tf.clear(); tf.margin_left = tf.margin_right = Inches(0.18)
    tf.margin_top = tf.margin_bottom = Inches(0.12)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    run = tf.paragraphs[0].add_run(); run.text = text
    set_run(run, size=typography.body, color=color, bold=True,
            font=typography.body_font)
    return box


def add_numbered_row(slide, x, y, w, number, title, body=None):
    typography = _type_for(slide)
    segments = [
        (f"{number:02}   ", {
            "size": typography.small, "color": PALETTE.blue,
            "bold": True, "font": typography.body_font,
        }),
        (title, {
            "size": typography.body, "color": PALETTE.text_primary,
            "bold": True, "font": typography.body_font,
        }),
    ]
    if body:
        segments.append((f"\n      {body}", {
            "size": typography.small, "color": PALETTE.text_secondary,
            "font": typography.body_font,
        }))
    add_paragraph_textbox(slide, x, y, w, Inches(0.76), [
        {"segments": segments, "space_after": 0}
    ])
    add_hairline(slide, x, y + Inches(0.84), w)


def add_item_list(slide, x, y, w, h, items, *, bullet="•", body_gap=3,
                  vertical_anchor=MSO_ANCHOR.TOP, adaptive=True):
    """複数項目を一つのtextboxとして配置し、手修正しやすく保つ。

    itemsは文字列、または {"title": ..., "body": ...} の配列。
    独立移動が必要な項目には使わない。section_leadの直下など見出しに
    連続させる場合はTOPのまま使う（既定）。項目数が領域に対して少ない
    場合、adaptive=True（既定）なら項目間の余白を上限付きで自動的に
    広げ、上詰めのまま下部の余白の割合を抑える。見出しを伴わない独立
    ブロック（カード、パネルなど）を領域全体で釣り合わせたい場合は
    vertical_anchor=MSO_ANCHOR.MIDDLEを指定する。
    """
    typography = _type_for(slide)
    if adaptive and items:
        content_pt = estimate_item_list_height_pt(
            typography, items, w / 12700, body_gap=0)
        body_gap = int(adaptive_gap_pt(
            content_pt, len(items), h / 12700, base_gap=body_gap))
    paragraphs = []
    for index, item in enumerate(items):
        if isinstance(item, str):
            segments = [(f"{bullet}  {item}", {
                "size": typography.body, "color": PALETTE.text_primary,
                "font": typography.body_font,
            })]
        else:
            title = item.get("title", "")
            body = item.get("body", "")
            segments = [(f"{bullet}  {title}", {
                "size": typography.body, "color": PALETTE.text_primary,
                "bold": True, "font": typography.body_font,
            })]
            if body:
                segments.append((f"\n    {body}", {
                    "size": typography.small, "color": PALETTE.text_secondary,
                    "font": typography.body_font,
                }))
        paragraphs.append({
            "segments": segments,
            "space_after": body_gap if index < len(items) - 1 else 0,
        })
    return add_paragraph_textbox(slide, x, y, w, h, paragraphs,
                                 vertical_anchor=vertical_anchor)
