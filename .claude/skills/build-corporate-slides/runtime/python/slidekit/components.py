from datetime import date

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from .icons import add_icon
from .layout import Region
from .textmetrics import (adaptive_gap_pt, estimate_item_list_height_pt,
                          estimate_paragraph_height_pt)
from .theme import LAYOUT, PALETTE, TYPE
from .typography import add_paragraph_textbox, add_textbox, set_run


def _type_for(slide):
    slide_typography = getattr(slide, "_slidekit_typography", None)
    if slide_typography is not None:
        return slide_typography
    presentation = slide.part.package.presentation_part.presentation
    return getattr(presentation, "_slidekit_typography", TYPE)


def _flat(shape, *, rounding=None, radius=None):
    shape.shadow.inherit = False
    effect_list = shape._element.spPr.find(qn("a:effectLst"))
    if effect_list is not None:
        shape._element.spPr.remove(effect_list)
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    if radius is not None and len(shape.adjustments):
        # ROUNDED_RECTANGLEのadj値は図形の短辺に対する割合として解釈されるため、
        # 同じ値でも縦横比が違う図形同士では角丸の見え方が揃わない。
        # 絶対の角丸半径を保つよう、短辺から都度adj値を逆算する。
        shape.adjustments[0] = min(radius / min(shape.width, shape.height), 0.5)
    elif rounding is not None and len(shape.adjustments):
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
            # 斜めの境界を一つだけにし、片側全体を単色の面にする
            # （PARALLELOGRAMは両辺が斜めになり境界が二重に見えるため使わない）。
            # 下側（ロゴを置く側）をやや広く、上側（タイトル側）をやや狭く
            # 取り、単調な平行四辺形にならないようにする。
            narrow_w = int(width * 0.76)
            wide_w = int(width * 1.18)
            if brand_side == "left":
                points = [(0, 0), (narrow_w, 0),
                         (wide_w, LAYOUT.slide_height),
                         (0, LAYOUT.slide_height)]
            else:
                top_x = LAYOUT.slide_width - narrow_w
                bottom_x = LAYOUT.slide_width - wide_w
                points = [(top_x, 0), (LAYOUT.slide_width, 0),
                         (LAYOUT.slide_width, LAYOUT.slide_height),
                         (bottom_x, LAYOUT.slide_height)]
            builder = slide.shapes.build_freeform(
                start_x=points[0][0], start_y=points[0][1])
            builder.add_line_segments(points[1:], close=True)
            brand_shapes.append(builder.convert_to_shape())
        for brand in brand_shapes:
            _flat(brand)
            brand.line.fill.background()
            if brand_shape == "diagonal":
                # 単色の面だけだと平板に見えるため、ごく控えめなグラデー
                # ションで奥行きを出す（表紙のブランド面に限定した使用）。
                brand.fill.gradient()
                brand.fill.gradient_angle = 115.0
                stops = brand.fill.gradient_stops
                stops[0].position = 0.0
                stops[0].color.rgb = RGBColor(0xE7, 0xF1, 0xFD)
                stops[1].position = 1.0
                stops[1].color.rgb = RGBColor(0xCC, 0xDB, 0xF8)
            else:
                brand.fill.solid()
                brand.fill.fore_color.rgb = PALETTE.surface_brand_soft

    if brand_side == "left":
        content_x = max(Inches(3.25), width + Inches(0.35))
        content_w = LAYOUT.slide_width - LAYOUT.margin_x - content_x
    else:
        content_x = LAYOUT.margin_x
        content_w = Inches(8.5)
    if eyebrow:
        add_textbox(slide, content_x, Inches(2.08), content_w, Inches(0.28),
                    eyebrow.upper(), size=typography.small, color=PALETTE.blue, bold=True)
    # 表紙はポスター的な大見出しのため、本文12pt基準で縮むtypography.titleは
    # 使わず、add_section_dividerと同じ固定サイズにする。
    add_textbox(slide, content_x, Inches(2.48), content_w, Inches(1.3),
                title, size=Pt(34), color=PALETTE.text_primary, bold=True,
                font=typography.headline_font, line_spacing=0.98)
    if subtitle:
        add_textbox(slide, content_x, Inches(4.03), content_w, Inches(0.62),
                    subtitle, size=typography.section, color=PALETTE.text_secondary)

    meta_x, meta_w = Inches(9.2), Inches(3.45)
    badge_x, badge_y, badge_w = Inches(11.42), Inches(0.32), Inches(1.22)
    if classification:
        label = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       badge_x, badge_y, badge_w, Inches(0.3))
        _flat(label, rounding=0.05)
        label.fill.background(); label.line.color.rgb = PALETTE.grey_500
        label.line.width = Pt(1.0)
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
        # 実際の右下の角へ、部外秘ボックスの右端・上端の余白を基準に置く
        # （同じ余白量で角に揃えると、意図的な配置に見える）。
        logo_w = Inches(0.55)
        corner_margin = badge_y  # 部外秘ボックスの上端の余白と揃える
        logo_x = Inches(0.55) if brand_side == "left" else badge_x + badge_w - logo_w
        picture = slide.shapes.add_picture(str(logo_path), logo_x, 0, width=logo_w)
        picture.top = LAYOUT.slide_height - corner_margin - picture.height


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


# add_section_leadの縦棒マーカー分として、呼び出し側が下に置く要素へ
# 空けるべき最小オフセット。マーカー高さが呼び出し側でも変わる場合は
# marker_h + SECTION_LEAD_GAPを使い、値がズレて重なるのを防ぐ。
SECTION_LEAD_GAP = Inches(0.1)


def add_section_lead(slide, x, y, w, text, *, color=PALETTE.line_brand,
                     size=None, marker_h=Inches(0.38)):
    typography = _type_for(slide)
    size = size or typography.section
    marker = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y,
                                    Inches(0.06), marker_h)
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
    _flat(zone, radius=LAYOUT.radius if rounded else None)
    zone.fill.solid(); zone.fill.fore_color.rgb = tones[tone]
    zone.line.fill.background()
    return zone


def add_panel(slide, x, y, w, h, *, tone="neutral", rounded=True,
             inset_x=Inches(0.32), inset_y=Inches(0.24)):
    """背景面を描き、内側の余白を差し引いたRegionを返す。

    見出し+リストや複数段落のブロックを「意図的な余白を持つ面」として
    見せたいときに使う。剥き出しの文字を隣接パネルと並べると、下部の
    余白が「欠けている」ように見えるため、そうした箇所で使う。
    呼び出し側は戻り値のRegion内へadd_section_lead/add_item_list/
    add_paragraph_textbox等を配置する。
    """
    add_background_zone(slide, x, y, w, h, tone=tone, rounded=rounded)
    return Region(x, y, w, h).inset(inset_x, inset_y)


def add_card(slide, x, y, w, h, *, fill=PALETTE.surface_base,
             line=PALETTE.line_neutral, elevated=False):
    if elevated:
        shadow = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        x + Pt(1.5), y + Pt(2), w, h)
        _flat(shadow, radius=LAYOUT.radius)
        shadow.fill.solid(); shadow.fill.fore_color.rgb = PALETTE.surface_subtle
        shadow.line.fill.background()
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    _flat(card, radius=LAYOUT.radius)
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
    _flat(panel, radius=LAYOUT.radius)
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
    _flat(box, radius=LAYOUT.radius)
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


def add_numbered_row(slide, x, y, w, number, title, body=None, *, row_h=None):
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
    # 罫線は固定オフセットではなく実際の文字高さに合わせる。行が短い場合、
    # 固定オフセットのままだと罫線が自分の行から離れ、次の行の文字に
    # 近づきすぎて見える（区切り線が次行に属して見える）ため。
    text_pt = w / 12700
    content_pt = estimate_paragraph_height_pt(title, typography.body.pt, text_pt,
                                              line_spacing=1.08)
    if body:
        # 実際の描画は本文の前に "      "（6スペース）の字下げが付くため、
        # 折り返し推定にも含めて過小評価を防ぐ。
        content_pt += estimate_paragraph_height_pt("      " + body, typography.small.pt,
                                                    text_pt, line_spacing=1.08)
    content_h = Inches(content_pt / 72)
    offset = content_h + Inches(0.14)
    if row_h is not None:
        # row_hが項目数に応じて縮められている場合、文字高さ基準のオフセット
        # のままだと次の行の番号・タイトルへ罫線が食い込む。次の行が始まる
        # 手前に収まるようクランプする。
        safe_offset = row_h - Inches(0.08)
        if safe_offset < content_h + Inches(0.04):
            # 本文自体が行の高さぎりぎりで、線を安全に置ける余白がない。
            # 中途半端な位置に引いて文字に被せるより、線を省略する。
            return
        offset = min(offset, safe_offset)
    add_hairline(slide, x, y + offset, w)


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


def add_icon_list(slide, x, y, w, h, items, *, icon="check",
                  icon_color=PALETTE.blue, icon_size=None,
                  text_gap=Inches(0.16), body_gap=14):
    """アイコン付き箇条書きを描く。

    文章はadd_item_listと同じく一つのtextboxへまとめ、行ごとに別shapeへ
    分割しない（編集性を保つため）。アイコンは装飾図形であり文章そのもの
    ではないため、行ごとに独立した図形として置く。itemsは文字列の配列。
    iconは全行共通の名前、またはitemsと同数のnameリストを渡す
    （行ごとに変える場合）。

    icon_sizeを指定しない場合、本文フォントサイズの約2倍を目安に自動計算
    する。固定インチ値のままだとlarge-roomモードなど本文が大きいmodeで
    アイコンが相対的に小さく見える（business比で本文が約1.4倍でも
    アイコンは同じ大きさのまま、というズレが生じる）ため。
    """
    typography = _type_for(slide)
    if icon_size is None:
        icon_size = Pt(typography.body.pt * 2.0)
    names = icon if isinstance(icon, list) else [icon] * len(items)
    text_x = x + icon_size + text_gap
    text_w = w - icon_size - text_gap
    text_w_pt = text_w / 12700
    heights_pt = [estimate_paragraph_height_pt(item, typography.body.pt, text_w_pt,
                                                line_spacing=1.15) for item in items]
    gap_pt = adaptive_gap_pt(sum(heights_pt), len(items), h / 12700, base_gap=body_gap)
    paragraphs, cursor = [], y
    for index, (item, name, item_h_pt) in enumerate(zip(items, names, heights_pt)):
        add_icon(slide, x, cursor, icon_size, name, color=icon_color)
        paragraphs.append({
            "segments": [(item, {
                "size": typography.body, "color": PALETTE.text_primary,
                "font": typography.body_font,
            })],
            "line_spacing": 1.15,
            "space_after": gap_pt if index < len(items) - 1 else 0,
        })
        cursor += Inches(item_h_pt / 72) + Pt(gap_pt)
    return add_paragraph_textbox(slide, text_x, y, text_w, h, paragraphs)
