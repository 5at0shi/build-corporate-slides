from datetime import date

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .atoms import Box, Marker, _flat, add_hairline
from .layout import Region
from .theme import LAYOUT, PALETTE
from .typography import (_type_for, add_text_list, add_textbox, set_run,
                         style_text_frame)


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
        tf = style_text_frame(label.text_frame, margin_x=Inches(0.05),
                              vertical_anchor=MSO_ANCHOR.MIDDLE)
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
    Marker(slide, LAYOUT.margin_x, center_y, Inches(0.09), Inches(0.6))
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
    marker = Marker(slide, x, y, Inches(0.06), marker_h, fill=color)
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
    # 画面の大部分を占める面のため、角丸を使う場合はradius_lg（Card等の
    # radius_baseより大きい半径）を使う。小さい半径だと丸みが足りずに見える。
    return Box(slide, x, y, w, h, rounded=rounded, radius=LAYOUT.radius_lg,
               fill=tones[tone], line=None)


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
             line=PALETTE.line_neutral, elevated=True):
    """独立した情報単位。既定で軽い影を付ける（elevated=Falseでフラットに）。"""
    return Box(slide, x, y, w, h, radius=LAYOUT.radius_base, fill=fill,
               line=line, line_width=Pt(0.7), elevated=elevated)


def add_focus_panel(slide, x, y, w, h, *, tone="solid"):
    if tone == "solid":
        fill, line = PALETTE.focus_primary, PALETTE.focus_primary
    elif tone == "brand":
        fill, line = PALETTE.surface_brand_soft, PALETTE.line_brand
    else:
        raise ValueError("toneは 'solid' または 'brand' を指定してください")
    return Box(slide, x, y, w, h, radius=LAYOUT.radius_base, fill=fill,
               line=line, line_width=Pt(1.0), elevated=False)


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
    if style == "solid":
        box = Box(slide, x, y, w, h, radius=LAYOUT.radius_base,
                  fill=PALETTE.navy, line=None)
        color = PALETTE.white
    else:
        box = Box(slide, x, y, w, h, radius=LAYOUT.radius_base,
                  fill=PALETTE.grey_100, line=PALETTE.grey_300, line_width=Pt(0.7))
        color = PALETTE.ink
    tf = style_text_frame(box.text_frame, margin_x=Inches(0.18), margin_y=Inches(0.12),
                          vertical_anchor=MSO_ANCHOR.MIDDLE)
    run = tf.paragraphs[0].add_run(); run.text = text
    set_run(run, size=typography.body, color=color, bold=True,
            font=typography.body_font)
    return box


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

    実体はadd_text_list（marker="bullet"）。
    """
    return add_text_list(slide, x, y, w, h, items, marker="bullet",
                         bullet_char=bullet, gap=body_gap,
                         vertical_anchor=vertical_anchor, adaptive=adaptive)


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

    実体はadd_text_list（marker="icon"）。
    """
    return add_text_list(slide, x, y, w, h, items, marker="icon", icon=icon,
                         icon_color=icon_color, icon_size=icon_size,
                         text_gap=text_gap, gap=body_gap)
