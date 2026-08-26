"""Fragment層: AtomとLayoutを組み合わせた、意味的にはまだ完結しない
再利用可能な構造パターン。

BoxGrid（N個のBoxを行/列に等間隔配置）、BandStack（BoxGridの縦積み版）、
MarkerOverlay（線上の位置にMarker＋ラベルを重ねる）、QuadrantGrid（軸を
持たない4分割）など、複数のrendererで同じ形が繰り返し必要になった構造
をここに集約する。Fragment自体はページの意味（何のためのグラフか等）
を持たず、rendererがこれらを組み合わせてページを構築する。

命名はビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。
形だけで再利用できることがFragment層の価値のため。
"""
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from .atoms import Marker, add_hairline
from .components import add_background_zone, add_card
from .theme import PALETTE
from .typography import _type_for, add_textbox


def BoxGrid(slide, region, items, *, skin="card", tones=None, direction="row",
           weights=None, gap="standard", inset_x=Inches(0.2), inset_y=Inches(0.2)):
    """N個のBox（Card/Zone）を行または列に配置し、内側の余白を差し引いた
    Regionのリストを返す（add_panelの複数版）。

    中身の描画（テキスト・Stat・アイコン等）は呼び出し側が返り値の
    Regionへ行う。Boxの見た目を揃えることだけに責務を絞ることで、
    中身の構造がrendererごとに大きく違っても対応できるようにする。

    skin: "card"（既定）または "zone"（tonesで色を指定・循環できる）。
    direction: "row"（既定、横並び）または "column"（縦並び）。
    weights未指定は等分割。
    """
    n = max(1, len(items))
    weights = weights or [1] * n
    cells = (region.columns(weights, gap=gap) if direction == "row"
            else region.rows(weights, gap=gap))
    tones = tones or ["neutral"]
    content_regions = []
    for index in range(len(items)):
        cell = cells[index]
        if skin == "zone":
            add_background_zone(slide, cell.x, cell.y, cell.w, cell.h,
                                tone=tones[index % len(tones)], rounded=True)
        else:
            add_card(slide, cell.x, cell.y, cell.w, cell.h)
        content_regions.append(cell.inset(inset_x, inset_y))
    return content_regions


def BandStack(slide, region, items, *, tones=None, weights=None,
             gap=Inches(0.16), inset_x=Inches(0.3), inset_y=Inches(0.14)):
    """BoxGridの縦積み版（帯構造）。org_layersの階層バンド等、上から順に
    積む縦のまとまりに使う。既定skinはzone（帯として塗りつぶす）。
    """
    return BoxGrid(slide, region, items, skin="zone", tones=tones,
                   direction="column", weights=weights, gap=gap,
                   inset_x=inset_x, inset_y=inset_y)


def MarkerOverlay(slide, region, points, *, track=True, track_y=None,
                  track_color=PALETTE.line_neutral, track_width=1,
                  shape="dot", marker_size=Inches(0.11),
                  marker_color=PALETTE.blue, label_size=None,
                  label_color=PALETTE.text_primary, label_bold=True,
                  label_w=Inches(1.75), label_gap=Inches(0.125)):
    """regionの横方向に0〜1のpositionでMarker（点/バー）とラベルを重ねる。

    track=Trueならregion内の水平線（track_y、未指定はregion中央）に沿って
    Markerを並べる（NodeTrackの土台を兼ねる）。pointsは
    {"position": 0.0-1.0, "title": ...}の配列。

    labelは中央揃えを基本としつつ、Markerが端に近い場合は中央揃えのまま
    領域外へはみ出さないよう、Marker位置を起点に片側へ揃える向きへ自動
    的に切り替える。中央揃えのままクランプすると、Markerの中心とラベル
    の中心がズレて「ラベルが点から離れて見える」ため（process_with_gates
    のゲートラベル不具合の修正をここに集約し、再発を防ぐ）。
    """
    typography = _type_for(slide)
    label_size = label_size or typography.small
    line_y = region.y + (track_y if track_y is not None else region.h // 2)
    if track:
        add_hairline(slide, region.x, line_y, region.w,
                     color=track_color, width=track_width)
    region_right = region.x + region.w
    for point in points:
        position = max(0.0, min(1.0, float(point.get("position", 0))))
        x = region.x + int(region.w * position)
        Marker(slide, x - marker_size // 2, line_y - marker_size // 2,
              marker_size, marker_size, shape=shape, fill=marker_color)
        title = point.get("title")
        if not title:
            continue
        if x - label_w // 2 < region.x:
            label_x, align = x, PP_ALIGN.LEFT
        elif x + label_w // 2 > region_right:
            label_x, align = x - label_w, PP_ALIGN.RIGHT
        else:
            label_x, align = x - label_w // 2, PP_ALIGN.CENTER
        add_textbox(slide, int(label_x), line_y + marker_size // 2 + label_gap,
                   label_w, Inches(0.34), title, size=label_size,
                   color=label_color, bold=label_bold, align=align)


def QuadrantGrid(slide, region, items, *, tones=None, emphasis_tone="brand-soft",
                 default_tone="neutral", gap=Inches(0.1),
                 inset_x=Inches(0.24), inset_y=Inches(0.2)):
    """軸を持たない独立4分割。itemsは4件、順序は[左上, 右上, 左下, 右下]。

    tones未指定時は各itemの"emphasis"キー（bool）でトーンを自動選択する
    （Trueならemphasis_tone、それ以外はdefault_tone）。SWOTのような
    4テーマ分類に使う。座標軸上の連続値で位置づけたい場合は使わない
    （chart_with_insightのscatterを使う）。
    """
    top_row, bottom_row = region.rows([1, 1], gap=gap)
    cells = list(top_row.columns([1, 1], gap=gap)) + \
        list(bottom_row.columns([1, 1], gap=gap))
    content_regions = []
    for index, cell in enumerate(cells[:len(items)]):
        if tones:
            tone = tones[index % len(tones)]
        else:
            item = items[index]
            emphasized = isinstance(item, dict) and item.get("emphasis")
            tone = emphasis_tone if emphasized else default_tone
        add_background_zone(slide, cell.x, cell.y, cell.w, cell.h,
                            tone=tone, rounded=True)
        content_regions.append(cell.inset(inset_x, inset_y))
    return content_regions
