"""Fragment層: AtomとLayoutを組み合わせた、意味的にはまだ完結しない
再利用可能な構造パターン。

BoxGrid（regionをR×Cのグリッドに分割しBoxを並べる）、MarkerOverlay
（線上の位置にMarker＋ラベルを重ねる）など、複数のrendererで同じ形が
繰り返し必要になった構造をここに集約する。Fragment自体はページの意味
（何のためのグラフか等）を持たず、rendererがこれらを組み合わせてページ
を構築する。

命名はビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。
形だけで再利用できることがFragment層の価値のため。同じ操作（矩形を
グリッドに分割する）に対して「横一列」「縦一列」「田の字」で別々の
名前・実装を用意しない。次元数(rows/cols)は同じ操作のパラメータに
過ぎず、本質的に異なる構造ではないため。
"""
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from .atoms import Marker, add_hairline
from .components import add_background_zone, add_card
from .theme import PALETTE
from .typography import _type_for, add_textbox


def BoxGrid(slide, region, items, *, rows=None, cols=None, skin="card",
           tones=None, row_weights=None, col_weights=None, gap="standard",
           inset_x=Inches(0.2), inset_y=Inches(0.2)):
    """regionをrows×colsのグリッドに分割し、各セルにBoxを描いて、内側の
    余白を差し引いたRegionのリストを行優先順（左上から右へ、次に下の行）
    で返す（add_panelの複数版）。

    中身の描画（テキスト・Stat・アイコン等）は呼び出し側が返り値の
    Regionへ行う。Boxの見た目を揃えることだけに責務を絞ることで、
    中身の構造がrendererごとに大きく違っても対応できるようにする。

    rows/colsは片方だけ指定すると、もう片方はitem数から自動計算される
    1次元の並び（横一列 or 縦一列）になる。両方指定すると2次元グリッド
    （例: 2x2の分割）になる。両方省略すると横一列（既定）。

    regionはスライド全体である必要はない。columns()/rows()で切り出した
    部分領域を渡せば、ページの左半分・右半分など任意のエリア内だけで
    独立してグリッドを組める。

    skin: "card"（既定）または "zone"（tonesで色を指定できる）。
    tones: 色名のリスト（インデックスで循環）、または
    tone(item, index) -> 色名 の関数（データに応じて選ぶ場合。例:
    「emphasisフラグを持つ項目だけ強調色にする」）。
    """
    n = max(1, len(items))
    if rows is None and cols is None:
        rows, cols = 1, n
    elif rows is None:
        rows = -(-n // cols)  # ceil
    elif cols is None:
        cols = -(-n // rows)  # ceil

    row_weights = row_weights or [1] * rows
    col_weights = col_weights or [1] * cols
    row_regions = region.rows(row_weights, gap=gap) if rows > 1 else [region]
    content_regions = []
    for row_index, row_region in enumerate(row_regions):
        cells = (row_region.columns(col_weights, gap=gap) if cols > 1
                else [row_region])
        for col_index, cell in enumerate(cells):
            index = row_index * cols + col_index
            if index >= len(items):
                break
            if skin == "zone":
                if callable(tones):
                    tone = tones(items[index], index)
                elif tones:
                    tone = tones[index % len(tones)]
                else:
                    tone = "neutral"
                add_background_zone(slide, cell.x, cell.y, cell.w, cell.h,
                                    tone=tone, rounded=True)
            else:
                add_card(slide, cell.x, cell.y, cell.w, cell.h)
            content_regions.append(cell.inset(inset_x, inset_y))
    return content_regions


def MarkerOverlay(slide, region, points, *, track=True, track_y=None,
                  track_color=PALETTE.line_neutral, track_width=1,
                  shape="dot", marker_size=Inches(0.11),
                  marker_color=PALETTE.blue, label_size=None,
                  label_color=PALETTE.text_primary, label_bold=True,
                  label_w=Inches(1.75), label_gap=Inches(0.125)):
    """regionの横方向に0〜1のpositionでMarker（点/バー）とラベルを重ねる。

    track=Trueならregion内の水平線（track_y、未指定はregion中央）に沿って
    Markerを並べる。pointsは{"position": 0.0-1.0, "title": ...}の配列。

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
