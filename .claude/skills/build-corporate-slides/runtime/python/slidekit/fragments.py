"""Fragment層: AtomとLayoutを組み合わせた、意味的にはまだ完結しない
再利用可能な構造パターン。

BoxGrid（regionをR×Cのグリッドに分割しBoxを並べる）、ProportionalStack
（値に比例した幅の帯を積む）、RadialLayout（円周上に配置し矢印で結ぶ）、
MarkerOverlay（線上の位置にMarker＋ラベルを重ねる）など、複数の
rendererで同じ形が繰り返し必要になった構造をここに集約する。Fragment
自体はページの意味（何のためのグラフか等）を持たず、rendererがこれら
を組み合わせてページを構築する。

命名はビジネス用語を使わない（「階層」「ゲート」等はrenderer側の語彙）。
形だけで再利用できることがFragment層の価値のため。同じ操作（矩形を
グリッドに分割する）に対して「横一列」「縦一列」「田の字」で別々の
名前・実装を用意しない。次元数(rows/cols)は同じ操作のパラメータに
過ぎず、本質的に異なる構造ではないため。
"""
import math

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from .atoms import Connector, Marker, add_hairline
from .components import add_background_zone, add_card
from .layout import Region
from .theme import PALETTE
from .typography import _type_for, add_textbox


def _skin_tone(items, index, tones):
    if callable(tones):
        return tones(items[index], index)
    if tones:
        return tones[index % len(tones)]
    return "neutral"


def _skinned_cell(slide, cell, items, index, *, skin, tones):
    """skin="zone"なら背景色付きの面、"card"なら白背景+枠線+影のCardを描く。

    BoxGridとProportionalStackが共有する「セルの見た目を決める」手順。
    どちらも中身（テキスト等）は呼び出し側が別途配置するため、ここでは
    面を描いて終わる。
    """
    if skin == "zone":
        tone = _skin_tone(items, index, tones)
        add_background_zone(slide, cell.x, cell.y, cell.w, cell.h,
                            tone=tone, rounded=True)
    else:
        add_card(slide, cell.x, cell.y, cell.w, cell.h)


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
            _skinned_cell(slide, cell, items, index, skin=skin, tones=tones)
            content_regions.append(cell.inset(inset_x, inset_y))
    return content_regions


def ProportionalStack(slide, region, items, *, value_key="value", min_ratio=0.18,
                      skin="zone", tones=None, gap="tight",
                      inset_x=Inches(0.28), inset_y=Inches(0.12)):
    """regionを縦に積み、各段の幅をvalue_keyの値に応じて描く。

    BoxGridが「等しい大きさのセルに分ける」操作なのに対し、これは
    「値に応じた幅の帯を順に積む」という別の操作（ファネル＝絞り込みの
    推移、ピラミッド＝下から積み上がる構造は、並び順と値の大小関係が
    違うだけの同じ操作のため、Fragmentは分けない）。

    幅は最大値に対する比率の平方根を使う（線形比率そのままだと、実際の
    ファネルによくある10〜100倍の落差で下位の段がほぼ潰れ、テキストが
    入らなくなるため）。正確な値の比率を厳密に伝えたい場合はこの
    Fragmentではなくchart_with_insightの棒グラフを使う。min_ratioは
    それでも潰れる最小段への下限。中央揃えで積み、各段の高さは均等。

    BoxGridと同じくskin/tonesを共有し、返り値も内側の余白を差し引いた
    Regionのリスト（items順）。
    """
    n = max(1, len(items))
    row_regions = region.rows([1] * n, gap=gap) if n > 1 else [region]
    values = [max(0.0, float(item.get(value_key, 0))) if isinstance(item, dict) else 0.0
             for item in items]
    max_value = max(values) if values else 0.0
    content_regions = []
    for index, (row_region, value) in enumerate(zip(row_regions, values)):
        ratio = max(min_ratio, (value / max_value) ** 0.5) if max_value else 1.0
        width = int(row_region.w * ratio)
        x = row_region.x + (row_region.w - width) // 2
        cell = Region(x, row_region.y, width, row_region.h)
        _skinned_cell(slide, cell, items, index, skin=skin, tones=tones)
        content_regions.append(cell.inset(inset_x, inset_y))
    return content_regions


def RadialLayout(slide, region, items, *, box_w=Inches(1.5), box_h=Inches(0.9),
                 skin="zone", tones=None, start_angle=-90, clockwise=True,
                 connector_color=PALETTE.grey_500, connector_width=Pt(1.75),
                 inset_x=Inches(0.16), inset_y=Inches(0.1)):
    """regionの中心を軸にitemsを円周上へ均等配置し、隣接する項目同士を
    最後尾から先頭へも矢印で結ぶ（循環プロセス: PDCA等）。

    BoxGrid（格子に等分割）・ProportionalStack（値に比例した幅で積む）
    とは別の、「円周上の位置へ配置する」という3つ目の配置操作（角度・
    半径という、他の2つには無い座標系を使うため）。矢印はConnectorの
    直線をそのまま使う（曲線の弧は環境間で見え方がぶれやすいため、
    まずは直線で多角形状のループを作る）。

    円周上の隣接2点は矢印で直接結ぶ。矢じり側がBoxに隠れて見えなく
    ならないよう、両端を「中心からその方向でBox（矩形）の外へ出る
    ちょうどの距離」だけ手前で止める（内接円半径のような一律の近似だと、
    横長のBoxに浅い角度で接続する矢印で不足し、Boxの下に完全に隠れて
    消えることがあるため。矩形と光線の交差なので方向ごとに正確に計算
    できる）。

    返り値はBoxGridと同じく内側の余白を差し引いたRegionのリスト
    （items順）。
    """
    n = max(1, len(items))
    cx = region.x + region.w // 2
    cy = region.y + region.h // 2
    radius = min(region.w, region.h) // 2 - max(box_w, box_h) // 2
    half_w, half_h = box_w / 2, box_h / 2
    direction = 1 if clockwise else -1

    centers = []
    for index in range(n):
        angle = math.radians(start_angle + direction * index * 360 / n)
        centers.append((cx + int(radius * math.cos(angle)),
                        cy + int(radius * math.sin(angle))))

    if n >= 2:
        for index in range(n):
            x1, y1 = centers[index]
            x2, y2 = centers[(index + 1) % n]
            dx, dy = x2 - x1, y2 - y1
            dist = math.hypot(dx, dy) or 1
            ux, uy = dx / dist, dy / dist
            # 中心から方向(ux, uy)へ進んだとき、矩形(half_w×half_h)の外へ
            # 出るのに必要な距離（光線と矩形の交差）。
            exits = []
            if abs(ux) > 1e-9:
                exits.append(half_w / abs(ux))
            if abs(uy) > 1e-9:
                exits.append(half_h / abs(uy))
            clearance = min(exits) if exits else half_w
            if dist <= clearance * 2:
                continue
            Connector(slide, x1 + int(ux * clearance), y1 + int(uy * clearance),
                     x2 - int(ux * clearance), y2 - int(uy * clearance),
                     color=connector_color, width=connector_width)

    content_regions = []
    for index, (px, py) in enumerate(centers):
        cell = Region(px - box_w // 2, py - box_h // 2, box_w, box_h)
        _skinned_cell(slide, cell, items, index, skin=skin, tones=tones)
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
