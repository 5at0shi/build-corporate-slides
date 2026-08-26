"""Atom層: これ以上分解できない描画の最小単位。

Boxは「矩形領域＋見た目（塗り・枠線・角丸・影）」という1つの構造を
表す。Card / Background Zone / Focus Panelは、この構造の見た目が違う
だけの同じもの（skinの違い）であり、幾何計算（角丸半径の絶対値換算、
影の作り方）は1箇所に集約する。components.pyのadd_card等は、この
Boxへ薄く委譲するラッパーとして再定義する。
"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Pt

from .theme import LAYOUT, PALETTE


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


def Box(slide, x, y, w, h, *, rounded=True, radius=None, fill=None,
        line=None, line_width=Pt(0.7), elevated=False, shadow_fill=None):
    """矩形コンテナを描く（Atom層の基本図形）。

    rounded=Trueかつradius未指定ならradius_base（標準コンテナ用）を使う。
    画面の大部分を占める面にはradius=LAYOUT.radius_lgを明示する。
    rounded=Falseは角丸なしの区切り用の面（Background Zoneの既定）。

    elevated=Trueの場合、オフセットした面を背後に重ねて擬似的な影を
    作る。PowerPoint/LibreOffice/Keynote間で描画結果が揺れるネイティブ
    shadowエフェクト（outerShdw）は使わない（LibreOffice側で想定より
    濃く描画され、環境間で見え方が揺れることを確認済み）。fill/line
    を指定しない場合はそれぞれ塗りなし/枠線なしになる（区切りのため
    の透明な領域としても使える）。

    elevated=Trueで作った影と本体は、PowerPoint上でグループ化して
    1つの編集単位にする。グループ化しないと、手作業でサイズ変更する
    際に本体と影を毎回2つ別々に動かす必要があり、編集性が落ちるため。
    グループ化後もbox（本体）への参照はそのまま有効で、fill/line等を
    個別に変更できる。
    """
    if rounded:
        radius = LAYOUT.radius_base if radius is None else radius
    else:
        radius = None
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE

    if elevated:
        shadow = slide.shapes.add_shape(shape_type, x + Pt(1.5), y + Pt(2), w, h)
        _flat(shadow, radius=radius)
        shadow.fill.solid()
        shadow.fill.fore_color.rgb = shadow_fill or PALETTE.surface_subtle
        shadow.line.fill.background()

    box = slide.shapes.add_shape(shape_type, x, y, w, h)
    _flat(box, radius=radius)
    if fill is not None:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    else:
        box.fill.background()
    if line is not None:
        box.line.color.rgb = line
        box.line.width = line_width
    else:
        box.line.fill.background()

    if elevated:
        slide.shapes.add_group_shape([shadow, box])
    return box


def add_hairline(slide, x, y, w, *, color=PALETTE.grey_300, width=0.75):
    """細い横罫線（区切り線）を引く。テキストを持たない純粋な装飾図形。"""
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Pt(width))
    _flat(line)
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()
    return line


def Marker(slide, x, y, w, h, *, shape="bar", fill=PALETTE.line_brand, rounding=0.16):
    """テキストを持たない小さな単色アクセント図形（縦棒・点など）を描く。

    見出し脇の縦棒（SectionLead・章扉）、プロセスのゲート点など、
    「小さく・単色塗り・枠線なし」という同じ構造の装飾がrenderer側に
    個別実装されがちだったため、1箇所に集約する。Boxと違い枠線・影の
    バリエーションは持たない（装飾はBoxほど作り込まない、という区別）。

    shape="bar"（既定）は角丸の細い帯。shape="dot"は円。
    """
    shape_type = MSO_SHAPE.OVAL if shape == "dot" else MSO_SHAPE.ROUNDED_RECTANGLE
    marker = slide.shapes.add_shape(shape_type, x, y, w, h)
    _flat(marker, rounding=None if shape == "dot" else rounding)
    marker.fill.solid()
    marker.fill.fore_color.rgb = fill
    marker.line.fill.background()
    return marker
