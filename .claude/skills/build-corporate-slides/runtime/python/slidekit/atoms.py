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
    shadowエフェクトは使わない。fill/lineを指定しない場合はそれぞれ
    塗りなし/枠線なしになる（区切りのための透明な領域としても使える）。
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
    return box
