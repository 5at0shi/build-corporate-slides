"""外部アセットに依存しない、ベクター描画のみのアイコンセット。

ライセンス上の懸念を避けるため、既存のアイコンフォント・SVGセットは使わず、
python-pptxの図形（Oval/RoundedRectangle/Freeform）だけで構成する。
座標はすべて0.0〜1.0のローカル単位（アイコンの正方形キャンバスに対する
比率）で書き、`add_icon`が実寸へスケール・配置する。

線（チェックマーク・時計の針）は、開いたfreeformのストロークではなく、
細い RoundedRectangle を回転させて作る（開いたfreeformは一部レンダラーで
意図通り描画されないため）。回転は図形自身の中心を軸に回るため、始点を
固定するための補正（_hand内のax/ay計算）を行っている。
"""
import math

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt

from .theme import PALETTE

_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

ICON_NAMES = {"check", "warning", "people", "document", "clock", "target"}


def _flat(shape):
    shape.shadow.inherit = False
    return shape


def _oval(slide, x0, y0, x1, y1, *, x, y, size, fill=None, line=None,
          line_width=Pt(1.5)):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, x + int(x0 * size), y + int(y0 * size),
        int((x1 - x0) * size), int((y1 - y0) * size))
    _flat(shape)
    if fill is not None:
        shape.fill.solid(); shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line is not None:
        shape.line.color.rgb = line; shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def _rounded_rect(slide, x0, y0, x1, y1, *, x, y, size, fill, rounding=0.15):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, x + int(x0 * size), y + int(y0 * size),
        int((x1 - x0) * size), int((y1 - y0) * size))
    _flat(shape)
    if len(shape.adjustments):
        shape.adjustments[0] = rounding
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def _preset(slide, shape_type, x0, y0, x1, y1, *, x, y, size, fill):
    """OVAL/ROUNDED_RECTANGLE以外の定型オートシェイプ用（三角形・台形等）。

    自由形状（freeform）はこの環境のレンダリングパイプラインで意図通りに
    描画されないことがあるため、アイコンでは定型プリセットだけを使う。
    """
    shape = slide.shapes.add_shape(
        shape_type, x + int(x0 * size), y + int(y0 * size),
        int((x1 - x0) * size), int((y1 - y0) * size))
    _flat(shape)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def _hand_between(slide, p0, p1, thickness, *, x, y, size, color):
    """p0からp1へ向かう、太さthicknessの棒を描く（p0を固定端として）。

    RoundedRectangleを回転させて作る。回転は図形自身の中心を軸に回るため、
    p0が回転後も動かないよう、開始位置をあらかじめ補正している。
    """
    dx, dy = p1[0] - p0[0], p1[1] - p0[1]
    length = math.hypot(dx, dy)
    ux, uy = dx / length, dy / length
    angle_deg = math.degrees(math.atan2(ux, -uy))
    theta = math.radians(angle_deg)
    anchor_x = p0[0] + (length / 2) * math.sin(theta)
    anchor_y = p0[1] + (length / 2) * (1 - math.cos(theta))

    w = max(1, int(thickness * size))
    h = max(1, int(length * size))
    left = x + int(anchor_x * size) - w // 2
    top = y + int(anchor_y * size) - h
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    _flat(shape)
    if len(shape.adjustments):
        shape.adjustments[0] = 0.5
    shape.fill.solid(); shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.rotation = angle_deg
    return shape


def _draw_check(slide, x, y, size, color):
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _hand_between(slide, (0.26, 0.5), (0.43, 0.68), 0.1,
                 x=x, y=y, size=size, color=_WHITE)
    _hand_between(slide, (0.43, 0.68), (0.77, 0.3), 0.1,
                 x=x, y=y, size=size, color=_WHITE)


def _draw_warning(slide, x, y, size, color):
    # 他の5種と同じ「色チップ＋白い図形」の言語に揃える
    # （三角形そのものを外形にすると、丸チップの並びの中で浮いて見えるため）。
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _preset(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, 0.2, 0.2, 0.8, 0.74,
           x=x, y=y, size=size, fill=_WHITE)
    _rounded_rect(slide, 0.465, 0.36, 0.535, 0.56, x=x, y=y, size=size,
                  fill=color, rounding=0.5)
    _oval(slide, 0.465, 0.62, 0.535, 0.67, x=x, y=y, size=size, fill=color)


def _draw_people(slide, x, y, size, color):
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _oval(slide, 0.36, 0.16, 0.64, 0.44, x=x, y=y, size=size, fill=_WHITE)
    _preset(slide, MSO_SHAPE.TRAPEZOID, 0.26, 0.58, 0.74, 0.92,
           x=x, y=y, size=size, fill=_WHITE)


def _draw_document(slide, x, y, size, color):
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _rounded_rect(slide, 0.3, 0.18, 0.7, 0.82, x=x, y=y, size=size,
                  fill=_WHITE, rounding=0.12)
    for line_y in (0.36, 0.48, 0.6):
        _rounded_rect(slide, 0.4, line_y, 0.6, line_y + 0.045, x=x, y=y,
                      size=size, fill=color, rounding=0.5)


def _draw_clock(slide, x, y, size, color):
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _oval(slide, 0.16, 0.16, 0.84, 0.84, x=x, y=y, size=size,
          line=_WHITE, line_width=Pt(size.pt * 0.07))
    _hand_between(slide, (0.5, 0.5), (0.5, 0.26), 0.08,
                 x=x, y=y, size=size, color=_WHITE)
    _hand_between(slide, (0.5, 0.5), (0.68, 0.42), 0.08,
                 x=x, y=y, size=size, color=_WHITE)
    _oval(slide, 0.46, 0.46, 0.54, 0.54, x=x, y=y, size=size, fill=_WHITE)


def _draw_target(slide, x, y, size, color):
    _oval(slide, 0, 0, 1, 1, x=x, y=y, size=size, fill=color)
    _oval(slide, 0.18, 0.18, 0.82, 0.82, x=x, y=y, size=size, fill=_WHITE)
    _oval(slide, 0.36, 0.36, 0.64, 0.64, x=x, y=y, size=size, fill=color)


_DRAWERS = {
    "check": _draw_check,
    "warning": _draw_warning,
    "people": _draw_people,
    "document": _draw_document,
    "clock": _draw_clock,
    "target": _draw_target,
}


def add_icon(slide, x, y, size, name, *, color=PALETTE.blue):
    """アイコンを(x, y)を左上、一辺sizeの正方形として描く。

    nameは check/warning/people/document/clock/target のいずれか。
    いずれも「colorの円チップ＋白い図形」で構成し、どんな背景の上でも
    視認できるようにしている。装飾ではなく、状態・分類を示す場合に使う。
    """
    if name not in _DRAWERS:
        raise ValueError(f"未定義のiconです: {name} (候補: {sorted(ICON_NAMES)})")
    _DRAWERS[name](slide, x, y, size, color)
