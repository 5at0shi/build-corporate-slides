"""Layout層: 矩形領域(Region)の分割・inset計算のみを行う、最下層の
純粋な幾何層。描画は一切行わない（塗り・線・文字を持たない）。

Atom以上のすべての層はRegionを受け取って初めて描画位置を持てる。
通常はrenderer層がcontent_region()を起点に内部で使うが、Regionと
content_regionは__all__に含まれ、該当rendererが無いページを個別構築
する場合（renderer-catalog.mdのEscape Hatch）は生成スクリプトから
直接使ってよい。層構成の全体像はARCHITECTURE.mdを参照。
"""
from dataclasses import dataclass

from pptx.util import Inches

from .theme import LAYOUT


def _gap(value):
    if isinstance(value, str):
        values = {
            "tight": Inches(0.14),
            "standard": LAYOUT.gap_x,
            "wide": Inches(0.42),
        }
        if value not in values:
            raise ValueError(f"未定義のgapです: {value}")
        return values[value]
    return value


def _positive_weights(weights):
    """重みの合計が0以下なら等分へ落とす。

    weightはYAMLから来る値（process_with_gatesのphases[].weight等）のため、
    全て0や負が渡されうる。合計0でそのまま除算するとZeroDivisionErrorで
    デッキ全体の生成が止まる。合計が正であれば個々の0はそのまま扱う
    （[0, 1]で片方を潰すのは意図した指定でありうるため）。
    """
    if not weights:
        return [1]
    return list(weights) if sum(weights) > 0 else [1] * len(weights)


@dataclass(frozen=True)
class Region:
    x: int
    y: int
    w: int
    h: int

    def inset(self, left=0, top=0, right=None, bottom=None):
        right = left if right is None else right
        bottom = top if bottom is None else bottom
        return Region(self.x + left, self.y + top,
                      self.w - left - right, self.h - top - bottom)

    def columns(self, weights, gap="standard"):
        gap = _gap(gap)
        weights = _positive_weights(weights)
        total_gap = gap * (len(weights) - 1)
        unit = (self.w - total_gap) / sum(weights)
        result, cursor = [], self.x
        for weight in weights:
            width = int(unit * weight)
            result.append(Region(int(cursor), self.y, width, self.h))
            cursor += width + gap
        return result

    def rows(self, weights, gap=Inches(0.18)):
        gap = _gap(gap)
        weights = _positive_weights(weights)
        total_gap = gap * (len(weights) - 1)
        unit = (self.h - total_gap) / sum(weights)
        result, cursor = [], self.y
        for weight in weights:
            height = int(unit * weight)
            result.append(Region(self.x, int(cursor), self.w, height))
            cursor += height + gap
        return result


def content_region(*, top=Inches(1.46), bottom=Inches(0.48)):
    return Region(
        LAYOUT.margin_x,
        top,
        LAYOUT.slide_width - 2 * LAYOUT.margin_x,
        LAYOUT.slide_height - top - bottom,
    )
