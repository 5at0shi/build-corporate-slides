from dataclasses import dataclass, replace

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def rgb(hex_value: str) -> RGBColor:
    value = hex_value.lstrip("#")
    return RGBColor(*(int(value[i:i + 2], 16) for i in (0, 2, 4)))


@dataclass(frozen=True)
class Palette:
    text_primary: RGBColor = rgb("102C49")
    text_secondary: RGBColor = rgb("53657A")
    line_neutral: RGBColor = rgb("D7DEE6")
    line_brand: RGBColor = rgb("2574D9")
    surface_base: RGBColor = rgb("FFFFFF")
    surface_subtle: RGBColor = rgb("E7EBF0")
    surface_brand_soft: RGBColor = rgb("E1EDFC")
    surface_teal_soft: RGBColor = rgb("E2F2F0")
    focus_primary: RGBColor = rgb("102C49")
    accent_secondary: RGBColor = rgb("148C9A")
    white: RGBColor = rgb("FFFFFF")

    # v0.4との後方互換。新規コードではsemantic roleを使う。
    navy: RGBColor = rgb("102C49")
    blue: RGBColor = rgb("2574D9")
    ink: RGBColor = rgb("102C49")
    grey_700: RGBColor = rgb("53657A")
    grey_500: RGBColor = rgb("738296")
    grey_300: RGBColor = rgb("D7DEE6")
    grey_100: RGBColor = rgb("E7EBF0")


@dataclass(frozen=True)
class BusinessTypography:
    """本文（body）は基本サイズ12ptを基準に、比率を保って設計する。

    注釈・キャプション（small）は元々12pt未満が前提の階層なので対象外。
    表紙・章扉などポスター的な大見出しはこのtitleを使わず、別途固定値を持つ
    （add_cover, add_section_divider参照）。
    """
    headline_font: str = "Yu Gothic"
    editorial_font: str = "Yu Mincho"
    body_font: str = "Yu Gothic"
    title: Pt = Pt(22)
    section: Pt = Pt(15)
    body: Pt = Pt(12)
    small: Pt = Pt(10)
    metric: Pt = Pt(20)


@dataclass(frozen=True)
class DenseTypography(BusinessTypography):
    title: Pt = Pt(20)
    section: Pt = Pt(15)
    body: Pt = Pt(11.5)
    small: Pt = Pt(9)
    metric: Pt = Pt(18)


@dataclass(frozen=True)
class LargeRoomTypography(BusinessTypography):
    title: Pt = Pt(32)
    section: Pt = Pt(20)
    body: Pt = Pt(17)
    small: Pt = Pt(12)
    metric: Pt = Pt(30)


@dataclass(frozen=True)
class Layout:
    """余白・角丸は4pt刻みの基準単位で揃える（デザイントークンの一貫性のため）。

    角丸はスケール化する: radius_base（Card等の標準コンテナ）と
    radius_lg（Background Zone等、画面の大部分を占める面）の2段階。
    同じ絶対半径でも小さい面では丸すぎ、大きい面では丸みが足りなく
    見える（知覚上のズレ）ため、役割に応じて使い分ける。完全な丸薬型
    （Tag等）は固定値を持たず、高さ/2から都度算出する。
    """
    slide_width = Inches(13.333333)
    slide_height = Inches(7.5)
    margin_x = Inches(48 / 72)
    margin_top = Inches(32 / 72)
    margin_bottom = Inches(32 / 72)
    gap_x = Inches(20 / 72)
    gap_y = Inches(16 / 72)
    radius_base = Pt(8)
    radius_lg = Pt(14)


PALETTE = Palette()
TYPE = BusinessTypography()
TYPE_BUSINESS = TYPE
TYPE_DENSE = DenseTypography()
TYPE_LARGE_ROOM = LargeRoomTypography()
# v0.6以前のimportとの互換性。新規コードではTYPE_BUSINESS等を使う。
TYPE_PRESENTATION = TYPE_LARGE_ROOM
LAYOUT = Layout()


def typography_for(mode: str = "business", fonts=None) -> BusinessTypography:
    aliases = {"presentation": "large-room", "document": "dense"}
    mode = aliases.get(mode, mode)
    modes = {
        "business": TYPE_BUSINESS,
        "dense": TYPE_DENSE,
        "large-room": TYPE_LARGE_ROOM,
    }
    if mode not in modes:
        raise ValueError("modeは 'business'、'dense'、'large-room' を指定してください")
    typography = modes[mode]
    if fonts:
        supported = {"headline_font", "body_font", "editorial_font"}
        values = {key: value for key, value in fonts.items()
                  if key in supported and value}
        if values:
            typography = replace(typography, **values)
    return typography
