"""アイコン機能の試作: カレーのレシピを4枚のスライドで作る。

構成（1: 表紙 / 2: 材料 / 3: 作り方 / 4: まとめ）のうち、材料とまとめは
renderer-catalogに一致する型がないため、DeckBuilderとRegionで個別構築する
（SKILL.mdのEscape Hatch）。既存のslides/work/generate_pptx.pyとは別の
一回限りの試作のため、curry-demo/ に分けて既存デッキを上書きしない。

使い方（ワークスペース直下から）:
    ./.venv/bin/python slides/work/curry-demo/generate_pptx.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
SKILL = ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from pptx.util import Inches  # noqa: E402

from slidekit import (DeckBuilder, PALETTE, Region, add_card, add_icon,  # noqa: E402
                      add_icon_list, add_item_list, add_key_message,
                      add_panel, add_section_lead, add_textbox)
from slidekit.renderers import RENDERERS  # noqa: E402
from slidekit.textmetrics import estimate_icon_list_height_pt  # noqa: E402

HEADING_BLOCK_H = Inches(0.62)
ICON_LIST_ICON_SIZE = Inches(0.34)
ICON_LIST_TEXT_GAP = Inches(0.16)


def add_stat_chip(slide, region, icon_name, label, typography):
    icon_size = Inches(0.5)
    add_icon(slide, region.x, region.y, icon_size, icon_name, color=PALETTE.blue)
    text_x = region.x + icon_size + Inches(0.18)
    add_textbox(slide, text_x, region.y + Inches(0.13),
                region.w - icon_size - Inches(0.18), Inches(0.3), label,
                size=typography.section, color=PALETTE.text_primary,
                bold=True, font=typography.body_font)


def add_ingredient_card(slide, region, heading, items, *, color=PALETTE.line_brand,
                        bullet="•"):
    add_card(slide, region.x, region.y, region.w, region.h)
    inner = region.inset(Inches(0.32), Inches(0.26))
    add_section_lead(slide, inner.x, inner.y, inner.w, heading, color=color)
    add_item_list(slide, inner.x, inner.y + HEADING_BLOCK_H, inner.w,
                  inner.h - HEADING_BLOCK_H, items, bullet=bullet)


def build_ingredients_page(builder, page):
    slide, area = builder.add_slide("材料", kicker="INGREDIENTS", page=page)
    typography = slide._slidekit_typography

    stat_h = Inches(0.9)
    stat_inner = add_panel(slide, area.x, area.y, area.w, stat_h,
                           tone="neutral", inset_y=Inches(0.2))
    chips = stat_inner.columns([1, 1, 1], gap="wide")
    add_stat_chip(slide, chips[0], "people", "4人分", typography)
    add_stat_chip(slide, chips[1], "clock", "調理時間 45分", typography)
    add_stat_chip(slide, chips[2], "document", "材料 10点", typography)

    cards_y = area.y + stat_h + Inches(0.28)
    cards_region = Region(area.x, cards_y, area.w, area.y + area.h - cards_y)
    left, right = cards_region.columns([1, 1], gap="wide")

    add_ingredient_card(slide, left, "肉・野菜", [
        "鶏もも肉 300g", "玉ねぎ 2個", "にんじん 1本",
        "じゃがいも 2個", "にんにく・しょうが 各1片",
    ])
    add_ingredient_card(slide, right, "スパイス・調味料", [
        "カレールー 1/2箱", "クミンシード 小さじ1", "サラダ油 大さじ1",
        "水 600ml", "塩 少々",
    ], color=PALETTE.accent_secondary, bullet="—")


def add_centered_icon_block(slide, region, heading, items, *, icon, icon_color,
                            heading_color, typography, body_gap=48):
    """見出し＋アイコンリストを一つのブロックとして領域内で上下中央に置く。

    見出しだけを上端に固定してリストの下だけに余白を残すと上寄りに見え、
    逆にリストだけを中央寄せすると見出しから浮いて見える（過去に検証済み）。
    見出しとリストを一体のブロックとして中央寄せすることで両方を避ける。

    body_gapは既定の14ptより広めにしてある。中央寄せで生まれる上下の
    余白に対し、項目同士の間隔が既定値のままだと狭く窮屈に見えるため
    （周囲の余白と項目間の余白の対比が大きすぎる、との指摘を受けて調整）。
    """
    list_h_pt = estimate_icon_list_height_pt(
        typography, items, region.w / 12700,
        icon_size_pt=ICON_LIST_ICON_SIZE / 12700,
        text_gap_pt=ICON_LIST_TEXT_GAP / 12700, body_gap=body_gap)
    list_h = Inches(list_h_pt / 72)
    block_h = HEADING_BLOCK_H + list_h
    y = region.y + max(0, int((region.h - block_h) / 2))
    add_section_lead(slide, region.x, y, region.w, heading, color=heading_color)
    add_icon_list(slide, region.x, y + HEADING_BLOCK_H, region.w, list_h, items,
                 icon=icon, icon_color=icon_color, icon_size=ICON_LIST_ICON_SIZE,
                 text_gap=ICON_LIST_TEXT_GAP, body_gap=body_gap)


def build_summary_page(builder, page):
    slide, area = builder.add_slide("まとめ：なぜカレーか", kicker="SUMMARY", page=page)
    typography = slide._slidekit_typography

    body, conclusion = area.rows([4.35, 0.72], gap=Inches(0.24))
    left, right = body.columns([1, 1], gap="wide")

    add_centered_icon_block(slide, left, "メリット", [
        "一度に大量に作れて、常備菜として作り置きできる",
        "野菜とたんぱく質を一皿で無理なく摂れる",
        "一晩置くと味がなじみ、翌日はさらに美味しくなる",
    ], icon="check", icon_color=PALETTE.blue, heading_color=PALETTE.line_brand,
       typography=typography)

    add_centered_icon_block(slide, right, "デメリット", [
        "煮込みに時間がかかり、45分程度は必要になる",
        "スパイス・ルーの在庫管理や買い足しが必要",
        "作り置き分は保存状態の管理に注意がいる",
    ], icon="warning", icon_color=PALETTE.accent_secondary,
       heading_color=PALETTE.accent_secondary, typography=typography)

    add_key_message(
        slide, conclusion.x, conclusion.y, conclusion.w,
        "手間はかかるが、作り置きのしやすさと栄養バランスの良さから定番になる。",
        style="subtle")


builder = DeckBuilder.from_workspace(ROOT)
builder.add_cover(
    "カレーの作り方", subtitle="基本の材料と手順で仕上げる、家庭のスパイスカレー",
    eyebrow="RECIPE GUIDE")
build_ingredients_page(builder, page=2)

RENDERERS["numbered_list"](builder, {
    "title": "作り方",
    "primary_message": "合計調理時間の目安は45分。弱火でじっくり煮込むと味がまとまりやすい。",
    "items": [
        {"title": "下ごしらえ",
         "body": "鶏肉と野菜を一口大に切り、にんにく・しょうがはみじん切りにする。"},
        {"title": "香りを立てる",
         "body": "鍋に油とクミンシードを入れて熱し、香りが立ったら玉ねぎを飴色になるまで炒める。"},
        {"title": "肉と野菜を炒める",
         "body": "鶏肉の色が変わったら、にんじん・じゃがいも・にんにく・しょうがを加えて炒め合わせる。"},
        {"title": "煮込む",
         "body": "水を加えて沸騰させ、アクを取りながら中火で15分煮込む。"},
        {"title": "ルーを溶かす",
         "body": "火を止めてカレールーを溶かし入れ、弱火でとろみがつくまで10分煮る。"},
    ],
}, 3)

build_summary_page(builder, page=4)

output = builder.save(builder.paths.output_dir / "curry-demo.pptx")
print("PPTX:", output)
