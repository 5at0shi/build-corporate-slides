"""build-corporate-slidesスキルの回帰確認デッキ。

renderer/componentの不具合を1件直すたびに、その不具合を再現する内容の
スライドを1枚追加していく「生きている」テストデッキ。過去に直したものが
再びスライド上で壊れていないかを、diffのレビューだけでなく実際の見た目
（PDF）で確認できるようにする。

各スライドには「何が壊れていて、どう直したか」の説明文を必ず入れる
（レビュー用のPDFだけを見て、コードを読まなくても確認できるように）。

使い方（ワークスペース直下から）:
    ./.venv/bin/python slides/work/skill-regression-check/generate_pptx.py
    ./.claude/skills/build-corporate-slides/scripts/render_and_check.py \\
        slides/output/skill-regression-check.pptx \\
        --pdf slides/work/skill-regression-check/render/deck.pdf
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
SKILL = ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from pptx.enum.text import MSO_ANCHOR  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from slidekit import DeckBuilder, LAYOUT, PALETTE  # noqa: E402
from slidekit.components import add_background_zone, add_card  # noqa: E402
from slidekit.renderers import RENDERERS  # noqa: E402
from slidekit.typography import add_textbox  # noqa: E402


def add_fix_note(slide, area, text, *, height=Inches(0.62)):
    """「何が壊れていて、どう直したか」を説明する固定位置のノートを描く。"""
    add_background_zone(slide, area.x, area.y, area.w, height,
                        tone="neutral", rounded=True)
    add_textbox(slide, area.x + Inches(0.22), area.y + Inches(0.08),
               area.w - Inches(0.44), height - Inches(0.12),
               text, size=Pt(12.5), color=PALETTE.text_secondary,
               line_spacing=1.15)
    return area.y + height + Inches(0.14)


def build_border_radius_page(builder, page):
    """FIX 1: add_card / add_background_zoneの角丸が図形の縦横比で揃わなかった不具合。

    ROUNDED_RECTANGLEのadj値は短辺に対する割合のため、同じ値を指定しても
    縦横比が違う図形同士では見た目の角丸が揃わなかった。_flatに絶対半径
    指定(radius)を追加し、短辺から都度adj値を逆算する形にした。
    """
    slide, area = builder.add_slide("FIX 1: 角丸を絶対半径で揃える",
                                    kicker="BORDER RADIUS", page=page)
    note_bottom = add_fix_note(
        slide, area,
        "修正前: add_card（画像枠等、縦横比が正方形に近い図形）と "
        "add_background_zone（横長で背が低いパネル）は同じ角丸の指定値でも "
        "見た目の丸みがバラバラだった。修正後: どちらもLAYOUT.radius(0.08in)"
        "という絶対半径に揃うようになった。下の3つの箱はすべて縦横比が"
        "違うが、四隅の丸みが同じに見えることを確認する。",
        height=Inches(0.95))

    row = area.inset(top=note_bottom - area.y, bottom=0)
    tall, wide, medium = row.columns([1, 1.3, 1], gap="standard")

    add_card(slide, tall.x, tall.y, tall.w, tall.h)
    add_textbox(slide, tall.x, tall.y + tall.h + Inches(0.06), tall.w,
               Inches(0.3), "add_card（縦長）", size=Pt(11),
               color=PALETTE.text_secondary, align=1)

    stat_h = Inches(0.85)
    add_background_zone(slide, wide.x, wide.y, wide.w, stat_h,
                        tone="neutral", rounded=True)
    time_h = Inches(1.5)
    add_background_zone(slide, wide.x, wide.y + stat_h + Inches(0.18),
                        wide.w, time_h, tone="brand-soft", rounded=True)
    add_textbox(slide, wide.x, wide.y + stat_h + time_h + Inches(0.24), wide.w,
               Inches(0.3), "add_background_zone（横長・高さ違い2種）",
               size=Pt(11), color=PALETTE.text_secondary, align=1)

    add_card(slide, medium.x, medium.y, medium.w, medium.h / 2 - Inches(0.09))
    add_background_zone(slide, medium.x, medium.y + medium.h / 2 + Inches(0.09),
                        medium.w, medium.h / 2 - Inches(0.09),
                        tone="teal-soft", rounded=True)
    add_textbox(slide, medium.x, medium.y + medium.h + Inches(0.06), medium.w,
               Inches(0.3), "上下で違うtype（同じ半径）", size=Pt(11),
               color=PALETTE.text_secondary, align=1)


def build_process_with_gates_page(builder, page):
    """FIX 2: process_with_gatesでphasesが空だとZeroDivisionErrorで生成不能だった不具合。

    org_layers/stage_track/priority_actions等は必須リストが空ならpreflightで
    エラーにする仕組みがあったが、process_with_gatesだけphasesのチェックが
    漏れていた。preflightにチェックを追加し、レンダラー側もmax(1, len(phases))
    パターンで防御した（Escape Hatchからpreflightを経由せず直接呼ばれても
    落ちないようにするため）。
    """
    spec = {
        "title": "FIX 2: phases未指定時のクラッシュを解消",
        "kicker": "PROCESS WITH GATES",
        "primary_message": "修正前: phases:[] のときZeroDivisionErrorで生成が"
                           "止まっていた。修正後: preflightで明確なエラーメッ"
                           "セージを出し、レンダラー側も落ちなくなった（本"
                           "スライドは6フェーズ・5ゲートの通常構成が壊れて"
                           "いないことの確認用）。",
        "phases": [
            {"title": f"フェーズ{i + 1}", "label": f"STEP{i + 1}", "weight": 1,
             "items": [f"作業{i + 1}-A", f"作業{i + 1}-B"]}
            for i in range(6)
        ],
        "gates": [{"title": f"承認{i + 1}", "position": i / 4} for i in range(5)],
    }
    RENDERERS["process_with_gates"](builder, spec, page)


def build_org_layers_page(builder, page):
    """FIX 3: org_layersでlayersが3件以上だと本文がはみ出していた不具合。

    見出し分の固定オフセット(0.46in)を、layer数が増えて行の高さが縮んでも
    引き続き固定のまま使っていたため、行が縮むほど本文用の残りスペースを
    ほぼ食い潰していた。行の高さに対する上限付き割合に変更し、3件までは
    はみ出さず収まるようにした（4件以上は行の高さが構造的に足りないため、
    preflightで警告を出す運用に変更）。
    """
    spec = {
        "title": "FIX 3: layers 3件でのはみ出しを解消",
        "kicker": "ORG LAYERS",
        "primary_message": "修正前: layersが3件になった時点で本文の推定必要"
                           "高さ32ptに対し実際の枠が16pt程度しかなく、確実に"
                           "はみ出していた。修正後: 見出し分のオフセットを"
                           "行の高さに応じて可変にし、3件までは収まるように"
                           "した（4件以上はpreflightで警告）。",
        "layers": [
            {"heading": f"階層{i + 1}", "title": f"意思決定レイヤー{i + 1}",
             "body": "この階層が担う責任範囲の説明文がここに入ります"}
            for i in range(3)
        ],
        "execution_heading": "実行",
        "execution": [
            {"title": "部門A", "body": "現場での利用"},
            {"title": "部門B", "body": "現場での利用"},
        ],
    }
    RENDERERS["org_layers"](builder, spec, page)


def build_numbered_list_page(builder, page):
    """FIX 4: numbered_listで項目数が多いとスライド外へはみ出していた不具合。

    row_h(行の高さ)を0.9in固定で項目数分積み上げていたため、8項目目から
    ブロックの下端がスライドの外へ完全にはみ出し、該当行が非表示になって
    いた（validate_pptx.pyの「スライド外のオブジェクト」検知で確認済み）。
    利用可能な高さに収まるようrow_hを自動で縮めるよう修正した。
    """
    spec = {
        "title": "FIX 4: 項目多数時のスライド外はみ出しを解消",
        "kicker": "NUMBERED LIST",
        "primary_message": "修正前: 8項目目からブロック下端がスライドの外に"
                           "出て、該当行が完全に見えなくなっていた。修正後: "
                           "行の高さを利用可能な範囲に収まるよう自動で縮め、"
                           "項目数によらずスライド内に収まるようにした（本"
                           "スライドは10項目でも全行が見えることの確認用）。",
        "message_position": "bottom",
        "items": [
            {"title": f"確認項目{i + 1}", "body": "この行が最後まで見えていればOK"}
            for i in range(10)
        ],
    }
    RENDERERS["numbered_list"](builder, spec, page)


builder = DeckBuilder.from_workspace(ROOT)
builder.add_cover(
    "スキル回帰確認デッキ",
    subtitle="build-corporate-slidesスキルの修正のたびに1枚追加し、"
             "過去の不具合が再発していないか目視で確認する",
    eyebrow="SKILL REGRESSION CHECK")

build_border_radius_page(builder, page=2)
build_process_with_gates_page(builder, page=3)
build_org_layers_page(builder, page=4)
build_numbered_list_page(builder, page=5)

output = builder.save(builder.paths.output_dir / "skill-regression-check.pptx")
print("PPTX:", output)
