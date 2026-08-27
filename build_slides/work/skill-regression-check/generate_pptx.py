"""build-corporate-slidesスキルの回帰確認デッキ。

renderer/componentの不具合を1件直すたびに、その不具合を再現する内容の
スライドを1枚追加していく「生きている」テストデッキ。過去に直したものが
再びスライド上で壊れていないかを、diffのレビューだけでなく実際の見た目
（PDF）で確認できるようにする。

各スライドには「何が壊れていて、どう直したか」の説明文を必ず入れる
（レビュー用のPDFだけを見て、コードを読まなくても確認できるように）。

使い方（ワークスペース直下から）:
    ./.venv/bin/python build_slides/work/skill-regression-check/generate_pptx.py
    ./.claude/skills/build-corporate-slides/scripts/render_and_check.py \\
        build_slides/output/skill-regression-check.pptx \\
        --pdf build_slides/work/skill-regression-check/render/deck.pdf
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
    """FIX 2: process_with_gates。

    (a) phasesが空だとZeroDivisionErrorで生成不能だった。preflightに
    必須チェックを追加し、レンダラー側もmax(1, len(phases))パターンで
    防御した（Escape Hatchからpreflightを経由せず直接呼ばれても落ちない
    ようにするため）。
    (b) ゲートの点が端(position 0 / 1付近)にあると、ラベルを中央揃えの
    まま行の外へクランプしていたためラベルだけ点から離れて見えた。端に
    近い点はラベルを点の位置を起点に片側へ伸ばす（左端はLEFT揃え、右端は
    RIGHT揃え）よう修正し、常に点とラベルが対応するようにした。
    """
    spec = {
        "title": "FIX 2: phasesクラッシュとゲートラベルのズレを解消",
        "kicker": "PROCESS WITH GATES",
        "primary_message": "(a) phases:[]でのクラッシュをpreflightで防止。"
                           "(b) 端の承認ラベルが点から離れていたのを、点を"
                           "起点に片側へ伸ばす揃え方に変更して解消（全ての"
                           "承認ラベルが点の真下・真上に揃っていることを"
                           "確認する）。",
        "phases": [
            {"title": f"フェーズ{i + 1}", "label": f"STEP{i + 1}", "weight": 1,
             "items": [f"作業{i + 1}-A", f"作業{i + 1}-B"]}
            for i in range(6)
        ],
        "gates": [{"title": f"承認{i + 1}", "position": i / 4} for i in range(5)],
    }
    RENDERERS["process_with_gates"](builder, spec, page)


def build_org_layers_page(builder, page):
    """FIX 3: org_layersのsection_leadマーカーが本文に被っていた不具合。

    layer数に応じて行の高さが縮むケースで見出し分のオフセットも縮める
    ようにしたが、初回修正では見出しの縦棒マーカー自体は固定0.38inの
    ままで、オフセットだけそれより短く縮めてしまい、マーカーの下側が
    本文へ被っていた。add_section_leadにmarker_hを渡せるようにし、
    「オフセット = マーカー高さ + 固定ギャップ」という一つの式で連動
    させることで、マーカーと本文が構造的に重ならないようにした。
    その結果、3layersでは本文が枠に収まりきらないことが分かったため、
    preflightの警告閾値も2layers超に修正した（本スライドは2layersで
    重なりなく収まることの確認用）。
    """
    spec = {
        "title": "FIX 3: 見出しマーカーが本文に被る不具合を解消",
        "kicker": "ORG LAYERS",
        "primary_message": "修正前: 見出しの縦棒マーカー(固定0.38in)より"
                           "オフセットの方が短くなり得て、マーカー下側が"
                           "本文に被っていた。修正後: オフセットをマーカー"
                           "高さ+ギャップで連動させ、構造的に被らないよう"
                           "にした（3layers以上は本文が収まらないため"
                           "preflightで警告する運用に変更）。",
        "layers": [
            {"heading": f"階層{i + 1}", "title": f"意思決定レイヤー{i + 1}",
             "body": "この階層が担う責任範囲の説明文がここに入ります"}
            for i in range(2)
        ],
        "execution_heading": "実行",
        "execution": [
            {"title": "部門A", "body": "現場での利用"},
            {"title": "部門B", "body": "現場での利用"},
        ],
    }
    RENDERERS["org_layers"](builder, spec, page)


def build_numbered_list_page(builder, page):
    """FIX 4: numbered_listの罫線が本文に被っていた不具合。

    項目多数時のスライド外はみ出しをrow_hの自動縮小で直したが、初回修正
    では罫線の位置を「次の行に食い込まない」よう次行開始位置手前へ
    クランプするだけで、行の高さそのものが本文2行分に足りない場合に
    「自分の行の本文」に罫線が被る問題が残っていた。行の高さに罫線を
    安全に置ける余白が無い場合は、中途半端な位置に引いて文字に被せる
    より罫線そのものを省略するよう修正した。
    """
    spec = {
        "title": "FIX 4: 罫線が本文に被る不具合を解消",
        "kicker": "NUMBERED LIST",
        "primary_message": "修正前: 罫線の位置を次の行に被らないようクラン"
                           "プするだけで、自分の行の本文に被るケースが残"
                           "っていた。修正後: 罫線を安全に置ける余白が無い"
                           "行では罫線自体を省略する（本スライドは10項目"
                           "で、文字に重なる罫線が無いことの確認用）。",
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
