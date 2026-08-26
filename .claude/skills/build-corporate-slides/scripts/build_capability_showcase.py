"""スキルが標準搭載する部品を一覧できるショーケースデッキを生成する。

配色・タイポグラフィ・基本コンポーネント・アイコン・ネイティブチャート・
renderer 14種を、実際の描画関数を直接呼び出して本物の出力として並べる。
内容は常に実装と一致する（手書きの説明画像ではない）。デザインシステムや
renderer/icon/chartを追加・変更したら、このスクリプトを再実行して
参照を最新化する。

使い方（ワークスペース直下から）:
    ./.venv/bin/python .claude/skills/build-corporate-slides/scripts/build_capability_showcase.py
    ./.venv/bin/python .claude/skills/build-corporate-slides/scripts/render_and_check.py \\
        .claude/skills/build-corporate-slides/user-guide/capability-showcase.pptx \\
        --pdf .claude/skills/build-corporate-slides/user-guide/capability-showcase.pdf
"""
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
WORKSPACE_ROOT = Path.cwd()
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from slidekit import DeckBuilder, PALETTE  # noqa: E402
from slidekit.charts import add_native_chart  # noqa: E402
from slidekit.components import (add_background_zone, add_card,  # noqa: E402
                                 add_focus_panel, add_icon_list,
                                 add_item_list, add_key_message,
                                 add_numbered_row, add_section_lead)
from slidekit.icons import ICON_NAMES, add_icon  # noqa: E402
from slidekit.layout import Region, content_region  # noqa: E402
from slidekit.renderers import RENDERERS  # noqa: E402
from slidekit.theme import TYPE_BUSINESS, TYPE_DENSE, TYPE_LARGE_ROOM  # noqa: E402
from slidekit.typography import add_paragraph_textbox, add_textbox  # noqa: E402

FONT = "Hiragino Sans"


def caption(slide, x, y, w, text, *, size=10.5, bold=False, italic=False,
           color=PALETTE.text_secondary, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, Inches(0.4))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    return box


def swatch(slide, x, y, w, h, color, name, sub):
    box = slide.shapes.add_shape(1, x, y, w, h)  # 1 = MSO_SHAPE.RECTANGLE
    box.shadow.inherit = False
    box.fill.solid()
    box.fill.fore_color.rgb = color
    box.line.fill.background()
    caption(slide, x, y + h + Inches(0.06), w, name, size=11, bold=True,
           color=PALETTE.text_primary)
    caption(slide, x, y + h + Inches(0.3), w, sub, size=9.5)


def centered_block(area, block_h, *, top_offset=Inches(0)):
    """area内、intro見出し等の下（top_offset）から、block_hの高さの帯を
    残りスペースの上下中央に配置したRegionを返す。

    項目数が少ないカタログページをtopのまま余白いっぱいへ伸ばすと、
    下部だけ大きく空いて間延びして見えるため（visual-quality.md）。
    """
    top = area.y + top_offset + max(0, (area.h - top_offset - block_h) // 2)
    return Region(area.x, top, area.w, block_h)


# ---------------------------------------------------------------- builder
builder = DeckBuilder.from_workspace(WORKSPACE_ROOT)

# ================================================================ 1. Cover
builder.add_cover(
    "build-corporate-slidesスキル 機能一覧",
    subtitle="配色・コンポーネント・アイコン・チャート・renderer 14種の実物見本",
    eyebrow="CAPABILITY SHOWCASE",
    brand_side="right", brand_shape="diagonal")

# ======================================================== 2. Color Palette
slide, area = builder.add_slide("配色（Semantic Color Roles）",
                                kicker="THEME.PY: PALETTE", page=2)
caption(slide, area.x, area.y, area.w,
       "色は役割（semantic role）で選ぶ。同じ役割は常に同じ色を使い、"
       "デッキ全体の一貫性を保つ。", size=12)
block = centered_block(area, Inches(2.9), top_offset=Inches(0.5))
grid = block.rows([1, 1], gap=Inches(0.5))
colors = [
    ("text_primary", "見出し・本文の主文字"),
    ("text_secondary", "注記・補足"),
    ("line_neutral", "表・区切り・非強調の接続"),
    ("line_brand", "タイトル下線・セクションマーカー"),
    ("surface_subtle", "ニュートラルな補助面"),
    ("surface_brand_soft", "青系のまとまり・選択範囲"),
    ("surface_teal_soft", "異なる構造の補助系列"),
    ("focus_primary", "最重要項目"),
    ("accent_secondary", "構造を区別する限定的アクセント"),
]
for row_index, row in enumerate(grid):
    cols = row.columns([1] * 5, gap="standard")
    for col_index, region in enumerate(cols):
        index = row_index * 5 + col_index
        if index >= len(colors):
            break
        name, desc = colors[index]
        swatch(slide, region.x, region.y, region.w, Inches(0.7),
              getattr(PALETTE, name), name, desc)

# ======================================================= 3. Typography
slide, area = builder.add_slide("タイポグラフィ（3モード）",
                                kicker="THEME.PY: typography_for()", page=3)
caption(slide, area.x, area.y, area.w,
       "businessの本文は12ptを基準にする。denseは情報量が"
       "多いページだけに使い、large-roomは遠距離投影が明示された場合だけ"
       "使う。実物大の比較は mode-guide.pdf を参照。", size=12)
modes = [("business（標準）", TYPE_BUSINESS), ("dense", TYPE_DENSE),
        ("large-room", TYPE_LARGE_ROOM)]
mode_block = centered_block(area, Inches(2.2), top_offset=Inches(0.85))
mode_cols = mode_block.columns([1, 1, 1], gap="wide")
for (label, type_), region in zip(modes, mode_cols):
    add_background_zone(slide, region.x, region.y, region.w, Inches(2.2),
                        tone="neutral", rounded=True)
    inner = region.inset(Inches(0.24), Inches(0.2))
    add_paragraph_textbox(slide, inner.x, inner.y, inner.w, inner.h, [
        {"segments": [(label, {"size": Pt(13), "color": PALETTE.text_primary,
                               "bold": True, "font": FONT})], "space_after": 8},
        {"segments": [(f"title {type_.title.pt:g}pt", {
            "size": Pt(11), "color": PALETTE.text_secondary, "font": FONT})],
         "space_after": 4},
        {"segments": [(f"section {type_.section.pt:g}pt", {
            "size": Pt(11), "color": PALETTE.text_secondary, "font": FONT})],
         "space_after": 4},
        {"segments": [(f"body {type_.body.pt:g}pt", {
            "size": Pt(11), "color": PALETTE.text_secondary, "font": FONT})],
         "space_after": 4},
        {"segments": [(f"small {type_.small.pt:g}pt", {
            "size": Pt(11), "color": PALETTE.text_secondary, "font": FONT})]},
    ])

# =============================================== 4. Components: 面・パネル
slide, area = builder.add_slide("コンポーネント（面・パネル）",
                                kicker="COMPONENTS.PY", page=4)
block4 = centered_block(area, Inches(4.0), top_offset=Inches(0.15))
row = block4.rows([2.0, 1.5], gap=Inches(0.5))
top_cols = row[0].columns([1, 1, 1], gap="standard")
add_card(slide, top_cols[0].x, top_cols[0].y, top_cols[0].w, Inches(1.5))
caption(slide, top_cols[0].x, top_cols[0].y + Inches(1.58), top_cols[0].w,
       "add_card()", size=11, bold=True, color=PALETTE.text_primary)
caption(slide, top_cols[0].x, top_cols[0].y + Inches(1.8), top_cols[0].w,
       "独立して比較・操作する単位。細線または軽い影。", size=9.5)

add_background_zone(slide, top_cols[1].x, top_cols[1].y, top_cols[1].w,
                    Inches(1.5), tone="brand-soft", rounded=True)
caption(slide, top_cols[1].x, top_cols[1].y + Inches(1.58), top_cols[1].w,
       'add_background_zone(tone="brand-soft")', size=11, bold=True,
       color=PALETTE.text_primary)
caption(slide, top_cols[1].x, top_cols[1].y + Inches(1.8), top_cols[1].w,
       "緩やかなまとまり・範囲・フェーズ。境界・影なし。", size=9.5)

add_focus_panel(slide, top_cols[2].x, top_cols[2].y, top_cols[2].w, Inches(1.5),
                tone="solid")
caption(slide, top_cols[2].x, top_cols[2].y + Inches(1.58), top_cols[2].w,
       'add_focus_panel(tone="solid")', size=11, bold=True,
       color=PALETTE.text_primary)
caption(slide, top_cols[2].x, top_cols[2].y + Inches(1.8), top_cols[2].w,
       "ページ内の重点対象。色・線・サイズで明確化。", size=9.5)

bottom_cols = row[1].columns([1, 1, 1], gap="standard")
tones = ["neutral", "teal-soft"]
for i, tone in enumerate(tones):
    region = bottom_cols[i]
    add_background_zone(slide, region.x, region.y, region.w, Inches(1.1),
                        tone=tone, rounded=True)
    caption(slide, region.x, region.y + Inches(1.18), region.w,
           f'add_background_zone(tone="{tone}")', size=11, bold=True,
           color=PALETTE.text_primary)
region = bottom_cols[2]
add_focus_panel(slide, region.x, region.y, region.w, Inches(1.1), tone="brand")
caption(slide, region.x, region.y + Inches(1.18), region.w,
       'add_focus_panel(tone="brand")', size=11, bold=True,
       color=PALETTE.text_primary)

# ============================================ 5. Components: テキスト・リスト
slide, area = builder.add_slide("コンポーネント（テキスト・リスト）",
                                kicker="COMPONENTS.PY", page=5)
block5 = centered_block(area, Inches(4.0), top_offset=Inches(0.15))
row5 = block5.rows([2.5, 1.1], gap=Inches(0.4))
top = row5[0]
lead_col, list_col, icon_col = top.columns([0.8, 1, 1], gap="wide")
add_section_lead(slide, lead_col.x, lead_col.y, lead_col.w, "add_section_lead()")
add_item_list(slide, lead_col.x, lead_col.y + Inches(0.6), lead_col.w, Inches(1.6), [
    {"title": "見出し + 項目", "body": "add_item_listは複数項目を一つのtextboxへまとめる"}])
caption(slide, lead_col.x, lead_col.y + Inches(2.3), lead_col.w,
       "add_section_lead() + add_item_list()", size=10, bold=True,
       color=PALETTE.text_primary)

add_numbered_row(slide, list_col.x, list_col.y, list_col.w, 1,
                 "add_numbered_row()", "番号付きの単一行。numbered_listで使用")
add_numbered_row(slide, list_col.x, list_col.y + Inches(0.9), list_col.w, 2,
                 "row_hを渡すと", "項目数に応じて罫線の安全な位置を計算する")
caption(slide, list_col.x, list_col.y + Inches(1.9), list_col.w,
       "add_numbered_row()", size=10, bold=True, color=PALETTE.text_primary)

add_icon_list(slide, icon_col.x, icon_col.y, icon_col.w, Inches(1.7), [
    "add_icon_listはアイコンを行ごとに独立させる", "文章は一つのtextboxへまとめる"],
    icon="check")
caption(slide, icon_col.x, icon_col.y + Inches(1.9), icon_col.w,
       "add_icon_list()", size=10, bold=True, color=PALETTE.text_primary)

msg_row = row5[1]
msg_cols = msg_row.columns([1, 1, 1, 1], gap="standard")
for (style, label), region in zip(
        [("editorial", "editorial"), ("subtle", "subtle"),
         ("solid", "solid"), ("card", "card")], msg_cols):
    h = Inches(0.82)
    add_key_message(slide, region.x, region.y, region.w,
                    f"Key Message: {label}", style=style)
    caption(slide, region.x, region.y + h + Inches(0.08), region.w,
           f'add_key_message(style="{style}")', size=9.5)

# ==================================================================== 6. Icons
slide, area = builder.add_slide("アイコン（14種）", kicker="ICONS.PY: add_icon()",
                                page=6)
caption(slide, area.x, area.y, area.w,
       "外部アセットを使わず、python-pptxの図形だけで描くベクターアイコン。"
       "状態・分類を一目で区別させたい場合だけ使う。", size=12)
names = sorted(ICON_NAMES)
icon_block = centered_block(area, Inches(3.0), top_offset=Inches(0.55))
icon_grid = icon_block.rows([1, 1], gap=Inches(0.5))
for row_index, row in enumerate(icon_grid):
    cols = row.columns([1] * 7, gap="tight")
    for col_index, region in enumerate(cols):
        index = row_index * 7 + col_index
        if index >= len(names):
            break
        name = names[index]
        cx = region.x + region.w // 2 - Inches(0.4)
        add_icon(slide, cx, region.y, Inches(0.8), name, color=PALETTE.blue)
        caption(slide, region.x, region.y + Inches(0.95), region.w, name,
               size=10.5, align=PP_ALIGN.CENTER)

# ============================================================ 7. Charts
slide, area = builder.add_slide("ネイティブチャート（6種）",
                                kicker="CHARTS.PY: add_native_chart()", page=7)
caption(slide, area.x, area.y, area.w,
       "PNG画像と異なり、貼り付け後もPowerPoint上で数値・系列名を直接編集"
       "できる。chart_with_insightのchartフィールドから使う。", size=12)
chart_grid = area.inset(top=Inches(0.55)).rows([1, 1], gap=Inches(0.35))
chart_specs = [
    ("column", {"chart_type": "column", "categories": ["Q1", "Q2", "Q3"],
               "series": [{"name": "s1", "values": [4, 7, 6]}]}),
    ("stacked_column", {"chart_type": "stacked_column",
                        "categories": ["Q1", "Q2", "Q3"],
                        "series": [{"name": "既存", "values": [4, 5, 6]},
                                   {"name": "新規", "values": [2, 3, 4]}]}),
    ("bar", {"chart_type": "bar", "categories": ["A", "B", "C"],
            "series": [{"name": "s1", "values": [5, 8, 3]}]}),
    ("line", {"chart_type": "line", "categories": ["4月", "5月", "6月"],
             "series": [{"name": "s1", "values": [10, 14, 12]}]}),
    ("pie", {"chart_type": "pie", "categories": ["A", "B", "C"],
            "series": [{"name": "s1", "values": [5, 3, 2]}]}),
    ("scatter", {"chart_type": "scatter", "series": [
        {"name": "s1", "points": [{"x": 1, "y": 2}, {"x": 2, "y": 3.4},
                                  {"x": 3, "y": 1.8}]}]}),
]
for row_index, row in enumerate(chart_grid):
    cols = row.columns([1, 1, 1], gap="standard")
    for col_index, region in enumerate(cols):
        index = row_index * 3 + col_index
        if index >= len(chart_specs):
            break
        label, kwargs = chart_specs[index]
        chart_h = region.h - Inches(0.3)
        add_native_chart(slide, region.x, region.y, region.w, chart_h,
                         typography=TYPE_BUSINESS, **kwargs)
        caption(slide, region.x, region.y + chart_h + Inches(0.04), region.w,
               f'chart.type: "{label}"', size=10, bold=True,
               color=PALETTE.text_primary, align=PP_ALIGN.CENTER)


# =================================================== 8-21. Renderer Catalog
# renderer-catalog.mdの「使う状況」列と同じ文言を使い、ドキュメントと
# 実物を対応させる。1ページのcontent_region()内に収まる、控えめな
# 現実的な分量で作る（stress-test用の極端な分量ではない）。
FOOTER_Y = Inches(7.5) - Inches(0.32)


def renderer_footer(type_name, desc):
    slide = builder.presentation.slides[-1]
    caption(slide, Inches(0.67), FOOTER_Y, Inches(11.99),
           f"renderer: “{type_name}” — {desc}", size=10,
           italic=True)


renderer_specs = [
    ("comparison", "二つの観点、価値とリスク、現状と将来", {
        "title": "なぜ今、生成AIを検証するのか", "density": "standard",
        "primary_message": "PoCで、会社として安全に成果を出せる使い方を決める",
        "left": {"heading": "期待される業務価値", "items": [
            {"title": "作業時間の短縮", "body": "文書作成や要約にかかる時間を削減"},
            {"title": "品質の標準化", "body": "担当者による成果物のばらつきを抑制"}]},
        "right": {"heading": "放置した場合のリスク", "items": [
            {"title": "個別利用の拡大", "body": "管理されない利用が先行する"},
            {"title": "判断基準の不在", "body": "部門ごとに異なる運用が定着する"}]},
    }),
    ("evidence_and_decision", "根拠から一つの推奨判断を導く", {
        "title": "初回PoCは文書業務から始める", "density": "standard",
        "primary_message": "文書要約とドラフト作成を、最初の検証対象とする",
        "evidence_heading": "選定の根拠", "evidence": [
            {"title": "利用頻度が高い", "body": "日常業務で繰り返し発生し、効果を継続的に測れる"},
            {"title": "比較条件をそろえやすい", "body": "現行工数と生成結果を同じ案件で比較できる"}],
        "decision_heading": "推奨する初回スコープ",
        "decision_detail": "効果測定とリスク管理を両立しやすく、10週間で継続判断に必要な材料をそろえられる。",
    }),
    ("scope_and_exclusions", "対象範囲と対象外を同時に示す", {
        "title": "検証範囲を限定し、安全性と効果を同時に確認する", "density": "dense",
        "primary_message": "AIは下書きまで。判断と外部利用は人が担う",
        "scope_heading": "PoC実施範囲", "scope": [
            {"label": "INPUT", "title": "限定した社内文書", "body": "機密区分と利用者を事前に定める"},
            {"label": "OUTPUT", "title": "人が最終確認", "body": "判断責任と外部送信は担当者に残す"}],
        "period": "期間：準備2週｜実証6週｜評価2週",
        "exclusions_heading": "今回の対象外",
        "exclusions": ["自動意思決定", "顧客への直接送信", "個人・要配慮情報"],
    }),
    ("process_with_gates", "フェーズ、作業、判断時点", {
        "title": "10週間で検証し、段階的に継続可否を判断する", "density": "dense",
        "primary_message": "各ゲートで安全性と効果を確認し、条件を満たす場合だけ次へ進む",
        "phases": [
            {"title": "準備", "label": "SETUP", "weight": 2,
             "items": ["対象業務と利用者を確定", "基準値と運用ルールを設定"]},
            {"title": "実証", "label": "EXECUTION", "weight": 3,
             "items": ["実案件で生成AIを利用", "時間・品質・修正内容を記録"]},
            {"title": "評価", "label": "EVALUATION", "weight": 2,
             "items": ["効果とリスクを総合評価", "継続・改善・終了を判断"]}],
        "gates": [{"title": "利用開始承認", "position": 0.28},
                  {"title": "最終判定会議", "position": 1.0}],
    }),
    ("table_with_conclusion", "条件、比較、評価を表で読む", {
        "title": "効果と安全性の両方を満たす場合だけ継続する", "density": "dense",
        "primary_message": "必須条件を一つでも満たさない場合は、改善または終了を判断する",
        "columns": [{"key": "criterion", "label": "評価項目", "weight": 1.6},
                    {"key": "measure", "label": "確認方法", "weight": 2.5},
                    {"key": "decision", "label": "判断", "weight": 1, "align": "center"}],
        "rows": [
            {"criterion": "時間短縮", "measure": "現行業務との所要時間比較",
             "decision": "必須", "_highlight": True},
            {"criterion": "品質維持", "measure": "担当者レビューと修正量", "decision": "必須"},
            {"criterion": "利用継続性", "measure": "利用率と担当者ヒアリング", "decision": "参考"}],
    }),
    ("org_layers", "意思決定・運営など縦の責任階層と、横に並ぶ実行部門", {
        "title": "意思決定と実行の役割を分けて運用する", "density": "standard",
        "primary_message": "現場の裁量と全社統制のバランスを取る",
        "layers": [
            {"heading": "意思決定", "title": "運営委員会", "body": "投資判断と全社ルールを決定"},
            {"heading": "運営", "title": "推進事務局", "body": "各部門の進捗と課題を集約"}],
        "execution_heading": "業務実行",
        "execution": [{"title": "営業部門", "body": "提案書作成での試行"},
                      {"title": "管理部門", "body": "社内文書要約での試行"}],
    }),
    ("priority_actions", "優先度付きの課題と、対応する方針", {
        "title": "重大リスクから優先して対応する", "density": "standard",
        "primary_message": "情報漏えいリスクへの対応を最優先で整備する",
        "issues": [
            {"priority": "最優先", "title": "情報漏えい", "body": "機密情報の入力を防ぐ運用が未整備"},
            {"priority": "中", "title": "誤情報の混入", "body": "生成結果を検証せず利用する懸念"}],
        "actions": ["入力可能な情報区分を明文化する", "生成結果は担当者が必ず確認する"],
    }),
    ("stage_track", "現在から将来への段階的な進行（ロードマップ等）", {
        "title": "PoCから全社展開まで段階的に広げる", "density": "standard",
        "primary_message": "各段階の基準を満たしてから次へ進む",
        "stages": [
            {"label": "STEP1", "title": "PoC", "body": "対象業務を限定して検証"},
            {"label": "STEP2", "title": "部門展開", "body": "関連部門へ展開し運用を固める"},
            {"label": "STEP3", "title": "全社展開", "body": "全部門へ展開し効果を測定"}],
    }),
    ("numbered_list", "アジェンダ、依頼事項など番号付きの単列項目", {
        "title": "ご承認のお願い", "density": "standard",
        "primary_message": "PoC実施予算と実施期間のご承認をお願いします",
        "message_position": "bottom",
        "items": [
            {"title": "実施予算の承認", "body": "ツール利用料・導入支援費用"},
            {"title": "対象部門の確定", "body": "営業部門・管理部門の2部門"},
            {"title": "実施期間の承認", "body": "準備2週・実証6週・評価2週"}],
    }),
    ("matrix_2x2", "2軸で選択肢を4象限に整理する（ポートフォリオ分析等）", {
        "title": "効果とリスクで施策の優先度を整理する", "density": "standard",
        "primary_message": "効果が高くリスクが低い施策から着手する",
        "x_axis": {"low": "低リスク", "high": "高リスク"},
        "y_axis": {"low": "低効果", "high": "高効果"},
        "quadrants": [
            {"label": "優先", "title": "文書要約", "body": "効果大・統制しやすい", "emphasis": True},
            {"label": "検討", "title": "外部連携文書", "body": "効果はあるが要件整理が必要"},
            {"label": "保留", "title": "個人メモ支援", "body": "効果は限定的"},
            {"label": "対象外", "title": "顧客への自動返信", "body": "現時点ではリスクが高い"}],
    }),
    ("stat_highlight", "単一の実績数値を主役に、補足指標と結論を示す", {
        "title": "PoCで作業時間の削減効果を確認した", "density": "standard",
        "primary_message": "本格導入に向けた投資対効果は十分に見込める",
        "stat": {"value": "-42%", "label": "対象業務の作業時間削減率",
                "detail": "文書要約・ドラフト作成の合算値"},
        "supporting": [{"value": "94%", "label": "利用者の継続利用意向"},
                       {"value": "0件", "label": "情報漏えい・誤送信"}],
    }),
]

page = 8
RENDERERS["section_divider"](builder, {
    "title": "renderer カタログ（14種）", "kicker": "SECTION DIVIDER",
    "subtitle": "renderer-catalog.mdの各typeを、実際の最小構成で見せる",
}, page)
renderer_footer("section_divider", "複数テーマを扱う資料の章区切り")
page += 1

for type_name, desc, spec in renderer_specs:
    RENDERERS[type_name](builder, spec, page)
    renderer_footer(type_name, desc)
    page += 1

# chart_with_insightは2 variantあるため2枚見せる
chart_with_insight_specs = [
    ("standard", "グラフ全体を公平に読む", {
        "title": "効果の全体像を確認し、継続判断につなげる", "density": "standard",
        "variant": "standard",
        "primary_message": "時間短縮だけでなく、品質と利用継続性を合わせて評価する",
        "insight_heading": "確認するポイント",
        "insights": ["業務別の時間短縮率", "修正量と品質評価", "継続利用の意向"],
        "chart": {"type": "column", "categories": ["準備", "実証", "評価"],
                 "series": [{"name": "削減率(%)", "values": [12, 38, 42]}]},
    }),
    ("conclusion-led", "一つの主張を強く伝える", {
        "title": "実証期間で削減率が拡大した", "density": "standard",
        "variant": "conclusion-led",
        "primary_message": "実証が進むほど削減率が拡大し、定着の手応えが確認できた",
        "insight_heading": "読み取れること",
        "insights": ["準備期は効果が小さい", "実証後半で効果が加速"],
        "chart": {"type": "line", "categories": ["準備", "実証前半", "実証後半", "評価"],
                 "series": [{"name": "削減率(%)", "values": [5, 22, 38, 42]}]},
    }),
]
for variant, desc, spec in chart_with_insight_specs:
    RENDERERS["chart_with_insight"](builder, spec, page)
    renderer_footer(f"chart_with_insight ({variant})", desc)
    page += 1

output = builder.save(SKILL / "user-guide" / "capability-showcase.pptx")
print("PPTX:", output)
