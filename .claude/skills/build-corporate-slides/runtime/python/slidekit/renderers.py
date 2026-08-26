from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .atoms import Marker
from .builder import DeckBuilder
from .charts import add_native_chart
from .components import (SECTION_LEAD_GAP, add_background_zone, add_card,
                         add_hairline, add_item_list, add_key_message,
                         add_panel, add_section_lead)
from .preflight import require_valid_content
from .images import add_image_contain
from .tables import add_data_table
from .textmetrics import adaptive_gap_pt, estimate_item_list_height_pt
from .theme import PALETTE
from .typography import (Stat, add_paragraph_textbox, add_text_list,
                         add_textbox)


HEADING_BLOCK_H = Inches(0.62)


def _type(slide):
    return slide._slidekit_typography


def _items(section):
    if isinstance(section, list):
        return section
    return section.get("items", [])


def _heading(section, fallback):
    return section.get("heading", fallback) if isinstance(section, dict) else fallback


def _adaptive_gap(items, available_h, *, base_gap, per_item_pt):
    if not items:
        return base_gap
    content_pt = per_item_pt * len(items)
    return int(adaptive_gap_pt(content_pt, len(items), available_h / 12700,
                               base_gap=base_gap))


def render_cover(builder, spec, page):
    brand_width = spec.get("brand_width")
    builder.add_cover(
        spec["title"],
        subtitle=spec.get("subtitle"),
        eyebrow=spec.get("eyebrow"),
        brand_side=spec.get("brand_side", "right"),
        brand_shape=spec.get("brand_shape", "diagonal"),
        brand_width=Inches(brand_width) if brand_width else None,
        classification=spec.get("classification"),
        created=spec.get("created"),
    )


def render_comparison(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    body, conclusion = area.rows([4.35, 0.72], gap=Inches(0.24))
    variant = spec.get("variant", "balanced")
    weights = [1, 1] if variant == "balanced" else [1.08, 0.92]
    left, right = body.columns(weights, gap="wide")
    if variant == "asymmetric":
        add_background_zone(slide, right.x - Inches(0.18), right.y - Inches(0.08),
                            right.w + Inches(0.36), right.h + Inches(0.12),
                            tone="neutral", rounded=True)
    left_items = _items(spec.get("left", {}))
    right_items = _items(spec.get("right", {}))
    add_section_lead(slide, left.x, left.y, left.w,
                     _heading(spec.get("left", {}), "左側"))
    add_section_lead(slide, right.x, right.y, right.w,
                     _heading(spec.get("right", {}), "右側"),
                     color=PALETTE.accent_secondary)
    add_item_list(slide, left.x, left.y + HEADING_BLOCK_H, left.w,
                  left.h - HEADING_BLOCK_H - Inches(0.06), left_items)
    add_item_list(slide, right.x, right.y + HEADING_BLOCK_H, right.w,
                  right.h - HEADING_BLOCK_H - Inches(0.06),
                  right_items, bullet="—")
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def render_evidence_and_decision(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    left, right = area.columns([1.55, 1], gap="wide")
    evidence = spec.get("evidence", [])

    # 右の推奨パネルと同じ「意図的な余白を持つ面」に見せるため、左も淡色の
    # 面で包み、見出し+リストのブロックをその中で釣り合わせる。背景のない
    # 剥き出しの文字のまま下部が空くと、右のパネルとの対比で「欠けている」
    # ように見えるため。
    left_inner = add_panel(slide, left.x - Inches(0.12), left.y - Inches(0.06),
                           left.w + Inches(0.24), left.h + Inches(0.12),
                           tone="neutral")
    list_pt = estimate_item_list_height_pt(
        typography, evidence, left_inner.w / 12700)
    block_h = HEADING_BLOCK_H + int(list_pt * 12700)
    top = left_inner.y + max(0, (left_inner.h - block_h) // 2)
    add_section_lead(slide, left_inner.x, top, left_inner.w,
                     spec.get("evidence_heading", "判断の根拠"))
    add_item_list(slide, left_inner.x, top + HEADING_BLOCK_H, left_inner.w,
                  left_inner.y + left_inner.h - (top + HEADING_BLOCK_H), evidence,
                  adaptive=False)

    inner = add_panel(slide, right.x, right.y, right.w, right.h,
                      tone="brand-soft", inset_x=Inches(0.34), inset_y=Inches(0.3))
    paragraphs = [
        {"segments": [(spec.get("decision_heading", "推奨方針"), {
            "size": typography.small, "color": PALETTE.blue, "bold": True,
            "font": typography.body_font,
        })], "space_after": 10},
        {"segments": [(spec["primary_message"], {
            "size": typography.section, "color": PALETTE.text_primary,
            "bold": True, "font": typography.headline_font,
        })], "space_after": 14 if spec.get("decision_detail") else 0,
         "line_spacing": 1.05},
    ]
    if spec.get("decision_detail"):
        paragraphs.append({"segments": [(spec["decision_detail"], {
            "size": typography.body, "color": PALETTE.text_secondary,
            "font": typography.body_font,
        })], "line_spacing": 1.15})
    add_paragraph_textbox(slide, inner.x, inner.y, inner.w, inner.h,
                          paragraphs, vertical_anchor=MSO_ANCHOR.MIDDLE)


def render_scope_and_exclusions(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    scope, exclusions = area.columns([2.2, 1], gap="wide")
    add_background_zone(slide, scope.x, scope.y, scope.w, scope.h,
                        tone="brand-soft", rounded=True)
    scope_inner = scope.inset(Inches(0.32), Inches(0.32))
    add_section_lead(slide, scope_inner.x, scope_inner.y, scope_inner.w,
                     spec.get("scope_heading", "対象範囲"))
    items = spec.get("scope", [])
    columns = scope_inner.columns([1] * max(1, len(items)), gap="standard")
    for index, item in enumerate(items):
        region = columns[index]
        label = item.get("label", f"{index + 1:02}")
        add_textbox(slide, region.x, region.y + Inches(0.78), region.w,
                    Inches(0.28), label, size=_type(slide).small,
                    color=PALETTE.blue, bold=True)
        add_paragraph_textbox(slide, region.x, region.y + Inches(1.22),
                              region.w, Inches(2.5), [
            {"segments": [(item.get("title", ""), {
                "size": _type(slide).section,
                "color": PALETTE.text_primary,
                "bold": True,
                "font": _type(slide).body_font,
            })], "space_after": 7},
            {"segments": [(item.get("body", ""), {
                "size": _type(slide).small,
                "color": PALETTE.text_secondary,
                "font": _type(slide).body_font,
            })]},
        ])
    if spec.get("period"):
        add_textbox(slide, scope_inner.x, scope_inner.y + scope_inner.h - Inches(0.42),
                    scope_inner.w, Inches(0.28), spec["period"],
                    size=_type(slide).small, color=PALETTE.text_primary,
                    bold=True)

    add_section_lead(slide, exclusions.x, exclusions.y, exclusions.w,
                     spec.get("exclusions_heading", "対象外"),
                     color=PALETTE.text_secondary)
    exclusions_items = spec.get("exclusions", [])
    add_item_list(slide, exclusions.x, exclusions.y + Inches(0.7),
                  exclusions.w, exclusions.h - Inches(1.35),
                  exclusions_items, bullet="—", body_gap=7)
    add_textbox(slide, exclusions.x, exclusions.y + exclusions.h - Inches(0.42),
                exclusions.w, Inches(0.34), spec["primary_message"],
                size=_type(slide).small, color=PALETTE.text_primary, bold=True)


def render_process_with_gates(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    phase_row, work_row, gate_row, conclusion = area.rows(
        [0.72, 2.35, 1.1, 0.62], gap=Inches(0.19))
    phases = spec.get("phases", [])
    phase_weights = [phase.get("weight", 1) for phase in phases] or [1]
    columns = phase_row.columns(phase_weights, gap="tight")
    tones = ["brand-soft", "neutral", "teal-soft"]
    for index, (phase, region) in enumerate(zip(phases, columns)):
        add_background_zone(slide, region.x, region.y, region.w, region.h,
                            tone=tones[index % len(tones)], rounded=True)
        add_textbox(slide, region.x + Inches(0.15), region.y + Inches(0.15),
                    region.w - Inches(0.3), Inches(0.3), phase.get("title", ""),
                    size=_type(slide).body, color=PALETTE.text_primary,
                    bold=True, align=PP_ALIGN.CENTER)

    work_columns = work_row.columns(phase_weights, gap="tight")
    for index, (phase, region) in enumerate(zip(phases, work_columns)):
        add_textbox(slide, region.x, region.y + Inches(0.18), region.w,
                    Inches(0.28), phase.get("label", f"{index + 1:02}"),
                    size=_type(slide).small, color=PALETTE.blue, bold=True)
        # phasesは最大6分割まで想定する狭い列のため、他の箇条書きと同じ
        # 全角ダッシュ「—」だと項目テキストに対して不自然に長く見える。
        # 中黒「・」は幅が狭く、狭い列でも項目とのバランスが崩れない。
        add_item_list(slide, region.x, region.y + Inches(0.62), region.w,
                      region.h - Inches(0.65), phase.get("items", []),
                      bullet="・", body_gap=5)

    add_hairline(slide, gate_row.x, gate_row.y + Inches(0.4), gate_row.w,
                 color=PALETTE.line_neutral, width=1)
    gates = spec.get("gates", [])
    gate_row_right = gate_row.x + gate_row.w
    label_w = Inches(1.75)
    for gate in gates:
        position = max(0.0, min(1.0, float(gate.get("position", 0))))
        x = gate_row.x + int(gate_row.w * position)
        Marker(slide, x - Inches(0.055), gate_row.y + Inches(0.345),
              Inches(0.11), Inches(0.11), shape="dot", fill=PALETTE.blue)
        # 中央揃えのまま単純に行の外へクランプすると、点の中心とラベルの
        # 中心がズレて「点から離れて見える」（端の点で顕著）。端に近い点は
        # 点の位置を起点にラベルを片側へ伸ばし、常に点と対応が付くようにする。
        if x - label_w // 2 < gate_row.x:
            label_x, align = x, PP_ALIGN.LEFT
        elif x + label_w // 2 > gate_row_right:
            label_x, align = x - label_w, PP_ALIGN.RIGHT
        else:
            label_x, align = x - label_w // 2, PP_ALIGN.CENTER
        add_textbox(slide, int(label_x), gate_row.y + Inches(0.58), label_w,
                    Inches(0.34), gate.get("title", ""),
                    size=_type(slide).small, color=PALETTE.text_primary,
                    bold=True, align=align)
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "editorial"))


def render_table_with_conclusion(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "dense"), page=page)
    table_region, conclusion = area.rows([4.35, 0.72], gap=Inches(0.24))
    add_data_table(slide, table_region, spec.get("columns", []),
                   spec.get("rows", []))
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def _render_chart_visual(builder, slide, spec, region):
    """imageが指定されればPNGを、chartが指定されればネイティブグラフを描く。"""
    chart_spec = spec.get("chart")
    if chart_spec:
        add_native_chart(
            slide, region.x, region.y, region.w, region.h,
            chart_type=chart_spec.get("type", "column"),
            categories=chart_spec.get("categories", []),
            series=chart_spec.get("series", []),
            typography=_type(slide),
            value_format=chart_spec.get("value_format"))
    else:
        image_path = builder.paths.input_dir / spec["image"]
        add_image_contain(slide, image_path, region)


def render_chart_with_insight(builder, spec, page):
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    variant = spec.get("variant", "standard")
    if variant == "conclusion-led":
        insight, chart = area.columns([0.82, 1.7], gap="wide")
        add_background_zone(slide, insight.x, insight.y, insight.w, insight.h,
                            tone="brand-soft", rounded=True)
        inner = insight.inset(Inches(0.3), Inches(0.34))
        add_textbox(slide, inner.x, inner.y, inner.w, Inches(0.3),
                    spec.get("insight_heading", "読み取れること"),
                    size=_type(slide).small, color=PALETTE.blue, bold=True)
        add_textbox(slide, inner.x, inner.y + Inches(0.58), inner.w,
                    Inches(1.55), spec["primary_message"],
                    size=_type(slide).section, color=PALETTE.text_primary,
                    bold=True, line_spacing=1.02)
        if spec.get("insights"):
            add_item_list(slide, inner.x, inner.y + Inches(2.35), inner.w,
                          inner.h - Inches(2.4), spec["insights"], bullet="—")
        _render_chart_visual(builder, slide, spec, chart)
    else:
        chart_and_notes, conclusion = area.rows([4.35, 0.72], gap=Inches(0.24))
        chart, notes = chart_and_notes.columns([1.8, 0.8], gap="wide")
        _render_chart_visual(builder, slide, spec, chart)
        add_section_lead(slide, notes.x, notes.y, notes.w,
                         spec.get("insight_heading", "読み取れること"))
        add_item_list(slide, notes.x, notes.y + Inches(0.65), notes.w,
                      notes.h - Inches(0.7), spec.get("insights", []), bullet="—")
        add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                        spec["primary_message"],
                        style=spec.get("conclusion_style", "subtle"))


def render_org_layers(builder, spec, page):
    """意思決定・運営など縦の責任階層と、横に並ぶ実行部門を分けて示す。

    layout-patterns.md「組織: 階層と横の役割」に対応する。
    """
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    layers = spec.get("layers", [])
    execution = spec.get("execution", [])
    weights = [1.3] * len(layers) + [2.1, 0.62]
    *layer_rows, execution_row, conclusion = area.rows(weights, gap=Inches(0.16))
    tones = ["brand-soft", "neutral", "teal-soft"]
    for index, (layer, region) in enumerate(zip(layers, layer_rows)):
        add_background_zone(slide, region.x, region.y, region.w, region.h,
                            tone=tones[index % len(tones)], rounded=True)
        inner = region.inset(Inches(0.3), Inches(0.14))
        # 見出し分のオフセットを固定0.46inのままにすると、layers数が増えて
        # 行の高さが縮んだ際に本文用の残りスペースをほぼ食い潰してしまう。
        # 行の高さに対する上限付き割合で縮めるが、add_section_leadの縦棒
        # マーカー自体も同じ値で縮め、offset = marker_h + GAPで揃えることで
        # マーカーが本文へ被らないことを構造的に保証する（マーカーだけ
        # 固定のまま offset だけ縮めると、マーカーが本文へ被る）。
        marker_h = min(Inches(0.38), max(Inches(0.22), int(inner.h * 0.34)))
        add_section_lead(slide, inner.x, inner.y, inner.w,
                         layer.get("heading", ""), marker_h=marker_h)
        header_offset = marker_h + SECTION_LEAD_GAP
        add_paragraph_textbox(slide, inner.x, inner.y + header_offset,
                              inner.w, inner.h - header_offset, [
            {"segments": [(layer.get("title", ""), {
                "size": typography.body, "color": PALETTE.text_primary,
                "bold": True, "font": typography.body_font,
            })], "space_after": 3},
            {"segments": [(layer.get("body", ""), {
                "size": typography.small, "color": PALETTE.text_secondary,
                "font": typography.body_font,
            })]},
        ], vertical_anchor=MSO_ANCHOR.MIDDLE)

    add_background_zone(slide, execution_row.x, execution_row.y,
                        execution_row.w, execution_row.h,
                        tone=tones[len(layer_rows) % len(tones)], rounded=True)
    execution_inner = execution_row.inset(Inches(0.3), Inches(0.16))
    add_section_lead(slide, execution_inner.x, execution_inner.y,
                     execution_inner.w,
                     spec.get("execution_heading", "業務実行"),
                     color=PALETTE.accent_secondary)
    cards_row = execution_inner.inset(top=Inches(0.48), bottom=0)
    columns = cards_row.columns([1] * max(1, len(execution)), gap="standard")
    for item, column in zip(execution, columns):
        add_card(slide, column.x, column.y, column.w, column.h)
        inner = column.inset(Inches(0.16), Inches(0.16))
        add_paragraph_textbox(slide, inner.x, inner.y, inner.w, inner.h, [
            {"segments": [(item.get("title", ""), {
                "size": typography.body, "color": PALETTE.text_primary,
                "bold": True, "font": typography.body_font,
            })], "space_after": 5},
            {"segments": [(item.get("body", ""), {
                "size": typography.small, "color": PALETTE.text_secondary,
                "font": typography.body_font,
            })]},
        ], vertical_anchor=MSO_ANCHOR.MIDDLE)
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def render_priority_actions(builder, spec, page):
    """優先度付きの課題と、対応する方針を左右に並べる。

    layout-patterns.md「課題と対応策」に対応する。
    """
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    body, conclusion = area.rows([4.35, 0.72], gap=Inches(0.24))
    left, right = body.columns([1.05, 1], gap="wide")

    issues = spec.get("issues", [])
    actions = spec.get("actions", [])

    add_section_lead(slide, left.x, left.y, left.w,
                     spec.get("issues_heading", "想定される課題"))
    top_priority = spec.get("top_priority", "最優先")
    gap = _adaptive_gap(issues, left.h - HEADING_BLOCK_H, base_gap=12,
                        per_item_pt=typography.body.pt * 1.2 +
                        typography.small.pt * 1.2 + 2)
    paragraphs = []
    for issue in issues:
        color = (PALETTE.blue if issue.get("priority") == top_priority
                 else PALETTE.text_secondary)
        paragraphs.append({"segments": [
            (f"{issue.get('priority', '')}  ", {
                "size": typography.small, "color": color, "bold": True,
                "font": typography.body_font,
            }),
            (issue.get("title", ""), {
                "size": typography.body, "color": PALETTE.text_primary,
                "bold": True, "font": typography.body_font,
            }),
        ], "space_after": 2})
        paragraphs.append({"segments": [
            (f"      {issue.get('body', '')}", {
                "size": typography.small, "color": PALETTE.text_secondary,
                "font": typography.body_font,
            }),
        ], "space_after": gap})
    add_paragraph_textbox(slide, left.x, left.y + HEADING_BLOCK_H, left.w,
                          left.h - HEADING_BLOCK_H - Inches(0.1), paragraphs)

    add_background_zone(slide, right.x - Inches(0.18), right.y - Inches(0.08),
                        right.w + Inches(0.36), right.h + Inches(0.12),
                        tone="neutral", rounded=True)
    add_section_lead(slide, right.x, right.y, right.w,
                     spec.get("actions_heading", "対応方針"),
                     color=PALETTE.accent_secondary)
    add_item_list(slide, right.x, right.y + HEADING_BLOCK_H, right.w,
                  right.h - HEADING_BLOCK_H - Inches(0.1),
                  actions, bullet="—")

    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def render_stage_track(builder, spec, page):
    """左から右への段階的な進行を、同格のステージカードで示す。"""
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    stages_row, note_row, conclusion = area.rows(
        [3.9, 0.34, 0.62], gap=Inches(0.16))
    stages = spec.get("stages", [])
    tones = ["neutral", "brand-soft", "teal-soft"]
    columns = stages_row.columns([1] * max(1, len(stages)), gap="wide")
    for index, (stage, column) in enumerate(zip(stages, columns)):
        add_background_zone(slide, column.x, column.y, column.w, column.h,
                            tone=tones[index % len(tones)], rounded=True)
        inner = column.inset(Inches(0.28), Inches(0.26))
        add_paragraph_textbox(slide, inner.x, inner.y, inner.w, inner.h, [
            {"segments": [(stage.get("label", f"STEP {index}"), {
                "size": typography.small, "color": PALETTE.blue,
                "bold": True, "font": typography.body_font,
            })], "space_after": 6},
            {"segments": [(stage.get("title", ""), {
                "size": typography.section, "color": PALETTE.text_primary,
                "bold": True, "font": typography.headline_font,
            })], "space_after": 8},
            {"segments": [(stage.get("body", ""), {
                "size": typography.small, "color": PALETTE.text_secondary,
                "font": typography.body_font,
            })], "line_spacing": 1.15},
        ], vertical_anchor=MSO_ANCHOR.MIDDLE)
    if spec.get("note"):
        add_textbox(slide, note_row.x, note_row.y, note_row.w, note_row.h,
                    spec["note"], size=typography.small,
                    color=PALETTE.text_secondary)
    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def render_numbered_list(builder, spec, page):
    """番号付きの項目群を、アジェンダや依頼事項のように単列で示す。

    項目数が少なくても余白が偏らないよう、リスト全体を利用可能な高さの
    中央へ配置する。
    """
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    items = spec.get("items", [])
    position = spec.get("message_position", "top")
    default_style = "plain" if position == "top" else "solid"
    style = spec.get("message_style", default_style)

    if position == "top":
        add_key_message(slide, area.x, area.y, area.w,
                        spec["primary_message"], style=style)
        list_top = area.y + Inches(0.62)
        list_bottom = area.y + area.h
    else:
        list_top = area.y
        list_bottom = area.y + area.h - Inches(1.1)

    available = max(0, list_bottom - list_top)
    # 項目が少なくても余白が偏らないよう、add_text_list側でブロックを
    # 利用可能な高さの中央へ配置する（項目数が多い場合はmax_row_hを
    # 上限に行の高さを縮め、スライド外へのはみ出しを防ぐ）。
    add_text_list(slide, area.x, list_top, area.w, available, items,
                 marker="number", divider=True, max_row_h=Inches(0.9))

    if position == "bottom":
        add_key_message(slide, area.x, list_bottom + Inches(0.25), area.w,
                        spec["primary_message"], style=style)


def render_section_divider(builder, spec, page):
    builder.add_section_divider(spec["title"], kicker=spec.get("kicker"),
                                subtitle=spec.get("subtitle"), page=page)


def render_matrix_2x2(builder, spec, page):
    """2軸で4象限に分け、各象限の位置づけを示す(ポートフォリオ分析等)。"""
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    grid_area, caption_row, conclusion = area.rows(
        [4.5, 0.34, 0.62], gap=Inches(0.16))
    axis_col, plot_col = grid_area.columns([0.14, 1], gap=Inches(0.14))

    y_axis = spec.get("y_axis", {})
    add_textbox(slide, axis_col.x, axis_col.y, axis_col.w, Inches(0.4),
                y_axis.get("high", ""), size=typography.small,
                color=PALETTE.text_secondary, bold=True)
    add_textbox(slide, axis_col.x, axis_col.y + axis_col.h - Inches(0.4),
                axis_col.w, Inches(0.4), y_axis.get("low", ""),
                size=typography.small, color=PALETTE.text_secondary, bold=True)

    quadrants = spec.get("quadrants", [])
    top_row, bottom_row = plot_col.rows([1, 1], gap=Inches(0.1))
    cells = list(top_row.columns([1, 1], gap=Inches(0.1))) + \
        list(bottom_row.columns([1, 1], gap=Inches(0.1)))
    for quadrant, cell in zip(quadrants, cells):
        tone = "brand-soft" if quadrant.get("emphasis") else "neutral"
        add_background_zone(slide, cell.x, cell.y, cell.w, cell.h,
                            tone=tone, rounded=True)
        inner = cell.inset(Inches(0.24), Inches(0.2))
        add_paragraph_textbox(slide, inner.x, inner.y, inner.w, inner.h, [
            {"segments": [(quadrant.get("label", ""), {
                "size": typography.small, "color": PALETTE.blue,
                "bold": True, "font": typography.body_font,
            })], "space_after": 4},
            {"segments": [(quadrant.get("title", ""), {
                "size": typography.body, "color": PALETTE.text_primary,
                "bold": True, "font": typography.body_font,
            })], "space_after": 3},
            {"segments": [(quadrant.get("body", ""), {
                "size": typography.small, "color": PALETTE.text_secondary,
                "font": typography.body_font,
            })]},
        ], vertical_anchor=MSO_ANCHOR.MIDDLE)

    x_axis = spec.get("x_axis", {})
    half = plot_col.w // 2
    add_textbox(slide, plot_col.x, caption_row.y, half, caption_row.h,
                x_axis.get("low", ""), size=typography.small,
                color=PALETTE.text_secondary, bold=True)
    add_textbox(slide, plot_col.x + half, caption_row.y, half, caption_row.h,
                x_axis.get("high", ""), size=typography.small,
                color=PALETTE.text_secondary, bold=True, align=PP_ALIGN.RIGHT)

    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


def render_stat_highlight(builder, spec, page):
    """単一の実績数値を主役にし、補足指標と結論を添える。"""
    slide, area = builder.add_slide(
        spec["title"], density=spec.get("density", "standard"), page=page)
    typography = _type(slide)
    stat = spec.get("stat", {})
    supporting = spec.get("supporting", [])
    if supporting:
        hero_row, supporting_row, conclusion = area.rows(
            [2.5, 2.05, 0.62], gap=Inches(0.2))
    else:
        hero_row, conclusion = area.rows([4.7, 0.62], gap=Inches(0.24))
        supporting_row = None

    add_background_zone(slide, hero_row.x, hero_row.y, hero_row.w, hero_row.h,
                        tone="brand-soft", rounded=True)
    inner = hero_row.inset(Inches(0.5), Inches(0.3))
    Stat(slide, inner.x, inner.y, inner.w, inner.h,
        stat.get("value", ""), stat.get("label", ""), detail=stat.get("detail"),
        value_size=Pt(56), label_size=typography.section,
        vertical_anchor=MSO_ANCHOR.MIDDLE)

    if supporting_row is not None:
        columns = supporting_row.columns([1] * max(1, len(supporting)),
                                         gap="standard")
        for item, column in zip(supporting, columns):
            add_card(slide, column.x, column.y, column.w, column.h)
            inner = column.inset(Inches(0.2), Inches(0.2))
            Stat(slide, inner.x, inner.y, inner.w, inner.h,
                item.get("value", ""), item.get("label", ""),
                value_size=typography.metric, value_color=PALETTE.blue,
                label_size=typography.small, label_color=PALETTE.text_secondary,
                label_bold=False, vertical_anchor=MSO_ANCHOR.MIDDLE)

    add_key_message(slide, conclusion.x, conclusion.y, conclusion.w,
                    spec["primary_message"],
                    style=spec.get("conclusion_style", "subtle"))


RENDERERS = {
    "cover": render_cover,
    "comparison": render_comparison,
    "evidence_and_decision": render_evidence_and_decision,
    "scope_and_exclusions": render_scope_and_exclusions,
    "process_with_gates": render_process_with_gates,
    "table_with_conclusion": render_table_with_conclusion,
    "chart_with_insight": render_chart_with_insight,
    "org_layers": render_org_layers,
    "priority_actions": render_priority_actions,
    "stage_track": render_stage_track,
    "numbered_list": render_numbered_list,
    "section_divider": render_section_divider,
    "matrix_2x2": render_matrix_2x2,
    "stat_highlight": render_stat_highlight,
}


def render_deck(content, root, output_path=None):
    warnings = require_valid_content(content)
    builder = DeckBuilder(root, mode=content.get("deck", {}).get("mode"))
    for page, spec in enumerate(content["slides"], 1):
        RENDERERS[spec["type"]](builder, spec, page)
    return builder.save(output_path), warnings
