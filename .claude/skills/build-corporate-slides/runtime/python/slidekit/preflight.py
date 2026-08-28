from collections import Counter


KNOWN_TYPES = {
    "cover", "comparison", "evidence_and_decision",
    "scope_and_exclusions", "process_with_gates",
    "table_with_conclusion", "table_with_insight", "chart_with_insight",
    "org_layers", "priority_actions", "stage_track", "numbered_list",
    "section_divider", "matrix", "stat_highlight", "funnel", "cycle",
    "timeline", "waterfall", "issue_tree",
}


def _walk_text(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from _walk_text(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_text(child)


def inspect_content(content, *, extra_types=()):
    """contentを事前診断する。

    extra_typesには、生成スクリプト側で個別構築するページのtype名を渡す
    （renderer-catalog.mdのEscape Hatch）。そのtypeは未対応として弾かず、
    共通の検査（title・primary_message・density・文字量・項目数）だけを
    適用する。個別構築ページも、内容はYAMLへ残すのがcontent-model.mdの
    方針であり、共通の契約はrendererのページと同じだけ効かせる。
    """
    known_types = KNOWN_TYPES | set(extra_types)
    errors, warnings = [], []
    deck = content.get("deck", {})
    if deck.get("mode", "business") not in {"business", "dense", "large-room"}:
        errors.append("deck.modeは business / dense / large-room のいずれかです")
    slides = content.get("slides")
    if not isinstance(slides, list) or not slides:
        return ["slidesが空です"], warnings
    ids = [slide.get("id") for slide in slides if slide.get("id")]
    duplicates = [key for key, count in Counter(ids).items() if count > 1]
    if duplicates:
        errors.append("slide idが重複しています: " + ", ".join(duplicates))

    for index, slide in enumerate(slides, 1):
        label = f"slide {index} ({slide.get('id', 'idなし')})"
        slide_type = slide.get("type")
        if slide_type not in known_types:
            errors.append(f"{label}: 未対応typeです: {slide_type}")
            continue
        if not slide.get("title"):
            errors.append(f"{label}: titleがありません")
        density = slide.get("density", "standard")
        if density not in {"standard", "dense"}:
            errors.append(f"{label}: densityはstandardまたはdenseです")
        # cover/section_divider以外の全rendererは spec["primary_message"] を
        # 直接参照する（結論用の領域を必ず確保するため）。欠けたまま描画へ
        # 進むとKeyErrorで落ちるので、警告ではなくerrorとして生成前に止める。
        if (slide_type not in {"cover", "section_divider"} and
                not slide.get("primary_message")):
            errors.append(f"{label}: primary_messageがありません")
        if slide_type in ("table_with_conclusion", "table_with_insight"):
            if not slide.get("columns") or not slide.get("rows"):
                errors.append(f"{label}: columnsとrowsが必要です")
        if slide_type == "chart_with_insight":
            chart = slide.get("chart")
            if not slide.get("image") and not chart:
                errors.append(f"{label}: imageまたはchartが必要です")
            elif chart:
                if chart.get("type") == "scatter":
                    # scatterはcategoriesを使わず、seriesの各要素がpoints
                    # (x/y座標)を持つ形式のため、他typeとは別に検証する。
                    if not chart.get("series"):
                        errors.append(f"{label}: chart.seriesが必要です")
                    elif any(not s.get("points") for s in chart["series"]):
                        errors.append(
                            f"{label}: chart.series各要素にpointsが必要です")
                elif not chart.get("categories") or not chart.get("series"):
                    errors.append(f"{label}: chart.categoriesとchart.seriesが必要です")
        if slide_type == "org_layers":
            if not slide.get("layers") or not slide.get("execution"):
                errors.append(f"{label}: layersとexecutionが必要です")
            elif len(slide["layers"]) > 2:
                warnings.append(
                    f"{label}: layersが多く({len(slide['layers'])}件)、各層の"
                    "本文がはみ出す可能性があります。2件以下に抑えるか、"
                    "階層を統合してください")
        if slide_type == "priority_actions":
            if not slide.get("issues") or not slide.get("actions"):
                errors.append(f"{label}: issuesとactionsが必要です")
        if slide_type == "stage_track" and not slide.get("stages"):
            errors.append(f"{label}: stagesが必要です")
        if slide_type == "process_with_gates" and not slide.get("phases"):
            errors.append(f"{label}: phasesが必要です")
        if slide_type == "numbered_list" and not slide.get("items"):
            errors.append(f"{label}: itemsが必要です")
        if slide_type == "matrix":
            rows, cols = slide.get("rows", 2), slide.get("cols", 2)
            cells = slide.get("cells")
            if not cells or len(cells) != rows * cols:
                errors.append(f"{label}: cellsは{rows * cols}件（rows×cols）必要です")
            if bool(slide.get("x_axis")) != bool(slide.get("y_axis")):
                errors.append(f"{label}: x_axisとy_axisは両方指定するか両方省略します"
                              "（片方だけでは軸を描けません）")
        if slide_type == "stat_highlight":
            if not slide.get("stat") and not slide.get("supporting"):
                errors.append(f"{label}: statまたはsupportingが必要です")
            elif slide.get("stat") and not slide["stat"].get("value"):
                errors.append(f"{label}: stat.valueが必要です")
        if slide_type == "funnel" and not slide.get("stages"):
            errors.append(f"{label}: stagesが必要です")
        if slide_type == "timeline":
            periods, timeline_rows = slide.get("periods"), slide.get("rows")
            if not periods or not timeline_rows:
                errors.append(f"{label}: periodsとrowsが必要です")
            elif len(periods) > 8:
                warnings.append(
                    f"{label}: periodsが多く({len(periods)}件)、1目盛りの幅が"
                    "狭くなります。8件以下にするか、期間の粒度を上げてください")
        if slide_type == "waterfall":
            bars = slide.get("bars")
            if not bars or len(bars) < 3:
                # 開始・増減・終了の3本が揃って初めて「AからBへの変化」に
                # なる。2本以下は増減の分解ではなく単なる比較のため、
                # comparisonやchart_with_insightの方が適切。
                errors.append(f"{label}: barsは3件以上必要です（開始・増減・終了）")
            elif not any(bar.get("kind") == "total" for bar in bars):
                warnings.append(
                    f"{label}: kind=\"total\"の棒がありません。開始値・終了値を"
                    "totalとして置かないと、増減が何に対する増減か読めません")
            if bars and len(bars) > 9:
                warnings.append(
                    f"{label}: barsが多く({len(bars)}件)、1本の幅が狭くなります。"
                    "要因をまとめるか、ページを分割してください")
        if slide_type == "issue_tree":
            branches = slide.get("branches")
            if not slide.get("root") or not branches:
                errors.append(f"{label}: rootとbranchesが必要です")
            elif len(branches) > 5:
                warnings.append(
                    f"{label}: branchesが多く({len(branches)}件)、枝1つの高さが"
                    "狭くなります。5件以下へまとめてください")
            elif sum(len(b.get("items", [])) for b in branches) > 12:
                warnings.append(
                    f"{label}: 内訳(items)の合計が多く、1行の高さが狭くなります。"
                    "12件以下へ絞るか、枝ごとにページを分けてください")
        if slide_type == "cycle":
            steps = slide.get("steps")
            if not steps or len(steps) < 2:
                errors.append(f"{label}: stepsは2件以上必要です")
            elif len(steps) > 6:
                warnings.append(
                    f"{label}: stepsが多く({len(steps)}件)、隣接するCard同士の"
                    "間隔が狭くなる可能性があります。6件以下を目安にしてください")

        text_values = list(_walk_text(slide))
        total_chars = sum(len(value) for value in text_values)
        budget = 620 if density == "dense" else 430
        if total_chars > budget:
            warnings.append(
                f"{label}: 文字量が目安を超えています ({total_chars}/{budget})。"
                "縮小前に重複削除、統合、表化、ページ分割を検討してください")
        for value in text_values:
            if len(value) > 90:
                warnings.append(f"{label}: 90文字を超える文章があります: {value[:30]}…")
                break

        primary = slide.get("primary_message", "").strip()
        if primary:
            occurrences = sum(primary in value for value in text_values)
            if occurrences > 1:
                warnings.append(f"{label}: primary_messageが本文でも重複しています")

        item_count = 0
        for key in ("items", "left", "right", "scope", "exclusions",
                    "phases", "gates", "evidence", "insights", "rows",
                    "layers", "execution", "issues", "actions", "stages",
                    "cells", "supporting", "steps", "bars", "branches"):
            value = slide.get(key)
            if isinstance(value, list):
                item_count += len(value)
            elif isinstance(value, dict) and isinstance(value.get("items"), list):
                item_count += len(value["items"])
        if item_count > (10 if density == "dense" else 7):
            warnings.append(f"{label}: 項目数が多いです ({item_count})")
    return errors, warnings


def require_valid_content(content, *, extra_types=()):
    errors, warnings = inspect_content(content, extra_types=extra_types)
    if errors:
        raise ValueError("\n".join(errors))
    return warnings
