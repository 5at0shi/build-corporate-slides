from collections import Counter


KNOWN_TYPES = {
    "cover", "comparison", "evidence_and_decision",
    "scope_and_exclusions", "process_with_gates",
    "table_with_conclusion", "chart_with_insight",
    "org_layers", "priority_actions", "stage_track", "numbered_list",
    "section_divider", "matrix_2x2", "stat_highlight",
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


def inspect_content(content):
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
        if slide_type not in KNOWN_TYPES:
            errors.append(f"{label}: 未対応typeです: {slide_type}")
            continue
        if not slide.get("title"):
            errors.append(f"{label}: titleがありません")
        density = slide.get("density", "standard")
        if density not in {"standard", "dense"}:
            errors.append(f"{label}: densityはstandardまたはdenseです")
        if (slide_type not in {"cover", "section_divider"} and
                not slide.get("primary_message")):
            warnings.append(f"{label}: primary_messageがありません")
        if slide_type == "table_with_conclusion":
            if not slide.get("columns") or not slide.get("rows"):
                errors.append(f"{label}: columnsとrowsが必要です")
        if slide_type == "chart_with_insight":
            chart = slide.get("chart")
            if not slide.get("image") and not chart:
                errors.append(f"{label}: imageまたはchartが必要です")
            if chart and (not chart.get("categories") or not chart.get("series")):
                errors.append(f"{label}: chart.categoriesとchart.seriesが必要です")
        if slide_type == "org_layers":
            if not slide.get("layers") or not slide.get("execution"):
                errors.append(f"{label}: layersとexecutionが必要です")
        if slide_type == "priority_actions":
            if not slide.get("issues") or not slide.get("actions"):
                errors.append(f"{label}: issuesとactionsが必要です")
        if slide_type == "stage_track" and not slide.get("stages"):
            errors.append(f"{label}: stagesが必要です")
        if slide_type == "numbered_list" and not slide.get("items"):
            errors.append(f"{label}: itemsが必要です")
        if slide_type == "matrix_2x2":
            quadrants = slide.get("quadrants")
            if not quadrants or len(quadrants) != 4:
                errors.append(f"{label}: quadrantsは4件必要です")
            if not slide.get("x_axis") or not slide.get("y_axis"):
                errors.append(f"{label}: x_axisとy_axisが必要です")
        if slide_type == "stat_highlight":
            if not slide.get("stat") or not slide["stat"].get("value"):
                errors.append(f"{label}: stat.valueが必要です")

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
                    "quadrants", "supporting"):
            value = slide.get(key)
            if isinstance(value, list):
                item_count += len(value)
            elif isinstance(value, dict) and isinstance(value.get("items"), list):
                item_count += len(value["items"])
        if item_count > (10 if density == "dense" else 7):
            warnings.append(f"{label}: 項目数が多いです ({item_count})")
    return errors, warnings


def require_valid_content(content):
    errors, warnings = inspect_content(content)
    if errors:
        raise ValueError("\n".join(errors))
    return warnings
