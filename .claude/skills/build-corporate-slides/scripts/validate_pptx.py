#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "runtime" / "python"))
from slidekit.textmetrics import char_width_factor  # noqa: E402


def _paragraph_line_sizes(paragraph, available_w: float) -> list[float] | None:
    """runをまたいだ折り返しを再現し、各行に使われた最大フォントサイズを返す。

    add_item_list・add_numbered_rowなどは、タイトルと本文を改行文字("\\n")
    で1つのparagraph内の別runへ連結する（見た目上の行ごとにtextboxを
    分割しないため）。単純に全runのテキストを結合して1つのフォントサイズで
    折り返し推定すると、明示的な改行を無視して過小評価する。runごとに
    改行で区切り、行に使われた文字の最大フォントサイズで高さを見積もる。
    """
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size]
    if not sizes:
        return None
    fallback_size = max(sizes)
    line_sizes = []
    current_w = 0.0
    current_max_size = 0.0
    started = False
    for run in paragraph.runs:
        size = run.font.size.pt if run.font.size else fallback_size
        segments = run.text.split("\n")
        for seg_index, segment in enumerate(segments):
            if seg_index > 0:
                line_sizes.append(current_max_size or size)
                current_w = 0.0
                current_max_size = 0.0
            for ch in segment:
                started = True
                ch_w = char_width_factor(ch) * size
                if current_w + ch_w > available_w and current_w > 0:
                    line_sizes.append(current_max_size or size)
                    current_w = 0.0
                    current_max_size = 0.0
                current_w += ch_w
                current_max_size = max(current_max_size, size)
    if current_w > 0 or not line_sizes:
        line_sizes.append(current_max_size or fallback_size)
    return line_sizes if started else None


def _estimate_text_height_pt(text_frame, width_pt: float) -> float | None:
    """textboxの折り返し後の推定高さ(pt)。フォントサイズ未指定の場合はNone。

    生成時のレイアウト判断(slidekit.textmetrics)と同じ文字幅ヒューリス
    ティックを使い、runをまたいだ改行も再現して折り返し行数を見積もる。
    """
    margin_left = Emu(text_frame.margin_left or 0).pt
    margin_right = Emu(text_frame.margin_right or 0).pt
    margin_top = Emu(text_frame.margin_top or 0).pt
    margin_bottom = Emu(text_frame.margin_bottom or 0).pt
    available_w = max(1.0, width_pt - margin_left - margin_right)
    total = margin_top + margin_bottom
    for paragraph in text_frame.paragraphs:
        if not paragraph.runs or not any(run.text for run in paragraph.runs):
            continue
        line_sizes = _paragraph_line_sizes(paragraph, available_w)
        if line_sizes is None:
            return None
        spacing = paragraph.line_spacing
        multiplier = spacing if isinstance(spacing, (int, float)) else 1.0
        space_before = Emu(paragraph.space_before or 0).pt
        space_after = Emu(paragraph.space_after or 0).pt
        total += sum(size * 1.2 * multiplier for size in line_sizes)
        total += space_before + space_after
    return total


def _check_text_overflow(slide, slide_no, warnings):
    for shape in slide.shapes:
        if not shape.has_text_frame or not shape.text.strip():
            continue
        if shape.height <= 0 or shape.width <= 0:
            continue
        estimated = _estimate_text_height_pt(shape.text_frame, Emu(shape.width).pt)
        if estimated is None:
            continue
        actual = Emu(shape.height).pt
        if estimated > actual * 1.15:
            snippet = shape.text.strip().replace("\n", " ")[:24]
            warnings.append(
                f"slide {slide_no}: テキストが枠からはみ出す可能性 "
                f"'{snippet}' (推定{estimated:.0f}pt / 枠{actual:.0f}pt)")


def validate(path: Path) -> tuple[list[str], list[str]]:
    issues = []
    warnings = []
    if not path.is_file() or path.stat().st_size == 0:
        return [f"PPTXが存在しないか空です: {path}"], warnings
    try:
        prs = Presentation(path)
    except Exception as exc:
        return [f"PPTXを開けません: {exc}"], warnings
    if not prs.slides:
        issues.append("スライドがありません")
    for slide_no, slide in enumerate(prs.slides, 1):
        visible = 0
        text_shapes = []
        for shape in slide.shapes:
            is_line = shape.shape_type == MSO_SHAPE_TYPE.LINE
            invalid_size = ((shape.width <= 0 and shape.height <= 0) if is_line
                            else (shape.width <= 0 or shape.height <= 0))
            if invalid_size:
                issues.append(f"slide {slide_no}: サイズ0のオブジェクト '{shape.name}'")
            if (shape.left + shape.width < 0 or shape.top + shape.height < 0 or
                    shape.left > prs.slide_width or shape.top > prs.slide_height):
                issues.append(f"slide {slide_no}: スライド外のオブジェクト '{shape.name}'")
            if shape.has_text_frame and shape.text.strip():
                visible += 1
                text_shapes.append(shape)
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size and run.font.size.pt < 7:
                            issues.append(
                                f"slide {slide_no}: 7pt未満の文字 '{run.text[:24]}'")
            elif not shape.has_text_frame:
                visible += 1
        if visible == 0:
            issues.append(f"slide {slide_no}: 空ページの可能性")
        _check_text_overflow(slide, slide_no, warnings)
        short_shapes = [s for s in text_shapes if len(s.text.strip()) <= 24]
        if len(text_shapes) >= 16 and len(short_shapes) / len(text_shapes) >= 0.65:
            warnings.append(
                f"slide {slide_no}: 短いtextboxが多く、文章の過剰分割の可能性 "
                f"({len(text_shapes)} text shapes)")
    return issues, warnings


def main() -> int:
    parser = argparse.ArgumentParser(description="PPTXの基本構造を検証します")
    parser.add_argument("pptx", type=Path)
    args = parser.parse_args()
    issues, warnings = validate(args.pptx)
    if warnings:
        print("編集性の警告:")
        for warning in warnings:
            print(f"- {warning}")
    if issues:
        print("検証で問題を検出しました:")
        for issue in issues:
            print(f"- {issue}")
        return 1
    print(f"OK: {args.pptx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
