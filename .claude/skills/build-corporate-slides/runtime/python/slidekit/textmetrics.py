"""テキストが必要とする概算の高さ(pt)を見積もる。

python-pptxは実測レイアウトを持たないため、CJKは概ね1em・半角英数は
概ね0.55emという簡易ヒューリスティックで折り返し行数を概算する。
生成時のレイアウト判断（見出し+リストのブロックを領域内で中央寄せする等）
と、validate_pptx.pyのはみ出し検知は、この一つの推定ロジックを共有する。
"""


def char_width_factor(char: str) -> float:
    code = ord(char)
    wide = (0x1100 <= code <= 0x115F or 0x2E80 <= code <= 0xA4CF or
            0xAC00 <= code <= 0xD7A3 or 0xF900 <= code <= 0xFAFF or
            0xFF00 <= code <= 0xFF60 or 0xFFE0 <= code <= 0xFFE6 or
            0x20000 <= code <= 0x3FFFD)
    return 1.0 if wide else 0.55


def estimate_line_count(text: str, font_size_pt: float, width_pt: float) -> int:
    if not text:
        return 0
    if width_pt <= 0:
        return 1
    char_width = sum(char_width_factor(ch) for ch in text) * font_size_pt
    return max(1, -(-int(char_width) // int(max(1, width_pt))))


def estimate_paragraph_height_pt(text: str, font_size_pt: float, width_pt: float, *,
                                 line_spacing=1.0, space_before=0.0,
                                 space_after=0.0) -> float:
    lines = estimate_line_count(text, font_size_pt, width_pt)
    return lines * font_size_pt * 1.2 * line_spacing + space_before + space_after


def adaptive_gap_pt(content_pt: float, item_count: int, available_pt: float, *,
                    base_gap: float = 3, max_extra: float = 24) -> float:
    """項目間の余白(pt)を、領域の余りに応じて上限付きで広げる。

    content_ptはbody_gapを含まない本文だけの高さ。見出し直下に項目を
    連続させたまま、下部にできる余白の割合を抑える（過剰に広げて
    間延びしないよう上限max_extraを設ける）。
    """
    if item_count <= 1:
        return base_gap
    slack = max(0.0, available_pt - content_pt - base_gap * (item_count - 1))
    extra = min(max_extra, slack / (item_count - 1))
    return base_gap + extra


def estimate_item_list_height_pt(typography, items, width_pt, *, body_gap=3) -> float:
    """add_item_listが描く内容のおおよその高さ(pt)。レイアウト判断用の概算。"""
    total = 0.0
    for index, item in enumerate(items):
        if isinstance(item, str):
            title, body = item, None
        else:
            title, body = item.get("title", ""), item.get("body")
        total += estimate_paragraph_height_pt(
            title, typography.body.pt, width_pt, line_spacing=1.08)
        if body:
            # add_item_listは実際には本文の前に "    "（4スペース）の字下げを
            # 付けて描画するため、折り返し推定もそれを含めないと実際より
            # 過小評価し、はみ出しの原因になる。
            total += estimate_paragraph_height_pt(
                "    " + body, typography.small.pt, width_pt, line_spacing=1.08)
        if index < len(items) - 1:
            total += body_gap
    return total
