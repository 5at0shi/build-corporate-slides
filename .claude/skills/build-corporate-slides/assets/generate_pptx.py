from pathlib import Path
import sys

import yaml


WORKSPACE_ROOT = Path.cwd()
SKILL = WORKSPACE_ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from slidekit import ContentError, render_deck  # noqa: E402


content_path = WORKSPACE_ROOT / "build_slides" / "work" / "slide_content.yaml"
content = yaml.safe_load(content_path.read_text(encoding="utf-8")) or {}

# 該当するrendererが無いページを個別構築する場合は、描画関数をここで定義し
# render_deck(..., extra_renderers={"型名": 関数}) へ渡す。内容はYAMLへ
# 残したまま、他のページと同じデッキに含められる（手順はrenderer-catalog.md
# の「個別構築ページをYAMLと同居させる」）。

try:
    output, warnings = render_deck(content, WORKSPACE_ROOT)
except ContentError as error:
    # YAMLの事前診断(preflight)で止まった場合。原因を1件ずつ示して終了する
    # （tracebackを出すより、直すべき箇所が読み取りやすいため）。描画中の
    # エラーはここで捕まえない。tracebackがそのまま出たほうが原因へ辿れる。
    print("内容の検証で問題が見つかりました。修正してから再実行してください:",
          file=sys.stderr)
    for line in str(error).splitlines():
        print(f"  - {line}", file=sys.stderr)
    raise SystemExit(1)

for warning in warnings:
    print(f"警告: {warning}")
print(f"PPTX: {output}")
