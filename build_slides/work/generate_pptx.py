from pathlib import Path
import sys

import yaml


WORKSPACE_ROOT = Path.cwd()
SKILL = WORKSPACE_ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from slidekit import render_deck  # noqa: E402


content_path = WORKSPACE_ROOT / "build_slides" / "work" / "slide_content.yaml"
content = yaml.safe_load(content_path.read_text(encoding="utf-8")) or {}

try:
    output, warnings = render_deck(content, WORKSPACE_ROOT)
except ValueError as error:
    # YAMLの事前診断(preflight)で止まった場合。原因を1件ずつ示して終了する
    # （tracebackを出すより、直すべき箇所が読み取りやすいため）。
    print("内容の検証で問題が見つかりました。修正してから再実行してください:",
          file=sys.stderr)
    for line in str(error).splitlines():
        print(f"  - {line}", file=sys.stderr)
    raise SystemExit(1)

for warning in warnings:
    print(f"警告: {warning}")
print(f"PPTX: {output}")
