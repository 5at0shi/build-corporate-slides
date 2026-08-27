from pathlib import Path
import sys

import yaml


WORKSPACE_ROOT = Path.cwd()
SKILL = WORKSPACE_ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from slidekit import render_deck  # noqa: E402


content_path = WORKSPACE_ROOT / "build_slides" / "work" / "slide_content.yaml"
content = yaml.safe_load(content_path.read_text(encoding="utf-8")) or {}
output, warnings = render_deck(content, WORKSPACE_ROOT)
for warning in warnings:
    print(f"警告: {warning}")
print(f"PPTX: {output}")
