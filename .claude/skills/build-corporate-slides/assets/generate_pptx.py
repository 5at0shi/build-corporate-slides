from pathlib import Path
import sys

import yaml


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / ".agents" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))

from slidekit import render_deck  # noqa: E402


content_path = ROOT / "work" / "slide_content.yaml"
content = yaml.safe_load(content_path.read_text(encoding="utf-8")) or {}
output, warnings = render_deck(content, ROOT)
for warning in warnings:
    print(f"警告: {warning}")
print(f"PPTX: {output}")
