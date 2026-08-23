#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path

import yaml


def main() -> int:
    parser = argparse.ArgumentParser(
        description="slide_content.yamlから編集可能なPowerPointを生成します")
    parser.add_argument("content", type=Path)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(skill_root / "runtime" / "python"))
    from slidekit import render_deck  # noqa: E402

    try:
        content = yaml.safe_load(args.content.read_text(encoding="utf-8")) or {}
        output, warnings = render_deck(content, args.root, args.output)
    except (OSError, ValueError, KeyError) as exc:
        print(f"生成できません: {exc}", file=sys.stderr)
        return 1
    for warning in warnings:
        print(f"警告: {warning}")
    print(f"PPTX: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
