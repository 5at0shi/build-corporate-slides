#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path

import yaml

REQUIRED = (
    ("python", "executable"),
    ("paths", "input_dir"),
    ("paths", "work_dir"),
    ("paths", "output_dir"),
    ("organization", "department"),
    ("organization", "classification"),
    ("branding", "logo", "enabled"),
)


def has_path(data, keys):
    current = data
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            return False
        current = current[key]
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="slide skill configを検証します")
    parser.add_argument("config", type=Path)
    args = parser.parse_args()
    try:
        config = yaml.safe_load(args.config.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:
        print(f"configを読めません: {exc}", file=sys.stderr)
        return 1
    missing = [".".join(keys) for keys in REQUIRED if not has_path(config, keys)]
    if missing:
        print("不足キー: " + ", ".join(missing), file=sys.stderr)
        return 1
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
