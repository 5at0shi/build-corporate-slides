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

    # python.executableが実在しない場合に黙って別環境へ切り替えるのを防ぐ
    # （powerpoint-production.md「実行環境」）。フォールバック自体は許容
    # されるため停止はせず、報告だけを必ず出す。
    executable = config.get("python", {}).get("executable")
    if executable:
        resolved = Path(executable)
        if not resolved.is_absolute():
            resolved = args.config.resolve().parent / resolved
        if not resolved.is_file():
            print(f"警告: python.executable が見つかりません: {executable}",
                  file=sys.stderr)
            print("  別のPythonを使う場合は、どれを使ったかを必ず報告すること"
                  "（黙って切り替えない）。", file=sys.stderr)

    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
