#!/usr/bin/env python3
import importlib
import sys


REQUIRED = {
    "pptx": "python-pptx",
    "yaml": "PyYAML",
    "PIL": "Pillow",
}


def main() -> int:
    missing = []
    for module, package in REQUIRED.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(package)
    if missing:
        print("不足依存: " + ", ".join(missing), file=sys.stderr)
        print("runtime/python/requirements.txt と照合してください", file=sys.stderr)
        return 1
    print(f"OK: {sys.executable}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
