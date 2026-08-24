from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

import yaml


DEFAULT_CONFIG = {
    "paths": {
        "input_dir": "./slides/input",
        "work_dir": "./slides/work",
        "output_dir": "./slides/output",
    },
    "organization": {
        "department": "",
        "classification": "",
    },
    "branding": {
        "logo": {"enabled": False, "path": None},
    },
    "deck": {"mode": "business"},
    "typography": {
        "headline_font": "Yu Gothic",
        "body_font": "Yu Gothic",
        "editorial_font": "Yu Mincho",
    },
}


def _merge(base, overlay):
    result = deepcopy(base)
    for key, value in (overlay or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _merge(result[key], value)
        else:
            result[key] = value
    return result


@dataclass(frozen=True)
class WorkspacePaths:
    root: Path
    input_dir: Path
    work_dir: Path
    output_dir: Path


def load_workspace_config(root):
    root = Path(root).resolve()
    path = root / ".slide-skill-config.yaml"
    if path.is_file():
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    else:
        data = {}
    return _merge(DEFAULT_CONFIG, data)


def workspace_paths(root, config):
    root = Path(root).resolve()

    def resolve(value):
        path = Path(value)
        return path.resolve() if path.is_absolute() else (root / path).resolve()

    settings = config["paths"]
    return WorkspacePaths(
        root=root,
        input_dir=resolve(settings["input_dir"]),
        work_dir=resolve(settings["work_dir"]),
        output_dir=resolve(settings["output_dir"]),
    )
