#!/usr/bin/env python3
import sys
import tempfile
from pathlib import Path

import yaml
from pptx import Presentation


def main() -> int:
    skill_root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(skill_root / "runtime" / "python"))
    from slidekit import inspect_content, logo_path_from_config, render_deck

    config = {
        "python": {"executable": sys.executable},
        "paths": {
            "input_dir": "./input", "work_dir": "./work",
            "output_dir": "./output",
        },
        "organization": {
            "department": "テスト部", "classification": "部外秘",
        },
        "deck": {"mode": "business"},
        "typography": {
            "headline_font": "Arial Unicode MS",
            "body_font": "Arial Unicode MS",
            "editorial_font": "Arial Unicode MS",
        },
        "branding": {"logo": {"enabled": True, "path": None}},
    }
    content = {
        "deck": {"mode": "business"},
        "slides": [
            {"id": "cover", "type": "cover", "title": "検証資料",
             "subtitle": "再現性テスト"},
            {"id": "comparison", "type": "comparison",
             "title": "比較して判断する", "density": "standard",
             "primary_message": "二つの観点を同時に確認する",
             "left": {"heading": "期待", "items": [
                 {"title": "効果", "body": "作業を短縮する"},
                 {"title": "品質", "body": "ばらつきを抑える"}]},
             "right": {"heading": "リスク", "items": [
                 {"title": "管理", "body": "無秩序な利用を防ぐ"},
                 {"title": "判断", "body": "基準を明確にする"}]}}
        ],
    }
    errors, warnings = inspect_content(content)
    assert not errors, errors
    assert logo_path_from_config(config, Path.cwd()).is_file()

    with tempfile.TemporaryDirectory(prefix="slidekit-test-") as temp:
        root = Path(temp)
        (root / ".slide-skill-config.yaml").write_text(
            yaml.safe_dump(config, allow_unicode=True), encoding="utf-8")
        output, render_warnings = render_deck(content, root)
        assert not render_warnings
        presentation = Presentation(output)
        assert len(presentation.slides) == 2
        cover_text = [shape.text for shape in presentation.slides[0].shapes
                      if shape.has_text_frame]
        assert "部外秘" in cover_text and "テスト部" in cover_text
        comparison = presentation.slides[1]
        text_shapes = [shape for shape in comparison.shapes
                       if shape.has_text_frame and shape.text.strip()]
        assert len(text_shapes) <= 8, len(text_shapes)
        assert any(len(shape.text_frame.paragraphs) >= 2 for shape in text_shapes)
    print("OK: slidekit self test")
    return 0


if __name__ == "__main__":
    sys.exit(main())
