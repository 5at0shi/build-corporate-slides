#!/usr/bin/env python3
"""samples/配下のYAMLからPDFを作り直す。

スキル側のレンダリングを変更したら、これを実行してサンプルPDFを最新の
実装へ揃える（実際に、文字階層のレンダリング不具合を直した際に7件すべて
作り直しが必要になった）。生成物はPDFのみを追跡し、PPTXは中間物として
build_slides/output/へ置く（YAMLから何度でも作り直せるため）。

使い方（リポジトリ直下から）:
    python3 samples/render_samples.py            # 全件
    python3 samples/render_samples.py 04 05      # 番号で絞る
"""
import subprocess
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / ".claude" / "skills" / "build-corporate-slides"
sys.path.insert(0, str(SKILL / "runtime" / "python"))
sys.path.insert(0, str(SKILL / "scripts"))

from slidekit import inspect_content, render_deck  # noqa: E402
from validate_pptx import validate  # noqa: E402


def render_pdf(pptx: Path, pdf: Path) -> bool:
    """PPTXをPDFへ変換する。レンダラーが無い環境ではPDFを更新せず知らせる。"""
    for binary in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [binary, "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf.parent), str(pptx)],
                check=True, capture_output=True)
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
        produced = pdf.parent / f"{pptx.stem}.pdf"
        if produced != pdf:
            produced.replace(pdf)
        return True
    return False


def main() -> int:
    wanted = sys.argv[1:]
    samples = sorted(Path(__file__).resolve().parent.glob("sample-*.yaml"))
    if wanted:
        samples = [s for s in samples
                   if any(key in s.name for key in wanted)]
    if not samples:
        print("対象のサンプルがありません", file=sys.stderr)
        return 1

    output_dir = ROOT / "build_slides" / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    failed, unrendered = [], []
    for sample in samples:
        content = yaml.safe_load(sample.read_text(encoding="utf-8"))
        errors, warnings = inspect_content(content)
        if errors:
            failed.append(sample.name)
            print(f"NG {sample.name}: " + " / ".join(errors))
            continue
        pptx, _ = render_deck(content, ROOT, output_dir / f"{sample.stem}.pptx")
        issues, edit_warnings = validate(Path(pptx))
        if issues:
            failed.append(sample.name)
        if not render_pdf(Path(pptx), sample.with_suffix(".pdf")):
            unrendered.append(sample.name)
        pages = len(content.get("slides", []))
        print(f"{'NG' if issues else 'OK'} {sample.stem}  "
              f"{pages}ページ / 警告{len(warnings) + len(edit_warnings)}件"
              + ("".join(f"\n     - {i}" for i in issues) if issues else ""))

    if unrendered:
        print("\nPDF未更新（LibreOfficeが見つかりません）: "
              + ", ".join(unrendered), file=sys.stderr)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
