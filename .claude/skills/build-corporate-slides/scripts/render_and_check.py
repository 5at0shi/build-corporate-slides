#!/usr/bin/env python3
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def run(command):
    subprocess.run(command, check=True)


def try_render_keynote(pptx: Path, pdf: Path) -> bool:
    if sys.platform != "darwin":
        return False
    osascript = shutil.which("osascript")
    if not osascript or not Path("/Applications/Keynote.app").exists():
        return False
    def apple_string(value):
        return str(value).replace("\\", "\\\\").replace('"', '\\"')

    input_path = apple_string(pptx.resolve())
    output_path = apple_string(pdf.resolve())
    script = f'''
tell application "Keynote"
    activate
    open POSIX file "{input_path}"
    set sourceDocument to front document
    export sourceDocument to POSIX file "{output_path}" as PDF
    close sourceDocument saving no
end tell
'''
    try:
        run([osascript, "-e", script])
    except subprocess.CalledProcessError:
        return False
    return pdf.is_file() and pdf.stat().st_size > 0


def render_pdf(pptx: Path, pdf: Path) -> None:
    pdf.parent.mkdir(parents=True, exist_ok=True)
    if try_render_keynote(pptx, pdf):
        return
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if office:
        with tempfile.TemporaryDirectory(prefix="slide-render-") as profile:
            profile_uri = Path(profile).resolve().as_uri()
            run([office, f"-env:UserInstallation={profile_uri}", "--headless",
                 "--convert-to", "pdf", "--outdir", str(pdf.parent),
                 str(pptx.resolve())])
        generated = pdf.parent / f"{pptx.stem}.pdf"
        if generated != pdf:
            generated.replace(pdf)
        return
    if sys.platform == "win32":
        try:
            import win32com.client  # type: ignore
            app = win32com.client.Dispatch("PowerPoint.Application")
            deck = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
            deck.SaveAs(str(pdf.resolve()), 32)
            deck.Close(); app.Quit()
            return
        except ImportError:
            pass
    raise RuntimeError("LibreOffice/soffice または Windows PowerPoint COM が必要です")


def render_pages(pdf: Path, pages: Path) -> None:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("ページPNG生成にはPopplerの pdftoppm が必要です")
    pages.mkdir(parents=True, exist_ok=True)
    for old_page in pages.glob("slide-*.png"):
        old_page.unlink()
    prefix = pages / "slide"
    run([pdftoppm, "-png", "-r", "150", str(pdf), str(prefix)])
    for index, source in enumerate(sorted(pages.glob("slide-*.png")), 1):
        target = pages / f"slide-{index:02}.png"
        if source != target:
            source.replace(target)


def main() -> int:
    parser = argparse.ArgumentParser(description="PPTXをPDFとページPNGへレンダリングします")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--pdf", type=Path, default=Path("work/render/deck.pdf"))
    parser.add_argument("--pages", type=Path, default=Path("work/render/pages"))
    args = parser.parse_args()
    try:
        render_pdf(args.pptx, args.pdf)
        render_pages(args.pdf, args.pages)
    except (RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"レンダリング未完了: {exc}", file=sys.stderr)
        return 2
    print(f"PDF: {args.pdf}\nPNG: {args.pages}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
