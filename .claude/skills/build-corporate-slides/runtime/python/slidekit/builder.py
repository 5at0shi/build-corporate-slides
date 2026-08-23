from datetime import date
from pathlib import Path

from .components import add_cover, add_section_divider, add_slide_title
from .config import load_workspace_config, workspace_paths
from .layout import content_region
from .theme import typography_for
from . import logo_path_from_config, new_presentation


class DeckBuilder:
    def __init__(self, root, config=None, mode=None):
        self.root = Path(root).resolve()
        self.config = config or load_workspace_config(self.root)
        self.paths = workspace_paths(self.root, self.config)
        self.mode = mode or self.config.get("deck", {}).get("mode", "business")
        self.fonts = self.config.get("typography", {})
        self.presentation = new_presentation(mode=self.mode, fonts=self.fonts)
        self.blank_layout = self.presentation.slide_layouts[6]

    @classmethod
    def from_workspace(cls, root):
        return cls(root)

    def add_slide(self, title, *, density="standard", kicker=None, page=None):
        slide = self.presentation.slides.add_slide(self.blank_layout)
        slide._slidekit_typography = typography_for(
            "dense" if density == "dense" else self.mode, fonts=self.fonts)
        add_slide_title(slide, title, kicker=kicker, page=page)
        return slide, content_region()

    def add_cover(self, title, *, subtitle=None, eyebrow=None,
                  brand_side="right", brand_shape="diagonal",
                  brand_width=None, classification=None, created=None):
        slide = self.presentation.slides.add_slide(self.blank_layout)
        slide._slidekit_typography = typography_for(self.mode, fonts=self.fonts)
        organization = self.config.get("organization", {})
        add_cover(
            slide,
            title,
            subtitle=subtitle,
            eyebrow=eyebrow,
            department=organization.get("department", ""),
            classification=(organization.get("classification", "")
                            if classification is None else classification),
            created=created or date.today().isoformat(),
            logo_path=logo_path_from_config(self.config, self.root),
            brand_side=brand_side,
            brand_shape=brand_shape,
            brand_width=brand_width,
        )
        return slide

    def add_section_divider(self, title, *, kicker=None, subtitle=None,
                            page=None):
        slide = self.presentation.slides.add_slide(self.blank_layout)
        slide._slidekit_typography = typography_for(self.mode, fonts=self.fonts)
        add_section_divider(slide, title, kicker=kicker, subtitle=subtitle,
                            page=page)
        return slide

    def save(self, path=None):
        path = Path(path) if path else self.paths.output_dir / "deck.pptx"
        if not path.is_absolute():
            path = (self.root / path).resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        self.presentation.save(path)
        return path
