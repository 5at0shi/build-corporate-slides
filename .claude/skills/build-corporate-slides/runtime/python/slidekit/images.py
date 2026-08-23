from pathlib import Path

from PIL import Image


def add_image_contain(slide, path, region):
    """画像を欠落させず、region中央へアスペクト比維持で配置する。"""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"画像が見つかりません: {path}")
    with Image.open(path) as image:
        image_ratio = image.width / image.height
    region_ratio = region.w / region.h
    if image_ratio >= region_ratio:
        width = region.w
        height = int(width / image_ratio)
    else:
        height = region.h
        width = int(height * image_ratio)
    x = region.x + int((region.w - width) / 2)
    y = region.y + int((region.h - height) / 2)
    return slide.shapes.add_picture(str(path), x, y, width=width, height=height)
