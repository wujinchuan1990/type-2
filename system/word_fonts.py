"""Measure the fonts used by editable Word output without changing its text."""

import os
import re
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont


def font(name):
    return _font(name, os.environ.get("REPORT_WORD_FONT_DIR"))


@lru_cache(maxsize=4)
def _font(name, directory):
    filename = {"Arial": "arial.ttf", "Arial Italic": "ariali.ttf", "宋体": "Simsun.ttc"}[name]
    if directory is not None and (not Path(directory).is_absolute() or not Path(directory).is_dir()):
        raise ValueError("REPORT_WORD_FONT_DIR必须指向已存在的绝对字体目录。")
    candidates = [Path(directory) / filename] if directory is not None else [
        Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts") / filename,
        Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / filename.lower(),
        Path("/System/Library/Fonts/Supplemental") /
        {"Arial": "Arial.ttf", "Arial Italic": "Arial Italic.ttf", "宋体": filename}[name],
    ]
    for path in candidates:
        if path.is_file():
            return ImageFont.truetype(str(path), size=1000)
    if directory is not None:
        raise ValueError(f"已配置的字体目录缺少{name}字体文件{filename}；未使用替代字体生成文档。")
    purpose = "燃气表曲线图绘制" if name == "Arial Italic" else "文档文字宽度计算"
    raise ValueError(f"{purpose}需要本机{name}字体，请安装该字体后重新生成。")


def text_width(text, size):
    return sum(font("Arial" if part.isascii() and not part.isspace() else "宋体").getlength(part)
               for part in re.findall(r"[\x21-\x7f]+|\s+|[^\x00-\x7f\s]+", text)) * size / 1000
