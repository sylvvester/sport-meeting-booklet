from __future__ import annotations

from dataclasses import dataclass
from math import floor
from unicodedata import east_asian_width

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


@dataclass(frozen=True, slots=True)
class SingleLineFit:
    font_size: float
    adjusted: bool
    fits: bool


def normalize_single_line_text(value: str) -> str:
    """去掉复制粘贴带入的换行和多余空白，避免显式断行绕过版式检查。"""
    return " ".join(value.split())


def visual_text_units(value: str) -> float:
    """按全角汉字宽度估算 Word 单元格中的文本宽度。"""
    units = 0.0
    for character in value:
        if character.isspace():
            units += 0.4
        elif east_asian_width(character) in {"W", "F", "A"}:
            units += 1.0
        else:
            units += 0.55
    return units


def fit_single_line_font(
    value: str,
    capacity: float,
    base_size: float = 10.5,
    minimum_size: float = 7.0,
) -> SingleLineFit:
    """计算保持单行所需字号；字号按 0.5 磅向下取整。"""
    units = visual_text_units(value)
    if not value or units <= capacity:
        return SingleLineFit(base_size, False, True)
    calculated = floor((base_size * capacity / units) * 2) / 2
    font_size = max(minimum_size, calculated)
    fits = units * font_size / base_size <= capacity + 0.01
    return SingleLineFit(font_size, font_size < base_size, fits)


def prevent_cell_text_wrap(cell) -> None:
    properties = cell._tc.get_or_add_tcPr()
    existing = properties.find(qn("w:noWrap"))
    if existing is None:
        properties.append(OxmlElement("w:noWrap"))
