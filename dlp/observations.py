import math
from collections import Counter
from dataclasses import dataclass

from openpyxl.cell.rich_text import CellRichText, TextBlock


def _rgb(color):
    if color is not None and color.type == 'rgb':
        return int(color.rgb[-6:], 16)
    return None


def _runs(cell):
    value = cell.value
    blocks = value if isinstance(value, CellRichText) else [value]
    for block in blocks:
        if isinstance(block, TextBlock):
            font = block.font
            yield block.text, font.rFont or cell.font.name, _rgb(font.color or cell.font.color), bool(font.b)
        else:
            yield str(block), cell.font.name, _rgb(cell.font.color), bool(cell.font.b)


def histogram_entropy(counts):
    total = sum(counts.values())
    return -sum(n / total * math.log2(n / total) for n in counts.values()) if total else 0


@dataclass
class CellObservation:
    address: str
    units: int
    text_bytes: int
    run_lengths: tuple
    colors: Counter
    fonts: Counter
    rotation: int
    hidden: bool
    hidden_number_format: bool


def observe(cell, *, hidden=False):
    if cell.data_type in ('f', 'e') or not isinstance(cell.value, (str, CellRichText)) or not str(cell.value):
        return None
    colors, fonts = Counter(), Counter()
    lengths = []
    for text, font, color, _ in _runs(cell):
        units = len(text.encode('utf-16-le', errors='surrogatepass')) // 2
        if not units:
            continue
        lengths.append(units)
        if color is not None:
            colors[color] += units
        fonts[str(font).lower()] += units
    rotation = cell.alignment.textRotation or 0
    if 90 < rotation < 255:
        rotation = 90 - rotation
    return CellObservation(cell.coordinate, sum(lengths),
                           len(str(cell.value).encode('utf-8', errors='surrogatepass')),
                           tuple(lengths) if isinstance(cell.value, CellRichText) else (),
                           colors, fonts, rotation,
                           hidden,
                           cell.number_format == ';;;')


