#!/usr/bin/env python3
"""图片文字修改引擎：在 JPG / PNG 上「擦除旧字 + 原地重绘新字」。

设计要点：
    1. 不依赖 OCR。本机没有 Tesseract，纯图片读不到文字内容，所以定位全部基于
       像素：先按横向投影切出“文字行”，再按纵向空白切出“文字块”，用户在页面上
       填的「旧文字 -> 新文字」按字宽与文字块匹配，只替换命中的那一块。
    2. 擦除用“背景色覆盖块”，颜色直接采样文字周边的原始像素，做到与底图零色差；
       覆盖块按字高向外扩一点，压掉旧字边缘的抗锯齿残留。
    3. 新字沿用原字高、原墨色，长度超出时自动缩小字号，保证不溢出覆盖块。
    4. 统一输出 PNG，避免 JPG 二次压缩带来的块状噪点。
    5. 带文字层的“图片型 PDF”仍优先走 PDF 文字层定位（更精确）。

主要入口：
    inspect_image(data)                     -> ImageInfo（尺寸 / 预览 PNG）
    replace_image_text(data, region, rules) -> ImageEditResult（output / 预览）
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Sequence

import pymupdf

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------
PREVIEW_MAX_WIDTH = 1200    # 页面预览图的最大宽度（像素）
EDGE_MARGIN = 4             # 覆盖块比文字框外扩的像素（压掉抗锯齿与相邻字的边缘）

INK_DARK = 110              # 亮度低于它视为笔画
INK_LIGHT = 205             # 亮度高于它视为背景

# 无系统字体时的兜底（PyMuPDF 内置 CJK 字体，任何机器都有）
FALLBACK_FONT = "china-s"
SYSTEM_FONT_CANDIDATES = (
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simsun.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
)


class ImageEditError(ValueError):
    """图片无法解析或规则不合法。"""


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------
@dataclass
class TextLine:
    """一行可编辑文字：图片像素坐标（左上原点）。"""

    text: str
    rect: tuple[float, float, float, float]

    @property
    def height(self) -> float:
        return max(1.0, self.rect[3] - self.rect[1])

    @property
    def width(self) -> float:
        return max(1.0, self.rect[2] - self.rect[0])


@dataclass
class ImageInfo:
    """图片概览：页面预览用 PNG、像素尺寸、能读出的文字行。"""

    width: int
    height: int
    preview_png: bytes = b""
    background: tuple[int, int, int] = (255, 255, 255)
    lines: list[TextLine] = field(default_factory=list)
    note: str = ""


@dataclass
class ImageEditResult:
    """一次图片改字的结果。"""

    ok: bool = False
    message: str = ""
    output: bytes | None = None          # PNG 字节（下载用）
    preview_png: bytes | None = None     # 修改后的预览图
    lines: list[TextLine] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------
def normalize(text: str) -> str:
    """全角 -> 半角并去掉所有空白，用于宽松比对。"""
    table = {ord(c): ord(d) for c, d in zip("０１２３４５６７８９ＡＢＣａｂｃ", "0123456789ABCabc")}
    table.update({ord("："): ord(":"), ord("，"): ord(","), ord("（"): ord("("), ord("）"): ord(")")})
    return "".join((text or "").translate(table).split())


def _luma(pixel: Sequence[int]) -> float:
    return 0.299 * pixel[0] + 0.587 * pixel[1] + 0.114 * pixel[2]


def _to_png(pixmap: pymupdf.Pixmap) -> bytes:
    if pixmap.colorspace is None or pixmap.colorspace.n != 3:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
    return pixmap.tobytes("png")


def open_image(data: bytes) -> tuple[pymupdf.Document, int, int]:
    """把图片包成单页 PDF 打开，返回 (文档, 像素宽, 像素高)。

    页面按“1 像素 = 1pt”建立，渲染时用矩阵放大到工作分辨率，
    这样引擎内部坐标可以一直用像素表示，和前端框选坐标一致。
    """
    if not data:
        raise ImageEditError("图片内容为空，请重新上传。")
    try:
        pixmap = pymupdf.Pixmap(data)
    except Exception as exc:  # noqa: BLE001
        raise ImageEditError(f"图片无法解析：{exc}") from exc
    width, height = pixmap.width, pixmap.height
    if width < 8 or height < 8:
        raise ImageEditError("图片尺寸过小，无法修改。")
    png = _to_png(pixmap)
    del pixmap  # 大图占内存，及时释放

    document = pymupdf.open()
    page = document.new_page(width=width, height=height)
    page.insert_image(page.rect, stream=png)
    return document, width, height


def render_pixels(document: pymupdf.Document) -> pymupdf.Pixmap:
    """渲染整页，1pt = 1 像素（即原始图片分辨率）。"""
    return document[0].get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)


def make_preview(pixmap: pymupdf.Pixmap, max_width: int = PREVIEW_MAX_WIDTH) -> bytes:
    """生成页面预览 PNG；超宽时等比缩小，避免浏览器卡顿。"""
    if pixmap.width <= max_width:
        return _to_png(pixmap)
    zoom = max_width / pixmap.width
    doc = pymupdf.open()
    page = doc.new_page(width=pixmap.width, height=pixmap.height)
    page.insert_image(page.rect, stream=_to_png(pixmap))
    small = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    doc.close()
    return _to_png(small)


def sample_pixel(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect) -> tuple[float, float, float]:
    """直接取区域内“代表像素”的原始值作为底色，做到覆盖块与原图零色差。"""
    area = pymupdf.Rect(rect) & pymupdf.Rect(0, 0, pixmap.width, pixmap.height)
    x0, y0, x1, y1 = (int(v) for v in (area.x0, area.y0, area.x1, area.y1))
    if x1 <= x0 or y1 <= y0:
        return (1.0, 1.0, 1.0)
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    history: dict[tuple[int, int, int], int] = {}
    quality: dict[tuple[int, int, int], float] = {}
    step = max(1, int(math.sqrt((x1 - x0) * (y1 - y0) / 2000)))
    for y in range(y0, y1, step):
        row = y * stride
        for x in range(x0, x1, step):
            offset = row + x * n
            pixel = (data[offset], data[offset + 1], data[offset + 2])
            history[pixel] = history.get(pixel, 0) + 1
            quality[pixel] = _luma(pixel)
    if not history:
        return (1.0, 1.0, 1.0)
    # 取出现最多的一档，再在该档里挑最亮的代表像素（纯背景，避开笔画抗锯齿）
    top = max(history.values())
    candidates = [pixel for pixel, count in history.items() if count >= top * 0.6]
    best = max(candidates, key=lambda pixel: quality[pixel])
    return tuple(channel / 255 for channel in best)


def background_color(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect) -> tuple[float, float, float]:
    """取区域内“偏亮的多数派”像素均值作为底色（文字笔画是少数深色像素）。"""
    area = pymupdf.Rect(rect) & pymupdf.Rect(0, 0, pixmap.width, pixmap.height)
    x0, y0, x1, y1 = (int(v) for v in (area.x0, area.y0, area.x1, area.y1))
    if x1 <= x0 or y1 <= y0:
        return (1.0, 1.0, 1.0)

    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    step = max(1, int(math.sqrt((x1 - x0) * (y1 - y0) / 8000)))
    total = [0, 0, 0]
    count = 0
    fallback: dict[tuple[int, int, int], int] = {}
    sampled = 0
    for y in range(y0, y1, step):
        row = y * stride
        for x in range(x0, x1, step):
            offset = row + x * n
            r, g, b = data[offset], data[offset + 1], data[offset + 2]
            sampled += 1
            key = (r // 8, g // 8, b // 8)
            fallback[key] = fallback.get(key, 0) + 1
            if _luma((r, g, b)) >= INK_LIGHT:
                total[0] += r
                total[1] += g
                total[2] += b
                count += 1

    if count * 2 >= sampled and count:  # 背景像素占一半以上：用亮像素均值
        # 量化到 8 位整数再还原，保证覆盖块与底图像素值完全一致（真正零色差）
        return tuple(round(value / count) / 255 for value in total)
    key = max(fallback.items(), key=lambda kv: kv[1])[0]  # 否则最常见的颜色
    return tuple((channel * 8 + 4) / 255 for channel in key)


# --------------------------------------------------------------------------
# 文字行识别（来自 PDF 文字层；纯图片没有文字层，返回空列表）
# --------------------------------------------------------------------------
def _group_lines(words: Sequence[Sequence[Any]], area: pymupdf.Rect | None) -> list[pymupdf.Rect]:
    """把同一条基线上的词合成行框，并按从上到下、从左到右排序。"""
    boxes: list[pymupdf.Rect] = []
    for word in words:
        box = pymupdf.Rect(word[:4])
        if area is not None:
            if not (box in area):  # 只保留完全落在框内的词，避免截断相邻文字
                continue
        boxes.append(box)
    if not boxes:
        return []

    sizes = sorted(box.height for box in boxes)
    tolerance = max(2.0, sizes[len(sizes) // 2] * 0.5)
    boxes.sort(key=lambda b: ((b.y0 + b.y1) / 2, b.x0))

    rows: list[list[pymupdf.Rect]] = []
    for box in boxes:
        center = (box.y0 + box.y1) / 2
        if rows:
            previous = rows[-1]
            ref = sum((b.y0 + b.y1) / 2 for b in previous) / len(previous)
            if abs(center - ref) <= tolerance:
                previous.append(box)
                continue
        rows.append([box])

    merged: list[pymupdf.Rect] = []
    for row in rows:
        rect = pymupdf.Rect(row[0])
        for box in row[1:]:
            rect |= box
        merged.append(rect)
    return merged


def text_lines(document: pymupdf.Document, area: pymupdf.Rect | None = None) -> list[TextLine]:
    """读出图片型 PDF 的文字层，坐标换算成像素（1pt = 1 像素）。"""
    page = document[0]
    try:
        words = page.get_text("words")
    except Exception:  # noqa: BLE001
        return []
    lines: list[TextLine] = []
    for rect in _group_lines(words, area):
        content = " ".join(page.get_textbox(rect).split()).strip()
        if not content:
            continue
        lines.append(TextLine(content, (rect.x0, rect.y0, rect.x1, rect.y1)))
    return lines


def matches(line_text: str, old: str) -> bool:
    """行内容是否就是要替换的那段文字（宽松匹配：忽略空白与全/半角差异）。"""
    target, content = normalize(old), normalize(line_text)
    if not target:
        return False
    return content == target or target in content or content in target


# --------------------------------------------------------------------------
# 字体：优先系统中文字体（字形更接近原图），没有就用内置 CJK 字体
# --------------------------------------------------------------------------
def pick_font() -> tuple[str, str | None]:
    """返回 (字体名, 字体文件)。字体文件为 None 时用 PyMuPDF 内置 CJK 字体。"""
    from pathlib import Path

    for candidate in SYSTEM_FONT_CANDIDATES:
        if Path(candidate).is_file():
            return "zjb-cjk", candidate
    return FALLBACK_FONT, None


def _font_object() -> pymupdf.Font | None:
    """取用于度量文字宽度的字体对象；不可用时返回 None（走粗估）。"""
    try:
        name, file = pick_font()
        return pymupdf.Font(fontfile=file) if file else pymupdf.Font(fontname=name)
    except Exception:  # noqa: BLE001
        return None


def split_rows(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect) -> list[pymupdf.Rect]:
    """按横向投影把区域切成若干“文字行”，每行再收紧到笔画范围。

    纯像素图片没有文字层，只能用像素判断：某一行只要出现过深色笔画，
    就属于正文；连续的空行把上下两段文字分开。
    """
    area = pymupdf.Rect(rect) & pymupdf.Rect(0, 0, pixmap.width, pixmap.height)
    x0, y0, x1, y1 = (int(v) for v in (area.x0, area.y0, area.x1, area.y1))
    if x1 <= x0 or y1 <= y0:
        return [area]

    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    threshold = _ink_threshold(pixmap, area)
    step_x = max(1, (x1 - x0) // 1000)

    bands: list[list[int]] = []          # [top, bottom]
    gap = 0
    for y in range(y0, y1):
        row = y * stride
        inked = any(
            _luma((data[row + x * n], data[row + x * n + 1], data[row + x * n + 2])) <= threshold
            for x in range(x0, x1, step_x)
        )
        if inked:
            if bands and gap <= 1:       # 紧邻上一段，视为同一行
                bands[-1][1] = y
            else:
                bands.append([y, y])
            gap = 0
        else:
            gap += 1

    rows: list[pymupdf.Rect] = []
    for top, bottom in bands:
        if bottom - top < 2:             # 噪点
            continue
        row_rect = pymupdf.Rect(x0, top, x1, bottom + 1)
        left, right = _horizontal_span(pixmap, row_rect, threshold)
        rows.append(pymupdf.Rect(left, top, right, bottom + 1))
    return rows or [area]


def ink_bbox(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect) -> pymupdf.Rect:
    """区域内所有笔画的外接框；没有笔画时返回原区域。"""
    rows = split_rows(pixmap, rect)
    merged = pymupdf.Rect(rows[0])
    for row in rows[1:]:
        merged |= row
    return merged


def _horizontal_span(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect, threshold: float) -> tuple[int, int]:
    """在一条文字带里找最左 / 最右的笔画像素。"""
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    x0, x1 = int(rect.x0), int(rect.x1)
    left, right = x1, x0
    for y in range(int(rect.y0), int(rect.y1)):
        row = y * stride
        for x in range(x0, x1):
            offset = row + x * n
            if _luma((data[offset], data[offset + 1], data[offset + 2])) <= threshold:
                left, right = min(left, x), max(right, x)
    if right < left:
        return x0, x1
    return left, right + 1


def text_width(text: str, height: float, font: pymupdf.Font) -> float:
    """用字体度量算文字在“字块高度 ≈ 字号”时的真实宽度。"""
    size = max(1.0, height)
    return font.text_length(text, fontsize=size)


def _size_score(old: str, blob: pymupdf.Rect, font: pymupdf.Font | None = None) -> float:
    """字块宽度与旧文字真实渲染宽度的接近程度（1 最好，0 表示差太远）。"""
    text = old.strip()
    if not text:
        return 0.0
    height = max(1.0, blob.y1 - blob.y0)
    actual = max(1.0, blob.x1 - blob.x0)
    if font is not None:
        expected = text_width(text, height, font)
    else:  # 兜底：汉字 1 个字宽，ASCII 约 0.55
        expected = sum(0.55 if ord(ch) < 128 else 1.0 for ch in text) * height
    ratio = actual / max(expected, 1.0)
    if not 0.7 <= ratio <= 1.4:
        return 0.0
    return 1.0 - abs(1.0 - ratio)


def find_blobs(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect, gap: int = 6) -> list[pymupdf.Rect]:
    """把一行文字切成若干个“文字块”：字块之间至少隔 gap 个空白列。

    相邻两个字块之间只隔一个窄空白时（例如「姓名： 张三」），后一块视为同一
    标签的值，单独成块，这样替换时不会连标签一起擦掉。
    """
    rows = split_rows(pixmap, rect)
    if not rows:
        return []
    band = rows[0]
    top, bottom = int(band.y0), int(band.y1)
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    threshold = _ink_threshold(pixmap, band)
    x0, x1 = int(band.x0), int(band.x1)

    columns = [
        any(
            _luma((data[y * stride + x * n], data[y * stride + x * n + 1], data[y * stride + x * n + 2]))
            <= threshold
            for y in range(top, bottom)
        )
        for x in range(x0, x1)
    ]

    segments: list[list[int]] = []   # [起, 止) 的列区间
    start = None
    blank = 0
    for offset, inked in enumerate(columns):
        if inked:
            if start is None:
                start = offset
            elif blank >= gap:
                segments.append([start, offset - blank])
                start = offset
            blank = 0
        else:
            blank += 1
    if start is not None:
        segments.append([start, len(columns) - blank if blank else len(columns)])

    blobs: list[list[int]] = []
    for left, right in segments:
        if right - left >= 2:            # 过滤噪点（1~2 列的零星笔画）
            blobs.append([left, right])

    result = [pymupdf.Rect(x0 + left, top, x0 + right, bottom) for left, right in blobs]
    return result or [band]


def find_blobs(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect, gap: int = 6) -> list[pymupdf.Rect]:
    """把一行文字切成若干个“文字块”（按列跳过空白），块之间至少隔 gap 个空白列。"""
    rows = split_rows(pixmap, rect)
    if not rows:
        return []
    band = rows[0] if len(rows) == 1 else None
    if band is None:  # 传入的是整张图：只取第一行
        band = rows[0]
    top, bottom = int(band.y0), int(band.y1)
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    threshold = _ink_threshold(pixmap, band)
    x0, x1 = int(band.x0), int(band.x1)

    columns: list[bool] = []
    for x in range(x0, x1):
        inked = any(
            _luma((data[y * stride + x * n], data[y * stride + x * n + 1], data[y * stride + x * n + 2]))
            <= threshold
            for y in range(top, bottom)
        )
        columns.append(inked)

    blobs: list[list[int]] = []
    start = None
    blank = 0
    for offset, inked in enumerate(columns):
        if inked:
            if start is None:
                start = offset
            elif blank >= gap:
                blobs.append([start, offset - blank])
                start = offset
            blank = 0
        else:
            blank += 1
    if start is not None:
        blobs.append([start, len(columns) - blank if blank else len(columns)])

    result: list[pymupdf.Rect] = []
    for left, right in blobs:
        if right - left < 2:
            continue
        box = pymupdf.Rect(x0 + left, top, x0 + right, bottom)
        result.append(box)
    return result or [band]


def _ink_threshold(pixmap: pymupdf.Pixmap, area: pymupdf.Rect) -> float:
    """笔画判定阈值：底色亮度与最暗像素亮度之间取一个偏暗的分界。"""
    background = _luma(tuple(int(v * 255) for v in background_color(pixmap, area)))
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    sample = 0
    darkest = 255.0
    for y in range(int(area.y0), int(area.y1), 3):
        row = y * stride
        for x in range(int(area.x0), int(area.x1), 3):
            offset = row + x * n
            darkest = min(darkest, _luma((data[offset], data[offset + 1], data[offset + 2])))
            sample += 1
    del sample
    return (background + darkest) / 2 if darkest < background - 20 else background - 40


def _ink_color(pixmap: pymupdf.Pixmap, rect: pymupdf.Rect) -> tuple[float, float, float]:
    """取区域内深色笔画像素的均值作为文字颜色；读不到时按黑字处理。"""
    area = pymupdf.Rect(rect) & pymupdf.Rect(0, 0, pixmap.width, pixmap.height)
    x0, y0, x1, y1 = (int(v) for v in (area.x0, area.y0, area.x1, area.y1))
    data, stride, n = pixmap.samples, pixmap.stride, pixmap.n
    darkest = min(x1 - x0, 400)
    step = max(1, (x1 - x0) // max(darkest, 1))
    total = [0, 0, 0]
    count = 0
    for y in range(y0, y1):
        row = y * stride
        for x in range(x0, x1, step):
            offset = row + x * n
            r, g, b = data[offset], data[offset + 1], data[offset + 2]
            if _luma((r, g, b)) <= INK_DARK:
                total[0] += r
                total[1] += g
                total[2] += b
                count += 1
    if not count:
        return (0.0, 0.0, 0.0)
    return tuple(round(value / count) / 255 for value in total)


def inspect_image(data: bytes) -> ImageInfo:
    """打开图片，给出尺寸、预览图和可读到的文字行（用于页面展示提示）。"""
    document, width, height = open_image(data)
    try:
        pixmap = render_pixels(document)
        info = ImageInfo(
            width=width,
            height=height,
            preview_png=make_preview(pixmap),
            lines=text_lines(document),
        )
        if info.lines:
            info.note = f"检测到 {len(info.lines)} 行可定位文字（图片自带文字层）。"
        else:
            info.note = "这是一张纯像素图片，请在预览图上框出要修改的位置。"
        return info
    finally:
        document.close()


def replace_image_text(
    data: bytes,
    region: tuple[float, float, float, float] | None,
    rules: Sequence[tuple[str, str]],
    *,
    replace_all: bool = False,
) -> ImageEditResult:
    """在框选区域里把旧文字擦掉、按原位置重绘新文字。

    region:      图片像素坐标 (x0, y0, x1, y1)；None 表示整张图。
    rules:       [(旧文字, 新文字), ...]。
    replace_all: 区域里有文字层且能定位时，是否替换全部匹配行；默认只改最匹配的一行。
    """
    result = ImageEditResult()
    pairs = [(old.strip(), new.strip()) for old, new in rules if old.strip()]
    if not pairs:
        result.message = "请至少填写一条「旧文字 -> 新文字」规则。"
        return result

    try:
        document, width, height = open_image(data)
    except ImageEditError as exc:
        result.message = str(exc)
        return result

    try:
        pixmap = render_pixels(document)
        area = pymupdf.Rect(0, 0, width, height)
        if region:
            area = pymupdf.Rect(*region) & pymupdf.Rect(0, 0, width, height)
            if area.is_empty:
                result.message = "框选区域超出图片范围，请在预览图上重新框选。"
                return result

        targets = text_lines(document, area)
        converted: list[TextLine] = []
        for old, new in pairs:
            found = [line for line in targets if matches(line.text, old)]
            if not found:
                if not targets:
                    # 纯像素图没有文字层：先按行切分，再按列切成文字块。
                    # 优先用旧文字长度匹配字形宽度最接近的块，匹配不到再按顺序补齐。
                    rows = split_rows(pixmap, area)
                    blobs: list[pymupdf.Rect] = []
                    for row in rows:
                        blobs.extend(find_blobs(pixmap, row))

                    used: set[int] = set()
                    chosen: list[tuple[str, pymupdf.Rect]] = []
                    metric = _font_object()
                    for old, new in pairs:
                        candidates = [
                            index for index, blob in enumerate(blobs)
                            if index not in used and _size_score(old, blob, metric) > 0
                        ]
                        if not candidates:
                            candidates = [i for i in range(len(blobs)) if i not in used]
                        if not candidates:
                            break
                        pick = max(candidates, key=lambda i: _size_score(old, blobs[i], metric))
                        used.add(pick)
                        chosen.append((new, blobs[pick]))

                    converted.extend(
                        TextLine(new, (blob.x0, blob.y0, blob.x1, blob.y1))
                        for new, blob in chosen
                    )
                    result.warnings.append(
                        f"图片没有文字层，已自动识别 {len(blobs)} 段文字，"
                        f"按字形宽度匹配并替换 {len(chosen)} 段。"
                        + (f" 未匹配的 {len(pairs) - len(chosen)} 条规则已跳过。"
                           if len(chosen) < len(pairs) else "")
                    )
                    break
                result.message = f"框选区域内没有找到「{old}」，请检查文字或改用带文字层的 PDF。"
                return result

            chosen = found if replace_all else [min(found, key=lambda l: abs(l.width - len(old) * l.height))]
            for line in chosen:
                if line not in converted:
                    converted.append(TextLine(new, line.rect))

        if targets:
            # 带文字层的图片（如“图片型 PDF”）：按 PDF 坐标拿到文字框后，再收紧到像素笔画
            refined: list[TextLine] = []
            for item in converted:
                box = ink_bbox(pixmap, pymupdf.Rect(item.rect))
                refined.append(TextLine(item.text, (box.x0, box.y0, box.x1, box.y1)))
            converted = refined

        if not converted:
            result.message = "没有找到可以替换的文字。"
            return result

        # 逐处：取底色 -> 覆盖 -> 重绘
        font_name, font_file = pick_font()
        for item in converted:
            rect = pymupdf.Rect(item.rect)
            ink = _ink_color(pixmap, rect)
            page = document[0]
            # 覆盖色取文字正下方紧邻的原始像素：那里是纯背景，和底图完全一致（零色差）
            pad = EDGE_MARGIN + 2
            if rect.y1 + pad <= pixmap.height:
                probe = pymupdf.Rect(rect.x0, rect.y1 + 1, rect.x1, rect.y1 + pad)
            elif rect.y0 - pad >= 1:
                probe = pymupdf.Rect(rect.x0, rect.y0 - pad, rect.x1, rect.y0 - 1)
            else:
                probe = pymupdf.Rect(rect)
            fill_color = sample_pixel(pixmap, probe)
            # 覆盖块向外扩：两侧多留出“新字可能超出旧字框”的余量，压掉旧字边缘残留
            pad_x = max(EDGE_MARGIN, rect.height * 0.45)
            pad_y = max(EDGE_MARGIN, rect.height * 0.2)
            cover = pymupdf.Rect(
                rect.x0 - pad_x,
                rect.y0 - pad_y,
                rect.x1 + pad_x,
                rect.y1 + pad_y,
            )
            page.draw_rect(cover, color=None, fill=fill_color, width=0)
            # 字号：以“旧字墨迹高度”为基准，再按可用宽度收一点，防止新字超出覆盖块
            avail = cover.width - 2
            size = max(6.0, (rect.height - 1) / 0.95)
            font_obj = _font_object()
            if font_obj is not None and item.text:
                measured = font_obj.text_length(item.text, fontsize=size)
                if measured > avail > 0:
                    size = max(6.0, size * avail / measured)
            page.insert_text(
                (rect.x0, rect.y1 - size * 0.2),
                item.text,
                fontname=font_name,
                fontfile=font_file,
                fontsize=size,
                color=ink,
            )
            result.lines.append(TextLine(item.text, (rect.x0, rect.y0, rect.x1, rect.y1)))

        changed = document[0].get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)
        result.output = _to_png(changed)
        result.preview_png = make_preview(changed)
        result.ok = True
        result.message = "".join(
            f"{old} → {new}；" for old, new in pairs
        ).rstrip("；") + f"（共 {len(result.lines)} 处）"
        return result
    except ImageEditError as exc:
        result.message = str(exc)
        return result
    except Exception as exc:  # noqa: BLE001 - 任何底层异常都转成中文提示
        result.message = f"图片修改失败：{exc}"
        return result
    finally:
        document.close()
