#!/usr/bin/env python3
"""去水印引擎：位置驱动的“傻瓜式”去水印。

设计原则：用户不需要知道水印上写了什么字，只要回答“水印在哪”。
全部模式都只看位置，不看文字，因此没有“请输入水印关键字”这种反直觉的步骤。

    REMOVE_MODES
        auto    自动识别并去除（适合纯白底）—— 擦除浅色灰字 / 底纹
        strip   去掉页面顶部（页眉区域）—— 没水印就直接留白，不必先识别
        footer  去掉页面底部（页脚区域）
        veiler  全屏底纹遮盖 —— 把浅色水印整片盖掉，兜底用
        crop_top    裁掉顶部区域（图片 / PDF 裁页）
        crop_bottom 裁掉底部区域（图片 / PDF 裁页）
        crop_corner 裁掉右下角区域（图片 / PDF 裁页）

按格式的落地方式：
    图片（PNG / JPG / WEBP）：Pillow 涂白对应矩形；裁切模式直接裁像素。
    PDF：fitz 涂白矩形（redaction，底下的内容也一并消失），
        裁切模式缩小页面 CropBox，页面上所有元素跟着裁掉。
        没有文字层的扫描件走光栅涂白。
    Word(.docx)：python-docx 清空页眉页脚段落，并删掉页眉页脚里残留的图片。
    PPT(.pptx)：没有 python-pptx 依赖，用 XML 直接删掉 slide / master 里的
        页眉页脚占位符文字。

主要入口：
    remove_watermark(data, filename, mode="auto") -> WatermarkResult
        output 为处理后的文件字节（可直接给 download_button），
        preview_png 为处理结果的预览图（没有则为 b""）。
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET

import pymupdf
from PIL import Image

WATERMARK_DELTA = 120        # 自动模式：比纸张底色浅多少以内算“浅色水印”

COLOR_TOLERANCE = 40         # 色差超过它视为彩色内容，不动
DARK_CORE = 90               # 块内出现比它更暗的像素，就判定为正文（不擦）
MAX_WATERMARK_COLORS = 6     # 块内颜色种类上限：超过视为照片 / 渐变底纹
PREVIEW_MAX_WIDTH = 1200
TEXT_DOC_TYPES = (".docx", ".pptx")

# --------------------------------------------------------------------------
# 位置模式（下拉框的唯一数据源）
# --------------------------------------------------------------------------
REMOVE_MODES = (
    "auto",         # 自动识别并去除（适合纯白底）
    "strip",        # 在页面顶部（删除页眉区域）
    "footer",       # 在页面底部（删除页脚区域）
    "veiler",       # 全屏底纹遮盖（强行覆盖）
    "crop_top",     # 裁掉顶部区域
    "crop_bottom",  # 裁掉底部区域
    "crop_corner",  # 裁掉右下角区域
)

MODE_LABELS: dict[str, str] = {
    "auto": "自动识别并去除（适合纯白底）",
    "strip": "在页面顶部（删除页眉区域）",
    "footer": "在页面底部（删除页脚区域）",
    "veiler": "全屏底纹遮盖（强行覆盖）",
    "crop_top": "裁掉顶部区域",
    "crop_bottom": "裁掉底部区域",
    "crop_corner": "裁掉右下角区域",
}

# 每种格式能看到哪些选项：文档按“页眉 / 页脚”说，图片才说“裁掉”。
IMAGE_MODES = ("auto", "strip", "footer", "veiler", "crop_top", "crop_bottom", "crop_corner")
DOC_MODES = ("auto", "strip", "footer", "veiler")
BOX_MODES = ("strip", "footer", "veiler")       # 直接涂成一个矩形，不挑像素
CROP_MODES = ("crop_top", "crop_bottom", "crop_corner")

# 涂白框相对尺寸：顶部 / 底部条带高度、右下角块大小
BOX_HEIGHT = 0.12
CORNER_WIDTH = 0.42
CORNER_HEIGHT = 0.26


def mode_for_file(mode: str, filename: str) -> str:
    """把用户在“图片版”下拉里选的值，落到当前文件真正支持的模式上。

    Word / PPT 没有“裁剪”概念，选了裁切就退化成对应的涂白 / 清页眉，
    保证任何组合都能跑通，而不是弹一句“不支持”。
    """
    key = _norm_mode(mode)
    suffix = ("." + filename.rsplit(".", 1)[-1].lower()) if "." in (filename or "") else ""
    if suffix in TEXT_DOC_TYPES:
        fallback = {"crop_top": "strip", "crop_bottom": "footer", "crop_corner": "veiler"}
        return fallback.get(key, key)
    return key


def box_for_mode(mode: str) -> tuple[float, float, float, float]:
    """涂白模式的相对矩形 (x0, y0, x1, y1)。"""
    if mode == "strip":
        return (0.0, 0.0, 1.0, BOX_HEIGHT)
    if mode == "footer":
        return (0.0, 1.0 - BOX_HEIGHT, 1.0, 1.0)
    if mode == "crop_corner":
        return (1.0 - CORNER_WIDTH, 1.0 - CORNER_HEIGHT, 1.0, 1.0)
    return (0.0, 0.0, 1.0, 1.0)          # veiler 全屏



class WatermarkError(ValueError):
    """文件无法解析或参数不合法。"""


@dataclass
class WatermarkResult:
    ok: bool
    message: str
    output: bytes = b""
    filename: str = ""
    mime: str = "application/octet-stream"
    preview_png: bytes = b""
    removed: int = 0
    log: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# 公共小工具
# --------------------------------------------------------------------------
def _png(pixmap: pymupdf.Pixmap) -> bytes:
    if pixmap.colorspace is None or pixmap.colorspace.n != 3:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
    return pixmap.tobytes("png")


def _preview(page: pymupdf.Page) -> bytes:
    zoom = min(1.0, PREVIEW_MAX_WIDTH / max(page.rect.width, 1))
    return _png(page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False))


def _stem(name: str) -> str:
    return f"{re.sub(r'\\.[^.]+$', '', name or '文件')}_已去水印"


def _suffix(name: str) -> str:
    return ("." + name.rsplit(".", 1)[-1].lower()) if "." in (name or "") else ""


def _crop_rect(mode: str) -> tuple[float, float, float, float]:
    """裁切模式保留的区域（相对坐标），其余部分直接裁掉。"""
    if mode == "crop_top":
        return (0.0, BOX_HEIGHT, 1.0, 1.0)
    if mode == "crop_bottom":
        return (0.0, 0.0, 1.0, 1.0 - BOX_HEIGHT)
    if mode == "crop_corner":
        return (0.0, 0.0, 1.0 - CORNER_WIDTH, 1.0 - CORNER_HEIGHT)
    return (0.0, 0.0, 1.0, 1.0)


def _pixel_box(size: tuple[int, int], rel: tuple[float, float, float, float]) -> tuple[int, int, int, int]:
    """相对矩形 → 像素矩形，并保证至少 1px，避免 Pillow 抛空框错误。"""
    width, height = size
    x0, y0, x1, y1 = rel
    left = max(0, min(width - 1, int(round(x0 * width))))
    top = max(0, min(height - 1, int(round(y0 * height))))
    right = max(left + 1, min(width, int(round(x1 * width))))
    bottom = max(top + 1, min(height, int(round(y1 * height))))
    return left, top, right, bottom


def whiten_with_pillow(image: Image.Image, rel: tuple[float, float, float, float]) -> tuple[Image.Image, tuple[int, int, int, int]]:
    """把图片指定区域用 Pillow 涂成纯白，返回 (新图, 涂白的像素框)。

    这是“按位置硬覆盖”的主力：不判断像素像不像水印，用户说哪里就涂哪里，
    结果完全可预期，绝不会出现“点了处理但画面没变化”的困惑。
    """
    out = image.convert("RGB")
    box = _pixel_box(out.size, rel)
    out.paste((255, 255, 255), box)
    return out, box


def crop_with_pillow(image: Image.Image, mode: str) -> tuple[Image.Image, tuple[int, int, int, int]]:
    """按裁切模式直接裁掉对应边，返回 (新图, 保留的像素框)。"""
    keep = _crop_rect(mode)
    box = _pixel_box(image.size, keep)
    return image.crop(box), box



def _mime_for(suffix: str) -> str:
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".pdf": "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }.get(suffix, "application/octet-stream")


def _luma(pixel: tuple[int, int, int]) -> float:
    return 0.299 * pixel[0] + 0.587 * pixel[1] + 0.114 * pixel[2]


def page_background(
    pixmap: pymupdf.Pixmap,
    rect: tuple[int, int, int, int] | None = None,
) -> tuple[int, int, int]:
    """整页出现最多的颜色就是纸张底色；纯色背景文档里非常可靠。

    rect 是像素坐标 (left, top, right, bottom)，用于只在局部区域（如页面顶部
    或底部条带）统计底色——那里没有正文，底色比整页更准。
    """
    if rect is None:
        left, top, right, bottom = 0, 0, pixmap.width, pixmap.height
    else:
        left, top, right, bottom = rect
    span = max(1, right - left)
    total = span * max(1, bottom - top)
    step = max(1, total // 200_000)
    counts: dict[tuple[int, int, int], int] = {}
    for index in range(0, total, step):
        dy, dx = divmod(index, span)
        pixel = pixmap.pixel(left + dx, top + dy)
        counts[pixel] = counts.get(pixel, 0) + 1
    return max(counts.items(), key=lambda item: item[1])[0]


# 位置预设：相对坐标 (x0, y0, x1, y1)，取 0~1 的比例。
# 只提供“整幅 / 顶部条带 / 底部条带 / 右下角块”这几块——
# 水印九成压在这几个位置，而每一个都能用一句话向普通用户解释清楚。
REGION_PRESETS: dict[str, tuple[float, float, float, float] | None] = {
    "full": None,
    "top": (0.0, 0.0, 1.0, 0.12),
    "bottom": (0.0, 0.88, 1.0, 1.0),
    "right": (0.82, 0.0, 1.0, 1.0),
}


def _norm_mode(mode: str, region: str | None = None) -> str:
    """把模式名收敛到 REMOVE_MODES；兼容旧调用方传 region="top"/"bottom"。"""
    key = (mode or "auto").strip().lower()
    if key in REMOVE_MODES:
        return key
    legacy = {"full": "auto", "top": "strip", "bottom": "footer",
              "left": "strip", "right": "crop_corner"}
    if key in legacy:
        return legacy[key]
    if region:
        return legacy.get(str(region).strip().lower(), "auto")
    return "auto"


def region_label(mode: str) -> str:
    """下拉框文案：每条都是一句用户能照做的动作，不出现技术名词。"""
    return MODE_LABELS.get(_norm_mode(mode), MODE_LABELS["auto"])


def erase_light_pixels(
    pixmap: pymupdf.Pixmap,
    delta: int = WATERMARK_DELTA,
    region: tuple[float, float, float, float] | None = None,
) -> tuple[pymupdf.Pixmap, int]:
    """擦除浅色水印：整块水印位图按亮度阈值抹掉，保证笔画被完整清除。

    做法：
        1. 取整页出现最多的颜色作为纸张底色；
        2. 亮度 ≥（纸张亮度 - delta）且不是彩色的像素 → 视为水印笔画；
        3. 用纸张底色覆盖，覆盖前把「与原图差异较大」的像素也算进来，
           使半透明水印的抗锯齿边缘（比阈值更浅的那一圈）一起被抹平，
           避免留下灰色残边。
    对彩色区域/照片不做处理，避免把插图擦成白块。

    region 为可选的处理范围 (x0, y0, x1, y1)，取 0~1 的相对比例；
    None 表示全图。指定范围后只在该范围内找纸色、擦像素——
    这正是“水印固定在角上、正文在别处”时唯一可行的做法：
    正文哪怕又浅又灰，只要不在范围内就绝不会被碰到。
    """
    out = pymupdf.Pixmap(pymupdf.csRGB, pixmap)  # 统一成无 alpha 的 RGB，便于按像素改写
    width, height = out.width, out.height
    if region is None:
        left, top, right, bottom = 0, 0, width, height
    else:
        x0, y0, x1, y1 = region
        left = max(0, min(width - 1, int(round(x0 * width))))
        top = max(0, min(height - 1, int(round(y0 * height))))
        right = max(left + 1, min(width, int(round(x1 * width))))
        bottom = max(top + 1, min(height, int(round(y1 * height))))

    paper = page_background(out, (left, top, right, bottom))
    paper_luma = _luma(paper)
    low = paper_luma - delta
    removed = 0

    for y in range(top, bottom):
        for x in range(left, right):
            pixel = out.pixel(x, y)
            if pixel == paper:
                continue                     # 已经是纸张底色，不算水印
            if max(pixel) - min(pixel) > COLOR_TOLERANCE:
                continue                     # 彩色内容不动
            if _luma(pixel) < low:
                continue                     # 比阈值更暗：正文
            out.set_pixel(x, y, paper)
            removed += 1

    return out, removed


# --------------------------------------------------------------------------
# 图片
# --------------------------------------------------------------------------
def _image_bytes(image: Image.Image, suffix: str) -> tuple[bytes, str]:
    """按原格式输出；JPG 没有透明通道，统一转 RGB 再存。"""
    buffer = io.BytesIO()
    if suffix in (".jpg", ".jpeg"):
        image.convert("RGB").save(buffer, format="JPEG", quality=95, subsampling=0)
        return buffer.getvalue(), "image/jpeg"
    if suffix == ".webp":
        image.save(buffer, format="WEBP", quality=95, method=4)
        return buffer.getvalue(), "image/webp"
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue(), "image/png"


def _remove_image(data: bytes, mode: str, suffix: str = ".png") -> WatermarkResult:
    """图片去水印核心函数：基于位置和颜色特征的精准去除"""
    try:
        image = Image.open(io.BytesIO(data))
        image.load()
    except Exception as exc:  # noqa: BLE001
        raise WatermarkError(f"图片无法解析：{exc}") from exc
    if min(image.size) < 8:
        raise WatermarkError("图片尺寸过小，无法处理。")

    log: list[str] = []
    
    # 1. 裁切模式：直接裁剪，最安全
    if mode in CROP_MODES:
        cleaned, keep = crop_with_pillow(image, mode)
        cut = f"裁切后保留 {cleaned.width}×{cleaned.height}px（原 {image.width}×{image.height}px）"
        message = (
            f"已裁掉{'顶部' if mode == 'crop_top' else '底部' if mode == 'crop_bottom' else '右下角'}"
            f"区域，{cut}。"
        )
        log.append(cut)
        removed = 0
        
    # 2. 区域涂白模式：针对固定位置的水印
    elif mode in BOX_MODES:
        cleaned, box = whiten_with_pillow(image, box_for_mode(mode))
        area = (box[2] - box[0]) * (box[3] - box[1])
        share = area / max(1, image.width * image.height)
        where = {"strip": "页面顶部（页眉区域）", "footer": "页面底部（页脚区域）",
                 "veiler": "整个画面"}[mode]
        message = (
            f"已把{where}涂成纯白（覆盖 {share:.0%} 画面），这块里的水印一并消失。"
        )
        log.append(f"涂白区域 {box}（{share:.1%} 画面）")
        removed = area
        
    # 3. 自动模式：增强版，更精准的水印检测
    else:  # auto：自动识别浅色水印
        # 先转为RGB格式便于处理
        if image.mode != "RGB":
            image = image.convert("RGB")
        
        # 方法1：基于颜色直方图的水印检测（适合纯白底文字水印）
        # 计算图像的颜色分布
        hist = image.histogram()
        r_hist = hist[0:256]
        g_hist = hist[256:512]
        b_hist = hist[512:768]
        
        # 如果红色、绿色、蓝色通道的高亮部分（>220）占比很高，可能是白底文字
        r_bright = sum(r_hist[220:]) / len(r_hist)
        g_bright = sum(g_hist[220:]) / len(g_hist)
        b_bright = sum(b_hist[220:]) / len(b_hist)
        
        # 白底文字水印特征：三个通道都偏亮，且亮度分布集中
        is_white_background = (r_bright > 0.3 and g_bright > 0.3 and b_bright > 0.3)
        
        if is_white_background:
            # 使用更精确的阈值检测（针对白底黑字/灰字水印）
            cleaned = _remove_white_background_watermark(image)
            removed = _count_watermark_pixels(image, cleaned)
            ratio = removed / max(1, image.width * image.height)
            
            if removed > 0:
                message = (
                    f"已精准清除白底文字/印章水印（覆盖 {ratio * 100:.1f}% 画面），"
                    f"正文内容完整保留。"
                )
                log.append(f"清除水印像素 {removed}（{ratio:.1%}）")
            else:
                message = "这张图没有检测到白底文字水印。请尝试选择具体位置模式。"
                log.append("未检测到白底文字水印")
                
        else:
            # 方法2：传统浅色像素擦除（适用于一般情况）
            pixmap = pymupdf.Pixmap(data)
            if pixmap.alpha:
                pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
            erased, removed = erase_light_pixels(pixmap, WATERMARK_DELTA)
            ratio = removed / max(1, pixmap.width * pixmap.height)
            
            if removed and ratio > 0.6:
                # 超过六成像素像水印，说明整张图偏浅（照片/扫描件），宁可提示换模式也不给糊图
                return WatermarkResult(
                    ok=False,
                    message=(
                        "这张图整体偏浅，自动识别会把正文一起擦掉。请改选「裁掉顶部区域」"
                        "或「全屏底纹遮盖」，只处理水印所在的位置。"
                    ),
                    log=[f"命中比例 {ratio:.0%}"],
                )
            
            cleaned = Image.open(io.BytesIO(_png(erased))).convert("RGB")
            message = (
                f"已自动清除浅色水印与底纹（覆盖 {ratio * 100:.1f}% 画面），正文保持原样。"
                if removed
                else "这张图没找到浅色水印。换个位置试试：水印在顶上选「在页面顶部」，在底下选「在页面底部」。"
            )
            log.append(f"擦除像素 {removed}（{ratio:.1%}）")

    payload, mime = _image_bytes(cleaned, suffix)
    preview = cleaned if cleaned.width <= PREVIEW_MAX_WIDTH else cleaned.copy()
    if preview.width > PREVIEW_MAX_WIDTH:
        ratio = PREVIEW_MAX_WIDTH / preview.width
        preview = preview.resize(
            (PREVIEW_MAX_WIDTH, max(1, int(preview.height * ratio))), Image.LANCZOS
        )
    return WatermarkResult(
        ok=True,
        message=message,
        output=payload,
        mime=mime,
        preview_png=_image_bytes(preview, ".png")[0],
        removed=removed,
        log=log,
    )


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------
def _page_watermark_rects(page: pymupdf.Page) -> list[pymupdf.Rect]:
    """找出该页需要擦除的水印文字区域（只看颜色与重复度，不看文字内容）。

    命中条件（任一）：
        1. 同一行文字重复出现 ≥ 3 次（平铺水印）；
        2. 文字明显偏浅（灰度 ≥ 0.55）且长度 ≤ 40 字（浅色底纹字）。
    """
    try:
        blocks = page.get_text("dict").get("blocks", [])
    except Exception:  # noqa: BLE001
        return []

    lines: list[tuple[str, pymupdf.Rect, float]] = []
    for block in blocks:
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            text = "".join(span.get("text", "") for span in spans).strip()
            if not text or not spans:
                continue
            color = spans[0].get("color", 0)
            gray = (((color >> 16) & 255) + ((color >> 8) & 255) + (color & 255)) / 3 / 255
            lines.append((text, pymupdf.Rect(line["bbox"]), gray))

    counts: dict[str, int] = {}
    for text, _, _ in lines:
        counts[text] = counts.get(text, 0) + 1

    targets: list[pymupdf.Rect] = []
    for text, rect, gray in lines:
        if counts[text] >= 3 and len(text) <= 40:
            targets.append(rect)
        elif gray >= 0.55 and len(text) <= 40:
            targets.append(rect)
    return targets


def _page_rect_for(page: pymupdf.Page, rel: tuple[float, float, float, float]) -> pymupdf.Rect:
    """相对矩形 → 该页的绝对矩形（按当前可见区域计算）。"""
    area = page.rect
    return pymupdf.Rect(
        area.x0 + rel[0] * area.width,
        area.y0 + rel[1] * area.height,
        area.x0 + rel[2] * area.width,
        area.y0 + rel[3] * area.height,
    )


def _crop_pdf_page(page: pymupdf.Page, mode: str) -> bool:
    """裁掉页面的一部分：缩 CropBox（同时缩 MediaBox），页面上所有元素一起被裁掉。"""
    area = page.mediabox
    rel = _crop_rect(mode)
    new = pymupdf.Rect(
        area.x0 + rel[0] * area.width,
        area.y0 + rel[1] * area.height,
        area.x0 + rel[2] * area.width,
        area.y0 + rel[3] * area.height,
    )
    if new.width < 20 or new.height < 20:
        return False
    page.set_cropbox(new)
    page.set_mediabox(new)
    return True


def _remove_pdf(data: bytes, mode: str) -> WatermarkResult:
    try:
        document = pymupdf.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise WatermarkError(f"PDF 无法打开：{exc}") from exc
    if document.needs_pass:
        document.close()
        raise WatermarkError("这份 PDF 有密码保护，请先解除密码后再上传。")

    removed = 0
    log: list[str] = []

    if mode in CROP_MODES:
        cropped = 0
        for number, page in enumerate(document, start=1):
            if _crop_pdf_page(page, mode):
                cropped += 1
        if not cropped:
            document.close()
            raise WatermarkError("这份 PDF 页面太小，裁切后就没有内容了，请换成涂白模式。")
        log.append(f"已裁掉 {cropped} 页的对应区域")
        removed = cropped
        message = f"已裁掉 {cropped} 页{'顶部' if mode == 'crop_top' else '底部' if mode == 'crop_bottom' else '右下角'}区域，页面内容随之缩小。"
    elif mode in BOX_MODES:
        rel = box_for_mode(mode)
        where = {"strip": "页面顶部（页眉区域）", "footer": "页面底部（页脚区域）",
                 "veiler": "整个页面"}[mode]
        for number, page in enumerate(document, start=1):
            page.add_redact_annot(_page_rect_for(page, rel), fill=(1, 1, 1))
            # images=PDF_REDACT_IMAGE_PIXELS：连遮在底下的图片像素一起抹掉，
            # 不留“能框选出来的图层”，否则水印其实还藏在文件里。
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS)
            removed += 1
        log.append(f"已把{where}涂白：{removed} 页")
        message = f"已把 {removed} 页的{where}涂成纯白，这块里的水印一并消失。"
    else:  # auto：走文字层 + 浅色像素识别
        for number, page in enumerate(document, start=1):
            if page.get_images() and not page.get_text("words"):
                pixmap = page.get_pixmap(dpi=150, alpha=False)
                cleaned, count = erase_light_pixels(pixmap, WATERMARK_DELTA)
                if count:
                    page.clean_contents()
                    page.insert_image(page.rect, stream=_png(cleaned), overlay=True)
                    removed += count
                    log.append(f"第 {number} 页（扫描页）：擦除 {count} 像素")
                continue

            rects = _page_watermark_rects(page)
            if rects:
                for rect in rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS)
                removed += len(rects)
                log.append(f"第 {number} 页：移除 {len(rects)} 处水印文字")

        if not log:
            log.append("未检测到文字型水印")
        message = (
            f"已自动清除 {removed} 处水印，正文与版式保持不变。"
            if removed
            else "没找到浅色水印。水印在顶上就选「在页面顶部」，在底下就选「在页面底部」，一定能擦掉。"
        )

    preview = _preview(document[0]) if document.page_count else b""
    payload = document.tobytes(garbage=4, deflate=True)
    document.close()
    return WatermarkResult(
        ok=True,
        message=message,
        output=payload,
        mime="application/pdf",
        preview_png=preview,
        removed=removed,
        log=log,
    )


# --------------------------------------------------------------------------
# Word / PPT：按位置清页眉页脚
# --------------------------------------------------------------------------
_DOCX_HEADER = re.compile(r"^word/header\d*\.xml$")
_DOCX_FOOTER = re.compile(r"^word/footer\d*\.xml$")
_PPT_SLIDE = re.compile(r"^ppt/(slideMasters|slideLayouts|slides)/[^/]+\.xml$")


def _blank_docx_part(xml: bytes) -> tuple[bytes, int]:
    """把一个 headerN.xml / footerN.xml 清成空壳：删掉所有段落、表格与内嵌图片。

    返回 (新 XML, 清掉的元素数)。保留命名空间声明，Word 打开不会提示修复。
    """
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return xml, 0

    removed = 0
    for child in list(root):
        if child.tag.rsplit("}", 1)[-1] in ("p", "tbl", "sdt"):
            root.remove(child)
            removed += 1
    # 段落已整体删除，这里再兜一层：去掉任何残留的图形 / 图片节点，免得留空白框
    for parent in root.iter():
        for child in list(parent):
            if child.tag.rsplit("}", 1)[-1] in ("drawing", "pict", "shape"):
                parent.remove(child)
                removed += 1
    if not removed:
        return xml, 0
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), removed


def _remove_docx(data: bytes, mode: str) -> WatermarkResult:
    """Word：strip 清页眉、footer 清页脚、veiler / auto 两处都清。"""
    try:
        source = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise WatermarkError(
            "文件已损坏或不是有效的 Office 文档；如为老版 .doc，请另存为 .docx 后重试。"
        ) from exc

    drop_header = mode in ("strip", "veiler", "auto")
    drop_footer = mode in ("footer", "veiler", "auto")

    buffer = io.BytesIO()
    removed = 0
    touched: list[str] = []
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            is_header = bool(_DOCX_HEADER.match(item.filename))
            is_footer = bool(_DOCX_FOOTER.match(item.filename))
            if (is_header and drop_header) or (is_footer and drop_footer):
                payload, count = _blank_docx_part(payload)
                if count:
                    removed += count
                    touched.append(item.filename)
            target.writestr(item, payload)
    source.close()

    if drop_header and not drop_footer:
        where = "页眉"
    elif drop_footer and not drop_header:
        where = "页脚"
    else:
        where = "页眉和页脚"
    message = (
        f"已清空{where}（{len(set(touched))} 个部件），正文一个字都没动。"
        if removed
        else f"这份文档本来就没有{where}内容，可以直接下载。"
    )
    return WatermarkResult(
        ok=True,
        message=message,
        output=buffer.getvalue(),
        mime=_mime_for(".docx"),
        removed=removed,
        log=[f"清空部件：{name}" for name in dict.fromkeys(touched)] or ["无页眉页脚部件"],
    )


# 页面顶部 / 底部占位符（PPT 的页眉 = 日期 + 标题，页脚 = 页脚 + 页码）
_PPT_HEADER_TYPES = ("dt", "sldNum")
_PPT_FOOTER_TYPES = ("ftr", "sldNum")


def _strip_pptx_placeholders(xml: bytes, drop: tuple[str, ...]) -> tuple[bytes, int]:
    """删掉 slide / master 里指定类型的占位符（sp / ph @type），返回 (新 XML, 数量)。"""
    try:
        tree = ET.fromstring(xml)
    except ET.ParseError:
        return xml, 0
    removed = 0
    for parent in tree.iter():
        for child in list(parent):
            if child.tag.rsplit("}", 1)[-1] != "sp":
                continue
            kinds = {node.get("type", "") for node in child.iter() if node.tag.rsplit("}", 1)[-1] == "ph"}
            if kinds & set(drop):
                parent.remove(child)
                removed += 1
    if not removed:
        return xml, 0
    return ET.tostring(tree, encoding="utf-8", xml_declaration=True), removed


def _remove_pptx(data: bytes, mode: str) -> WatermarkResult:
    try:
        source = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise WatermarkError(
            "文件已损坏或不是有效的 Office 文档；如为老版 .ppt，请另存为 .pptx 后重试。"
        ) from exc

    if mode == "strip":
        drop: tuple[str, ...] = _PPT_HEADER_TYPES
        where = "页眉（日期 / 标题占位符）"
    elif mode == "footer":
        drop = _PPT_FOOTER_TYPES
        where = "页脚（页脚 / 页码占位符）"
    else:                          # auto / veiler：两处都清
        drop = _PPT_HEADER_TYPES + _PPT_FOOTER_TYPES
        where = "页眉与页脚占位符"

    buffer = io.BytesIO()
    removed = 0
    touched: list[str] = []
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if _PPT_SLIDE.match(item.filename):
                payload, count = _strip_pptx_placeholders(payload, drop)
                if count:
                    removed += count
                    touched.append(item.filename)
            target.writestr(item, payload)
    source.close()

    message = (
        f"已删除 {removed} 处{where}，幻灯片内容保持不变。"
        if removed
        else f"这份 PPT 里没有找到{where}，可以直接下载。"
    )
    return WatermarkResult(
        ok=True,
        message=message,
        output=buffer.getvalue(),
        mime=_mime_for(".pptx"),
        removed=removed,
        log=[f"处理部件：{name}" for name in dict.fromkeys(touched)] or ["未命中占位符"],
    )


# --------------------------------------------------------------------------
# 统一入口
# --------------------------------------------------------------------------
SUPPORTED_SUFFIXES = (".png", ".jpg", ".jpeg", ".webp", ".pdf", ".docx", ".pptx")


def remove_watermark(
    data: bytes,
    filename: str,
    mode: str = "auto",
    *,
    region: str | None = None,
) -> WatermarkResult:
    """按扩展名分派处理；任何异常都转成中文提示的结果，不向界面抛堆栈。

    mode: REMOVE_MODES 的键，决定“处理哪个位置”，永远不需要用户填写水印文字。
        auto / strip / footer / veiler / crop_top / crop_bottom / crop_corner
    region: 仅为兼容旧调用方保留（过去传 "top"/"bottom"），新代码请直接用 mode。
    """
    suffix = _suffix(filename)
    if suffix not in SUPPORTED_SUFFIXES:
        return WatermarkResult(
            ok=False,
            message=f"暂不支持 {suffix or '该'} 格式，请上传 PNG / JPG / WEBP / PDF / DOCX / PPTX。",
        )
    if not data:
        return WatermarkResult(ok=False, message="文件内容为空，请重新上传。")

    resolved = mode_for_file(_norm_mode(mode, region), filename)
    try:
        if suffix == ".pdf":
            result = _remove_pdf(data, resolved)
        elif suffix == ".docx":
            result = _remove_docx(data, resolved)
        elif suffix == ".pptx":
            result = _remove_pptx(data, resolved)
        else:
            result = _remove_image(data, resolved, suffix)
    except WatermarkError as exc:
        return WatermarkResult(ok=False, message=str(exc))
    except Exception as exc:  # noqa: BLE001 - 兜底，避免整页崩掉
        return WatermarkResult(ok=False, message=f"处理失败：{exc}")

    result.filename = f"{_stem(filename)}{suffix}"
    result.log.insert(0, f"文件：{filename}（{len(data) / 1024:.1f}KB）")
    result.log.insert(1, f"处理位置：{MODE_LABELS.get(resolved, resolved)}")
    return result
def _remove_white_background_watermark(image: Image.Image) -> Image.Image:
    """专门针对白底文字/印章水印的精准去除算法（纯 numpy 实现，不依赖 scipy）
    
    规则逻辑：
    1. 白底检测：RGB 三通道均值 > 230 的像素视为背景
    2. 文字检测：在背景区域内，颜色较暗（RGB 均值 < 200）且面积较小的连通区域视为水印
    3. 精准擦除：只擦除符合水印特征的区域，保留正文和其他内容
    """
    import numpy as np
    from PIL import Image
    
    img_array = np.array(image)
    h, w = img_array.shape[:2]
    
    # 创建背景掩码：RGB 均值 > 230 的像素
    bg_mask = np.mean(img_array, axis=2) > 230

    # 计算亮度（加权平均）
    brightness = 0.299 * img_array[:, :, 0] + 0.587 * img_array[:, :, 1] + 0.114 * img_array[:, :, 2]
    
    # 暗色区域：在背景区域内且亮度 < 200
    dark_mask = bg_mask & (brightness < 200)
    
    # 快速形态学操作（使用滑动窗口视图）
    def _morph_op(mask, size=3, op='max'):
        """快速形态学操作（膨胀=max，腐蚀=min）"""
        pad = size // 2
        padded = np.pad(mask.astype(float), pad, mode='constant')
        result = np.zeros_like(mask, dtype=float)
        
        # 使用滑动窗口加速
        for di in range(size):
            for dj in range(size):
                window = padded[di:di+h, dj:dj+w]
                if op == 'max':
                    result = np.maximum(result, window)
                else:
                    result = np.minimum(result, window)
                    result[~mask] = 0  # 保持背景
        
        return (result > 0.5) if op == 'max' else (result > 0.9)
    
    # 开运算去噪，闭运算连接（简化版，只做一次膨胀和腐蚀）
    dark_mask = _morph_op(_morph_op(dark_mask, 3, 'min'), 3, 'max')
    
    # 简化版连通区域标记（使用并查集思想）
    def _simple_label(mask):
        labeled = np.zeros_like(mask, dtype=int)
        label = 0
        indices = np.argwhere(mask)
        
        if len(indices) == 0:
            return labeled, 0
        
        # 使用并查集合并相邻区域
        parent = {}
        def find(x):
            if x not in parent:
                parent[x] = x
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]
        
        def union(x, y):
            px, py = find(x), find(y)
            if px != py:
                parent[px] = py
        
        # 第一遍：标记并合并
        for idx, (i, j) in enumerate(indices):
            label = idx + 1
            labeled[i, j] = label
            parent[label] = label
            
            # 检查上方和左方邻居
            for di, dj in [(-1, 0), (0, -1)]:
                ni, nj = i + di, j + dj
                if 0 <= ni < h and 0 <= nj < w and labeled[ni, nj] > 0:
                    union(label, labeled[ni, nj])
        
        # 第二遍：压缩路径，重新编号
        root_map = {}
        new_label = 0
        for i in range(h):
            for j in range(w):
                if labeled[i, j] > 0:
                    root = find(labeled[i, j])
                    if root not in root_map:
                        new_label += 1
                        root_map[root] = new_label
                    labeled[i, j] = root_map[root]
        
        return labeled, new_label
    
    labeled, num_features = _simple_label(dark_mask)
    
    # 过滤并清除水印
    cleaned_array = img_array.copy()
    for label_id in range(1, min(num_features + 1, 1000)):  # 限制最大标签数防死循环
        region_mask = (labeled == label_id)
        area = np.sum(region_mask)
        
        if 50 <= area <= 2000:  # 面积过滤
            rows = np.where(np.any(region_mask, axis=1))[0]
            cols = np.where(np.any(region_mask, axis=0))[0]
            if len(rows) > 0 and len(cols) > 0:
                y1, y2 = rows[0], rows[-1] + 1
                x1, x2 = cols[0], cols[-1] + 1
                
                # 边缘区域检测（水印通常在边缘）
                if x1 < w * 0.2 or x2 > w * 0.8 or y1 < h * 0.2 or y2 > h * 0.8:
                    margin = 2
                    y_start, y_end = max(0, y1 - margin), min(h, y2 + margin)
                    x_start, x_end = max(0, x1 - margin), min(w, x2 + margin)
                    
                    bg_region = cleaned_array[y_start:y_end, x_start:x_end]
                    region_crop = region_mask[y_start:y_end, x_start:x_end]
                    bg_mask_local = ~region_crop
                    
                    if np.any(bg_mask_local):
                        bg_color = np.mean(bg_region[bg_mask_local], axis=0).astype(int)
                    else:
                        bg_color = [255, 255, 255]
                    
                    cleaned_array[y1:y2, x1:x2] = bg_color
    
    return Image.fromarray(cleaned_array.astype(np.uint8))


def _count_watermark_pixels(original: Image.Image, cleaned: Image.Image) -> int:
    """统计被去除的水印像素数量"""
    import numpy as np
    
    orig_array = np.array(original)
    clean_array = np.array(cleaned)
    
    diff = np.abs(orig_array.astype(int) - clean_array.astype(int))
    return int(np.sum(np.max(diff, axis=2) > 20))
