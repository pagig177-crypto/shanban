"""PDF 压缩：在不降低可读性/画质的前提下减小文件体积。

压缩策略（全部无损）：
    1. 对象流 + 交叉引用流（object/xref streams）：合并小对象，通常能省 5%~30%；
    2. 无损重编码：把未压缩的 Flate 图像重新压缩，并去掉元数据、孤立对象；
    3. 重复资源合并（clean / sanitize）；
    4. 可选：删除缩略图、未使用的外部链接目标等非正文数据。

默认不重采样任何图像、不改 DPI、不改分辨率，所以文字清晰度、图片画质、
页面内容都不变，阅读体验与原件一致；只砍掉冗余的存储开销。

目标大小模式：先走无损流程；仍达不到目标时，才按“最小必要”逐步降级
（仅对照片类图片降低 JPEG 质量，文字/矢量/页面尺寸始终不动），
每一步都重新校验，达到目标即停止，并在结果里说明用了哪一级。
"""

from __future__ import annotations

import contextlib
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

import pymupdf

ProgressCallback = Callable[[str], None]


@dataclass
class CompressResult:
    source: Path
    output: Path | None = None
    ok: bool = False
    original_size: int = 0
    final_size: int = 0
    level: str = "无损"
    target_bytes: int | None = None
    message: str = ""
    traceback_text: str = ""

    @property
    def saved_bytes(self) -> int:
        return max(0, self.original_size - self.final_size)

    @property
    def ratio(self) -> float:
        """压缩率，例如 0.25 表示体积减少 25%。"""
        if not self.original_size:
            return 0.0
        return self.saved_bytes / self.original_size

    @property
    def reached_target(self) -> bool:
        return self.target_bytes is not None and self.final_size <= self.target_bytes

    def describe(self) -> str:
        if not self.ok:
            return self.message
        text = (
            f"{human_size(self.original_size)} → {human_size(self.final_size)}"
            f"（减少 {self.ratio * 100:.1f}%，{self.level}）"
        )
        if self.target_bytes is not None:
            text += (
                f"，目标 {human_size(self.target_bytes)}："
                + ("已达成" if self.reached_target else "未能达成")
            )
        return text


def human_size(num: int | float) -> str:
    value = float(num)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{value:.1f} GB"


def _log(callback: ProgressCallback | None, message: str) -> None:
    if callback is not None:
        callback(message)


# 无损保存参数：garbage=4 回收所有孤立对象，clean 合并重复资源，
# deflate=True 压缩内容流，deflate_images 只重压“未被压缩”的图像数据。
_LOSSLESS_SAVE = dict(
    garbage=4,
    clean=True,
    deflate=True,
    deflate_images=True,
    deflate_fonts=True,
    use_objstms=1,
    expand=0,
)

# 重压图片时不能再用 deflate_images：已编码的 JPEG(DCTDecode) 数据被 Flate
# 二次压缩后阅读器无法解码，必须让图像流保持原样写入。
_JPEG_SAVE_OVERRIDES = dict(deflate_images=False)
# 达到目标时才逐级启用的降级（仅影响照片类位图，不影响文字/页面尺寸）
_FALLBACK_STEPS: list[tuple[str, dict]] = [
    ("高质量重压（照片图片 JPEG 92）", {"drop_thumbnails": True, "jpeg_quality": 92}),
    ("中高质量重压（照片图片 JPEG 85）", {"jpeg_quality": 85}),
    ("中等重压（照片图片 JPEG 78）", {"jpeg_quality": 78}),
    ("较高压缩（照片图片 JPEG 70）", {"jpeg_quality": 70}),
]


def parse_target_size(text: str | None) -> int | None:
    """把 “2MB”“1.5m”“800kb”“900000” 解析成字节数；空值返回 None。"""
    if text is None:
        return None
    value = str(text).strip().lower().replace(" ", "")
    if not value:
        return None
    # 注意：必须按后缀长度从长到短匹配，否则 "1.5mb" 会先命中 "b"
    units = (
        ("gb", 1024 ** 3), ("mb", 1024 ** 2), ("kb", 1024), ("b", 1),
        ("g", 1024 ** 3), ("m", 1024 ** 2), ("k", 1024),
    )
    number = value
    factor = 1
    for suffix, unit_factor in units:
        if value.endswith(suffix) and len(value) > len(suffix):
            number, factor = value[: -len(suffix)], unit_factor
            break
    try:
        return max(1, int(float(number) * factor))
    except ValueError as exc:
        raise ValueError(f"无法识别的大小：{text}") from exc


# --------------------------------------------------------------------------
# 内部步骤
# --------------------------------------------------------------------------
def _save(doc: pymupdf.Document, dst: Path, *, lossless: bool = True, **extra) -> int:
    """保存到目标路径并返回字节数（失败会尝试退回较保守的参数）。"""
    options = {**_LOSSLESS_SAVE, **extra}
    if not lossless:
        # 续压已优化过的文件：不再做结构级重排，只写回改动
        options.pop("use_objstms", None)
        options.pop("clean", None)
    try:
        doc.save(str(dst), **options)
    except (RuntimeError, ValueError):
        # 某些文档不支持对象流（例如已加密/线性化结构），退回基础参数
        options.pop("use_objstms", None)
        options.pop("clean", None)
        doc.save(str(dst), **options)
    return dst.stat().st_size


def _drop_thumbnails(doc: pymupdf.Document) -> None:
    """删除页面缩略图（阅读器会自行生成），不改变正文。"""
    for page in doc:
        try:
            doc.xref_set_key(page.xref, "Thumb", "null")
        except Exception:  # noqa: BLE001
            continue


def _photo_images(doc: pymupdf.Document, min_pixels: int = 200_000) -> list[int]:
    """找出“照片类”位图 xref：尺寸较大且不是二值/线条图（不做蒙版处理）。"""
    xrefs: set[int] = set()
    for page in doc:
        for info in page.get_images(full=True):
            xref = info[0]
            try:
                width, height = info[2], info[3]
            except IndexError:
                continue
            if width * height >= min_pixels:
                xrefs.add(xref)
    return sorted(xrefs)


def _recompress_photos(doc: pymupdf.Document, quality: int) -> int:
    """按指定 JPEG 质量重压照片图片，返回处理的图片数。

    只处理原本就是 JPEG 的大图（照片），保留原分辨率与宽高比；
    PNG/Smask/图标等一律不动，避免出现锯齿或透明背景变黑。
    """
    processed = 0
    for xref in _photo_images(doc):
        try:
            raw = doc.xref_stream_raw(xref)
            if not raw or not raw.startswith(b"\xff\xd8"):
                continue  # 只处理 JPEG，天然“照片”
            pix = pymupdf.Pixmap(doc, xref)
            if pix.n - pix.alpha >= 4:  # CMYK 转 RGB，避免阅读器偏色
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            data = pix.tobytes("jpeg", jpg_quality=quality)
            if len(data) < len(raw):
                doc.update_stream(xref, data, compress=False)
                processed += 1
        except Exception:  # noqa: BLE001 - 单张图片失败不影响整份文件
            continue
    return processed


def _compress_once(
    src: Path,
    dst: Path,
    *,
    drop_thumbnails: bool = False,
    jpeg_quality: int | None = None,
    lossless: bool = True,
) -> tuple[int, int]:
    """执行一次完整压缩，返回 (输出字节数, 重压图片数)。

    lossless=False 时省略对象流/clean 等结构优化，仅用于“续压”上一级已优化
    的结果，避免重复处理。
    """
    with pymupdf.open(str(src)) as doc:
        if doc.needs_pass:
            raise RuntimeError("PDF 已加密，请先解密后再压缩")
        if drop_thumbnails:
            _drop_thumbnails(doc)
        touched = 0
        if jpeg_quality is not None:
            touched = _recompress_photos(doc, jpeg_quality)
        extra = _JPEG_SAVE_OVERRIDES if jpeg_quality is not None else {}
        size = _save(doc, dst, lossless=lossless, **extra)
    return size, touched


# --------------------------------------------------------------------------
# 对外 API
# --------------------------------------------------------------------------
def compress_pdf(
    src: str | Path,
    out_dir: str | Path | None = None,
    *,
    target_size: int | str | None = None,
    overwrite: bool = False,
    callback: ProgressCallback | None = None,
    suffix: str = "-压缩",
    allow_quality_fallback: bool = True,
) -> CompressResult:
    """无损压缩单个 PDF；给定目标大小时尽量达到。

    Args:
        target_size: 目标字节数，或 "2MB"/"800kb" 这样的字符串；None 表示尽量小。
        allow_quality_fallback: 无损压缩达不到目标时，是否允许对照片图片降级。
        suffix: 输出文件名后缀，默认 "原名-压缩.pdf"。

    Returns:
        CompressResult，包含前后体积、使用的级别、是否达成目标。
    """
    import traceback

    src_path = Path(src).expanduser().resolve()
    result = CompressResult(source=src_path, target_bytes=None)
    if isinstance(target_size, str):
        try:
            target_size = parse_target_size(target_size)
        except ValueError as exc:
            result.message = str(exc)
            return result
    result.target_bytes = target_size

    if not src_path.is_file():
        result.message = "文件不存在"
        return result

    result.original_size = src_path.stat().st_size
    target_dir = Path(out_dir).expanduser().resolve() if out_dir else src_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    dst = target_dir / f"{src_path.stem}{suffix}.pdf"
    if dst.exists():
        if overwrite:
            dst.unlink()
        else:
            dst = _unique_path(target_dir, f"{src_path.stem}{suffix}")
    result.output = dst

    _log(callback, f"压缩：{src_path.name}（原始 {human_size(result.original_size)}）")
    try:
        with tempfile.TemporaryDirectory(prefix="pdf2word-compress-") as tmp:
            step_path = Path(tmp) / "step.pdf"
            # 第 1 级：无损
            _log(callback, "  [无损] 合并对象流、清理冗余数据…")
            size, touched = _compress_once(src_path, step_path)
            best_source = step_path if size < result.original_size else src_path
            best_size = min(size, result.original_size)
            level = "无损"
            if touched:
                level = "无损（含无损图片重压）"
            _log(callback, f"  → {human_size(best_size)}")

            # 第 2..n 级：仅在设置了目标且尚未达成时逐级降级
            if (
                allow_quality_fallback
                and target_size is not None
                and best_size > target_size
            ):
                for label, options in _FALLBACK_STEPS:
                    _log(callback, f"  [{label}] 仍未达标，继续压缩…")
                    next_path = Path(tmp) / f"step-{options.get('jpeg_quality', 0)}.pdf"
                    # 在上一级结果的基础上继续压，避免重复劳动
                    size, touched = _compress_once(
                        best_source, next_path, lossless=False, **options
                    )
                    if size < best_size:
                        best_size, best_source, level = size, next_path, label
                    _log(callback, f"  → {human_size(best_size)}")
                    if best_size <= target_size:
                        break

            if best_source != src_path:
                shutil.copyfile(best_source, dst)
            else:
                # 压不动（原文已是优化过的），直接复制，保证输出文件存在
                shutil.copyfile(src_path, dst)
                best_size = result.original_size
                level = "文件已是优化状态，无进一步压缩空间"
    except Exception as exc:  # noqa: BLE001
        result.traceback_text = traceback.format_exc()
        result.message = f"压缩失败：{exc}"
        _log(callback, f"  ✘ {result.message}")
        return result

    result.final_size = dst.stat().st_size if dst.exists() else best_size
    result.ok = True
    result.level = level
    result.message = result.describe()
    _log(callback, f"  ✔ {dst.name}：{result.message}")
    return result


def compress_many(
    sources: Iterable[str | Path],
    out_dir: str | Path | None = None,
    **kwargs,
) -> list[CompressResult]:
    """批量无损压缩，参数同 compress_pdf（target_size 对每个文件生效）。"""
    return [compress_pdf(src, out_dir=out_dir, **kwargs) for src in sources]


def estimate_lossless_ratio(path: str | Path) -> float:
    """快速估算无损压缩能省多少（用于界面提示，不做实际输出）。"""
    src = Path(path)
    if not src.is_file():
        return 0.0
    with tempfile.TemporaryDirectory(prefix="pdf2word-estimate-") as tmp:
        dst = Path(tmp) / "probe.pdf"
        try:
            size, _ = _compress_once(src, dst)
        except Exception:  # noqa: BLE001
            return 0.0
    return max(0.0, 1 - size / src.stat().st_size)


def _unique_path(directory: Path, stem: str, suffix: str = ".pdf") -> Path:
    candidate = directory / f"{stem}{suffix}"
    counter = 1
    while candidate.exists():
        candidate = directory / f"{stem}({counter}){suffix}"
        counter += 1
    return candidate
