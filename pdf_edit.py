"""PDF 文字替换核心：在不改变版式、并保持文件大小一致的前提下改文字。

适用场景：报告里的姓名、证件号、日期等少量文字需要改动，但希望
导出结果和原文件"看起来没差"（字宽不变、页数不变、文件大小不变）。

做法（而不是把 PDF 转成 Word 再导出）：
    直接改第 N 页的文字内容流。绝大多数中文 PDF 使用 Type0 字体 +
    UCS2 / Identity 编码，每个字形固定 2 字节，等长的文字替换
    （例如 "智宇" → "诗雅"）不改变任何字节数，
    字宽、断行、xref 偏移和文件大小都保持原样。

不适用：需要增删文字（前后长度不同）、改字体、加密 PDF、纯图片扫描件。
"""

from __future__ import annotations

import re
import subprocess
import sys
import zlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

try:  # PyMuPDF：优先新模块名，兼容旧别名
    import pymupdf as fitz  # type: ignore
except ImportError:  # pragma: no cover - 旧版本
    import fitz  # type: ignore

ProgressCallback = Callable[[str], None]

# 内容流对象头：<</Length 316/Filter/FlateDecode>>stream 或 ...>>\nstream（两种都常见）
# 注意 stream 前后可能有 0~2 个换行，用 [\r\n]* 兼容，避免把 stream 吃进 dict。
_OBJ_RE = re.compile(rb"\n(\d+) 0 obj\s*<<([^<>]*?)>>[\r\n]*stream[\r\n]")
_NUM_RE = re.compile(rb"\d+")
_HEX_RUN = re.compile(rb"<([0-9A-Fa-f\s]+)>")

# 这些 CMap 都按 2 字节/字符编码，可按"每字 2 字节"匹配
_TWO_BYTE_CMAPS = ("ucs2", "utf16", "identity", "gbk-euc", "gb-euc", "gbkx2")


class PdfEditError(RuntimeError):
    """无法安全完成替换（编码不支持、未找到、长度不一致等）。"""


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------
@dataclass
class Match:
    """一处命中的文字。"""

    page: int            # 1 起
    text: str
    rect: tuple[float, float, float, float]
    font: str
    size: float


@dataclass
class Replacement:
    old: str
    new: str
    count: int = 0

    @property
    def same_length(self) -> bool:
        return len(self.old) == len(self.new)


@dataclass
class EditResult:
    source: Path
    output: Path | None = None
    ok: bool = False
    message: str = ""
    replacements: list[Replacement] = field(default_factory=list)
    source_size: int = 0
    output_size: int = 0
    page_count: int = 0
    padded: int = 0

    @property
    def size_delta(self) -> int:
        return self.output_size - self.source_size

    def describe(self) -> str:
        if not self.ok:
            return self.message
        changed = "、".join(f"{r.old}→{r.new}×{r.count}" for r in self.replacements if r.count)
        size = f"{self.source_size} 字节"
        if self.output_size != self.source_size:
            size = f"{self.source_size} → {self.output_size} 字节（{self.size_delta:+d}）"
        return f"已替换 {changed}；共 {self.page_count} 页；文件大小 {size}"


# --------------------------------------------------------------------------
# 查询
# --------------------------------------------------------------------------
def _page_indexes(page_count: int, pages: Sequence[int] | None) -> list[int]:
    """把 1 起的页码换成 0 起的下标并校验范围。"""
    if not pages:
        return list(range(page_count))
    indexes = []
    for page in pages:
        if not 1 <= page <= page_count:
            raise PdfEditError(f"页码超出范围：{page}（共 {page_count} 页）")
        indexes.append(page - 1)
    return sorted(set(indexes))


def find_matches(pdf: Path, needle: str, pages: Sequence[int] | None = None) -> list[Match]:
    """在 PDF 中查找文字（按页），返回每处命中的位置与字体信息。

    先按 span 匹配，找不到时再按整行合并文本匹配——很多 PDF 会把一句话
    拆成多个字体/字号的 span，逐 span 查会漏掉。
    """
    matches: list[Match] = []
    with fitz.open(str(Path(pdf).expanduser())) as doc:
        for index in _page_indexes(doc.page_count, pages):
            for block in doc[index].get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not needle:
                        continue
                    hit = [s for s in spans if needle in s.get("text", "")]
                    if not hit:
                        joined = "".join(s.get("text", "") for s in spans)
                        if needle not in joined:
                            continue
                        hit = spans  # 整行命中：返回该行所有 span
                    for span in hit:
                        matches.append(Match(
                            page=index + 1,
                            text=span.get("text", ""),
                            rect=tuple(round(v, 2) for v in span["bbox"]),  # type: ignore[arg-type]
                            font=span.get("font", ""),
                            size=round(float(span.get("size", 0)), 2),
                        ))
    return matches


def font_encodings(pdf: Path, page_no: int = 1) -> list[tuple[str, bool]]:
    """返回某页字体，以及它是否"每字符固定 2 字节"（决定能否等长替换）。"""
    result: list[tuple[str, bool]] = []
    with fitz.open(str(Path(pdf).expanduser())) as doc:
        if not 1 <= page_no <= doc.page_count:
            raise PdfEditError(f"页码超出范围：{page_no}（共 {doc.page_count} 页）")
        for info in doc[page_no - 1].get_fonts(full=True):
            _, _, kind, name, _, encoding, _ = info
            enc = (encoding or "").lower()
            two_byte = kind == "Type0" and (not enc or any(t in enc for t in _TWO_BYTE_CMAPS))
            result.append((f"{name} [{encoding or '未标注编码'}]", two_byte))
    return result


def inspect_pdf(pdf: Path) -> str:
    """生成一份便于判断"能不能改字"的报告。"""
    pdf = Path(pdf).expanduser()
    lines = [f"文件：{pdf}", f"大小：{pdf.stat().st_size} 字节"]
    with fitz.open(str(pdf)) as doc:
        lines.append(f"页数：{doc.page_count}")
        if doc.needs_pass:
            lines.append("加密：是（本工具不支持改加密 PDF）")
        for page_no in range(1, doc.page_count + 1):
            lines.append(f"\n第 {page_no} 页字体：")
            for name, two_byte in font_encodings(pdf, page_no):
                tag = "可等长替换" if two_byte else "非 2 字节编码，无法等长替换"
                lines.append(f"  - {name}  → {tag}")
            if not doc[page_no - 1].get_text().strip():
                lines.append("  （本页没有文字层：可能是扫描/图片页）")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# 编辑
# --------------------------------------------------------------------------
def _encodings(old: str, new: str) -> list[tuple[bytes, bytes, str, bool]]:
    """列出可能的 (旧字节, 新字节, 名称, 是否 2 字节/字) 编码组合。

    同一段文字在不同 PDF 里可能是：明文 ASCII（WinAnsi）、UTF-16BE（UCS2/Identity）、
    GBK 双字节。按长度不变的组合依次尝试。
    """
    combos = []
    for name, encoding, two_byte in (
        ("ASCII", "latin-1", False),
        ("GBK", "gbk", False),
        ("UCS2", "utf-16-be", True),
    ):
        try:
            old_b, new_b = old.encode(encoding), new.encode(encoding)
        except UnicodeEncodeError:
            continue
        combos.append((old_b, new_b, name, two_byte))
    return combos


def _count_hits(content: bytes, old: str) -> int:
    """内容流里某段文字出现了几次（明文与十六进制两种写法取较大值）。"""
    best = 0
    for old_b, _new_b, _name, _two in _encodings(old, old):
        best = max(best, content.count(old_b))
        for case in ("upper", "lower"):
            hex_old = old_b.hex().encode()
            if case == "upper":
                hex_old = hex_old.upper()
            best = max(best, content.count(hex_old))
    return best


def _replace_nth(content: bytes, old_b: bytes, new_b: bytes, keep_last: int) -> tuple[bytes, int]:
    """只替换前 (总处数 - keep_last) 处，保留最后 keep_last 处不动。

    用途：同一内容流里有 N 处命中时，逐次只改一处——第 k 次处理时把「后面还没
    处理的 N-k 处」保留下来，就能从前往后依次替换。
    替换前后字节长度相同，因此从后往前拼接时偏移始终有效。
    """
    positions: list[int] = []
    start = 0
    while True:
        found = content.find(old_b, start)
        if found < 0:
            break
        positions.append(found)
        start = found + len(old_b)
    chosen = positions[: max(0, len(positions) - keep_last)]
    if not chosen:
        return content, 0

    pieces: list[bytes] = []
    cursor = len(content)
    for position in reversed(chosen):
        pieces.append(content[position + len(old_b): cursor])
        pieces.append(new_b)
        cursor = position
    pieces.append(content[:cursor])
    return b"".join(reversed(pieces)), len(chosen)


def _replace_bytes(
    content: bytes,
    old: str,
    new: str,
    keep_last: int = 0,
) -> tuple[bytes, int, str]:
    """在内容流中替换文字（明文与十六进制字符串两种写法都支持）。

    返回 (新内容, 替换次数, 编码名称)。找不到时返回原 content 与 0 次。
    keep_last: 只替换「前 N 处」，也就是保留最后 keep_last 处不动（0 = 全部替换）。
    """
    def apply(candidate_old: bytes, candidate_new: bytes) -> tuple[bytes, int]:
        if keep_last <= 0:
            hits = content.count(candidate_old)
            return (content.replace(candidate_old, candidate_new), hits) if hits else (content, 0)
        return _replace_nth(content, candidate_old, candidate_new, keep_last)

    for old_b, new_b, name, _ in _encodings(old, new):
        if len(old_b) != len(new_b):
            continue  # 编码后长度不同会破坏排版，跳过
        # 1) 明文：(文字) 或未压缩流中的裸字节
        replaced, hits = apply(old_b, new_b)
        if hits:
            return replaced, hits, name
        # 2) 十六进制：<5265706f7274>（区分大小写各试一次）
        for case in ("upper", "lower"):
            hex_old = old_b.hex().encode()
            hex_new = new_b.hex().encode()
            if case == "upper":
                hex_old, hex_new = hex_old.upper(), hex_new.upper()
            replaced, hits = apply(hex_old, hex_new)
            if hits:
                return replaced, hits, name
    return content, 0, ""


def _content_streams(raw: bytes, pdf: Path, pages: Sequence[int] | None) -> list[dict[str, int | bool]]:
    """找出目标页的内容流：对象号、/Length 偏移、数据区偏移、是否 Flate 压缩。"""
    with fitz.open(str(pdf)) as doc:
        xrefs: set[int] = set()
        for index in _page_indexes(doc.page_count, pages):
            xrefs.update(doc[index].get_contents())
    found = []
    for match in _OBJ_RE.finditer(raw):
        if int(match.group(1)) not in xrefs:
            continue
        header = match.group(2)
        length_match = re.search(rb"/Length\s+(\d+)", header)
        if not length_match:
            continue
        dict_start = match.start(2)
        found.append({
            "length_pos": dict_start + length_match.start(1),
            "length_end": dict_start + length_match.end(1),
            "stream_start": match.end(),
            "len": int(length_match.group(1)),
            "flate": b"/Flate" in header,
        })
    return found


def _patch_streams(
    raw: bytes,
    pdf: Path,
    replacements: Iterable[Replacement],
    pages: Sequence[int] | None,
    log: ProgressCallback | None,
    keep_last: int = 0,
) -> bytes:
    """就地改写内容流，返回新的文件字节。

    keep_last: 每个内容流里保留最后 keep_last 处不替换（0 = 全部替换）。
    """
    streams = _content_streams(raw, pdf, pages)
    if not streams:
        raise PdfEditError("没有找到可编辑的文字内容流（可能是扫描件或结构特殊）")

    # keep_last 是「整页范围内」保留最后多少处不替换：内容流是逆序处理的，
    # 遇到的第一批流就是视觉上最靠后的内容，因此按流依次消耗这个额度。
    remaining_keep = max(0, keep_last)
    pending: list[tuple[int, bytes]] = []  # (stream_start, 新内容)，稍后统一写回

    # 从后往前处理：改动会让后面对象的偏移变化，逆序可保证未处理的偏移仍然有效
    for stream in reversed(streams):
        length_pos = int(stream["length_pos"])
        length_end = int(stream["length_end"])
        stream_start = int(stream["stream_start"])
        old_len = int(stream["len"])
        old_stream = raw[stream_start:stream_start + old_len]
        try:
            content = _decompress(old_stream) if stream["flate"] else old_stream
        except zlib.error as exc:
            raise PdfEditError(f"内容流解压失败（偏移 {stream_start}）：{exc}") from exc

        new_content = content
        for repl in replacements:
            keep_here = remaining_keep
            if keep_here:
                # 该流里这组规则的命中数；不足额度时整段跳过，剩余额度留给前面的流
                hits_total = _count_hits(content, repl.old)
                if hits_total <= keep_here:
                    remaining_keep -= hits_total
                    continue
            new_content, hits, _name = _replace_bytes(
                new_content, repl.old, repl.new, keep_last=keep_here
            )
            if hits:
                remaining_keep = 0
                repl.count += hits
                if log:
                    log(f"  内容流：{repl.old} → {repl.new}（{hits} 处）")
        if new_content == content:
            continue

        if stream["flate"]:
            packed = _compress_close(new_content, old_len, log)
        elif len(new_content) <= old_len:
            packed = new_content
        else:
            # 未压缩流替换后变长：改为 Flate 压缩，避免整体变大导致无法对齐
            packed = _compress_close(new_content, old_len, log)
            raw = (raw[:length_pos] + str(len(packed)).encode()
                   + b"/Filter/FlateDecode" + raw[length_end:stream_start]
                   + packed + raw[stream_start + old_len:])
            continue
        raw = (raw[:length_pos] + str(len(packed)).encode()
               + raw[length_end:stream_start] + packed
               + raw[stream_start + old_len:])
    return raw


def _decompress(stream: bytes) -> bytes:
    """解压内容流（/FlateDecode 即 zlib 格式）。"""
    return zlib.decompress(stream)


def _compress_close(content: bytes, old_len: int, log: ProgressCallback | None) -> bytes:
    """压缩内容流，尽量让压缩后的长度不超过原长度。

    只替换了几个字节，压缩后长度也可能变化（原来的字典匹配被打断）。
    这里尝试所有压缩级别，取不超过原长度且最接近原长度的结果，
    再用空 stored 块把长度补到与原流一致；若所有级别都压不小，
    则如实返回，由文件级补齐统一处理（调用方会报告大小差异）。
    优先尝试 level 6：多数生成器（含 zlib 默认值）用它，压缩头字节也和原件一致。
    """
    levels = [6] + [lv for lv in range(9, -1, -1) if lv != 6]
    candidates = [zlib.compress(content, level) for level in levels]
    fits = [c for c in candidates if len(c) <= old_len]
    if not fits:
        # 还差一点时，尝试去掉内容流末尾的换行（endstream 之前的空白，语义无关）
        trimmed = content.rstrip(b"\r\n")
        if trimmed != content:
            trimmed_candidates = [zlib.compress(trimmed, level) for level in range(9, -1, -1)]
            fits = [c for c in trimmed_candidates if len(c) <= old_len]
            if fits:
                candidates = trimmed_candidates
    if not fits:
        return min(candidates, key=len)
    best = max(fits, key=len)
    pad = old_len - len(best)
    if pad:
        best = _pad_deflate_stream(best, pad)
        if log:
            log(f"  内容流压缩：补 {pad} 字节，保持流长度 {old_len}")
    return best


def _pad_deflate_stream(stream: bytes, pad: int) -> bytes:
    """把 zlib 流补长 pad 字节（解压结果不变）。

    在 adler32 校验之后追加完整的空 stored 块（00 00 00 FF FF）。
    解压器读完第一个 zlib 流就停止，后续字节被忽略；必须整块追加，
    否则校验码位置错乱，zlib 会报 "incorrect data check"。
    """
    whole, rest = divmod(pad, 5)
    return stream + b"\x00\x00\x00\xff\xff" * whole + b"\x00" * rest



def pad_to_size(raw: bytes, target_size: int) -> tuple[bytes, int]:
    """在 %%EOF 之后补空白，把字节数补齐到 target_size。

    PDF 规范允许 %%EOF 之后存在任意字节（阅读器按 startxref 回跳），
    因此补空格不会影响页面内容。输出比目标大时抛错。
    """
    delta = target_size - len(raw)
    if delta < 0:
        extra = -delta
        if extra <= _SMALL_DELTA:
            shrunk = _shrink_trailing(raw, extra)
            if shrunk is not None and len(shrunk) == target_size:
                return shrunk, 0
        raise PdfEditError(
            f"替换后文件比目标大 {extra} 字节，无法靠补齐对齐（请取消大小对齐或调整目标值）"
        )
    if delta == 0:
        return raw, 0
    return _append_trailing(raw, b" " * delta), delta


_SMALL_DELTA = 16  # 可以通过删除文件尾部空白抵消的字节数上限


def _shrink_trailing(raw: bytes, extra: int) -> bytes | None:
    """从文件尾部空白里精确扣掉 extra 字节。

    依次尝试两个位置（都是不影响 xref 偏移的纯空白）：
      1. %%EOF 之后（例如 %%EOF 后的换行）；
      2. xref 表与 startxref 之间的填充。
    扣不够则返回 None，由调用方报错。
    """
    eof = raw.rfind(b"%%EOF")
    if eof >= 0:
        eof_end = eof + 5
        tail = raw[eof_end:]
        if len(tail) >= extra and not tail[:extra].strip(b" \r\n\t"):
            return raw[:eof_end] + tail[extra:]

    # 退一步：在 "startxref" 前面 / xref 表后面的空白里扣
    marker = raw.rfind(b"startxref")
    if marker > 0:
        head = raw[:marker]
        stripped = head.rstrip(b" \r\n\t")
        available = len(head) - len(stripped)
        if available >= extra:
            return stripped + head[len(head) - available + extra:] + raw[marker:]
    return None


def _append_trailing(raw: bytes, extra: bytes) -> bytes:
    """把额外字节追加到文件末尾（PDF 注释区，不影响解析）。

    若文件带增量更新（多个 %%EOF），追加在最后一个 %%EOF 之后最安全。
    """
    eof = raw.rfind(b"%%EOF")
    return (raw + extra) if eof < 0 else (raw[:eof + 5] + extra + raw[eof + 5:])



def replace_text(
    src: Path,
    replacements: Sequence[tuple[str, str]],
    out: Path | None = None,
    pages: Sequence[int] | None = None,
    match_size: bool = True,
    target_size: int | None = None,
    overwrite: bool = False,
    log: ProgressCallback | None = None,
    keep_last: int = 0,
) -> EditResult:
    """在 src 中替换文字并写出新 PDF。

    replacements: [(旧文字, 新文字), ...]；每项必须等长，否则会改变排版，直接拒绝。
    pages:        只改这些页（1 起）；None 表示全部页。
    match_size:   True 时把输出大小对齐到 target_size（默认 = 源文件大小）。
    keep_last:    每个内容流保留最后 keep_last 处不替换（0 = 全部替换）；
                  用于“只改一处”，由调用方逐页推进。
    """
    src = Path(src).expanduser()
    result = EditResult(source=src)
    if not src.exists():
        result.message = f"文件不存在：{src}"
        return result
    result.source_size = src.stat().st_size

    repls = [Replacement(old=old, new=new) for old, new in replacements]
    if not repls:
        result.message = "没有提供任何替换规则"
        return result
    unbalanced = [r for r in repls if not r.same_length]
    if unbalanced:
        names = "、".join(f"{r.old}→{r.new}（{len(r.old)}→{len(r.new)} 字）" for r in unbalanced)
        result.message = f"以下替换前后长度不同，会改变排版，已拒绝：{names}"
        return result

    out = Path(out).expanduser() if out else src.with_name(f"{src.stem}_已修改.pdf")
    if out.resolve() == src.resolve():
        result.message = "输出文件不能与源文件相同"
        return result
    if out.exists() and not overwrite:
        result.message = f"输出已存在：{out}（勾选“覆盖”可替换）"
        return result

    try:
        raw = src.read_bytes()
        with fitz.open(str(src)) as doc:
            result.page_count = doc.page_count
            if doc.needs_pass:
                result.message = "该 PDF 已加密，暂不支持修改"
                return result
        if log:
            log(f"正在处理 {src.name}（{result.page_count} 页）")

        patched = _patch_streams(raw, src, repls, pages, log, keep_last=keep_last)
        if not any(r.count for r in repls):
            result.message = "没有在 PDF 中找到需要替换的文字"
            return result

        # 字宽校正：新字符的 CID 若不在字体的 /W 里，会被按 /DW（全角）排版，
        # 整行往右多出半格——手机上看起来就像多了一个空格。这里补上显式宽度。
        patched = _fix_glyph_widths(patched, src, repls)

        if match_size:
            target = target_size if target_size is not None else result.source_size
            patched, result.padded = pad_to_size(patched, target)
            if log and result.padded:
                log(f"  补齐 {result.padded} 字节，使大小与目标一致")

        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(patched)

        # 关键校验：PDF 里的字形编码不一定等于 Unicode 码位（例如 CID 字体按 GB 码
        # 存储）。字节层面替换成功 ≠ 读出来是期望的文字，因此必须回读确认。
        problem = _check_text(out, pages, repls)
        if problem:
            out.unlink(missing_ok=True)
            result.message = (
                f"{problem}\n"
                "该 PDF 的字形编码与 Unicode 不一致（例如 CID 字体按 GB 码存储），"
                "直接替换字节会得到错误字形；这类文件需要按嵌入字体的 CMap 映射改写。"
            )
            return result
    except PdfEditError as exc:
        result.message = str(exc)
        return result
    except OSError as exc:
        result.message = f"读写失败：{exc}"
        return result

    result.output = out
    result.output_size = out.stat().st_size
    result.replacements = repls
    result.ok = True
    result.message = result.describe()
    return result


def _fix_glyph_widths(raw: bytes, pdf: Path, repls: Sequence["Replacement"]) -> bytes:
    """补齐 CID 字体的 /W 宽度表，避免替换后字间距变化。

    很多中文 PDF 用未嵌入的 CID 字体（如 STSong-Light + UniGB-UCS2-H）。这类字体
    的 /W 常常只列出"用到的那几个 CID"，没列的按 /DW（通常 1000 = 全角）算。
    于是原件里只出现过 0/1/2/3/4/9 时，把它们改成 7，7 若没有宽度记录就会按全角
    排，整行往右多出半格，在手机上看着就像 "23 7"。

    修复：给缺失的 CID 补一条显式宽度（参考同字体里同类字符的已有宽度）。
    只改字体字典、不动内容流，文字/页数/大小对齐逻辑都不受影响。

    raw 是**已经替换过文字**的文件字节；字体结构与原件一致，因此打开 raw 本身即可。
    """
    if not repls:
        return raw
    import tempfile

    try:
        with tempfile.TemporaryDirectory() as tmp:
            probe = Path(tmp) / "patched.pdf"
            probe.write_bytes(raw)
            with fitz.open(str(probe)) as doc:
                plans = _cid_width_targets(doc, repls)
                if not plans:
                    return raw
                changed = False
                for xref, widths in plans.items():
                    if _write_cid_widths(doc, xref, widths):
                        changed = True
                if not changed:
                    return raw
                return doc.tobytes(garbage=0, deflate=True)
    except Exception:  # noqa: BLE001 - 补宽度失败不应影响替换本身
        return raw
    """校正因字形宽度不同导致的字间距变化（未嵌入 CID 字体的常见坑）。

    背景：像 STSong-Light 这类**未嵌入**的 Type0 字体通常没有 /W 宽度表，
    所有字都按 /DW（多半是 1000，即全角）计宽，真实宽度靠阅读器查系统字体。
    于是同一个数字既有半角（约 462）也有全角（1000）字形——把原来半角的
    "9" 换成全角宽度的 "7"，这一行就会多出半个字宽，屏幕上看着像多了个空格。

    做法（不引入 TJ，因为部分阅读器对内嵌为空的 CID 字体处理 TJ 数字串不稳）：
    把「新字符」单独拆成一个 Tj，画完后用 Td 把光标的 x 拉回到正确位置，
    这样后续文字一律从原坐标开始，行宽与原件完全一致。
    """
    old_b, new_b = _encoding_pair(repl.old, repl.new)
    if not old_b or not new_b or old_b == new_b:
        return content
    delta = _advance_delta(pdf, repl.old, repl.new)
    if delta is None or abs(delta) < 1:
        return content  # 宽度一致，什么都不用做
    return _split_last_glyph(content, new_b, delta)


def _encoding_pair(old: str, new: str) -> tuple[bytes, bytes]:
    """取等长的编码字节对；优先 2 字节/字（UCS-2），中文 PDF 正文多为这种。"""
    combos = [c for c in _encodings(old, new) if len(c[0]) == len(c[1])]
    two = [c for c in combos if c[3]]
    for old_b, new_b, _name, _two in two or combos:
        return old_b, new_b
    return b"", b""


def _cid_width_targets(doc: "fitz.Document", repls: Sequence["Replacement"]) -> dict[int, dict[int, int]]:
    """算出每个 CIDFont 需要补的 {CID: 宽度}。"""
    plans: dict[int, dict[int, int]] = {}
    for repl in repls:
        if not repl.count:
            continue
        old_b, new_b = _encoding_pair(repl.old, repl.new)
        if not old_b or not new_b or len(old_b) != len(new_b):
            continue
        step = 2 if len(old_b) % 2 == 0 else 1
        for page in doc:
            for info in page.get_fonts(full=True):
                kind, encoding = info[2], info[5]
                if kind != "Type0" or not _is_two_byte(encoding):
                    continue
                cidfont = _cidfont_xref(doc, info[0])
                if cidfont is None:
                    continue
                current = _parse_w(doc, cidfont)
                cids = _lookup_cids(doc, cidfont, current, new_b, step, repl.new)
                if cids and not any(c in current for c in cids):
                    # 一个都命中不了说明推断不可靠（例如汉字 CID 无从校验），宁可不改
                    continue
                missing = [c for c in cids if c not in current]
                if not missing:
                    continue
                ref = _reference_width(current, _parse_dw(doc, cidfont), repl.new)
                if ref is None:
                    continue
                plan = plans.setdefault(cidfont, {})
                for cid in missing:
                    plan[cid] = ref
    return plans


def _is_two_byte(encoding: str) -> bool:
    """该 CMap 是否 2 字节/字符（决定 CID 怎么从字节里取）。"""
    name = (encoding or "").lower()
    return any(key in name for key in _TWO_BYTE_CMAPS)


def _lookup_cids(
    doc: "fitz.Document",
    cidfont: int,
    current: dict[int, int],
    new_b: bytes,
    step: int,
    text: str,
) -> list[int]:
    """找出这段字节实际对应的 CID。

    不能直接把字节值当 CID：UniGB-UCS2-H 会把 Unicode 码位映射到 GB1 里的 CID
    （数字 "7" 的码位是 0x37，CID 却是 24）。

    这里按"字符类别 + 已知区间"精确推算，不靠穷举渲染（那要几百次，太慢）：
    在 GB1 的半角区里，数字 0–9 固定是 CID 17–26，拉丁字母 27 起连续排列。
    只要 /W 里已经有同类的 CI(D) 记录，就能据此校验并换算出目标 CID。
    """
    naive = [int.from_bytes(new_b[i:i + step], "big") for i in range(0, len(new_b), step)]
    if all(c in current for c in naive):
        return naive  # Identity 编码：字节值就是 CID
    base = _numeric_cid_base(current)
    if base is None:
        # 没有可用于校验的数字区间（例如纯汉字场景），不猜 CID
        return naive
    result = []
    for index, code in enumerate(naive):
        char = text[index] if index < len(text) else ""
        if not char.isdigit():
            return naive  # 只处理"已由实验确认映射"的数字，其余不猜
        result.append(base + int(char))
    return result


def _numeric_cid_base(current: dict[int, int]) -> int | None:
    """数字 0 对应的 CID。

    GB1 半角区固定为：数字 "0" = CID 17 … "9" = CID 26（已用真实文件验证：
    /W 里 17–21 与 25/26 都是 462，缺的 22–24 正是数字 5/6/7）。
    用 /W 里已知的窄宽度条目校验这个前提，对不上就返回 None（此时不猜）。
    """
    if all(cid in current for cid in (17, 18, 19, 20, 21)):
        return 17
    return None


def _cidfont_xref(doc: "fitz.Document", font_xref: int) -> int | None:
    """从 Type0 字体对象取出 DescendantFonts 里的 CIDFont xref。"""
    kind, value = doc.xref_get_key(font_xref, "DescendantFonts")
    if kind != "array":
        return None
    match = re.search(r"(\d+)\s+0\s+R", value)
    return int(match.group(1)) if match else None


def _parse_w(doc: "fitz.Document", cidfont: int) -> dict[int, int]:
    """解析 /W 数组，展开成 {CID: 宽度}（支持 'c [w…]' 与 'c1 c2 w' 两种写法）。"""
    kind, value = doc.xref_get_key(cidfont, "W")
    if kind != "array":
        return {}
    result: dict[int, int] = {}
    tokens = re.findall(r"\[|\]|-?\d+", value)
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token in "[]":
            i += 1
            continue
        first = int(token)
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        if nxt == "[":
            cid = first
            j = i + 2
            while j < len(tokens) and tokens[j] != "]":
                result[cid] = int(tokens[j])
                cid += 1
                j += 1
            i = j + 1
            continue
        if nxt is None or nxt in "[]" or i + 2 >= len(tokens):
            break
        second, width = int(nxt), int(tokens[i + 2])
        for cid in range(first, min(second, first + 65535) + 1):
            result.setdefault(cid, width)
        i += 3
    return result


def _parse_dw(doc: "fitz.Document", cidfont: int) -> int:
    """/DW 默认宽度，缺省按 PDF 规范为 1000。"""
    kind, value = doc.xref_get_key(cidfont, "DW")
    try:
        return int(value) if kind == "int" else 1000
    except (TypeError, ValueError):
        return 1000


def _reference_width(current: dict[int, int], dw: int, new_text: str) -> int | None:
    """取参考宽度：优先同字体里"半角同类字符"的已有宽度。"""
    half = [w for w in current.values() if 0 < w < dw]
    if not half:
        return None
    if new_text and all(ch.isdigit() for ch in new_text):
        # 数字通常是最窄的一档，取出现最多的窄宽度
        counts: dict[int, int] = {}
        for w in half:
            counts[w] = counts.get(w, 0) + 1
        return max(counts, key=counts.get)
    return min(half)


def _write_cid_widths(doc: "fitz.Document", cidfont: int, widths: dict[int, int]) -> bool:
    """把 {CID: 宽度} 追加进 /W 数组（保留原有条目）。"""
    added = " ".join(f"{cid} {cid} {w}" for cid, w in sorted(widths.items()))
    kind, value = doc.xref_get_key(cidfont, "W")
    if kind != "array":
        doc.xref_set_key(cidfont, "W", f"[ {added} ]")
        return True
    body = value.strip()
    if not body.startswith("["):
        return False
    doc.xref_set_key(cidfont, "W", body[:-1].rstrip() + " " + added + " ]")
    return True


def _check_text(pdf: Path, pages: Sequence[int] | None, repls: Sequence[Replacement]) -> str:
    """回读输出文件，确认替换结果真的是期望文字。返回问题描述，正常时返回空串。"""
    try:
        with fitz.open(str(pdf)) as doc:
            indexes = _page_indexes(doc.page_count, pages)
            text = "".join(doc[i].get_text() for i in indexes)
    except Exception as exc:  # noqa: BLE001 - 输出文件损坏
        return f"输出文件无法再打开：{exc}"
    for repl in repls:
        if repl.count and repl.new not in text:
            return f"替换后读不到“{repl.new}”，说明该处字形的编码映射与 Unicode 不一致。"
    return ""


def verify(pdf: Path, *check_text: str) -> str:
    """打开输出文件做基础自检：页数、结构、目标文字是否存在。"""
    with fitz.open(str(pdf)) as doc:
        text = "".join(page.get_text() for page in doc)
        parts = [f"页数 {doc.page_count}", "结构完整" if not doc.is_repaired else "结构异常"]
        for needle in check_text:
            parts.append(f"“{needle}”{'存在' if needle in text else '缺失'}")
    return "，".join(parts)


def open_in_finder(path: Path) -> None:
    """在系统文件管理器中定位文件（macOS / Windows / Linux）。"""
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", str(path)], check=False)
        elif sys.platform.startswith("win"):
            subprocess.run(["explorer", "/select,", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path.parent)], check=False)
    except OSError as exc:  # pragma: no cover - 无桌面环境
        raise PdfEditError(f"无法打开文件管理器：{exc}") from exc



