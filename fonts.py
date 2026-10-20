"""字体识别与自动获取（供 PDF -> Word 转换使用）。

能力：
    1. 识别 PDF 实际使用的字体（TrueType / Type0(CID) 优先，跳过 Type3 等无法内嵌的字体）；
    2. 在本机已安装字体中按真实字体名匹配（Windows / macOS / Linux 常见目录）；
    3. 若本机缺失，可自动从 GitHub（开源字体仓库 / Google Fonts）下载对应字体文件；
    4. 字体以“字体名 + 文件路径”的形式交给 python-docx，并在 OOXML 中补齐
       eastAsia / cs 字体槽，保证 Word / WPS 显示一致。

下载只允许 https，写入用户缓存目录（可用环境变量 PDF2WORD_FONT_CACHE 覆盖）。
"""

from __future__ import annotations

import json
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

FONT_SUFFIXES = {".ttf", ".otf", ".ttc", ".otc"}
CACHE_DIR = Path(
    os.environ.get("PDF2WORD_FONT_CACHE")
    or (
        Path(os.environ.get("LOCALAPPDATA", Path.home() / ".cache")) / "pdf2word" / "fonts"
        if os.name == "nt"
        else Path.home() / ".cache" / "pdf2word" / "fonts"
    )
)

# 常见字体的本机别名：PDF 中的名字 -> 系统里可能出现的家族名/文件名。
_SYSTEM_ALIASES: dict[str, list[str]] = {
    "helvetica": ["arial", "helvetica", "liberationsans", "nimbussans"],
    "helv": ["arial", "helvetica"],
    "arial": ["arial", "liberationsans"],
    "times": ["timesnewroman", "times", "liberationserif", "nimbusroman"],
    "timesnewroman": ["timesnewroman", "liberationserif"],
    "courier": ["couriernew", "courier", "liberationmono"],
    "couriernew": ["couriernew", "liberationmono"],
    "symbol": ["symbol", "dejavusans"],
    "cambria": ["cambria", "caladea"],
    "calibri": ["calibri", "carlito"],
    "segoeui": ["segoeui", "dejavusans"],
    "arialunicodems": ["arialunicodems", "arial", "dejavusans"],
    "arialuni": ["arialunicodems", "arial", "dejavusans"],
    "songti": ["songti", "simsun", "stsong", "songtikaiti"],
    "simsun": ["simsun", "songti", "stsong"],
    "stsong": ["stsong", "songti", "simsun"],
    "heiti": ["heiti", "simhei", "stheiti", "pingfang"],
    "simhei": ["simhei", "heiti", "stheiti"],
    "pingfang": ["pingfangsc", "pingfang", "heiti"],
    "kaiti": ["kaiti", "stkaiti", "kaitisc"],
    "stkaiti": ["stkaiti", "kaiti"],
    "fangsong": ["fangsong", "stfangsong"],
    "hiraginosans": ["hiraginosans", "yugothic", "notosanscjk", "dejavusans"],
    "hiraginomincho": ["hiraginomincho", "yumincho", "notoserifcjk"],
    "notosanscjk": ["notosanscjk", "notosanssc", "sourcehansans"],
}

# 自动下载源：repo -> 文件名正则（尽量只下载单个字体文件，避免整包）。
_DOWNLOAD_SOURCES: dict[str, dict] = {
    "lxgw": {
        "label": "霞鹜文楷（开源中文字体，OFL）",
        "repo": "lxgw/LxgwWenKai",
        "pattern": r"^LXGWWenKai-Regular\.ttf$",
        "families": ["lxgwwenkai"],
        "cjk": True,
    },
    "notosans": {
        "label": "Noto Sans（开源西文字体，OFL）",
        "repo": "google/fonts",
        "branch": "main",
        "dir": "ofl/notosans",
        "pattern": r"^NotoSans\[wdth,wght\]\.ttf$",
        "families": ["notosans"],
        "cjk": False,
    },
    "notoserif": {
        "label": "Noto Serif（开源西文衬线字体，OFL）",
        "repo": "google/fonts",
        "branch": "main",
        "dir": "ofl/notoserif",
        "pattern": r"^NotoSerif\[wdth,wght\]\.ttf$",
        "families": ["notoserif"],
        "cjk": False,
    },
}

ProgressCallback = Callable[[str], None]


@dataclass
class FontRecord:
    """PDF 中出现的一种字体。"""

    pdf_name: str            # 内部资源名，如 /F1
    base_name: str           # 字体家族名，如 Arial Unicode MS Regular
    font_type: str           # Type1 / TrueType / Type0 ...
    encoding: str
    embedded: bool
    xref: int
    size: float = 0.0
    used: int = 0            # 出现次数（字符数）
    is_cjk: bool = False

    @property
    def safe(self) -> bool:
        """能否内嵌进 Word（Type3 / Type1 标准 14 字体无法内嵌）。"""
        return self.font_type.lower() in {"type0", "truetype", "cid"} or self.embedded

    @property
    def display(self) -> str:
        return f"{self.base_name}（{self.font_type}）"


@dataclass
class ResolvedFont:
    pdf_name: str
    family: str                     # 写入 Word 的字体名
    source: str                     # embedded / system / download / fallback
    path: Path | None = None
    cjk_family: str | None = None   # 东亚字体（宋体/黑体等）
    note: str = ""


@dataclass
class FontReport:
    records: list[FontRecord] = field(default_factory=list)
    resolved: dict[str, ResolvedFont] = field(default_factory=dict)
    downloaded: list[Path] = field(default_factory=list)

    def summary(self) -> str:
        if not self.records:
            return "未检测到字体信息"
        kinds = {"embedded": 0, "system": 0, "download": 0, "fallback": 0}
        for item in self.resolved.values():
            kinds[item.source] = kinds.get(item.source, 0) + 1
        parts = [f"共 {len(self.records)} 种字体"]
        for key, label in (
            ("embedded", "内嵌"),
            ("system", "本机已有"),
            ("download", "自动下载"),
            ("fallback", "替代"),
        ):
            if kinds.get(key):
                parts.append(f"{label} {kinds[key]}")
        return "，".join(parts)


def normalize(name: str) -> str:
    """去掉子集前缀与样式后缀，只留可比较的核心名。"""
    text = re.sub(r"^[A-Z]{6}\+", "", name or "")      # ABCDEF+Arial
    text = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", text).lower()
    for suffix in (
        "bolditalic", "boldoblique", "italic", "oblique", "bold", "regular",
        "medium", "light", "semibold", "black", "thin", "heavy",
    ):
        if text.endswith(suffix) and len(text) > len(suffix) + 2:
            text = text[: -len(suffix)]
    return text


def is_cjk_name(name: str) -> bool:
    core = normalize(name)
    return any(
        key in core
        for key in ("song", "hei", "kai", "fang", "ming", "gothic", "mincho", "cjk",
                    "notosanssc", "notoserifsc", "pingfang", "yahei", "simsun")
    ) or any(ch in (name or "") for ch in "宋黑楷仿明")


# --------------------------------------------------------------------------
# 1) 识别 PDF 字体
# --------------------------------------------------------------------------
def scan_pdf_fonts(pdf_path: str | Path, max_pages: int | None = None) -> list[FontRecord]:
    """扫描 PDF 每一页实际用到的字体，并统计使用次数与字号。"""
    import pymupdf

    records: dict[str, FontRecord] = {}
    with pymupdf.open(str(pdf_path)) as doc:
        pages = range(doc.page_count if max_pages is None else min(max_pages, doc.page_count))
        for index in pages:
            page = doc.load_page(index)
            for entry in page.get_fonts(full=True):
                xref, ext, ftype, base, name, encoding, *_rest = entry
                if not name:
                    continue
                record = records.get(name)
                if record is None:
                    record = FontRecord(
                        pdf_name=name,
                        base_name=base or name,
                        font_type=ftype,
                        encoding=encoding,
                        embedded=ext not in ("n/a", "", None),
                        xref=xref,
                        is_cjk=is_cjk_name(base or name),
                    )
                    records[name] = record
                for block in page.get_text("dict").get("blocks", []):
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("font") != base:
                                continue
                            record.used += len(span.get("text", ""))
                            record.size = max(record.size, float(span.get("size", 0)))
    return sorted(records.values(), key=lambda r: (-r.used, r.pdf_name))


# --------------------------------------------------------------------------
# 2) 本机字体索引
# --------------------------------------------------------------------------
def system_font_dirs() -> list[Path]:
    if os.name == "nt":
        roots = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
        local = os.environ.get("LOCALAPPDATA")
        if local:
            roots.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
        return [p for p in roots if p.is_dir()]
    dirs = [
        Path.home() / "Library" / "Fonts",
        Path("/Library/Fonts"),
        Path("/System/Library/Fonts"),
        Path("/System/Library/Fonts/Supplemental"),
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path.home() / ".fonts",
        Path.home() / ".local" / "share" / "fonts",
        CACHE_DIR,
    ]
    return [p for p in dirs if p.is_dir()]


def _ttc_stems(path: Path) -> list[str]:
    """TTC 集合通常含多个字体；只把文件名作为索引键，避免依赖第三方库。"""
    return [path.stem]


def build_system_index() -> dict[str, Path]:
    """返回 {规范化字体名: 字体文件路径}。"""
    index: dict[str, Path] = {}
    for folder in system_font_dirs():
        try:
            files = list(folder.rglob("*")) if folder == CACHE_DIR else list(folder.iterdir())
        except OSError:
            continue
        for path in files:
            if not path.is_file() or path.suffix.lower() not in FONT_SUFFIXES:
                continue
            for stem in _ttc_stems(path):
                index.setdefault(normalize(stem), path)
    return index


def find_in_system(name: str, index: dict[str, Path] | None = None) -> tuple[str, Path] | None:
    """按字体名 + 常见别名查找本机字体，返回 (家族名, 路径)。

    匹配优先级：完全同名 > 别名精确 > 更长的名字包含匹配（Arial Unicode MS
    应优先命中 "Arial Unicode"，而不是 "Arial Bold"）。
    """
    table = index if index is not None else build_system_index()
    core = normalize(name)
    candidates = [core] + _SYSTEM_ALIASES.get(core, [])
    for candidate in candidates:
        path = table.get(candidate)
        if path is not None:
            return (path.stem, path)

    if len(core) < 4:
        return None
    # 名字互相包含时，取重叠最长（最接近）的候选
    best: tuple[int, Path] | None = None
    for key, path in table.items():
        if len(key) < 4:
            continue
        if key in core or core in key:
            score = min(len(key), len(core))
            if best is None or score > best[0]:
                best = (score, path)
    if best is not None:
        return (best[1].stem, best[1])
    return None


# --------------------------------------------------------------------------
# 3) 自动下载
# --------------------------------------------------------------------------
def _http_json(url: str, timeout: float = 20.0) -> object:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "pdf2word-fontfetcher", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:  # noqa: S310 - 仅 https
        return json.loads(response.read().decode("utf-8"))


def download_targets(source_key: str) -> list[tuple[str, str]]:
    """返回 [(字体文件名, 下载地址)]，不实际下载。"""
    if source_key not in _DOWNLOAD_SOURCES:
        raise ValueError(f"未知字体来源：{source_key}")
    info = _DOWNLOAD_SOURCES[source_key]
    repo = info["repo"]
    pattern = re.compile(str(info["pattern"]))

    if info.get("branch"):
        # Google Fonts 仓库：直接拼 raw 地址，避免 API 限流
        folder = info["dir"]
        listing = _http_json(
            f"https://api.github.com/repos/{repo}/contents/{folder}?ref={info['branch']}"
        )
        raw = f"https://raw.githubusercontent.com/{repo}/{info['branch']}/{folder}"
        return [
            (item["name"], f"{raw}/{item['name']}")
            for item in listing
            if isinstance(item, dict) and pattern.match(item.get("name", ""))
        ]

    release = _http_json(f"https://api.github.com/repos/{repo}/releases/latest")
    assets = release.get("assets", []) if isinstance(release, dict) else []
    return [
        (asset["name"], asset["browser_download_url"])
        for asset in assets
        if pattern.match(asset.get("name", ""))
    ]


def _looks_like_font(head: bytes) -> bool:
    return (
        head.startswith(b"\x00\x01\x00\x00")       # TrueType
        or head.startswith(b"OTTO")                 # OpenType/CFF
        or head.startswith(b"true")                 # Apple TrueType
        or head.startswith(b"ttcf")                 # TrueType Collection
    )


def download_font(source_key: str, callback: ProgressCallback | None = None) -> list[Path]:
    """下载某个来源的字体到缓存目录，返回本地路径列表（已存在则跳过）。"""
    if source_key not in _DOWNLOAD_SOURCES:
        raise ValueError(f"未知字体来源：{source_key}")
    label = _DOWNLOAD_SOURCES[source_key]["label"]
    _log(callback, f"  正在查询字体下载地址：{label}")
    targets = download_targets(source_key)
    if not targets:
        raise RuntimeError(f"没有找到可下载的字体文件：{label}")

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for filename, url in targets:
        target = CACHE_DIR / filename
        if target.exists() and target.stat().st_size > 1024:
            saved.append(target)
            _log(callback, f"  已存在：{filename}")
            continue
        _log(callback, f"  下载 {filename} …")
        request = urllib.request.Request(url, headers={"User-Agent": "pdf2word-fontfetcher"})
        with urllib.request.urlopen(request, timeout=60) as response:  # noqa: S310
            data = response.read()
        if not _looks_like_font(data[:4]):
            raise RuntimeError(f"下载内容不是字体文件：{filename}")
        temp = target.with_suffix(target.suffix + ".part")
        temp.write_bytes(data)
        temp.replace(target)
        saved.append(target)
        _log(callback, f"  ✔ {filename}（{len(data) / 1024 / 1024:.1f} MB）")
    return saved


def download_sources() -> dict[str, dict]:
    """给界面展示可下载字体列表。"""
    return dict(_DOWNLOAD_SOURCES)


# --------------------------------------------------------------------------
# 4) 解析：把 PDF 字体映射成 Word 可用字体
# --------------------------------------------------------------------------
def _extract_embedded(record: FontRecord, pdf_path: str | Path) -> Path | None:
    """尝试从 PDF 里抽出嵌入字体（本机缺失时使用）。"""
    import pymupdf

    try:
        with pymupdf.open(str(pdf_path)) as doc:
            for xref in range(1, doc.xref_length()):
                try:
                    name, ext, _kind, buffer = doc.extract_font(xref)
                except Exception:  # noqa: BLE001 - 非字体对象
                    continue
                if not name or not buffer:
                    continue
                if normalize(name) != normalize(record.base_name):
                    continue
                suffix = "." + (ext or "ttf").lower()
                if suffix not in FONT_SUFFIXES:
                    suffix = ".ttf"
                CACHE_DIR.mkdir(parents=True, exist_ok=True)
                safe = re.sub(r"[^0-9A-Za-z]+", "-", name).strip("-") or "font"
                target = CACHE_DIR / f"embedded-{safe}{suffix}"
                if not target.exists():
                    target.write_bytes(buffer)
                return target
    except Exception:  # noqa: BLE001
        return None
    return None


def _pick_source(record: FontRecord) -> str | None:
    """根据字体类型决定下载哪种开源替代字体。"""
    core = normalize(record.base_name)
    if record.is_cjk or core in {"songti", "simsun", "heiti", "kaiti", "ming"}:
        return "lxgw"
    if "serif" in core or "times" in core or "song" in core or "roman" in core:
        return "notoserif"
    return "notosans"


def resolve_fonts(
    records: list[FontRecord],
    pdf_path: str | Path,
    *,
    auto_download: bool = False,
    callback: ProgressCallback | None = None,
    index: dict[str, Path] | None = None,
    extract_embedded: bool = True,
) -> FontReport:
    """为每种 PDF 字体确定 Word 字体名与字体文件（下载失败会自动兜底）。

    extract_embedded=False 时跳过「从 PDF 抽出内嵌字体并写入缓存目录」这一步，
    用于「文件不留存」场景：不把用户文档里的字体字符集落盘。
    """
    report = FontReport(records=records)
    table = index

    def get_index() -> dict[str, Path]:
        nonlocal table
        if table is None:
            table = build_system_index()
        return table

    for record in records:
        # 1) 本机已安装
        found = find_in_system(record.base_name, get_index())
        if found is not None:
            family, path = found
            report.resolved[record.pdf_name] = ResolvedFont(
                record.pdf_name, family, "system", path,
                cjk_family=family if record.is_cjk else None,
            )
            continue

        # 2) PDF 内嵌字体（可关闭：避免把用户文档的字体子集写进本机缓存）
        embedded = _extract_embedded(record, pdf_path) if extract_embedded else None
        if embedded is not None:
            report.resolved[record.pdf_name] = ResolvedFont(
                record.pdf_name, record.base_name, "embedded", embedded,
                cjk_family=record.base_name if record.is_cjk else None,
            )
            continue

        # 3) 自动下载开源替代字体
        source_key = _pick_source(record) if auto_download else None
        if source_key is not None:
            try:
                files = download_font(source_key, callback)
            except (urllib.error.URLError, RuntimeError, OSError) as exc:
                _log(callback, f"  下载失败：{exc}")
            else:
                if files:
                    family = files[0].stem
                    report.downloaded.extend(files)
                    report.resolved[record.pdf_name] = ResolvedFont(
                        record.pdf_name, family, "download", files[0],
                        cjk_family=family if record.is_cjk else None,
                        note=f"下载自 {_DOWNLOAD_SOURCES[source_key]['label']}",
                    )
                    get_index()[normalize(family)] = files[0]
                    continue

        # 4) 兜底：常见系统字体
        core = normalize(record.base_name)
        if record.is_cjk:
            family, cjk = "SimSun", "SimSun"
        elif "times" in core or "serif" in core:
            family, cjk = "Times New Roman", None
        elif "courier" in core or "mono" in core:
            family, cjk = "Courier New", None
        else:
            family, cjk = "Arial", None
        report.resolved[record.pdf_name] = ResolvedFont(
            record.pdf_name, family, "fallback", None,
            cjk_family=cjk, note="本机未找到且未下载，已用通用字体替代",
        )
    return report


def _log(callback: ProgressCallback | None, message: str) -> None:
    if callback is not None:
        callback(message)


# --------------------------------------------------------------------------
# 5) 把识别结果写进 .docx
# --------------------------------------------------------------------------
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def dominant_fonts(report: FontReport) -> tuple[str, str, str]:
    """挑出正文最常用的西文字体、东亚字体、等宽字体。

    返回 (latin_family, east_asia_family, mono_family)。
    """
    by_usage = sorted(report.records, key=lambda r: -r.used)
    latin = ""
    east_asia = ""
    mono = ""
    for record in by_usage:
        resolved = report.resolved.get(record.pdf_name)
        if resolved is None:
            continue
        family = resolved.family
        if record.is_cjk:
            if not east_asia:
                east_asia = resolved.cjk_family or family
            continue
        core = normalize(record.base_name)
        if any(k in core for k in ("courier", "mono", "consol", "menlo")):
            if not mono:
                mono = family
            continue
        if not latin:
            latin = family
    if not latin:
        latin = next(
            (r.family for r in report.resolved.values() if r.source != "fallback"),
            "Arial",
        )
    return latin, east_asia or "宋体", mono or latin


def apply_font_report(docx_path: str | Path, report: FontReport) -> None:
    """把识别到的字体写入 docx 的默认样式 + Normal 样式（Word/WPS 都能识别）。"""
    from docx import Document
    from docx.oxml.ns import qn

    latin, east_asia, mono = dominant_fonts(report)
    document = Document(str(docx_path))

    def set_rfonts(name: str, family: str) -> None:
        fonts_element = name.get_or_add_rFonts()
        fonts_element.set(qn("w:ascii"), family)
        fonts_element.set(qn("w:hAnsi"), family)
        fonts_element.set(qn("w:eastAsia"), east_asia)
        fonts_element.set(qn("w:cs"), family)

    # 1) docDefaults（影响所有新建段落）
    styles_element = document.styles.element
    rpr_default = styles_element.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault"))
    if rpr_default is not None:
        rpr = rpr_default.find(qn("w:rPr"))
        if rpr is None:
            from docx.oxml import OxmlElement

            rpr = OxmlElement("w:rPr")
            rpr_default.append(rpr)
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            from docx.oxml import OxmlElement

            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), latin)
        rfonts.set(qn("w:hAnsi"), latin)
        rfonts.set(qn("w:eastAsia"), east_asia)
        rfonts.set(qn("w:cs"), latin)

    # 2) Normal 样式
    try:
        normal = document.styles["Normal"]
        normal.font.name = latin
        if normal.element.rPr is None:
            normal.element.get_or_add_rPr()
        set_rfonts(normal.element.rPr, latin)
    except (KeyError, AttributeError):
        pass

    # 3) 等宽字体（代码块）单独设置
    for style_name in ("Macro Text", "HTML Preformatted", "No Spacing"):
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        if mono:
            style.font.name = mono
            if style.element.rPr is not None:
                set_rfonts(style.element.rPr, mono)

    document.save(str(docx_path))

