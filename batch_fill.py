#!/usr/bin/env python3
"""批量生成引擎：Excel 一行 -> 一份 PDF，模板里用 {{字段名}} 占位。

能力边界（重要）：
    pdf_edit / pdf_replace 的核心卖点是“等长替换、版式零变化”，因此**拒绝**
    {{姓名}} -> 张三 这种不等长替换。批量填表天然是不等长的，所以这里单独实现
    “覆盖 + 重绘”（redaction 方式）：
        1. 用 PyMuPDF 在指定页找到 {{字段}} 的实际位置和字号；
        2. 擦掉占位符所占区域（同时移除原文字，避免复制 PDF 时残影）；
        3. 按原字号、原颜色、原基线重绘填值；
        4. 填值太长时自动缩小字号，保证不溢出到相邻内容。
    代价：重绘的字形来自系统字体，和模板原字体不会 100% 一致（这一点必须在
    界面和日志里如实告诉用户，而不是假装“完全无痕”）。

主要入口：
    read_table(data, filename)                 -> TableData（表头 + 行）
    inspect_template(data)                     -> list[str]（模板里的 {{字段}}）
    sample_rows(...) / sample_template_pdf()   -> 示例文件（下载用）
    render_one(template, row)                   -> bytes（单份 PDF）
    preview_batch(df, template)                -> BatchPreview（校验报告 + 首页预览图）
    build_zip(df, template)                    -> BatchResult（zip 字节 + 报告）
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Iterable, Sequence
from xml.etree import ElementTree as ET

import pymupdf
import pandas as pd

MAX_BATCH = 500                    # 单次上限，防止一次生成几百份把内存打满
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]{1,40}?)\s*\}\}")
FONT_CANDIDATES = (
    ("zjb-cjk", "/System/Library/Fonts/Supplemental/Songti.ttc"),
    ("zjb-cjk", "/System/Library/Fonts/STHeiti Light.ttc"),
    ("china-s", None),             # PyMuPDF 内置 CJK 字体，兜底
)
DEFAULT_FIRST_LINE = "姓名"


class BatchError(ValueError):
    """数据表 / 模板本身不可用（会在界面上以中文提示展示）。"""


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------
@dataclass
class TableData:
    """一张数据表：表头 + 数据行。"""

    headers: list[str] = field(default_factory=list)
    rows: list[dict[str, str]] = field(default_factory=list)
    note: str = ""

    @property
    def count(self) -> int:
        return len(self.rows)


@dataclass
class BatchPreview:
    """批量生成前的检查报告。"""

    ok: bool = False
    message: str = ""
    fields: list[str] = field(default_factory=list)     # 模板里的占位符
    headers: list[str] = field(default_factory=list)    # 表头
    missing: list[str] = field(default_factory=list)    # 模板要、表里没有
    empty_cells: int = 0                                # 空值格数
    total: int = 0
    preview_png: bytes = b""
    preview_name: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass
class BatchResult:
    """批量生成结果。"""

    ok: bool = False
    message: str = ""
    output: bytes | None = None
    filename: str = "批量生成结果.zip"
    generated: int = 0
    skipped: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# 读取数据表（xlsx 用内置 XML 解析，csv 用标准库；不依赖 openpyxl）
# --------------------------------------------------------------------------
_XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _xlsx_cell_column(ref: str) -> int:
    """把 A1 / BC12 这样的单元格引用换算成 0 起的列号。"""
    letters = "".join(ch for ch in ref if ch.isalpha()).upper()
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - 64)
    return max(value - 1, 0)


def _read_xlsx(data: bytes) -> TableData:
    """读取 xlsx 的第一个工作表（标准库 zipfile + ElementTree 实现）。"""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise BatchError("Excel 文件已损坏或不是 .xlsx 格式，请另存后再上传。") from exc

    with archive:
        names = set(archive.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{_XLSX_NS}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{_XLSX_NS}t")))

        # 找第一张工作表（workbook.xml 里的第一个 sheet + rels 映射）
        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            first = workbook.find(f"{_XLSX_NS}sheets/{_XLSX_NS}sheet")
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            if first is not None:
                rid = first.get(f"{_REL_NS}id")
                for rel in rels:
                    if rel.get("Id") == rid:
                        target = (rel.get("Target") or "").lstrip("/")
                        sheet_path = target if target.startswith("xl/") else f"xl/{target}"
                        break
        if sheet_path not in names:
            raise BatchError("Excel 里没有找到工作表，请检查文件内容。")

        sheet = ET.fromstring(archive.read(sheet_path))

    table: list[list[str]] = []
    for row in sheet.iter(f"{_XLSX_NS}row"):
        cells: dict[int, str] = {}
        for cell in row.findall(f"{_XLSX_NS}c"):
            ref = cell.get("r") or ""
            column = _xlsx_cell_column(ref)
            kind = cell.get("t")
            value = ""
            if kind == "inlineStr":
                value = "".join(node.text or "" for node in cell.iter(f"{_XLSX_NS}t"))
            else:
                node = cell.find(f"{_XLSX_NS}v")
                raw = node.text if node is not None and node.text else ""
                if kind == "s" and raw.isdigit():
                    index = int(raw)
                    value = shared[index] if index < len(shared) else ""
                else:
                    value = raw
            cells[column] = value.strip()
        if cells:
            width = max(cells) + 1
            table.append([cells.get(index, "") for index in range(width)])

    if not table:
        raise BatchError("Excel 是空的，请至少填写一行数据。")
    headers = [h for h in table[0]]
    if not any(headers):
        raise BatchError("第一行（表头）是空的，请把字段名写在第一行。")
    headers = [h or f"列{index + 1}" for index, h in enumerate(headers)]

    rows: list[dict[str, str]] = []
    for line in table[1:]:
        if not any(cell.strip() for cell in line):
            continue  # 跳过空行
        record = {headers[i]: (line[i] if i < len(line) else "") for i in range(len(headers))}
        rows.append(record)
    return TableData(headers=headers, rows=rows, note=f"读取到 {len(rows)} 行数据")


def _read_csv(data: bytes) -> TableData:
    """读取 csv（自动尝试 utf-8 / gbk，兼容 Excel 导出的中文表）。"""
    text = ""
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise BatchError("CSV 编码无法识别，请另存为 UTF-8 编码后再上传。")

    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    reader = list(csv.reader(io.StringIO(text), dialect))
    if not reader:
        raise BatchError("CSV 是空的，请至少填写一行数据。")

    headers = [cell.strip() or f"列{index + 1}" for index, cell in enumerate(reader[0])]
    rows: list[dict[str, str]] = []
    for line in reader[1:]:
        if not any(cell.strip() for cell in line):
            continue
        rows.append({headers[i]: (line[i].strip() if i < len(line) else "") for i in range(len(headers))})
    return TableData(headers=headers, rows=rows, note=f"读取到 {len(rows)} 行数据")


def read_table(data: bytes, filename: str) -> TableData:
    """按扩展名读取数据表。"""
    if not data:
        raise BatchError("数据表为空，请重新上传。")
    suffix = ("." + filename.rsplit(".", 1)[-1].lower()) if "." in filename else ""
    if suffix in (".xlsx", ".xlsm"):
        return _read_xlsx(data)
    if suffix == ".csv":
        return _read_csv(data)
    if suffix == ".xls":
        raise BatchError("旧版 .xls 暂不支持，请在 Excel 里另存为 .xlsx 后再上传。")
    raise BatchError(f"不支持的数据表格式：{suffix or '未知'}，请上传 .xlsx 或 .csv。")


# --------------------------------------------------------------------------
# 示例文件（下载用）
# --------------------------------------------------------------------------
SAMPLE_HEADERS = ["姓名", "岗位", "入职日期", "月薪", "公司名称", "签发日期"]
SAMPLE_ROWS = [
    ["张三", "前端工程师", "2026-03-01", "18000", "示例科技有限公司", "2026-02-20"],
    ["李四", "产品经理", "2026-03-08", "22000", "示例科技有限公司", "2026-02-20"],
    ["王五", "财务专员", "2026-03-15", "14000", "示例科技有限公司", "2026-02-20"],
]


SAMPLE_SHEET_TITLE = "批量生成数据"


def _column_letter(index: int) -> str:
    """0 起的列号 -> Excel 列名（0 -> A，26 -> AA）。"""
    letters = ""
    value = index + 1
    while value:
        value, rest = divmod(value - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _build_minimal_xlsx(headers: list[str], rows: list[list[str]], sheet_title: str) -> bytes:
    """纯标准库写出的最小可用 .xlsx（兜底用，pandas / openpyxl 不可用时仍能下载打开）。"""
    from xml.sax.saxutils import escape

    table = [headers] + rows
    body: list[str] = []
    for row_index, row in enumerate(table, start=1):
        cells = "".join(
            f'<c r="{_column_letter(col)}{row_index}" t="inlineStr">'
            f'<is><t xml:space="preserve">{escape(str(value))}</t></is></c>'
            for col, value in enumerate(row)
        )
        body.append(f'<row r="{row_index}">{cells}</row>')
    widths = "".join(
        f'<col min="{col}" max="{col}" width="{min(28, max(len(str(headers[col - 1])) * 2, 10) + 4)}" customWidth="1"/>'
        for col in range(1, len(headers) + 1)
    )

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_XLSX_NS[1:-1]}">'
        f"<cols>{widths}</cols>"
        f'<sheetData>{"".join(body)}</sheetData>'
        "</worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_XLSX_NS[1:-1]}" '
        f'xmlns:r="{_REL_NS[1:-1]}">'
        f'<sheets><sheet name="{escape(sheet_title)}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL_NS[1:-1]}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL_NS[1:-1]}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return buffer.getvalue()


def sample_table_xlsx() -> bytes:
    """示例数据表：标准 .xlsx（带表头行），WPS / Excel 可直接双击打开。

    注意：必须是真正的 xlsx（zip 包），不能拿 CSV 字节冒充——那样 WPS 会报文件损坏。
    优先用 pandas + openpyxl 生成；任一依赖不可用时自动退回标准库实现，
    保证「下载 Excel 示例」按钮永远有真实文件可下，不会抛 AttributeError / ImportError。
    """
    try:
        frame = pd.DataFrame(SAMPLE_ROWS, columns=SAMPLE_HEADERS)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            frame.to_excel(writer, index=False, sheet_name=SAMPLE_SHEET_TITLE)
            sheet = writer.sheets[SAMPLE_SHEET_TITLE]
            # 列宽按内容自适应，打开就整齐（中文字符按 2 个字符宽估算）
            for column, header in enumerate(SAMPLE_HEADERS, start=1):
                longest = max(
                    [len(str(header)) * 2]
                    + [len(str(row[column - 1])) * 2 for row in SAMPLE_ROWS]
                )
                letter = sheet.cell(row=1, column=column).column_letter
                sheet.column_dimensions[letter].width = min(28, longest + 4)
        payload = buffer.getvalue()
        if zipfile.is_zipfile(io.BytesIO(payload)):
            return payload
    except Exception:  # noqa: BLE001 - 没装 openpyxl / pandas 版本差异都走兜底
        pass
    return _build_minimal_xlsx(SAMPLE_HEADERS, SAMPLE_ROWS, SAMPLE_SHEET_TITLE)


def sample_table_csv() -> bytes:
    """示例数据表（CSV 版）：带 BOM，Excel 双击打开不乱码。"""
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(SAMPLE_HEADERS)
    writer.writerows(SAMPLE_ROWS)
    return buffer.getvalue().encode("utf-8-sig")


def sample_template_pdf() -> bytes:
    """示例模板：A4 一页的录用通知，含 {{姓名}} / {{岗位}} 等占位符。

    用 reportlab 生成标准 PDF（xref 表、字体子集化都由库处理），WPS / Adobe
    都能正常打开；不再手拼 PDF 结构，避免个别阅读器报“文件已损坏”。
    reportlab 缺失或字体全部不可用时返回 b""（界面据此显示中文提示，而不是崩溃）。
    """
    try:
        return _build_sample_template_pdf()
    except Exception:  # noqa: BLE001 - 示例文件不该拖垮整个页面
        return b""


def _build_sample_template_pdf() -> bytes:
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas

    font_name = "ZJB-CJK"
    if font_name not in pdfmetrics.getRegisteredFontNames():
        registered = False
        for _, path in FONT_CANDIDATES:
            if not path:
                continue
            from pathlib import Path

            if not Path(path).is_file():
                continue
            try:
                pdfmetrics.registerFont(TTFont(font_name, path, subfontIndex=0))
                registered = True
                break
            except Exception:  # noqa: BLE001 - 字体损坏就换下一个候选
                continue
        if not registered:  # 兜底：reportlab 自带的中文字体
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont

            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            font_name = "STSong-Light"
    if font_name not in pdfmetrics.getRegisteredFontNames():  # 字体注册彻底失败
        raise RuntimeError("没有可用的中文字体，无法生成示例模板")

    buffer = io.BytesIO()
    page = canvas.Canvas(buffer, pagesize=A4)
    page.setTitle("录用通知书（批量生成模板示例）")
    ink = HexColor("#1F2224")
    navy = HexColor("#0A2440")

    def top(distance: float) -> float:
        """距页面顶部 distance 磅的基线 y（reportlab 用左下角原点，换算一次就不会错）。"""
        return A4[1] - distance

    page.setFillColor(navy)
    page.setFont(font_name, 26)
    page.drawString(72, top(110), "录用通知书")
    page.setStrokeColor(navy)
    page.setLineWidth(1.2)
    page.line(72, top(128), 523, top(128))

    body = [
        ("尊敬的 {{姓名}}：", 15),
        ("", 10),
        ("您好！很高兴通知您，我司决定录用您担任 {{岗位}} 一职，", 13),
        ("入职报到日期为 {{入职日期}}，试用期 3 个月。", 13),
        ("", 10),
        ("转正后月度薪酬为人民币 {{月薪}} 元（税前），", 13),
        ("公司提供五险一金、带薪年假及年度体检。", 13),
        ("", 10),
        ("请于收到本通知后 5 个工作日内回复确认。期待您的加入！", 13),
    ]
    page.setFillColor(ink)
    offset = 175
    for text, size in body:
        if text:
            page.setFont(font_name, size)
            page.drawString(72, top(offset), text)
        offset += size * 1.9

    page.setFont(font_name, 13)
    page.drawString(72, top(600), "{{公司名称}}（盖章）")
    page.drawString(72, top(628), "{{签发日期}}")

    page.setStrokeColor(HexColor("#C9D2DE"))
    page.setLineWidth(0.6)
    page.rect(72, top(720), 90 * mm, 80, stroke=1, fill=0)  # 盖章框，包住左侧文字下方
    page.showPage()
    page.save()
    return buffer.getvalue()


# --------------------------------------------------------------------------
# 模板解析
# --------------------------------------------------------------------------
def _font() -> tuple[str, str | None]:
    from pathlib import Path

    for name, file in FONT_CANDIDATES:
        if file is None or Path(file).is_file():
            return name, file
    return FONT_CANDIDATES[-1]


def template_fields(data: bytes) -> list[str]:
    """扫描模板里出现的 {{字段}}，按首次出现顺序去重（支持跨行断开的占位符）。"""
    try:
        document = pymupdf.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise BatchError(f"模板 PDF 无法打开：{exc}") from exc

    found: list[str] = []
    try:
        if document.needs_pass:
            raise BatchError("模板 PDF 已加密，请先解除密码后再上传。")
        for page in document:
            for match in PLACEHOLDER_RE.finditer(page.get_text()):
                name = match.group(1).strip()
                if name and name not in found:
                    found.append(name)
    finally:
        document.close()
    return found


def prepare_preview(table: TableData, template: bytes) -> BatchPreview:
    """校验模板与数据表是否匹配，并生成第一份的预览图。"""
    preview = BatchPreview(total=table.count)
    try:
        preview.fields = template_fields(template)
    except BatchError as exc:
        preview.message = str(exc)
        return preview
    preview.headers = table.headers

    if not preview.fields:
        preview.message = (
            "模板里没有找到 {{字段}} 占位符。请把要填的地方写成 {{姓名}} 这种格式"
            "（花括号必须是英文半角）。"
        )
        return preview
    if not table.rows:
        preview.message = "数据表里只有表头、没有数据行，请补充后再试。"
        return preview

    preview.missing = [name for name in preview.fields if name not in table.headers]
    if preview.missing:
        preview.message = (
            "模板里的字段和数据表表头对不上：" + "、".join(preview.missing)
            + f"。数据表表头为：{'、'.join(table.headers)}。"
        )
        return preview

    preview.empty_cells = sum(
        1 for row in table.rows for name in preview.fields if not str(row.get(name, "")).strip()
    )
    if preview.empty_cells:
        preview.warnings.append(f"数据表里有 {preview.empty_cells} 个空格子，将按空白填入。")
    if preview.total > MAX_BATCH:
        preview.warnings.append(f"数据有 {preview.total} 行，单次最多生成 {MAX_BATCH} 份，超出部分会跳过。")

    try:
        rendered, _ = render_one(template, table.rows[0])
        document = pymupdf.open(stream=rendered, filetype="pdf")
        pixmap = document[0].get_pixmap(matrix=pymupdf.Matrix(1.4, 1.4), alpha=False)
        preview.preview_png = pixmap.tobytes("png")
        document.close()
        preview.preview_name = f"第 1 份预览（{preview.fields[0]}：{table.rows[0].get(preview.fields[0], '')}）"
    except BatchError as exc:
        preview.message = str(exc)
        return preview
    except Exception as exc:  # noqa: BLE001
        preview.message = f"预览生成失败：{exc}"
        return preview

    preview.ok = True
    preview.message = f"校验通过：{len(preview.fields)} 个字段 · {preview.total} 行数据"
    return preview


# --------------------------------------------------------------------------
# 渲染：定位 {{字段}} -> 擦除 -> 按原字号重绘
# --------------------------------------------------------------------------
def _find_placeholders(page: pymupdf.Page, name: str) -> list[pymupdf.Rect]:
    """在一页里找出某个占位符的所有位置（跨行时 PyMuPDF 会返回多个矩形）。"""
    return page.search_for("{{" + name + "}}")


def _fit_font_size(text: str, max_width: float, size: float, font: pymupdf.Font) -> float:
    """文字太长时逐步缩小字号，直到宽度放得下（最小 5pt）。"""
    current = size
    while current > 5 and font.text_length(text, fontsize=current) > max_width:
        current -= 0.5
    return max(current, 5)


def _style_for(page: pymupdf.Page, name: str) -> tuple[float, tuple[float, float, float]]:
    """取占位符所在的字号与颜色；找不到时按 12pt 黑字处理。"""
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if name in span.get("text", ""):
                    value = span.get("color", 0)
                    color = (
                        ((value >> 16) & 255) / 255,
                        ((value >> 8) & 255) / 255,
                        (value & 255) / 255,
                    )
                    return float(span.get("size", 12.0)), color
    return 12.0, (0.0, 0.0, 0.0)


def render_one(template: bytes, row: dict[str, str]) -> tuple[bytes, list[str]]:
    """用一行数据渲染一份 PDF，返回 (PDF 字节, 提示信息)。"""
    notes: list[str] = []
    try:
        document = pymupdf.open(stream=template, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise BatchError(f"模板 PDF 无法打开：{exc}") from exc

    font_name, font_file = _font()
    try:
        font = pymupdf.Font(fontfile=font_file) if font_file else pymupdf.Font(fontname=font_name)
    except Exception:  # noqa: BLE001 - 系统字体不可用时退回内置字体
        font_name, font_file = "china-s", None
        font = pymupdf.Font(fontname=font_name)

    try:
        for page in document:
            for name in dict.fromkeys(PLACEHOLDER_RE.findall(page.get_text())):
                value = str(row.get(name, "")).strip()
                rects = _find_placeholders(page, name)
                if not rects:
                    continue
                size, color = _style_for(page, name)

                # 先擦掉占位符（redaction 会同时移除原文字，避免复制 PDF 时残留）
                for rect in rects:
                    page.add_redact_annot(rect, fill=None)
                page.apply_redactions(
                    images=pymupdf.PDF_REDACT_IMAGE_NONE,
                    graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                )

                if not value:
                    continue
                for rect in rects:
                    used = _fit_font_size(value, rect.width * 2.5, size, font)
                    if used < size - 0.6:
                        notes.append(f"{name} 填值较长，字号由 {size:.0f}pt 缩到 {used:.0f}pt")
                    # 基线取占位符底部略上移；x 保持占位符左端，视觉上原地替换
                    page.insert_text(
                        (rect.x0, rect.y1 - size * 0.2),
                        value,
                        fontname=font_name,
                        fontfile=font_file,
                        fontsize=used,
                        color=color,
                        overlay=True,
                    )

            leftover = PLACEHOLDER_RE.findall(page.get_text())
            if leftover:
                notes.append("未替换的占位符：" + "、".join(sorted(set(leftover))))

        payload = document.tobytes(garbage=3, deflate=True)
    finally:
        document.close()
    return payload, list(dict.fromkeys(notes))


def safe_filename(value: str, fallback: str) -> str:
    """生成跨平台安全的文件名（去掉路径分隔符等非法字符）。"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(value or "")).strip(" ._")
    return cleaned[:60] or fallback


def build_zip(
    table: TableData,
    template: bytes,
    *,
    limit: int = MAX_BATCH,
    base_name: str = "批量生成",
) -> BatchResult:
    """按数据表逐行渲染并打包成 ZIP。"""
    result = BatchResult()
    if not table.rows:
        result.message = "数据表里没有数据行。"
        return result

    first_key = table.headers[0] if table.headers else DEFAULT_FIRST_LINE
    buffer = io.BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for index, row in enumerate(table.rows[:limit], start=1):
            label = safe_filename(str(row.get(first_key, "")), f"第{index}份")
            name = f"{label}.pdf"
            if name in used:  # 重名（例如同名同姓）加序号，避免互相覆盖
                name = f"{label}_{index}.pdf"
            used.add(name)
            try:
                payload, notes = render_one(template, row)
            except (BatchError, Exception) as exc:  # noqa: BLE001
                result.skipped.append(f"第 {index} 行：{exc}")
                continue
            archive.writestr(name, payload)
            result.generated += 1
            result.log.append(f"✔ {name}（{len(payload) / 1024:.0f}KB）")
            result.log.extend(f"   {note}" for note in notes)

        skipped = len(table.rows) - result.generated
        report = [f"共 {result.generated} 份生成成功，{skipped} 份失败"]
        report.extend(result.skipped)
        archive.writestr("生成报告.txt", batch_report(result, report))

    if not result.generated:
        result.message = "没有生成任何文件：" + ("；".join(result.skipped[:3]) or "请检查模板占位符")
        return result

    result.ok = True
    result.output = buffer.getvalue()
    result.filename = f"{base_name}_{datetime.now():%Y%m%d_%H%M}.zip"
    suffix = f"，跳过 {len(result.skipped)} 份" if result.skipped else ""
    result.message = f"已生成 {result.generated} 份 PDF{suffix}"
    return result


def batch_report(result: BatchResult, report: Iterable[str] = ()) -> bytes:
    """生成 ZIP 内的“生成报告.txt”内容（UTF-8）。"""
    lines = [f"闪办 · 批量生成报告（{datetime.now():%Y-%m-%d %H:%M}）", ""]
    lines.extend(report)
    lines.extend(result.log)
    return "\n".join(lines).encode("utf-8")
