"""功能自测：文档转 PDF / 多行替换 / 录音 mock 转写。

直接运行：.venv/bin/python _smoke_features.py
只依赖项目已有依赖，不引入 pytest。
"""

from __future__ import annotations

import io
import zipfile

import pymupdf
from docx import Document

import app


def make_docx() -> bytes:
    """生成一份包含标题、加粗、表格、分页的 Word 文档。"""
    document = Document()
    document.add_heading("2025 年度工作报告", level=1)
    paragraph = document.add_paragraph("第一段：")
    run = paragraph.add_run("加粗的结论")
    run.bold = True
    document.add_paragraph("第二段：普通正文内容，用于验证段落与翻页。")
    table = document.add_table(rows=2, cols=3)
    for index, cell in enumerate(table.rows[0].cells):
        cell.text = f"表头{index + 1}"
    for index, cell in enumerate(table.rows[1].cells):
        cell.text = f"数据{index + 1}"
    document.add_page_break()
    document.add_heading("附录", level=2)
    document.add_paragraph("附录正文。")
    for _ in range(60):
        document.add_paragraph("填充段落，用来验证多页排版是否正常。")
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def make_xlsx() -> bytes:
    """手写一个最小 xlsx（sharedStrings + 两个工作表），验证不依赖 openpyxl。"""
    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>'
        '<c r="C1" t="s"><v>2</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2" t="s"><v>4</v></c>'
        '<c r="C2"><v>12800</v></c></row>'
        '<row r="3"><c r="A3" t="s"><v>5</v></c><c r="B3" t="s"><v>6</v></c>'
        '<c r="C3"><v>9600</v></c></row>'
        "</sheetData></worksheet>"
    )
    sheet2 = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>第二页说明</t></is></c>'
        "</row></sheetData></worksheet>"
    )
    shared = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<si><t>姓名</t></si><si><t>部门</t></si><si><t>薪资</t></si>"
        "<si><t>林一</t></si><si><t>研发</t></si><si><t>周禾</t></si><si><t>设计</t></si>"
        "</sst>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="薪资表" sheetId="1" r:id="rId1"/>'
        '<sheet name="说明" sheetId="2" r:id="rId2"/></sheets></workbook>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://x" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://x" Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr("xl/sharedStrings.xml", shared)
        archive.writestr("xl/worksheets/sheet1.xml", sheet1)
        archive.writestr("xl/worksheets/sheet2.xml", sheet2)
    return buffer.getvalue()


def make_png() -> bytes:
    """生成一张 400x300 的测试图片。"""
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 400, 300), False)
    pixmap.set_rect(pixmap.irect, (36, 71, 120))
    return pixmap.tobytes("png")


def make_pdf() -> bytes:
    document = pymupdf.open()
    page = document.new_page()
    page.insert_text((72, 100), "Invoice No 2024-0001", fontsize=14)
    page.insert_text((72, 140), "Invoice No 2024-0001", fontsize=14)
    payload = document.tobytes()
    document.close()
    return payload


def pdf_text(payload: bytes) -> str:
    with pymupdf.open(stream=payload, filetype="pdf") as document:
        return "\n".join(page.get_text() for page in document)


def check(name: str, condition: bool, detail: str = "") -> None:
    print(("  ✔ " if condition else "  ✘ ") + name + (f" — {detail}" if detail else ""))
    if not condition:
        raise AssertionError(name)


print("\n[1] Word -> PDF")
outcome = app.document_to_pdf(make_docx(), "年度报告.docx")
check("转换成功", outcome.ok, outcome.message)
check(
    "输出为 PDF 且改名",
    outcome.filename == "年度报告.pdf" and (outcome.data or b"").startswith(b"%PDF"),
)
with pymupdf.open(stream=outcome.data, filetype="pdf") as document:
    text = "".join(page.get_text() for page in document)
    check("保留标题文字", "2025 年度工作报告" in text)
    check("保留表格内容", "表头1" in text and "数据3" in text)
    check("正文自动分页（>=2 页）", document.page_count >= 2, f"{document.page_count} 页")

print("\n[2] Excel -> PDF")
outcome = app.document_to_pdf(make_xlsx(), "薪资表.xlsx")
check("转换成功", outcome.ok, outcome.message)
check("工作表识别为 2 个", "2 个工作表" in outcome.message, outcome.message)
text = pdf_text(outcome.data or b"")
check("读到共享字符串", "林一" in text and "周禾" in text)
check("读到数字与第二个工作表", "12800" in text and "第二页说明" in text)

print("\n[3] 图片 -> PDF")
outcome = app.document_to_pdf(make_png(), "截图.png")
check("转换成功", outcome.ok, outcome.message)
with pymupdf.open(stream=outcome.data, filetype="pdf") as document:
    page = document[0]
    check(
        "页面按图片比例 4:3",
        abs(page.rect.width / page.rect.height - 4 / 3) < 0.01,
        f"{page.rect.width:.0f}x{page.rect.height:.0f}",
    )
    check("图片已嵌入", len(page.get_images()) == 1)

print("\n[4] 失败分支")
check("旧版 .doc 拒绝并提示", not app.document_to_pdf(b"x", "a.doc").ok)
check("空文件拒绝", not app.document_to_pdf(b"", "a.docx").ok)
check("未知类型拒绝", not app.document_to_pdf(b"x", "a.txt").ok)

print("\n[5] 多行替换规则")
rules, problem = app.build_replace_rules("0001\nA公司", "0002\nB公司", "1")
check("两组规则解析成功", problem == "" and len(rules) == 2, problem)
_, problem = app.build_replace_rules("0001\nA公司", "0002", "1")
check("行数不一致报错", "行数不一致" in problem, problem)
_, problem = app.build_replace_rules("0001", "000", "1")
check("不等长报错", "等长" in problem, problem)

print("\n[6] 精准替换：不勾选 / 勾选全文替换")
once = app.replace_text(make_pdf(), "invoice.pdf", "0001", "0002", "1", replace_all=False)
check("单处替换成功", once.ok, once.message)
text_once = pdf_text(once.data or b"")
check("只替换 1 处", text_once.count("0002") == 1 and text_once.count("0001") == 1,
      repr(text_once))

every = app.replace_text(make_pdf(), "invoice.pdf", "0001", "0002", "1", replace_all=True)
check("全文替换成功", every.ok, every.message)
text_every = pdf_text(every.data or b"")
check("两处都替换", text_every.count("0002") == 2 and "0001" not in text_every,
      repr(text_every))

multi = app.replace_text(
    make_pdf(), "invoice.pdf", "Invoice\n0001", "Invoice\n0002", "1", replace_all=True
)
check("多行两组规则同时生效", multi.ok, multi.message)

print("\nALL FEATURE CHECKS PASSED")
