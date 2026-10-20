#!/usr/bin/env python3
"""图片改字 / 批量生成 两个新功能的回归测试。

运行：.venv/bin/python _smoke_new_pages.py
"""
from __future__ import annotations

import io
import zipfile

import pymupdf
import openpyxl
from PIL import Image
from streamlit.testing.v1 import AppTest

import app as app_module
import batch_fill
import image_edit

FONT = "/System/Library/Fonts/Supplemental/Songti.ttc"


def make_image(lines: list[str], width: int = 420) -> bytes:
    """生成纯像素 PNG（模拟截图）：没有文字层，只能靠笔画定位。"""
    doc = pymupdf.open()
    height = 60 + 60 * len(lines)
    page = doc.new_page(width=600, height=height)
    page.draw_rect(page.rect, color=None, fill=(0.97, 0.96, 0.93))
    for index, text in enumerate(lines):
        page.insert_text((40, 60 + 60 * index), text, fontname="song", fontfile=FONT,
                         fontsize=30, color=(0.1, 0.1, 0.12))
    png = page.get_pixmap(alpha=False).tobytes("png")
    doc.close()
    buffer = io.BytesIO()
    Image.open(io.BytesIO(png)).resize((width, int(width * height / 600))).save(buffer, "PNG")
    return buffer.getvalue()


def check_engine() -> None:
    # 1. 单行：只改数字，标签必须保留
    src = make_image(["合同编号 2024-001"])
    res = image_edit.replace_image_text(src, None, [("2024-001", "2026-888")])
    assert res.ok, res.message
    assert res.lines[0].text == "2026-888", res.lines
    assert res.output.startswith(b"\x89PNG")
    print("✔ 单行图片改字:", res.message)

    # 2. 多行：两条规则，从上到下匹配
    src = make_image(["姓名：张三", "部门：技术部"])
    res = image_edit.replace_image_text(src, None, [("张三", "李四"), ("技术部", "市场部")])
    assert res.ok and [line.text for line in res.lines] == ["李四", "市场部"], res.lines
    print("✔ 多行图片改字:", res.message)

    # 3. 覆盖块与背景同色（零色差）
    after = pymupdf.Pixmap(res.output)
    assert after.pixel(5, 5) == after.pixel(int(res.lines[0].rect[2]) - 1, int(res.lines[0].rect[1])), (
        after.pixel(5, 5),
        after.pixel(int(res.lines[0].rect[2]) - 1, int(res.lines[0].rect[1])),
    )
    print("✔ 覆盖块与背景颜色一致")

    # 4. 规则为空 / 图片损坏 -> 中文提示，不抛异常
    bad = image_edit.replace_image_text(src, None, [(" ", "")])
    assert not bad.ok and "至少填写一条" in bad.message
    broken = image_edit.replace_image_text(b"not-an-image", None, [("a", "b")])
    assert not broken.ok and "无法解析" in broken.message
    print("✔ 异常输入返回中文提示")


def check_batch() -> None:
    template = batch_fill.sample_template_pdf()
    table = batch_fill.read_table(batch_fill.sample_table_csv(), "示例数据.csv")
    assert table.count == 3

    report = batch_fill.prepare_preview(table, template)
    assert report.ok, report.message
    assert report.preview_png.startswith(b"\x89PNG")
    print("✔ 批量预览:", report.message)

    result = batch_fill.build_zip(table, template, limit=3)
    assert result.ok and result.generated == 3, result.message
    names = zipfile.ZipFile(io.BytesIO(result.output)).namelist()
    assert names[:3] == ["张三.pdf", "李四.pdf", "王五.pdf"] and "生成报告.txt" in names
    print("✔ 批量生成:", result.message, names)

    # 生成的文件里不应残留占位符
    first = zipfile.ZipFile(io.BytesIO(result.output)).read("张三.pdf")
    text = pymupdf.open(stream=first, filetype="pdf")[0].get_text()
    assert "{{" not in text and "张三" in text
    print("✔ 占位符全部替换")


def check_ui() -> None:
    app = AppTest.from_file("app.py", default_timeout=180).run()
    assert not app.exception, app.exception

    def find(kind: str, key: str):
        for item in app.get(kind):
            if item.key == key:
                return item
        raise AssertionError(f"找不到 {kind}[{key}]")

    # 图片页
    find("file_uploader", "replace_upload").set_value(
        [("截图.png", make_image(["编号 2024-001"]), "image/png")]
    ).run()
    # 只有新文字、没有旧文字 -> 应给出校验提示
    find("text_input", "replace_old_0").set_value("").run()
    find("text_input", "replace_new_0").set_value("2026-888").run()
    assert any("请补上左边「原内容」" in c.value for c in app.caption), [
        c.value for c in app.caption
    ]
    find("text_input", "replace_old_0").set_value("2024-001").run()
    find("button", "replace_btn").click().run()
    assert not app.exception, app.exception
    assert any("2026-888" in s.value for s in app.success), [s.value for s in app.success]
    assert any(d.key == "replace_img_download" for d in app.get("download_button"))
    print("✔ 图片页端到端")

    # 批量页：两个示例下载必须是真实的 xlsx / pdf
    finds = {d.key: d for d in app.get("download_button")}
    assert "batch_sample_excel" in finds and "batch_sample_pdf" in finds, list(finds)
    excel_bytes = app_module.batch_fill.sample_table_xlsx()
    pdf_bytes = app_module.batch_fill.sample_template_pdf()
    assert zipfile.is_zipfile(io.BytesIO(excel_bytes))
    assert openpyxl.load_workbook(io.BytesIO(excel_bytes)).active["A1"].value == "姓名"
    assert pdf_bytes.startswith(b"%PDF")
    print("✔ 示例文件：标准 xlsx + 标准 PDF")

    # 批量页
    labels = [d.label for d in app.get("download_button")]
    assert any("Excel 示例" in label for label in labels)
    assert any("PDF 模板示例" in label for label in labels)
    find("button", "batch_preview_btn").click().run()
    assert any("请先上传名单" in w.value for w in app.warning), [w.value for w in app.warning]

    find("file_uploader", "batch_excel").set_value(
        [("示例数据.csv", batch_fill.sample_table_csv(), "text/csv")]
    )
    find("file_uploader", "batch_template").set_value(
        [("模板.pdf", batch_fill.sample_template_pdf(), "application/pdf")]
    )
    app.run()
    find("button", "batch_preview_btn").click().run()
    assert any("校验通过" in s.value for s in app.success), [s.value for s in app.success]
    find("button", "batch_btn").click().run()
    assert any("已生成 3 份" in s.value for s in app.success), [s.value for s in app.success]
    assert any(d.key == "batch_download" for d in app.get("download_button"))
    print("✔ 批量页端到端（单次 3 份额度生效）")


if __name__ == "__main__":
    check_engine()
    check_batch()
    check_ui()
    print("\nALL NEW PAGE CHECKS PASSED")
